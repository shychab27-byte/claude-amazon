const express = require('express');
const multer  = require('multer');
const xlsx    = require('xlsx');
const path    = require('path');

const app  = express();
const PORT = process.env.PORT || 3000;

app.use(express.static(path.join(__dirname, 'public')));

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    const ok  = ['.xlsx', '.xls', '.csv', '.txt'].includes(ext);
    cb(ok ? null : new Error(`Unsupported file type: ${ext}`), ok);
  }
});

// ─── Column helpers ──────────────────────────────────────────────────────────

// Normalize a column header: lowercase, dashes/underscores → space, trim
function norm(s) {
  return String(s || '').toLowerCase().replace(/[-_]/g, ' ').trim();
}

// Find a column value by trying multiple candidate names.
// Tries exact normalized match first, then substring match.
function findCol(row, candidates) {
  const keys = Object.keys(row);
  for (const candidate of candidates) {
    const c = norm(candidate);
    const exact = keys.find(k => norm(k) === c);
    if (exact !== undefined) return row[exact];
    const partial = keys.find(k => norm(k).includes(c));
    if (partial !== undefined) return row[partial];
  }
  return null;
}

// Parse a numeric value; returns null for empty / N/A strings
function toNum(val) {
  if (val === null || val === undefined) return null;
  const s = String(val).trim().toLowerCase();
  if (!s || s === 'n/a' || s === '-' || s === '--') return null;
  const n = parseFloat(s.replace(/,/g, ''));
  return isNaN(n) ? null : n;
}

// ─── Priority logic ──────────────────────────────────────────────────────────

function derivePriority(dosNum, stockNum) {
  if (stockNum !== null && stockNum <= 0) return 'OUT_OF_STOCK';
  if (dosNum === null)  return 'UNKNOWN';
  if (dosNum <= 0)      return 'OUT_OF_STOCK';
  if (dosNum <= 7)      return 'CRITICAL';
  if (dosNum <= 21)     return 'HIGH';
  if (dosNum <= 45)     return 'MEDIUM';
  return 'LOW';
}

const PRIORITY_ORDER = {
  OUT_OF_STOCK: 0, CRITICAL: 1, HIGH: 2, MEDIUM: 3, LOW: 4, UNKNOWN: 5
};

function worstOf(priorities) {
  return priorities.reduce((worst, p) => {
    return (PRIORITY_ORDER[p] ?? 99) < (PRIORITY_ORDER[worst] ?? 99) ? p : worst;
  }, 'UNKNOWN');
}

// ─── Routes ──────────────────────────────────────────────────────────────────

app.get('/api/health', (req, res) => res.json({ status: 'ok' }));

app.post('/api/parse', upload.single('file'), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded.' });

    const ext      = path.extname(req.file.originalname).toLowerCase();
    const workbook = (ext === '.csv' || ext === '.txt')
      ? xlsx.read(req.file.buffer.toString('utf8'), { type: 'string', raw: false })
      : xlsx.read(req.file.buffer, { type: 'buffer', raw: false });

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows  = xlsx.utils.sheet_to_json(sheet, { defval: '', raw: false, blankrows: false });

    if (rows.length === 0) {
      return res.status(400).json({ error: 'The file appears to be empty.' });
    }

    // Check whether a Parent ASIN column exists
    const headerNorms  = Object.keys(rows[0]).map(norm);
    const hasParentAsin = headerNorms.some(h => h.includes('parent') && h.includes('asin'));

    // ── Parse rows ────────────────────────────────────────────────────────────
    const items = rows.map(row => {
      const sku = String(
        findCol(row, ['sku', 'seller sku', 'seller-sku', 'msku']) ?? ''
      ).trim();
      if (!sku) return null;

      const parentAsin = String(
        findCol(row, ['parent asin', 'parent-asin', 'parentasin']) ?? ''
      ).trim();
      const asin = String(findCol(row, ['asin']) ?? '').trim();
      const name = String(
        findCol(row, ['product name', 'product-name', 'title', 'item name', 'description', 'product']) ?? ''
      ).trim();

      const stockRaw   = findCol(row, ['available', 'fulfillable qty', 'afn fulfillable', 'afn-fulfillable', 'sellable', 'current inventory']);
      const dosRaw     = findCol(row, ['days of supply at amazon', 'days of supply', 'days supply']);
      const reorderRaw = findCol(row, ['recommended replenishment qty', 'recommended order qty', 'recommended order quantity', 'suggested order quantity', 'replenishment qty']);
      const orderByRaw = findCol(row, ['recommended ship date', 'order by date', 'ship by date', 'restock by date', 'ship date']);

      const stockNum   = toNum(stockRaw) ?? 0;
      const dosNum     = toNum(dosRaw);
      const reorderNum = toNum(reorderRaw);

      return {
        sku,
        parent_asin:   parentAsin || null,
        asin:          asin       || null,
        product_name:  name,
        current_stock: stockNum,
        days_of_supply: dosNum,
        reorder_qty:   reorderNum != null ? Math.round(reorderNum) : null,
        order_by_date: orderByRaw ? String(orderByRaw).trim() : null,
        priority:      derivePriority(dosNum, stockNum),
      };
    }).filter(Boolean);

    if (items.length === 0) {
      return res.status(400).json({
        error: 'No valid rows found. Make sure this is the Restock Inventory report.'
      });
    }

    // ── Group by Parent ASIN → ASIN → SKU (fallback chain) ───────────────────
    const groupMap = new Map();
    for (const item of items) {
      const key = item.parent_asin || item.asin || item.sku;
      if (!groupMap.has(key)) {
        groupMap.set(key, { parent_asin: key, items: [] });
      }
      groupMap.get(key).items.push(item);
    }

    // ── Build group metadata ──────────────────────────────────────────────────
    const groups = [];
    for (const group of groupMap.values()) {
      // Sort items within group: worst priority first, then shortest days_of_supply
      group.items.sort((a, b) => {
        const pa = PRIORITY_ORDER[a.priority] ?? 99;
        const pb = PRIORITY_ORDER[b.priority] ?? 99;
        if (pa !== pb) return pa - pb;
        return (a.days_of_supply ?? 9999) - (b.days_of_supply ?? 9999);
      });

      const priorities       = group.items.map(i => i.priority);
      group.worst_priority   = worstOf(priorities);
      group.counts           = {};
      for (const p of priorities) group.counts[p] = (group.counts[p] || 0) + 1;

      // Group display name: product name of the first (worst) item, trimmed
      group.group_name = group.items[0]?.product_name || group.parent_asin;

      groups.push(group);
    }

    // ── Sort groups: worst overall priority first ──────────────────────────────
    groups.sort((a, b) => {
      const pa = PRIORITY_ORDER[a.worst_priority] ?? 99;
      const pb = PRIORITY_ORDER[b.worst_priority] ?? 99;
      return pa - pb;
    });

    res.json({
      success:      true,
      rowCount:     rows.length,
      itemCount:    items.length,
      groupCount:   groups.length,
      hasParentAsin,
      groups,
    });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || 'Failed to parse file.' });
  }
});

app.listen(PORT, () =>
  console.log(`\nRestock Tracker running at http://localhost:${PORT}\n`)
);
