const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const Anthropic = require('@anthropic-ai/sdk');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;

// Serve static files
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json({ limit: '50mb' }));

// Configure multer for file uploads (memory storage)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 }, // 50MB per file
  fileFilter: (req, file, cb) => {
    const allowed = ['.xlsx', '.xls', '.csv', '.txt'];
    const ext = path.extname(file.originalname).toLowerCase();
    if (allowed.includes(ext)) {
      cb(null, true);
    } else {
      cb(new Error(`File type not allowed: ${ext}. Please upload .xlsx, .xls, .csv, or .txt files.`));
    }
  }
});

// ─────────────────────────────────────────────
// Report Detection Logic
// ─────────────────────────────────────────────

const REPORT_SIGNATURES = {
  // ── 5 primary report formats (most specific checked first) ───────

  // Business Report: unique "Child ASIN" + "Units Ordered" columns
  'Business Report - By Child ASIN': [
    ['child asin', 'units ordered', 'sessions'],
    ['child asin', 'ordered product sales', 'buy box percentage']
  ],

  // Amazon Fulfilled Inventory: distinguished by hyphenated AFN + MFN columns
  'Amazon Fulfilled Inventory': [
    ['afn-fulfillable-qty', 'mfn-fulfillable-qty'],
    ['afn fulfillable qty', 'mfn listing exists']
  ],

  // Restock Inventory (Amazon's own suggestions) — multiple format variants
  'Restock Inventory': [
    ['recommended replenishment qty', 'days of supply'],         // current format
    ['recommended order quantity', 'reorder point'],             // alternate
    ['suggested order quantity', 'order by date', 'max inventory level'], // older
    ['units sold per day', 'suggested order quantity']           // legacy
  ],

  // Inventory Ledger: event-based, has "event type" + "fulfillment center"
  'Inventory Ledger': [
    ['event type', 'fulfillment center', 'disposition', 'fnsku'],
    ['event type', 'fnsku', 'quantity', 'msku']
  ],

  // Manage FBA Inventory: simple available/reserved/inbound columns
  'Manage FBA Inventory': [
    ['sku', 'fnsku', 'asin', 'available'],
    ['seller-sku', 'asin', 'condition', 'available']
  ],

  // ── Backward-compatible names (kept for older uploads) ───────────
  'Inventory Health Report': [
    ['inventory age', 'sell through', 'days of supply'],
    ['inv age', 'units sold (last 30 days)', 'days of supply']
  ],
  'All Orders Report': [
    ['amazon-order-id', 'purchase-date', 'order-status', 'fulfillment-channel'],
    ['amazon-order-id', 'asin', 'quantity', 'item-price']
  ],

  // ── Season Map ───────────────────────────────────────────────────
  'Season Map': [
    ['sku', 'season'],
    ['asin', 'season']
  ]
};

// Valid season values (normalized)
const VALID_SEASONS = ['spring/summer', 'fall/winter', 'year-round', 'all'];

// Batching config
const BATCH_SIZE        = 150; // SKUs per Claude call
const BATCH_CONCURRENCY = 3;   // max parallel Claude calls

// In-memory progress tracker keyed by requestId
const progressMap = new Map();
function setProgress(requestId, data) {
  if (requestId) progressMap.set(requestId, { ...data, ts: Date.now() });
}

function normalizeSeasonValue(val) {
  const v = String(val || '').toLowerCase().trim();
  if (v.includes('spring') || v.includes('summer') || v === 'ss') return 'Spring/Summer';
  if (v.includes('fall') || v.includes('autumn') || v.includes('winter') || v === 'fw') return 'Fall/Winter';
  if (v.includes('year') || v.includes('all') || v.includes('evergreen')) return 'Year-Round';
  return null;
}

// Build a SKU→season lookup map from parsed Season Map rows
function buildSeasonMap(rows) {
  const map = {};
  for (const row of rows) {
    // Support both "sku" and "asin" as the key column
    const key = String(row['sku'] || row['SKU'] || row['asin'] || row['ASIN'] || '').trim();
    const rawSeason = row['season'] || row['Season'] || row['SEASON'] || '';
    const season = normalizeSeasonValue(rawSeason);
    if (key && season) map[key.toUpperCase()] = season;
  }
  return map;
}

function detectReportType(headers) {
  const normalizedHeaders = headers.map(h =>
    String(h || '').toLowerCase().trim()
  );

  for (const [reportType, signatureSets] of Object.entries(REPORT_SIGNATURES)) {
    for (const signatures of signatureSets) {
      const matches = signatures.filter(sig =>
        normalizedHeaders.some(h => h.includes(sig.toLowerCase()))
      );
      if (matches.length >= Math.ceil(signatures.length * 0.7)) {
        return reportType;
      }
    }
  }
  return 'Unknown Report';
}

// ─────────────────────────────────────────────
// Excel / CSV Parsing
// ─────────────────────────────────────────────

function parseFile(buffer, filename) {
  try {
    const ext = path.extname(filename).toLowerCase();
    let workbook;

    if (ext === '.csv' || ext === '.txt') {
      const text = buffer.toString('utf8');
      workbook = xlsx.read(text, { type: 'string', raw: false });
    } else {
      workbook = xlsx.read(buffer, { type: 'buffer', raw: false });
    }

    const results = [];

    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const rows = xlsx.utils.sheet_to_json(sheet, {
        defval: '',
        raw: false,
        blankrows: false
      });

      if (rows.length === 0) continue;

      const headers = Object.keys(rows[0]);
      const reportType = detectReportType(headers);

      results.push({
        sheetName,
        reportType,
        headers,
        rowCount: rows.length,
        data: rows.slice(0, 5000) // load up to 5000 rows; batching controls what's sent to Claude
      });
    }

    return results;
  } catch (err) {
    throw new Error(`Failed to parse "${filename}": ${err.message}`);
  }
}

// ─────────────────────────────────────────────
// Data Summarization (before sending to Claude)
// ─────────────────────────────────────────────

// Extract season map from parsed files (returns combined lookup object)
function extractSeasonMap(parsedFiles) {
  let combined = {};
  for (const file of parsedFiles) {
    for (const sheet of file.sheets) {
      if (sheet.reportType === 'Season Map') {
        Object.assign(combined, buildSeasonMap(sheet.data));
      }
    }
  }
  return combined;
}

function summarizeData(parsedFiles) {
  const summary = [];

  for (const file of parsedFiles) {
    for (const sheet of file.sheets) {
      // Skip season map from main data blocks — handled separately
      if (sheet.reportType === 'Season Map') continue;

      const { reportType, data, headers, rowCount } = sheet;

      const block = {
        reportType,
        fileName: file.fileName,
        rowCount,
        headers,
        sample: data.slice(0, 3),
        data
      };

      // Add computed stats for known report types
      if (reportType === 'Restock Inventory' || reportType === 'Restocking Report') {
        const getDOS = r => parseFloat(
          r['Days of Supply at Amazon'] || r['Days of Supply'] || r['days of supply'] || r['Days of supply'] || 0
        );
        block.criticalItems = data.filter(r => { const d = getDOS(r); return d >= 0 && d <= 14; }).length;
        block.lowItems      = data.filter(r => { const d = getDOS(r); return d > 14 && d <= 30; }).length;
      }

      if (reportType === 'Inventory Health Report') {
        block.zeroSellThrough = data.filter(row => {
          const st = parseFloat(row['Sell Through'] || row['sell through'] || 0);
          return st === 0;
        }).length;
      }

      if (reportType === 'Amazon Fulfilled Inventory') {
        block.outOfStock = data.filter(row => {
          const qty = parseInt(row['afn-fulfillable-qty'] || row['AFN Fulfillable Qty'] || 0);
          return qty === 0;
        }).length;
        block.unsellableUnits = data.reduce((sum, row) => {
          return sum + (parseInt(row['afn-unsellable-qty'] || row['AFN Unsellable Qty'] || 0) || 0);
        }, 0);
      }

      if (reportType === 'Manage FBA Inventory' || reportType === 'FBA Manage Inventory') {
        block.outOfStock = data.filter(row => {
          const avail = parseInt(row['Available'] || row['available'] || 0);
          return avail === 0;
        }).length;
      }

      // Business Report is compact (1 row per Child ASIN) — include in EVERY batch as global context
      if (reportType === 'Business Report - By Child ASIN') {
        block.isGlobalContext = true;
      }

      // Aggregate multi-row-per-SKU reports to 1 summary row per SKU (major token reduction)
      if (reportType === 'All Orders Report') {
        const agg = aggregateOrdersReport(data);
        block.isAggregated   = true;
        block.aggregatedFrom = data.length;   // original event count
        block.data           = agg;
        block.headers        = ['sku', 'asin', 'units_sold_30d', 'units_sold_90d',
                                 'order_count_30d', 'revenue_30d', 'last_order_date'];
      }

      if (reportType === 'Inventory Ledger Report') {
        const agg = aggregateLedgerReport(data);
        block.isAggregated   = true;
        block.aggregatedFrom = data.length;
        block.data           = agg;
        block.headers        = ['msku', 'fnsku', 'asin', 'received', 'returns', 'removals', 'adjustments'];
      }

      summary.push(block);
    }
  }

  return summary;
}

// ─────────────────────────────────────────────
// Data Compaction (key fields only, per report type)
// ─────────────────────────────────────────────

// Fields to keep per report type (lowercase fragments for flexible matching)
const REPORT_KEY_FIELDS = {
  // Primary 5 report types
  'Restock Inventory':               ['sku', 'asin', 'fnsku', 'product', 'days of supply', 'recommended', 'reorder point', 'units sold', 'max inventory', 'ship date'],
  'Amazon Fulfilled Inventory':      ['sku', 'asin', 'fnsku', 'product', 'afn-fulfillable', 'afn-unsellable', 'afn-warehouse', 'mfn-fulfillable', 'afn-inbound', 'afn-reserved'],
  'Manage FBA Inventory':            ['sku', 'asin', 'fnsku', 'product', 'available', 'reserved', 'inbound'],
  'Inventory Ledger':                ['fnsku', 'asin', 'msku', 'event type', 'quantity', 'fulfillment center', 'date'],
  'Business Report - By Child ASIN': ['child asin', 'title', 'units ordered', 'sessions', 'buy box', 'ordered product sales', 'unit session'],
  // Backward-compatible names
  'Restocking Report':       ['sku', 'asin', 'fnsku', 'product', 'days of supply', 'suggested order', 'order by date', 'units sold per day', 'max inventory'],
  'Inventory Health Report': ['sku', 'asin', 'product', 'days of supply', 'sell through', 'units sold', 'inv age'],
  'FBA Manage Inventory':    ['sku', 'asin', 'fnsku', 'product', 'available', 'reserved', 'inbound'],
  'Inventory Ledger Report': ['fnsku', 'asin', 'msku', 'event type', 'quantity', 'fulfillment center', 'date'],
  'All Orders Report':       ['order-id', 'asin', 'sku', 'quantity', 'item-price', 'purchase-date', 'order-status'],
};

function compactRows(block) {
  // Aggregated blocks are already compact — send them as-is
  if (block.isAggregated) return block.data;

  const wantedFragments = REPORT_KEY_FIELDS[block.reportType];
  const rows = block.data; // no hard row cap — batching controls volume

  if (!wantedFragments) return rows.slice(0, 50);

  return rows.map(row => {
    const compact = {};
    for (const [key, val] of Object.entries(row)) {
      const keyLower = key.toLowerCase().trim();
      if (wantedFragments.some(f => keyLower.includes(f) || f.includes(keyLower))) {
        compact[key] = val;
      }
    }
    // Always keep at least the raw row if nothing matched (fallback)
    return Object.keys(compact).length > 0 ? compact : row;
  });
}

// ─────────────────────────────────────────────
// Batching helpers
// ─────────────────────────────────────────────

function extractRowKey(row) {
  const sku = String(
    row['sku'] || row['SKU'] || row['seller-sku'] || row['Seller SKU'] ||
    row['Seller-SKU'] || row['MSKU'] || row['msku'] || ''
  ).trim().toUpperCase();
  if (sku) return sku;

  // Also check "Child ASIN" column from Business Reports
  const asin = String(
    row['asin'] || row['ASIN'] || row['Child ASIN'] || row['child asin'] || ''
  ).trim().toUpperCase();
  if (asin) return asin;

  const fnsku = String(row['fnsku'] || row['FNSKU'] || '').trim().toUpperCase();
  return fnsku || null;
}

// ─────────────────────────────────────────────
// Server-side report aggregation (token reduction)
// ─────────────────────────────────────────────

// All Orders Report: many rows per SKU → 1 aggregated row per SKU
function aggregateOrdersReport(data) {
  const now        = Date.now();
  const ms30       = 30 * 24 * 60 * 60 * 1000;
  const ms90       = 90 * 24 * 60 * 60 * 1000;
  const byKey      = new Map();

  for (const row of data) {
    const key = extractRowKey(row);
    if (!key) continue;

    const status = String(row['order-status'] || row['Order Status'] || '').toLowerCase();
    if (status.includes('cancel')) continue;   // skip cancelled

    const dateStr = row['purchase-date'] || row['Purchase Date'] || '';
    const ts      = dateStr ? new Date(dateStr).getTime() : NaN;
    const qty     = Math.max(0, parseInt(row['quantity'] || row['Quantity'] || 0) || 0);
    const price   = parseFloat(row['item-price'] || row['item-price'] || 0) || 0;

    if (!byKey.has(key)) {
      byKey.set(key, {
        sku:             row['sku'] || row['SKU'] || '',
        asin:            row['asin'] || row['ASIN'] || '',
        units_sold_30d:  0,
        units_sold_90d:  0,
        order_count_30d: 0,
        revenue_30d:     0,
        last_order_date: dateStr,
      });
    }

    const e = byKey.get(key);
    if (!isNaN(ts)) {
      const age = now - ts;
      if (age <= ms30) { e.units_sold_30d += qty; e.order_count_30d++; e.revenue_30d += price * qty; }
      if (age <= ms90)   e.units_sold_90d += qty;
      if (dateStr > e.last_order_date) e.last_order_date = dateStr;
    }
  }

  return [...byKey.values()];
}

// Inventory Ledger Report: many event rows per SKU → 1 summary row per SKU
function aggregateLedgerReport(data) {
  const byKey = new Map();

  for (const row of data) {
    const fnsku = String(row['fnsku'] || row['FNSKU'] || '').trim().toUpperCase();
    const msku  = String(row['msku']  || row['MSKU']  || '').trim().toUpperCase();
    const asin  = String(row['asin']  || row['ASIN']  || '').trim().toUpperCase();
    const key   = msku || fnsku || asin;
    if (!key) continue;

    const eventType = String(row['event type'] || row['Event Type'] || '').toLowerCase();
    const qty       = parseInt(row['quantity']  || row['Quantity']  || 0) || 0;

    if (!byKey.has(key)) {
      byKey.set(key, { msku, fnsku, asin, received: 0, returns: 0, removals: 0, adjustments: 0 });
    }

    const e = byKey.get(key);
    if      (eventType.includes('receipt'))         e.received    += qty;
    else if (eventType.includes('customer return')) e.returns     += qty;
    else if (eventType.includes('removal') || eventType.includes('dispos')) e.removals += qty;
    else if (eventType.includes('adjustment'))      e.adjustments += qty;
  }

  return [...byKey.values()];
}

function getAllIdentifiers(summaryData) {
  const seen = new Set();
  for (const block of summaryData) {
    for (const row of block.data) {
      const key = extractRowKey(row);
      if (key) seen.add(key);
    }
  }
  return [...seen];
}

function filterSummaryDataForBatch(summaryData, batchIds) {
  const batchSet = new Set(batchIds);
  return summaryData.map(block => {
    // Global-context blocks (e.g. Business Report By Child ASIN) go into every batch intact
    if (block.isGlobalContext) return block;

    const filteredData = block.data.filter(row => {
      const key = extractRowKey(row);
      return key && batchSet.has(key);
    });
    if (filteredData.length === 0) return null;
    return { ...block, data: filteredData, rowCount: filteredData.length };
  }).filter(Boolean);
}

// Run async tasks with a concurrency limit
async function runWithConcurrency(tasks, limit) {
  const results = new Array(tasks.length);
  let nextIndex = 0;
  async function worker() {
    while (nextIndex < tasks.length) {
      const i = nextIndex++;
      results[i] = await tasks[i]();
    }
  }
  await Promise.all(Array.from({ length: Math.min(limit, tasks.length) }, worker));
  return results;
}

// Retry a Claude API call up to maxRetries times on rate-limit (429) errors
async function callClaudeWithRetry(client, params, onWait = null, maxRetries = 3) {
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      return await client.messages.create(params);
    } catch (err) {
      if (err.status === 429 && attempt < maxRetries) {
        const waitSecs = 30;
        console.warn(`Rate limit hit — waiting ${waitSecs}s (attempt ${attempt + 1}/${maxRetries + 1})`);
        if (onWait) onWait(waitSecs, attempt + 1);
        await new Promise(r => setTimeout(r, waitSecs * 1000));
      } else {
        throw err;
      }
    }
  }
}

const PRIORITY_ORDER = { CRITICAL: 0, HIGH: 1, MEDIUM: 2, LOW: 3, HOLD: 4 };

function mergeAnalysisResults(results, totalSKUs, batchCount) {
  const allItems    = [];
  const allWatch    = [];
  const allInsights = new Set();
  const allNotes    = new Set();
  let urgent = false;

  for (const r of results) {
    if (!r) continue;
    if (r.urgent_action_required) urgent = true;
    (r.replenishment_items || []).forEach(i => allItems.push(i));
    (r.watch_list          || []).forEach(w => allWatch.push(w));
    (r.insights            || []).forEach(i => allInsights.add(i));
    (r.data_quality_notes  || []).forEach(n => allNotes.add(n));
  }

  // Deduplicate by SKU — keep highest priority entry
  const itemMap = new Map();
  for (const item of allItems) {
    const key = String(item.sku || item.asin || '').toUpperCase();
    if (!key) continue;
    const existing = itemMap.get(key);
    const curRank  = PRIORITY_ORDER[item.priority]     ?? 99;
    const exRank   = PRIORITY_ORDER[existing?.priority] ?? 99;
    if (!existing || curRank < exRank) itemMap.set(key, item);
  }

  const sorted = [...itemMap.values()].sort((a, b) => {
    const pa = PRIORITY_ORDER[a.priority] ?? 99;
    const pb = PRIORITY_ORDER[b.priority] ?? 99;
    if (pa !== pb) return pa - pb;
    return (a.days_of_supply ?? 9999) - (b.days_of_supply ?? 9999);
  });

  const watchMap = new Map();
  for (const w of allWatch) {
    const key = String(w.sku || w.asin || '').toUpperCase();
    if (key) watchMap.set(key, w);
  }

  const criticalCount = sorted.filter(i => i.priority === 'CRITICAL').length;
  const highCount     = sorted.filter(i => i.priority === 'HIGH').length;

  return {
    summary: `Analyzed ${totalSKUs} SKUs across ${batchCount} batch${batchCount !== 1 ? 'es' : ''}. ` +
             `Found ${criticalCount} CRITICAL and ${highCount} HIGH priority items requiring attention.`,
    urgent_action_required: urgent,
    replenishment_items:    sorted,
    watch_list:             [...watchMap.values()],
    insights:               [...allInsights],
    data_quality_notes:     [...allNotes],
  };
}

async function analyzeInBatches(apiKey, summaryData, userNotes, targetSeason, seasonMap, requestId = null) {
  const allIds = getAllIdentifiers(summaryData);

  // Small enough to run in one shot
  if (allIds.length <= BATCH_SIZE) {
    setProgress(requestId, { stage: 'analyzing', batchCurrent: 0, batchTotal: 1, skuCount: allIds.length,
      message: `Analyzing ${allIds.length} SKUs…` });
    const result = await analyzeWithClaude(apiKey, summaryData, userNotes, targetSeason, seasonMap,
      null, null, requestId);
    setProgress(requestId, { stage: 'done', batchCurrent: 1, batchTotal: 1, skuCount: allIds.length,
      message: 'Analysis complete.' });
    return result;
  }

  // Split into batches
  const batches = [];
  for (let i = 0; i < allIds.length; i += BATCH_SIZE) {
    batches.push(allIds.slice(i, i + BATCH_SIZE));
  }
  console.log(`Batching: ${allIds.length} SKUs → ${batches.length} batches (concurrency ${BATCH_CONCURRENCY})`);
  setProgress(requestId, { stage: 'analyzing', batchCurrent: 0, batchTotal: batches.length,
    skuCount: allIds.length, message: `Splitting ${allIds.length} SKUs into ${batches.length} batches…` });

  let completed = 0;
  const tasks = batches.map((batchIds, idx) => async () => {
    const batchData = filterSummaryDataForBatch(summaryData, batchIds);
    console.log(`  Batch ${idx + 1}/${batches.length}: ${batchIds.length} SKUs`);
    if (batchData.length === 0) return null;
    const result = await analyzeWithClaude(apiKey, batchData, userNotes, targetSeason, seasonMap,
      idx + 1, batches.length, requestId);
    completed++;
    setProgress(requestId, {
      stage: 'analyzing', batchCurrent: completed, batchTotal: batches.length, skuCount: allIds.length,
      message: `Analyzed batch ${completed} of ${batches.length} (${Math.min(completed * BATCH_SIZE, allIds.length)} of ${allIds.length} SKUs)…`,
    });
    return result;
  });

  const results = await runWithConcurrency(tasks, BATCH_CONCURRENCY);
  setProgress(requestId, { stage: 'merging', batchCurrent: batches.length, batchTotal: batches.length,
    skuCount: allIds.length, message: 'Merging results from all batches…' });
  return mergeAnalysisResults(results.filter(Boolean), allIds.length, batches.length);
}

// ─────────────────────────────────────────────
// Claude API Analysis
// ─────────────────────────────────────────────

async function analyzeWithClaude(apiKey, summaryData, userNotes, targetSeason, seasonMap, batchNum = null, totalBatches = null, requestId = null) {
  const client = new Anthropic({ apiKey });

  const today = new Date().toISOString().slice(0, 10); // e.g. 2026-02-24

  // Build the season map section for the prompt
  const seasonMapEntries = Object.entries(seasonMap);
  const seasonMapSection = seasonMapEntries.length > 0
    ? `\n## Seller-Provided Season Map (${seasonMapEntries.length} SKUs mapped)\n` +
      seasonMapEntries.slice(0, 300).map(([sku, s]) => `${sku}: ${s}`).join('\n')
    : '';

  // Build a compact representation — key fields only, capped rows
  const reportSections = summaryData.map(block => {
    const compacted = compactRows(block);
    const dataStr   = JSON.stringify(compacted, null, 0);
    const rowDesc   = block.isAggregated
      ? `aggregated ${block.aggregatedFrom} events → ${compacted.length} SKUs`
      : block.isGlobalContext
        ? `global context, ${compacted.length} rows`
        : `${block.rowCount} total rows, showing ${compacted.length}`;
    return `
## ${block.reportType} (${block.fileName}, ${rowDesc})
Headers: ${block.headers.join(', ')}
Data: ${dataStr}
${block.criticalItems   !== undefined ? `Critical (≤14 days supply): ${block.criticalItems}` : ''}
${block.lowItems        !== undefined ? `Low (15-30 days supply): ${block.lowItems}` : ''}
${block.outOfStock      !== undefined ? `Out of Stock: ${block.outOfStock}` : ''}
${block.unsellableUnits !== undefined ? `Unsellable Units: ${block.unsellableUnits}` : ''}
    `.trim();
  }).join('\n\n');

  const systemPrompt = `You are an Amazon FBA footwear inventory management expert. Your job is to analyze Amazon Seller Central reports and provide clear, prioritized replenishment recommendations with full awareness of product seasonality.

You will receive data from these five Amazon Seller Central report types:
- **Inventory Ledger** (90-day): All inventory movements pre-aggregated to one row per SKU — shows net received, returns, removals, and adjustments.
- **Restock Inventory**: Amazon's own restocking suggestions. Key fields: the recommended quantity column (may be named "Recommended replenishment qty", "Recommended Order Quantity", or "Suggested Order Quantity"), "Days of Supply at Amazon", and "Recommended ship date" / "Order by date".
- **Amazon Fulfilled Inventory**: Current FBA stock snapshot. Key fields: afn-fulfillable-qty (sellable), afn-unsellable-qty (stranded/damaged), afn-warehouse-qty (total at FC).
- **Manage FBA Inventory**: Fulfillable, reserved, and inbound unit breakdown.
- **Business Report - By Child ASIN** (90-day): Sales velocity by ASIN. The "Units Ordered" column is the total units sold in the 90-day window. Match to SKUs via ASIN.

## Velocity Calculation
Use the **Business Report's "Units Ordered"** as your primary velocity source — it is the most accurate. Divide by 90 for daily velocity, then multiply by 30 for a 30-day estimate.
- Match Business Report rows to SKUs/ASINs from other reports using the Child ASIN column.
- If the Business Report is not available or an ASIN doesn't appear, fall back to "Units sold per day" from the Restock Inventory report, or derive from the Inventory Ledger.

## Seasonality Rules for Footwear
The seller operates in footwear where seasonality is critical. You MUST apply the following logic:

1. **Season assignment** — For each item, determine its season using this priority order:
   a. Use the seller-provided season map if the SKU/ASIN appears there.
   b. Otherwise, infer from the product name/description using common footwear signals:
      - Spring/Summer: sandals, flip flops, water shoes, aqua shoes, beach shoes, slides, open-toe, pool shoes, sport sandals, canvas sneakers, espadrilles, boat shoes, mesh sneakers
      - Fall/Winter: boots, snow boots, winter boots, ankle boots, Chelsea boots, lined shoes, shearling, insulated, waterproof boots, clogs, slippers, fur-lined
      - Year-Round: sneakers (without seasonal qualifier), athletic shoes, dress shoes, loafers, oxfords, mules (when not sandal-style)
   c. If you cannot determine the season, mark as "Unknown" and note it.

2. **Season-aware priority boosting** — Today's date is ${today}. The seller is preparing for: **${targetSeason || 'both seasons'}**.
   - If the target season is Spring/Summer: BOOST priority for Spring/Summer items (they need stock now). SUPPRESS priority for Fall/Winter items (don't order off-season inventory unnecessarily — only flag CRITICAL stockouts for those).
   - If the target season is Fall/Winter: apply the reverse logic.
   - If "Both / Year-Round": use standard priority rules for all seasons.
   - Year-Round items always use standard priority rules.

3. **Seasonal timing context**:
   - Spring/Summer season typically runs March–August. FBA lead times mean sellers should start shipping water shoes, sandals etc. by early February for Spring.
   - Fall/Winter season typically runs September–February. Sellers should start shipping boots, winter footwear by July/August.

## Cross-referencing Amazon's Suggestions
The Restock Inventory report contains Amazon's own replenishment recommendation per SKU. For each item:
1. Set **amazon_suggested_qty** = the value from the Restock Inventory report's recommended quantity column (null if the SKU isn't in that report).
2. Calculate **our_suggested_qty** = your own recommendation based on: velocity (from Business Report), current fulfillable stock (from Amazon Fulfilled Inventory or Manage FBA), days of supply, seasonal urgency, and a target of 60 days of supply for in-season items.
3. If amazon_suggested_qty and our_suggested_qty differ by more than 25%, set **qty_discrepancy_flag** to a concise explanation, e.g.:
   - "Amazon suggests 2× more — may reflect Amazon's safety stock model or a promotion"
   - "We suggest 40% more due to strong in-season velocity not yet in Amazon's model"
   - "Amazon suggests ordering but days of supply is 90+ — verify velocity data"
4. If they're within 25% of each other, set qty_discrepancy_flag to null.

Your response MUST be structured as valid JSON with this exact schema:
{
  "summary": "2-3 sentence executive summary mentioning target season and overall inventory health",
  "urgent_action_required": boolean,
  "replenishment_items": [
    {
      "sku": "string",
      "asin": "string",
      "product_name": "string",
      "season": "Spring/Summer | Fall/Winter | Year-Round | Unknown",
      "season_source": "mapped | inferred | unknown",
      "priority": "CRITICAL | HIGH | MEDIUM | LOW | HOLD",
      "priority_reason": "string — include season reasoning if relevant",
      "current_stock": number_or_null,
      "days_of_supply": number_or_null,
      "units_sold_30d": number_or_null,
      "amazon_suggested_qty": number_or_null,
      "our_suggested_qty": number_or_null,
      "qty_discrepancy_flag": "string_or_null",
      "order_by_date": "string_or_null",
      "action": "string describing what to do"
    }
  ],
  "watch_list": [
    {
      "sku": "string",
      "asin": "string",
      "product_name": "string",
      "season": "string",
      "concern": "string",
      "recommendation": "string"
    }
  ],
  "insights": ["array of key observations — include seasonal readiness observations"],
  "data_quality_notes": ["any missing data, unmapped seasons, or report issues noticed"]
}

Priority definitions (after seasonal adjustment):
- CRITICAL: Stockout imminent (≤7 days supply) or already out of stock for IN-SEASON items → Order immediately
- HIGH: Low stock (8-21 days supply) for in-season items → Order this week
- MEDIUM: Getting low (22-45 days supply) for in-season items, OR CRITICAL/HIGH for off-season items → Plan accordingly
- LOW: Suggested by Amazon but not urgent (>45 days supply) for in-season; or monitor for off-season
- HOLD: Off-season item with sufficient stock — do not order, just hold

Sort replenishment_items: CRITICAL first, then HIGH, then MEDIUM, then LOW, then HOLD. Within each priority, sort by days_of_supply ascending.
Include HOLD items so the seller sees their full picture, but clearly mark them.
Only include items that appear in the uploaded data.`;

  const batchNote = batchNum
    ? `\nNote: This is batch ${batchNum} of ${totalBatches}. Analyze ONLY the SKUs in this batch's data — do not comment on missing data from other batches.\n`
    : '';

  const userMessage = `Please analyze my Amazon Seller Central footwear inventory reports.

Today's date: ${today}
Target season I'm preparing for: ${targetSeason || 'Both / Year-Round'}
${userNotes ? `Additional context from seller: ${userNotes}` : ''}${batchNote}
${seasonMapSection}

${reportSections}

Provide your analysis as valid JSON matching the schema in the system prompt.`;

  const onRateLimitWait = (waitSecs, attempt) => {
    const label = batchNum ? `batch ${batchNum}/${totalBatches}` : 'request';
    console.warn(`  Rate limit on ${label} — waiting ${waitSecs}s (attempt ${attempt})`);
    setProgress(requestId, {
      stage: 'rate_limited', batchCurrent: batchNum ?? 0, batchTotal: totalBatches ?? 1,
      message: `Rate limit hit on ${label}. Retrying in ${waitSecs}s… (attempt ${attempt})`,
    });
  };

  const response = await callClaudeWithRetry(client, {
    model: 'claude-sonnet-4-6',
    max_tokens: 16000,
    system: systemPrompt,
    messages: [{ role: 'user', content: userMessage }]
  }, onRateLimitWait);

  const text = response.content[0].text;

  // Extract JSON from response (handle markdown code blocks)
  const jsonMatch = text.match(/```(?:json)?\s*([\s\S]*?)\s*```/) || text.match(/(\{[\s\S]*\})/);
  if (!jsonMatch) {
    throw new Error('Could not parse JSON response from Claude. Raw response: ' + text.slice(0, 500));
  }

  let jsonStr = jsonMatch[1];

  // Detect and repair truncated JSON caused by token limits
  if (response.stop_reason === 'max_tokens') {
    console.warn('Warning: Claude response hit max_tokens limit — attempting JSON repair');
    // Close any open arrays/objects by counting brackets
    const opens = (jsonStr.match(/[\[{]/g) || []).length;
    const closes = (jsonStr.match(/[\]}]/g) || []).length;
    const deficit = opens - closes;
    if (deficit > 0) {
      // Trim to last complete object in the replenishment_items array if possible
      const lastCompleteObj = jsonStr.lastIndexOf('},');
      if (lastCompleteObj > 0) {
        jsonStr = jsonStr.slice(0, lastCompleteObj + 1); // end after last }
      }
      // Re-close all open structures
      for (let i = 0; i < deficit; i++) {
        jsonStr += (jsonStr.trimEnd().endsWith(']') || jsonStr.trimEnd().endsWith('}')) ? '' : '';
      }
      // Best-effort: close arrays then object
      jsonStr = jsonStr.replace(/,\s*$/, ''); // remove trailing comma
      jsonStr += ']}'; // close replenishment_items array + root object
      console.warn('Attempted JSON repair. Result may be partial.');
    }
  }

  try {
    return JSON.parse(jsonStr);
  } catch (parseErr) {
    throw new Error(
      `Failed to parse Claude's JSON response: ${parseErr.message}. ` +
      `This usually means your report has too many SKUs. Try uploading one report at a time, or filter your report to fewer SKUs before uploading.`
    );
  }
}

// ─────────────────────────────────────────────
// API Routes
// ─────────────────────────────────────────────

// Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Progress polling endpoint
app.get('/api/progress/:requestId', (req, res) => {
  const data = progressMap.get(req.params.requestId);
  res.json(data || { stage: 'unknown', message: 'No progress info yet.' });
});

// Main analyze endpoint
app.post('/api/analyze', upload.array('files', 10), async (req, res) => {
  try {
    const { apiKey, userNotes, targetSeason, requestId } = req.body;

    if (!apiKey || !apiKey.startsWith('sk-')) {
      return res.status(400).json({ error: 'A valid Anthropic API key is required (starts with sk-).' });
    }

    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: 'Please upload at least one report file.' });
    }

    // Parse all uploaded files
    const parsedFiles = [];
    const parseErrors = [];

    for (const file of req.files) {
      try {
        const sheets = parseFile(file.buffer, file.originalname);
        parsedFiles.push({ fileName: file.originalname, sheets });
      } catch (err) {
        parseErrors.push({ file: file.originalname, error: err.message });
      }
    }

    if (parsedFiles.length === 0) {
      return res.status(400).json({
        error: 'Could not parse any of the uploaded files.',
        details: parseErrors
      });
    }

    // Extract season map from any uploaded season map files
    const seasonMap = extractSeasonMap(parsedFiles);

    // Summarize and analyze (auto-batches if SKU count > BATCH_SIZE)
    const summaryData  = summarizeData(parsedFiles);
    const totalSKUs    = getAllIdentifiers(summaryData).length;
    const batchCount   = Math.ceil(totalSKUs / BATCH_SIZE);
    console.log(`Total unique SKUs detected: ${totalSKUs} → ${batchCount} batch(es)`);
    setProgress(requestId, {
      stage: 'preparing', batchCurrent: 0, batchTotal: batchCount, skuCount: totalSKUs,
      message: `Found ${totalSKUs} SKUs across ${batchCount} batch${batchCount !== 1 ? 'es' : ''}. Starting analysis…`,
    });

    const analysis = await analyzeInBatches(apiKey, summaryData, userNotes, targetSeason, seasonMap, requestId);

    // Include detection metadata in response
    const filesMeta = parsedFiles.map(f => ({
      fileName: f.fileName,
      sheets: f.sheets.map(s => ({
        sheetName: s.sheetName,
        reportType: s.reportType,
        rowCount: s.rowCount
      }))
    }));

    res.json({
      success: true,
      filesProcessed: filesMeta,
      parseErrors,
      seasonMapCount:   Object.keys(seasonMap).length,
      targetSeason:     targetSeason || 'Both / Year-Round',
      totalSKUsAnalyzed: totalSKUs,
      batchCount,
      analysis
    });

    // Clean up progress entry after 2 min
    if (requestId) setTimeout(() => progressMap.delete(requestId), 2 * 60 * 1000);

  } catch (err) {
    console.error('Analysis error:', err);

    if (err.status === 401) {
      return res.status(401).json({ error: 'Invalid Anthropic API key. Please check your key and try again.' });
    }
    if (err.status === 429) {
      return res.status(429).json({ error: 'API rate limit reached. Please wait a moment and try again.' });
    }

    res.status(500).json({ error: err.message || 'An unexpected error occurred.' });
  }
});

// Start server
app.listen(PORT, () => {
  console.log(`\n🚀 Amazon Inventory Replenishment Tracker running at http://localhost:${PORT}\n`);
});
