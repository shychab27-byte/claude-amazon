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
  'FBA Manage Inventory': [
    ['sku', 'asin', 'fnsku', 'available', 'fulfillable'],
    ['seller-sku', 'asin', 'condition', 'available']
  ],
  'Inventory Health Report': [
    ['inventory age', 'sell through', 'days of supply'],
    ['inv age', 'units sold (last 30 days)', 'days of supply']
  ],
  'Restocking Report': [
    ['suggested order quantity', 'order by date', 'max inventory level'],
    ['units sold per day', 'suggested order quantity']
  ],
  'Inventory Ledger Report': [
    ['event type', 'fulfillment center', 'disposition', 'reconciled quantity'],
    ['fnsku', 'event type', 'quantity', 'fulfillment center']
  ],
  'All Orders Report': [
    ['amazon-order-id', 'purchase-date', 'order-status', 'fulfillment-channel'],
    ['amazon-order-id', 'asin', 'quantity', 'item-price']
  ],
  'Season Map': [
    ['sku', 'season'],
    ['asin', 'season']
  ]
};

// Valid season values (normalized)
const VALID_SEASONS = ['spring/summer', 'fall/winter', 'year-round', 'all'];

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
        data: rows.slice(0, 500) // cap at 500 rows to keep API payload reasonable
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
      if (reportType === 'Restocking Report') {
        block.criticalItems = data.filter(row => {
          const dos = parseFloat(row['Days of Supply'] || row['days of supply'] || 0);
          return dos >= 0 && dos <= 14;
        }).length;
        block.lowItems = data.filter(row => {
          const dos = parseFloat(row['Days of Supply'] || row['days of supply'] || 0);
          return dos > 14 && dos <= 30;
        }).length;
      }

      if (reportType === 'Inventory Health Report') {
        block.zeroSellThrough = data.filter(row => {
          const st = parseFloat(row['Sell Through'] || row['sell through'] || 0);
          return st === 0;
        }).length;
      }

      if (reportType === 'FBA Manage Inventory') {
        block.outOfStock = data.filter(row => {
          const avail = parseInt(row['Available'] || row['available'] || 0);
          return avail === 0;
        }).length;
      }

      summary.push(block);
    }
  }

  return summary;
}

// ─────────────────────────────────────────────
// Claude API Analysis
// ─────────────────────────────────────────────

async function analyzeWithClaude(apiKey, summaryData, userNotes, targetSeason, seasonMap) {
  const client = new Anthropic({ apiKey });

  const today = new Date().toISOString().slice(0, 10); // e.g. 2026-02-24

  // Build the season map section for the prompt
  const seasonMapEntries = Object.entries(seasonMap);
  const seasonMapSection = seasonMapEntries.length > 0
    ? `\n## Seller-Provided Season Map (${seasonMapEntries.length} SKUs mapped)\n` +
      seasonMapEntries.slice(0, 300).map(([sku, s]) => `${sku}: ${s}`).join('\n')
    : '';

  // Build a compact but complete representation of the data
  const reportSections = summaryData.map(block => {
    const dataStr = JSON.stringify(block.data.slice(0, 150), null, 0);
    return `
## ${block.reportType} (${block.fileName}, ${block.rowCount} rows)
Headers: ${block.headers.join(', ')}
Data (up to 150 rows): ${dataStr}
${block.criticalItems !== undefined ? `Critical (≤14 days supply): ${block.criticalItems}` : ''}
${block.lowItems !== undefined ? `Low (15-30 days supply): ${block.lowItems}` : ''}
${block.outOfStock !== undefined ? `Out of Stock: ${block.outOfStock}` : ''}
    `.trim();
  }).join('\n\n');

  const systemPrompt = `You are an Amazon FBA footwear inventory management expert. Your job is to analyze Amazon Seller Central reports and provide clear, prioritized replenishment recommendations with full awareness of product seasonality.

You will receive data from one or more of these report types:
- FBA Manage Inventory: Current stock levels per SKU/ASIN
- Inventory Health Report: Days of supply, sell-through rate, aged inventory
- Restocking Report: Amazon's suggested quantities, order-by dates
- Inventory Ledger Report: Inventory movement history
- All Orders Report: Historical order data for velocity calculation

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
      "current_inventory": number_or_null,
      "days_of_supply": number_or_null,
      "units_sold_30d": number_or_null,
      "suggested_order_qty": number_or_null,
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

  const userMessage = `Please analyze my Amazon Seller Central footwear inventory reports.

Today's date: ${today}
Target season I'm preparing for: ${targetSeason || 'Both / Year-Round'}
${userNotes ? `Additional context from seller: ${userNotes}` : ''}
${seasonMapSection}

${reportSections}

Provide your analysis as valid JSON matching the schema in the system prompt.`;

  const response = await client.messages.create({
    model: 'claude-sonnet-4-6',
    max_tokens: 8096,
    system: systemPrompt,
    messages: [{ role: 'user', content: userMessage }]
  });

  const text = response.content[0].text;

  // Extract JSON from response (handle markdown code blocks)
  const jsonMatch = text.match(/```(?:json)?\s*([\s\S]*?)\s*```/) || text.match(/(\{[\s\S]*\})/);
  if (!jsonMatch) {
    throw new Error('Could not parse JSON response from Claude. Raw response: ' + text.slice(0, 500));
  }

  return JSON.parse(jsonMatch[1]);
}

// ─────────────────────────────────────────────
// API Routes
// ─────────────────────────────────────────────

// Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Main analyze endpoint
app.post('/api/analyze', upload.array('files', 10), async (req, res) => {
  try {
    const { apiKey, userNotes, targetSeason } = req.body;

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

    // Summarize and analyze
    const summaryData = summarizeData(parsedFiles);
    const analysis = await analyzeWithClaude(apiKey, summaryData, userNotes, targetSeason, seasonMap);

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
      seasonMapCount: Object.keys(seasonMap).length,
      targetSeason: targetSeason || 'Both / Year-Round',
      analysis
    });

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
