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
  ]
};

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

function summarizeData(parsedFiles) {
  const summary = [];

  for (const file of parsedFiles) {
    for (const sheet of file.sheets) {
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

async function analyzeWithClaude(apiKey, summaryData, userNotes) {
  const client = new Anthropic({ apiKey });

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

  const systemPrompt = `You are an Amazon FBA inventory management expert. Your job is to analyze uploaded Amazon Seller Central reports and provide clear, prioritized replenishment recommendations.

You will receive data from one or more of these report types:
- FBA Manage Inventory: Current stock levels per SKU/ASIN
- Inventory Health Report: Days of supply, sell-through rate, aged inventory
- Restocking Report: Amazon's suggested quantities, order-by dates
- Inventory Ledger Report: Inventory movement history
- All Orders Report: Historical order data for velocity calculation

Your response MUST be structured as valid JSON with this exact schema:
{
  "summary": "2-3 sentence executive summary of inventory situation",
  "urgent_action_required": boolean,
  "replenishment_items": [
    {
      "sku": "string",
      "asin": "string",
      "product_name": "string",
      "priority": "CRITICAL | HIGH | MEDIUM | LOW",
      "priority_reason": "string",
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
      "concern": "string",
      "recommendation": "string"
    }
  ],
  "insights": ["array of key observations about the inventory"],
  "data_quality_notes": ["any missing data or report issues noticed"]
}

Priority definitions:
- CRITICAL: Stockout imminent (≤7 days supply) or already out of stock → Order immediately
- HIGH: Low stock (8-21 days supply) → Order this week
- MEDIUM: Getting low (22-45 days supply) → Plan order soon
- LOW: Suggested by Amazon but not urgent (>45 days supply)

Sort replenishment_items by priority (CRITICAL first) then by days_of_supply ascending.
Only include items that genuinely need attention. Do not pad the list.`;

  const userMessage = `Please analyze my Amazon Seller Central inventory reports and tell me what I need to replenish.

${userNotes ? `Additional context from seller: ${userNotes}\n` : ''}

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
    const { apiKey, userNotes } = req.body;

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

    // Summarize and analyze
    const summaryData = summarizeData(parsedFiles);
    const analysis = await analyzeWithClaude(apiKey, summaryData, userNotes);

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
