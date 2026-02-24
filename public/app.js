/* ─────────────────────────────────────────────
   Amazon Inventory Replenishment Tracker
   Client-side logic
   ───────────────────────────────────────────── */

// ── State ──────────────────────────────────────
let selectedFiles = [];
let lastAnalysis  = null;
let activeFilter  = 'ALL';

// ── Element refs ───────────────────────────────
const $ = id => document.getElementById(id);

const apiKeyInput    = $('apiKey');
const toggleKeyBtn   = $('toggleKey');
const userNotesInput = $('userNotes');
const dropZone       = $('dropZone');
const fileInput      = $('fileInput');
const fileListEl     = $('fileList');
const analyzeBtn     = $('analyzeBtn');
const analyzeBtnText = $('analyzeBtnText');
const analyzeSpinner = $('analyzeBtnSpinner');
const resultsSection = $('resultsSection');
const errorBox       = $('errorBox');
const errorText      = $('errorText');
const statusBanner   = $('statusBanner');
const filesProcessed = $('filesProcessed');
const summaryText    = $('summaryText');
const replenishBody  = $('replenishBody');
const noResults      = $('noResults');
const exportBtn      = $('exportBtn');
const resetBtn       = $('resetBtn');

// ── API Key toggle ─────────────────────────────
toggleKeyBtn.addEventListener('click', () => {
  apiKeyInput.type = apiKeyInput.type === 'password' ? 'text' : 'password';
});

// ── Drop Zone ──────────────────────────────────
dropZone.addEventListener('click', () => fileInput.click());

dropZone.addEventListener('dragover', e => {
  e.preventDefault();
  dropZone.classList.add('over');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('over');
  addFiles(Array.from(e.dataTransfer.files));
});

fileInput.addEventListener('change', () => {
  addFiles(Array.from(fileInput.files));
  fileInput.value = '';
});

function addFiles(newFiles) {
  const allowed = ['.xlsx', '.xls', '.csv', '.txt'];
  const dupeNames = new Set(selectedFiles.map(f => f.name));

  for (const f of newFiles) {
    const ext = f.name.slice(f.name.lastIndexOf('.')).toLowerCase();
    if (!allowed.includes(ext)) {
      showError(`"${f.name}" is not a supported file type. Use .xlsx, .xls, .csv, or .txt`);
      continue;
    }
    if (dupeNames.has(f.name)) continue;
    selectedFiles.push(f);
    dupeNames.add(f.name);
  }

  renderFileList();
}

function removeFile(idx) {
  selectedFiles.splice(idx, 1);
  renderFileList();
}

function formatBytes(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function fileIcon(name) {
  const ext = name.slice(name.lastIndexOf('.')).toLowerCase();
  if (ext === '.csv' || ext === '.txt') return '📄';
  return '📊';
}

function renderFileList() {
  if (selectedFiles.length === 0) {
    fileListEl.classList.add('hidden');
    fileListEl.innerHTML = '';
    analyzeBtn.disabled = true;
    return;
  }

  fileListEl.classList.remove('hidden');
  fileListEl.innerHTML = selectedFiles.map((f, i) => `
    <div class="file-item">
      <span class="file-icon">${fileIcon(f.name)}</span>
      <div class="file-info">
        <div class="file-name" title="${esc(f.name)}">${esc(f.name)}</div>
        <div class="file-meta">${formatBytes(f.size)}</div>
      </div>
      <button class="file-remove" onclick="removeFile(${i})" title="Remove">✕</button>
    </div>
  `).join('');

  analyzeBtn.disabled = false;
}

// ── Analyze ────────────────────────────────────
analyzeBtn.addEventListener('click', runAnalysis);

async function runAnalysis() {
  hideError();

  const apiKey = apiKeyInput.value.trim();
  if (!apiKey) {
    showError('Please enter your Anthropic API key.');
    apiKeyInput.focus();
    return;
  }
  if (!apiKey.startsWith('sk-')) {
    showError('API key should start with "sk-". Please check your key.');
    apiKeyInput.focus();
    return;
  }
  if (selectedFiles.length === 0) {
    showError('Please upload at least one report file.');
    return;
  }

  setLoading(true);
  resultsSection.classList.add('hidden');

  try {
    const formData = new FormData();
    formData.append('apiKey', apiKey);
    formData.append('userNotes', userNotesInput.value.trim());
    for (const f of selectedFiles) formData.append('files', f);

    const resp = await fetch('/api/analyze', {
      method: 'POST',
      body: formData
    });

    const data = await resp.json();

    if (!resp.ok) {
      throw new Error(data.error || `Server error ${resp.status}`);
    }

    lastAnalysis = data.analysis;
    renderResults(data);
    resultsSection.classList.remove('hidden');
    resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });

  } catch (err) {
    showError(err.message);
  } finally {
    setLoading(false);
  }
}

function setLoading(loading) {
  analyzeBtn.disabled = loading;
  analyzeBtnText.textContent = loading ? 'Analyzing…' : 'Analyze Inventory';
  analyzeSpinner.classList.toggle('hidden', !loading);
}

// ── Render Results ─────────────────────────────
function renderResults(data) {
  const a = data.analysis;

  // Status banner
  if (a.urgent_action_required) {
    statusBanner.className = 'status-banner urgent';
    statusBanner.innerHTML = '⚠️ Urgent Action Required — Some items need immediate attention.';
  } else {
    statusBanner.className = 'status-banner ok';
    statusBanner.innerHTML = '✅ Inventory looks manageable — review the list below to stay ahead.';
  }

  // Files processed
  filesProcessed.innerHTML = (data.filesProcessed || []).flatMap(f =>
    f.sheets.map(s => `
      <span class="tag">
        <span class="tag-dot" style="background:${reportColor(s.reportType)}"></span>
        ${esc(s.reportType)} &nbsp;<span style="font-weight:400;color:var(--muted)">${s.rowCount} rows</span>
      </span>
    `)
  ).join('');

  // Summary
  summaryText.textContent = a.summary || '';

  // Replenishment table
  renderTable(a.replenishment_items || []);

  // Watch list
  const watchCard = $('watchCard');
  const watchItems = a.watch_list || [];
  if (watchItems.length > 0) {
    watchCard.classList.remove('hidden');
    $('watchList').innerHTML = watchItems.map(w => `
      <div class="watch-item">
        <div class="watch-item-header">
          <span class="watch-sku">${esc(w.sku || w.asin || '')}</span>
          <span class="watch-name">${esc(w.product_name || '')}</span>
        </div>
        <div class="watch-concern">⚠ ${esc(w.concern || '')}</div>
        <div class="watch-rec">→ ${esc(w.recommendation || '')}</div>
      </div>
    `).join('');
  } else {
    watchCard.classList.add('hidden');
  }

  // Insights
  const insightsCard = $('insightsCard');
  const insights = a.insights || [];
  if (insights.length > 0) {
    insightsCard.classList.remove('hidden');
    $('insightsList').innerHTML = insights.map(i => `<li>${esc(i)}</li>`).join('');
  } else {
    insightsCard.classList.add('hidden');
  }

  // Data notes
  const notesCard = $('notesCard');
  const notes = a.data_quality_notes || [];
  if (notes.length > 0) {
    notesCard.classList.remove('hidden');
    $('notesList').innerHTML = notes.map(n => `<li>${esc(n)}</li>`).join('');
  } else {
    notesCard.classList.add('hidden');
  }
}

function renderTable(items) {
  if (items.length === 0) {
    replenishBody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:32px;color:var(--muted)">No replenishment items found.</td></tr>';
    noResults.classList.add('hidden');
    return;
  }

  replenishBody.innerHTML = items.map((item, idx) => {
    const priority  = (item.priority || 'LOW').toUpperCase();
    const hidden    = activeFilter !== 'ALL' && priority !== activeFilter;
    const dos       = item.days_of_supply != null ? Number(item.days_of_supply).toFixed(0) : '—';
    const onHand    = item.current_inventory != null ? item.current_inventory : '—';
    const sold30    = item.units_sold_30d != null ? item.units_sold_30d : '—';
    const orderQty  = item.suggested_order_qty != null ? item.suggested_order_qty : '—';
    const orderBy   = item.order_by_date ? esc(item.order_by_date) : '—';

    return `
      <tr data-priority="${priority}" class="${hidden ? 'hidden-row' : ''}">
        <td><span class="badge badge-${priority}" title="${esc(item.priority_reason || '')}">${priority}</span></td>
        <td>
          <div class="product-sku">${esc(item.sku || '')}</div>
          <div class="product-sku" style="margin-top:2px;color:#9ca3af">${esc(item.asin || '')}</div>
        </td>
        <td class="product-cell">
          <div class="product-name">${esc(item.product_name || '—')}</div>
        </td>
        <td class="num">${onHand}</td>
        <td class="num ${dosClass(item.days_of_supply)}">${dos}</td>
        <td class="num">${sold30}</td>
        <td class="num"><strong>${orderQty}</strong></td>
        <td>${orderBy}</td>
        <td class="action-cell">${esc(item.action || '')}</td>
      </tr>
    `;
  }).join('');

  applyFilter();
}

function dosClass(dos) {
  if (dos == null) return '';
  if (dos <= 7)  return 'style="color:var(--red);font-weight:700"';
  if (dos <= 21) return 'style="color:var(--yellow)"';
  return '';
}

// ── Filter Tabs ────────────────────────────────
$('filterTabs').addEventListener('click', e => {
  const tab = e.target.closest('.filter-tab');
  if (!tab) return;
  activeFilter = tab.dataset.priority;
  document.querySelectorAll('.filter-tab').forEach(t => t.classList.remove('active'));
  tab.classList.add('active');
  applyFilter();
});

function applyFilter() {
  const rows = replenishBody.querySelectorAll('tr[data-priority]');
  let visible = 0;
  rows.forEach(row => {
    const match = activeFilter === 'ALL' || row.dataset.priority === activeFilter;
    row.classList.toggle('hidden-row', !match);
    if (match) visible++;
  });
  noResults.classList.toggle('hidden', visible > 0);
}

// ── Export CSV ────────────────────────────────
exportBtn.addEventListener('click', () => {
  if (!lastAnalysis || !lastAnalysis.replenishment_items) return;

  const headers = ['Priority','SKU','ASIN','Product Name','On Hand','Days of Supply','Units Sold 30d','Suggested Order Qty','Order By Date','Action','Priority Reason'];
  const rows = lastAnalysis.replenishment_items.map(item => [
    item.priority,
    item.sku,
    item.asin,
    item.product_name,
    item.current_inventory,
    item.days_of_supply,
    item.units_sold_30d,
    item.suggested_order_qty,
    item.order_by_date,
    item.action,
    item.priority_reason
  ].map(v => `"${String(v ?? '').replace(/"/g, '""')}"`));

  const csv = [headers.join(','), ...rows.map(r => r.join(','))].join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url;
  a.download = `replenishment-${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
});

// ── Reset ──────────────────────────────────────
resetBtn.addEventListener('click', () => {
  selectedFiles = [];
  lastAnalysis  = null;
  activeFilter  = 'ALL';
  renderFileList();
  resultsSection.classList.add('hidden');
  hideError();
  window.scrollTo({ top: 0, behavior: 'smooth' });
});

// ── Helpers ────────────────────────────────────
function esc(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function showError(msg) {
  errorText.textContent = msg;
  errorBox.classList.remove('hidden');
  errorBox.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function hideError() {
  errorBox.classList.add('hidden');
}

function reportColor(type) {
  const colors = {
    'FBA Manage Inventory':    '#146EB4',
    'Inventory Health Report': '#1a7f37',
    'Restocking Report':       '#FF9900',
    'Inventory Ledger Report': '#6f42c1',
    'All Orders Report':       '#0891b2',
  };
  return colors[type] || '#6b7280';
}
