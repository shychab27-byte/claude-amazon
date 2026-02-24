/* ─────────────────────────────────────────────
   Restock Tracker — grouped by Parent ASIN
   ───────────────────────────────────────────── */

// ── State ──────────────────────────────────────
let selectedFile = null;
let lastGroups   = [];
let activeFilter = 'ALL';

// ── Element refs ───────────────────────────────
const $ = id => document.getElementById(id);

const dropZone            = $('dropZone');
const fileInput           = $('fileInput');
const fileListEl          = $('fileList');
const parseBtn            = $('parseBtn');
const parseBtnText        = $('parseBtnText');
const parseBtnSpinner     = $('parseBtnSpinner');
const resultsSection      = $('resultsSection');
const errorBox            = $('errorBox');
const errorTextEl         = $('errorText');
const replenishBody       = $('replenishBody');
const noResults           = $('noResults');
const exportBtn           = $('exportBtn');
const resetBtn            = $('resetBtn');
const itemCountEl         = $('itemCount');
const noParentAsinWarning = $('noParentAsinWarning');

// ── Drop zone ──────────────────────────────────
dropZone.addEventListener('click', () => fileInput.click());

dropZone.addEventListener('dragover', e => {
  e.preventDefault();
  dropZone.classList.add('over');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('over');
  setFile(e.dataTransfer.files[0]);
});

fileInput.addEventListener('change', () => {
  setFile(fileInput.files[0]);
  fileInput.value = '';
});

function setFile(f) {
  if (!f) return;
  const ext = f.name.slice(f.name.lastIndexOf('.')).toLowerCase();
  if (!['.xlsx', '.xls', '.csv', '.txt'].includes(ext)) {
    showError(`"${f.name}" is not a supported file type.`);
    return;
  }
  selectedFile = f;
  renderFileItem();
}

function renderFileItem() {
  if (!selectedFile) {
    fileListEl.classList.add('hidden');
    fileListEl.innerHTML = '';
    parseBtn.disabled = true;
    return;
  }
  fileListEl.classList.remove('hidden');
  fileListEl.innerHTML = `
    <div class="file-item">
      <span class="file-icon">📄</span>
      <div class="file-info">
        <div class="file-name">${esc(selectedFile.name)}</div>
        <div class="file-meta">${formatBytes(selectedFile.size)}</div>
      </div>
      <button class="file-remove" onclick="clearFile()" title="Remove">✕</button>
    </div>
  `;
  parseBtn.disabled = false;
}

function clearFile() {
  selectedFile = null;
  renderFileItem();
}

function formatBytes(bytes) {
  if (bytes < 1024)    return bytes + ' B';
  if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / 1048576).toFixed(1) + ' MB';
}

// ── Parse ──────────────────────────────────────
parseBtn.addEventListener('click', runParse);

async function runParse() {
  if (!selectedFile) return;
  hideError();
  setLoading(true);
  resultsSection.classList.add('hidden');

  try {
    const fd = new FormData();
    fd.append('file', selectedFile);
    const resp = await fetch('/api/parse', { method: 'POST', body: fd });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.error || `Server error ${resp.status}`);

    lastGroups = data.groups || [];
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
  parseBtn.disabled = loading;
  parseBtnText.textContent = loading ? 'Loading…' : 'Load Report';
  parseBtnSpinner.classList.toggle('hidden', !loading);
}

// ── Render results ─────────────────────────────
const PRIORITY_LABEL = {
  OUT_OF_STOCK: 'Out of Stock',
  CRITICAL:     'Critical',
  HIGH:         'High',
  MEDIUM:       'Medium',
  LOW:          'Low',
  UNKNOWN:      'Unknown',
};

const PRIORITY_ORDER = {
  OUT_OF_STOCK: 0, CRITICAL: 1, HIGH: 2, MEDIUM: 3, LOW: 4, UNKNOWN: 5
};

function renderResults(data) {
  const totalItems  = data.itemCount  || 0;
  const totalStyles = data.groupCount || 0;
  const urgentCount = (data.groups || []).filter(
    g => g.worst_priority === 'OUT_OF_STOCK' || g.worst_priority === 'CRITICAL'
  ).length;

  itemCountEl.textContent =
    `${totalItems} SKUs · ${totalStyles} style${totalStyles !== 1 ? 's' : ''}` +
    (urgentCount > 0 ? ` · ${urgentCount} need attention` : '');

  noParentAsinWarning.classList.toggle('hidden', !!data.hasParentAsin);

  renderTable(data.groups || []);

  // Reset filter to ALL
  activeFilter = 'ALL';
  document.querySelectorAll('.filter-tab').forEach(t =>
    t.classList.toggle('active', t.dataset.priority === 'ALL')
  );
}

// ── Table rendering ────────────────────────────
function renderTable(groups) {
  if (groups.length === 0) {
    replenishBody.innerHTML =
      '<tr><td colspan="7" class="empty-cell">No items found in report.</td></tr>';
    return;
  }

  let html = '';

  groups.forEach((group, gi) => {
    const wp         = group.worst_priority || 'UNKNOWN';
    const itemCount  = group.items?.length || 0;
    const skuLabel   = itemCount === 1 ? 'SKU' : 'SKUs';
    const summaryTxt = buildSummaryText(group.counts || {});

    // ── Group header row ──────────────────────────────────────────────────────
    html += `
      <tr class="group-header" data-group-priority="${wp}" data-group-id="${gi}">
        <td colspan="7">
          <div class="group-header-inner" onclick="toggleGroup(${gi})">
            <span class="group-toggle" id="toggle-${gi}">▾</span>
            <span class="badge badge-${wp}">${PRIORITY_LABEL[wp] || wp}</span>
            <span class="group-asin">${esc(group.parent_asin || '')}</span>
            <span class="group-name">${esc(group.group_name || '')}</span>
            <span class="group-spacer"></span>
            <span class="group-meta">${itemCount} ${skuLabel}${summaryTxt ? ' · ' + summaryTxt : ''}</span>
          </div>
        </td>
      </tr>`;

    // ── Child SKU rows ────────────────────────────────────────────────────────
    (group.items || []).forEach(item => {
      const priority = item.priority || 'UNKNOWN';
      const dos      = item.days_of_supply != null
        ? Number(item.days_of_supply).toFixed(0) : null;
      const stock    = item.current_stock != null ? item.current_stock : '—';
      const reorder  = item.reorder_qty   != null ? item.reorder_qty   : '—';
      const orderBy  = item.order_by_date ? esc(item.order_by_date) : '—';

      html += `
        <tr class="item-row" data-group-id="${gi}">
          <td><span class="badge badge-${priority}">${PRIORITY_LABEL[priority] || priority}</span></td>
          <td><span class="sku-mono">${esc(item.sku)}</span></td>
          <td class="product-cell"><div class="product-name">${esc(item.product_name || '—')}</div></td>
          <td class="num">${stock}</td>
          <td class="num">${dosBadge(item.days_of_supply, dos)}</td>
          <td class="num"><strong>${reorder}</strong></td>
          <td class="order-by-cell">${orderBy}</td>
        </tr>`;
    });
  });

  replenishBody.innerHTML = html;
  applyFilter();
}

function buildSummaryText(counts) {
  const order  = ['OUT_OF_STOCK', 'CRITICAL', 'HIGH', 'MEDIUM', 'LOW'];
  const labels = {
    OUT_OF_STOCK: 'OOS',
    CRITICAL: 'critical',
    HIGH:     'high',
    MEDIUM:   'medium',
    LOW:      'low',
  };
  return order
    .filter(p => counts[p] > 0)
    .map(p => `<span class="meta-${p.toLowerCase()}">${counts[p]} ${labels[p]}</span>`)
    .join(' · ');
}

function dosBadge(dos, formatted) {
  if (dos == null || formatted == null) return '—';
  const n = Number(dos);
  if (n <= 0)  return `<span class="dos-oos">${formatted}d</span>`;
  if (n <= 7)  return `<span class="dos-critical">${formatted}d</span>`;
  if (n <= 21) return `<span class="dos-high">${formatted}d</span>`;
  return `${formatted}d`;
}

// ── Collapse / expand ──────────────────────────
function toggleGroup(gi) {
  const rows   = replenishBody.querySelectorAll(`.item-row[data-group-id="${gi}"]`);
  const toggle = $(`toggle-${gi}`);
  const isNowCollapsed = toggle?.textContent === '▸';
  rows.forEach(r => r.classList.toggle('collapsed', !isNowCollapsed));
  if (toggle) toggle.textContent = isNowCollapsed ? '▾' : '▸';
}

// ── Priority filter ────────────────────────────
$('filterTabs').addEventListener('click', e => {
  const tab = e.target.closest('.filter-tab');
  if (!tab) return;
  activeFilter = tab.dataset.priority;
  document.querySelectorAll('.filter-tab').forEach(t =>
    t.classList.toggle('active', t === tab)
  );
  applyFilter();
});

function applyFilter() {
  const headers = replenishBody.querySelectorAll('tr.group-header');
  let visibleGroups = 0;

  headers.forEach(header => {
    const gi      = header.dataset.groupId;
    const gp      = header.dataset.groupPriority;
    const visible = activeFilter === 'ALL' || gp === activeFilter;

    header.classList.toggle('hidden-row', !visible);

    // Show/hide child rows (respecting their individual collapsed state too)
    replenishBody.querySelectorAll(`.item-row[data-group-id="${gi}"]`).forEach(row =>
      row.classList.toggle('hidden-row', !visible)
    );

    if (visible) visibleGroups++;
  });

  noResults.classList.toggle('hidden', visibleGroups > 0);
}

// ── Export CSV ─────────────────────────────────
exportBtn.addEventListener('click', () => {
  if (!lastGroups.length) return;

  const headers = [
    'Parent ASIN', 'Style Worst Priority',
    'SKU', 'ASIN', 'Product Name',
    'Current Stock', 'Days of Supply', 'Reorder Qty', 'Order By Date',
    'Priority',
  ];

  const rows = [];
  for (const group of lastGroups) {
    for (const item of (group.items || [])) {
      rows.push([
        group.parent_asin     ?? '',
        group.worst_priority  ?? '',
        item.sku              ?? '',
        item.asin             ?? '',
        item.product_name     ?? '',
        item.current_stock    ?? '',
        item.days_of_supply   ?? '',
        item.reorder_qty      ?? '',
        item.order_by_date    ?? '',
        item.priority         ?? '',
      ].map(v => `"${String(v).replace(/"/g, '""')}"`));
    }
  }

  const csv  = [headers.join(','), ...rows.map(r => r.join(','))].join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = `restock-${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
});

// ── Reset ──────────────────────────────────────
resetBtn.addEventListener('click', () => {
  selectedFile = null;
  lastGroups   = [];
  activeFilter = 'ALL';
  renderFileItem();
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
  errorTextEl.textContent = msg;
  errorBox.classList.remove('hidden');
  errorBox.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function hideError() {
  errorBox.classList.add('hidden');
}
