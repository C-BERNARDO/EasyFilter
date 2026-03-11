/* ═══════════════════════════════════════
   STATE
═══════════════════════════════════════ */
let allData      = [];   // [{date, account, client, balance, ptpAmount, claimPaid, relation, callStatus, status}]
let uniqueDates   = [];
let uniqueClients = [];
let uniqueStatuses = [];
let selectedDates   = new Set();
let selectedClients = new Set();
let selectedStatuses = new Set();

/* ═══════════════════════════════════════
   DOM REFERENCES
═══════════════════════════════════════ */
const fileInput      = document.getElementById('fileInput');
const uploadZone     = document.getElementById('uploadZone');
const fileLoadedRow  = document.getElementById('fileLoadedRow');
const fileNameEl     = document.getElementById('fileName');
const dateBlock      = document.getElementById('dateBlock');
const clientBlock    = document.getElementById('clientBlock');
const statusBlock    = document.getElementById('statusBlock');
const dateList       = document.getElementById('dateList');
const clientList     = document.getElementById('clientList');
const statusList     = document.getElementById('statusList');
const dateSelCount   = document.getElementById('dateSelCount');
const clientSelCount = document.getElementById('clientSelCount');
const statusSelCount = document.getElementById('statusSelCount');
const clientSearch   = document.getElementById('clientSearch');
const emptyState     = document.getElementById('emptyState');
const dashboard      = document.getElementById('dashboard');
const topbarMeta     = document.getElementById('topbarMeta');
const metaRows       = document.getElementById('metaRows');
const metaDates      = document.getElementById('metaDates');
const metaClients    = document.getElementById('metaClients');
const resetBtn       = document.getElementById('resetBtn');
const kpiAccounts    = document.getElementById('kpiAccounts');
const kpiBalance     = document.getElementById('kpiBalance');
const kpiPtpSum      = document.getElementById('kpiPtpSum');
const kpiPtpCount    = document.getElementById('kpiPtpCount');
const kpiClaimSum    = document.getElementById('kpiClaimSum');
const kpiClaimCount  = document.getElementById('kpiClaimCount');
const kpiDebtors     = document.getElementById('kpiDebtors');
const kpiDebtorBal   = document.getElementById('kpiDebtorBal');
const kpiConnected    = document.getElementById('kpiConnected');
const kpiConnectedBal = document.getElementById('kpiConnectedBal');
const btnDownloadDebtors   = document.getElementById('btnDownloadDebtors');
const btnDownloadConnected = document.getElementById('btnDownloadConnected');
const kpiAccountsSub = document.getElementById('kpiAccountsSub');
const kpiBalanceSub  = document.getElementById('kpiBalanceSub');
const filterSummary  = document.getElementById('filterSummary');
const breakdownWrap  = document.getElementById('breakdownWrap');
const tableNote      = document.getElementById('tableNote');

/* ═══════════════════════════════════════
   FILE UPLOAD
═══════════════════════════════════════ */
// Note: the <label for="fileInput"> inside uploadZone handles click-to-browse natively.
// No extra uploadZone click listener needed — adding one caused the file dialog to open twice.
const loadingOverlay = document.getElementById('loadingOverlay');

fileInput.addEventListener('change', e => { if (e.target.files[0]) processFile(e.target.files[0]); });

uploadZone.addEventListener('dragover',  e => { e.preventDefault(); uploadZone.classList.add('drag-over'); });
uploadZone.addEventListener('dragleave', () => uploadZone.classList.remove('drag-over'));
uploadZone.addEventListener('drop', e => {
  e.preventDefault(); uploadZone.classList.remove('drag-over');
  const f = e.dataTransfer.files[0];
  if (f && /\.(xlsx|xls)$/i.test(f.name)) processFile(f);
});

resetBtn.addEventListener('click', resetAll);

/* ── Parse Excel ── */
function processFile(file) {
  loadingOverlay.style.display = 'flex';
  const reader = new FileReader();
  reader.onload = e => {
    // Small timeout lets the browser paint the overlay before heavy parsing
    setTimeout(() => {
    try {
      const wb   = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

      if (!rows.length) { loadingOverlay.style.display = 'none'; alert('The file appears to be empty.'); return; }

      const keys       = Object.keys(rows[0]);
      const findKey    = name => keys.find(k => k.trim().toLowerCase() === name.toLowerCase());
      const dateKey    = findKey('date');
      const accountKey = findKey('account no.');
      const clientKey  = findKey('client');
      const balanceKey = findKey('balance');
      const ptpKey     = findKey('ptp amount');
      const claimKey   = findKey('claim paid amount');
      const relationKey   = findKey('relation');
      const callStatusKey = findKey('call status');
      const statusKey     = findKey('status');

      const missing = [];
      if (!dateKey)       missing.push('"Date"');
      if (!accountKey)    missing.push('"Account No."');
      if (!clientKey)     missing.push('"Client"');
      if (!balanceKey)    missing.push('"Balance"');
      if (!ptpKey)        missing.push('"PTP Amount"');
      if (!claimKey)      missing.push('"Claim Paid Amount"');
      if (!relationKey)   missing.push('"Relation"');
      if (!callStatusKey) missing.push('"Call Status"');
      if (!statusKey)     missing.push('"Status"');
      if (missing.length) {
        loadingOverlay.style.display = 'none';
        alert(`Missing column(s): ${missing.join(', ')}\nPlease check your Excel file headers.`);
        return;
      }

      allData = rows.map(row => ({
        date:       formatDate(row[dateKey]),
        account:    String(row[accountKey]).trim(),
        client:     String(row[clientKey]).trim(),
        balance:    parseBalance(row[balanceKey]),
        ptpAmount:  parseBalance(row[ptpKey]),
        claimPaid:  parseBalance(row[claimKey]),
        relation:   String(row[relationKey]).trim(),
        callStatus: String(row[callStatusKey]).trim(),
        status:     String(row[statusKey]).trim()
      })).filter(r => r.date && r.account);

      fileNameEl.textContent = file.name;
      fileLoadedRow.style.display = 'flex';

      buildFilters();
      topbarMeta.style.display = 'flex';
      metaRows.textContent    = `${allData.length} rows`;
      metaDates.textContent   = `${uniqueDates.length} dates`;
      metaClients.textContent = `${uniqueClients.length} clients`;
      emptyState.style.display = 'none';
      dashboard.style.display  = 'flex';
      refreshResults();

    } catch (err) {
      alert('Error reading file: ' + err.message);
    } finally {
      loadingOverlay.style.display = 'none';
      fileInput.value = ''; // reset so same file can be re-uploaded
    }
    }, 50);
  };
  reader.readAsArrayBuffer(file);
}

/* ── Helpers ── */
function parseBalance(val) {
  if (typeof val === 'number') return val;
  const n = parseFloat(String(val).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function formatDate(val) {
  if (!val && val !== 0) return '';
  if (val instanceof Date) {
    if (isNaN(val)) return '';
    return val.toLocaleDateString('en-US', { year:'numeric', month:'short', day:'numeric' });
  }
  if (typeof val === 'number') {
    const d = XLSX.SSF.parse_date_code(val);
    if (!d) return String(val);
    const dt = new Date(d.y, d.m - 1, d.d);
    return dt.toLocaleDateString('en-US', { year:'numeric', month:'short', day:'numeric' });
  }
  const d = new Date(val);
  if (!isNaN(d)) return d.toLocaleDateString('en-US', { year:'numeric', month:'short', day:'numeric' });
  return String(val).trim();
}

/* ═══════════════════════════════════════
   BUILD FILTER LISTS
═══════════════════════════════════════ */
function buildFilters() {
  // Dates
  const ds = new Set(); allData.forEach(r => ds.add(r.date));
  uniqueDates = [...ds].sort((a,b) => new Date(a) - new Date(b));
  selectedDates = new Set(uniqueDates);

  // Clients
  const cs = new Set(); allData.forEach(r => cs.add(r.client));
  uniqueClients = [...cs].sort((a,b) => a.localeCompare(b));
  selectedClients = new Set(uniqueClients);

  // Statuses
  const ss = new Set(); allData.forEach(r => ss.add(r.status));
  uniqueStatuses = [...ss].sort((a,b) => a.localeCompare(b));
  selectedStatuses = new Set(uniqueStatuses);

  renderDateList();
  renderClientList('');
  renderStatusList();
  updateCountBadges();

  dateBlock.style.display   = 'block';
  clientBlock.style.display = 'block';
  statusBlock.style.display = 'block';
}

/* ── Date checkboxes ── */
function renderDateList() {
  dateList.innerHTML = '';
  uniqueDates.forEach(date => {
    const item = makeCheckItem(date, selectedDates.has(date), date, (checked) => {
      checked ? selectedDates.add(date) : selectedDates.delete(date);
      updateCountBadges();
      refreshResults();
    });
    dateList.appendChild(item);
  });
}

/* ── Client checkboxes (with search filter) ── */
function renderClientList(query) {
  clientList.innerHTML = '';
  const q = query.toLowerCase();
  const filtered = uniqueClients.filter(c => !q || c.toLowerCase().includes(q));
  if (!filtered.length) {
    const p = document.createElement('p');
    p.style.cssText = 'font-size:.78rem;color:var(--muted2);padding:8px;text-align:center;';
    p.textContent = 'No clients match';
    clientList.appendChild(p); return;
  }
  filtered.forEach(client => {
    const item = makeCheckItem(client, selectedClients.has(client), client, (checked) => {
      checked ? selectedClients.add(client) : selectedClients.delete(client);
      updateCountBadges();
      refreshResults();
    });
    clientList.appendChild(item);
  });
}

clientSearch.addEventListener('input', () => renderClientList(clientSearch.value));

/* ── Status checkboxes ── */
function renderStatusList() {
  statusList.innerHTML = '';
  uniqueStatuses.forEach(status => {
    const item = makeCheckItem(status, selectedStatuses.has(status), status, (checked) => {
      checked ? selectedStatuses.add(status) : selectedStatuses.delete(status);
      updateCountBadges();
      refreshResults();
    });
    statusList.appendChild(item);
  });
}

/* ── Checkbox factory ── */
function makeCheckItem(value, checked, labelText, onChange) {
  const label = document.createElement('label');
  label.className = 'chk-item';
  const inp = document.createElement('input');
  inp.type = 'checkbox'; inp.value = value; inp.checked = checked;
  inp.addEventListener('change', () => onChange(inp.checked));
  const box = document.createElement('div'); box.className = 'chk-box';
  const lbl = document.createElement('span'); lbl.className = 'chk-lbl'; lbl.textContent = labelText;
  label.appendChild(inp); label.appendChild(box); label.appendChild(lbl);
  return label;
}

function syncCheckboxes(listEl, selectedSet) {
  listEl.querySelectorAll('input[type=checkbox]').forEach(cb => {
    cb.checked = selectedSet.has(cb.value);
  });
}

function updateCountBadges() {
  dateSelCount.textContent   = selectedDates.size;
  clientSelCount.textContent = selectedClients.size;
  statusSelCount.textContent = selectedStatuses.size;
}

/* ── Select / None ── */
document.getElementById('dateAll').addEventListener('click', () => {
  selectedDates = new Set(uniqueDates);
  syncCheckboxes(dateList, selectedDates);
  updateCountBadges(); refreshResults();
});
document.getElementById('dateNone').addEventListener('click', () => {
  selectedDates.clear();
  syncCheckboxes(dateList, selectedDates);
  updateCountBadges(); refreshResults();
});
document.getElementById('clientAll').addEventListener('click', () => {
  selectedClients = new Set(uniqueClients);
  syncCheckboxes(clientList, selectedClients);
  updateCountBadges(); refreshResults();
});
document.getElementById('clientNone').addEventListener('click', () => {
  selectedClients.clear();
  syncCheckboxes(clientList, selectedClients);
  updateCountBadges(); refreshResults();
});
document.getElementById('statusAll').addEventListener('click', () => {
  selectedStatuses = new Set(uniqueStatuses);
  syncCheckboxes(statusList, selectedStatuses);
  updateCountBadges(); refreshResults();
});
document.getElementById('statusNone').addEventListener('click', () => {
  selectedStatuses.clear();
  syncCheckboxes(statusList, selectedStatuses);
  updateCountBadges(); refreshResults();
});

/* ═══════════════════════════════════════
   FILTERING & CALCULATION
═══════════════════════════════════════ */
function getFilteredRows() {
  return allData.filter(r =>
    selectedDates.has(r.date) &&
    selectedClients.has(r.client) &&
    selectedStatuses.has(r.status)
  );
}

/**
 * Deduplicate by Account No. — for each unique account keep the FIRST
 * occurrence to pick its balance (consistent rule).
 */
function deduplicateByAccount(rows) {
  const seen = new Map();
  rows.forEach(r => {
    if (!seen.has(r.account)) seen.set(r.account, r);
  });
  return [...seen.values()];
}

function refreshResults() {
  const filtered   = getFilteredRows();
  const unique     = deduplicateByAccount(filtered);
  const totalAccts = unique.length;
  const totalBal   = unique.reduce((s, r) => s + r.balance, 0);

  // PTP Amount and Claim Paid: calculated on ALL filtered rows (not deduped),
  // counting only rows where the value is non-zero
  const ptpRows   = filtered.filter(r => r.ptpAmount !== 0);
  const claimRows = filtered.filter(r => r.claimPaid !== 0);
  const ptpSum    = ptpRows.reduce((s, r) => s + r.ptpAmount, 0);
  const claimSum  = claimRows.reduce((s, r) => s + r.claimPaid, 0);

  // Unique Debtors: distinct Account No. values where Relation === 'Debtor'
  // Also sum their balance (first occurrence per account, same dedup rule)
  const debtorRows    = filtered.filter(r => r.relation.toLowerCase() === 'debtor');
  const debtorUnique  = deduplicateByAccount(debtorRows);
  const totalDebtors  = debtorUnique.length;
  const totalDebtorBal = debtorUnique.reduce((s, r) => s + r.balance, 0);

  // Unique Connected: distinct Account No. values where Call Status === 'CONNECTED'
  // Also sum their balance (first occurrence per account, same dedup rule)
  const connectedRows    = filtered.filter(r => r.callStatus.toUpperCase() === 'CONNECTED');
  const connectedUnique  = deduplicateByAccount(connectedRows);
  const totalConnected   = connectedUnique.length;
  const totalConnectedBal = connectedUnique.reduce((s, r) => s + r.balance, 0);

  animateNum(kpiAccounts, totalAccts, false);

  // Balance: display exact full number (no abbreviation)
  animateNumExact(kpiBalance, totalBal);
  animateNumExact(kpiPtpSum, ptpSum);
  animateNumExact(kpiClaimSum, claimSum);

  kpiPtpCount.textContent   = ptpRows.length.toLocaleString();
  kpiClaimCount.textContent = claimRows.length.toLocaleString();
  animateNum(kpiDebtors,    totalDebtors,    false);
  animateNumExact(kpiDebtorBal, totalDebtorBal);
  animateNum(kpiConnected,  totalConnected,  false);
  animateNumExact(kpiConnectedBal, totalConnectedBal);

  const allDatesSelected   = selectedDates.size === uniqueDates.length;
  const allClientsSelected = selectedClients.size === uniqueClients.length;
  const allStatusesSelected = selectedStatuses.size === uniqueStatuses.length;
  kpiAccountsSub.textContent = `of ${allData.length} total rows`;
  kpiBalanceSub.textContent  = `${totalAccts} unique account${totalAccts !== 1 ? 's' : ''} summed`;

  renderFilterSummary(allDatesSelected, allClientsSelected, allStatusesSelected);
  renderBreakdown(unique, filtered);
  tableNote.textContent = `${totalAccts} unique account${totalAccts!==1?'s':''} · ${formatCurrencyExact(totalBal)}`;
}

/* ═══════════════════════════════════════
   FILTER SUMMARY CHIPS
═══════════════════════════════════════ */
function renderFilterSummary(allDates, allClients, allStatuses) {
  filterSummary.innerHTML = '';

  if (allDates && allClients && allStatuses) {
    filterSummary.appendChild(chip_('All dates · All clients · All statuses', 'fchip fchip-all'));
    return;
  }

  if (allDates) {
    filterSummary.appendChild(chip_('All dates', 'fchip fchip-date'));
  } else {
    [...selectedDates].sort((a,b)=>new Date(a)-new Date(b)).slice(0, 6).forEach(d => {
      filterSummary.appendChild(chip_(d, 'fchip fchip-date', 'Date'));
    });
    if (selectedDates.size > 6)
      filterSummary.appendChild(chip_(`+${selectedDates.size - 6} more dates`, 'fchip fchip-date'));
  }

  if (allClients) {
    filterSummary.appendChild(chip_('All clients', 'fchip fchip-client'));
  } else {
    [...selectedClients].sort().slice(0, 4).forEach(c => {
      filterSummary.appendChild(chip_(c, 'fchip fchip-client', 'Client'));
    });
    if (selectedClients.size > 4)
      filterSummary.appendChild(chip_(`+${selectedClients.size - 4} more clients`, 'fchip fchip-client'));
  }

  if (allStatuses) {
    filterSummary.appendChild(chip_('All statuses', 'fchip fchip-status'));
  } else {
    [...selectedStatuses].sort().slice(0, 5).forEach(s => {
      filterSummary.appendChild(chip_(s, 'fchip fchip-status', 'Status'));
    });
    if (selectedStatuses.size > 5)
      filterSummary.appendChild(chip_(`+${selectedStatuses.size - 5} more statuses`, 'fchip fchip-status'));
  }
}

function chip_(text, cls, prefix) {
  const el = document.createElement('div');
  el.className = cls;
  el.innerHTML = prefix ? `<em class="fchip-label">${prefix}</em>${text}` : text;
  return el;
}

/* ═══════════════════════════════════════
   BREAKDOWN TABLE
═══════════════════════════════════════ */
function renderBreakdown(allUnique, filteredRows) {
  breakdownWrap.innerHTML = '';

  if (!allUnique.length) {
    const el = document.createElement('div');
    el.className = 'bd-empty'; el.textContent = 'No data matches the current filters.';
    breakdownWrap.appendChild(el); return;
  }

  [...selectedDates].sort((a,b) => new Date(a)-new Date(b)).forEach((date, i) => {
    const dateRows   = filteredRows.filter(r => r.date === date);
    const dateUnique = deduplicateByAccount(dateRows);
    const ptpRows    = dateRows.filter(r => r.ptpAmount !== 0);
    const claimRows  = dateRows.filter(r => r.claimPaid !== 0);

    const count    = dateUnique.length;
    const balance  = dateUnique.reduce((s, r) => s + r.balance, 0);
    const ptpSum   = ptpRows.reduce((s, r) => s + r.ptpAmount, 0);
    const claimSum = claimRows.reduce((s, r) => s + r.claimPaid, 0);

    const row = document.createElement('div');
    row.className = 'bd-row';
    row.style.animationDelay = `${i * 30}ms`;
    row.innerHTML = `
      <span class="bd-date">${date}</span>
      <span class="bd-count">${count.toLocaleString()}</span>
      <span class="bd-balance">${formatCurrencyExact(balance)}</span>
      <span class="bd-ptp">
        <span class="bd-val">${formatCurrencyExact(ptpSum)}</span>
        <span class="bd-cnt">${ptpRows.length} entries</span>
      </span>
      <span class="bd-claim">
        <span class="bd-val">${formatCurrencyExact(claimSum)}</span>
        <span class="bd-cnt">${claimRows.length} entries</span>
      </span>
    `;
    breakdownWrap.appendChild(row);
  });
}

/* ═══════════════════════════════════════
   ANIMATED COUNTER
═══════════════════════════════════════ */
const _counters = new WeakMap();

function animateNum(el, target, isCurrency) {
  const prev = _counters.get(el) || 0;
  _counters.set(el, target);
  if (prev === target) return;

  const duration = 500;
  const start    = performance.now();

  function step(now) {
    const t  = Math.min((now - start) / duration, 1);
    const e  = 1 - Math.pow(1 - t, 3); // ease-out cubic
    const v  = prev + (target - prev) * e;
    el.textContent = isCurrency ? formatCurrencyExact(v) : Math.round(v).toLocaleString();
    if (t < 1) requestAnimationFrame(step);
    else { el.textContent = isCurrency ? formatCurrencyExact(target) : target.toLocaleString(); }
  }
  requestAnimationFrame(step);
}

/* Animate to exact full decimal value (no abbreviation) */
function animateNumExact(el, target) {
  const prev = _counters.get(el) || 0;
  _counters.set(el, target);
  if (prev === target) return;

  const duration = 500;
  const start    = performance.now();

  function step(now) {
    const t = Math.min((now - start) / duration, 1);
    const e = 1 - Math.pow(1 - t, 3);
    const v = prev + (target - prev) * e;
    el.textContent = formatCurrencyExact(v);
    if (t < 1) requestAnimationFrame(step);
    else el.textContent = formatCurrencyExact(target);
  }
  requestAnimationFrame(step);
}

/* Always show full number with 2 decimal places, comma-separated */
function formatCurrencyExact(n) {
  return n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatCurrency(n) {
  return formatCurrencyExact(n);
}

/* ═══════════════════════════════════════
   DOWNLOAD DEBTORS
═══════════════════════════════════════ */
btnDownloadDebtors.addEventListener('click', downloadDebtors);
btnDownloadConnected.addEventListener('click', downloadConnected);

function downloadDebtors() {
  // Get filtered rows → keep only Debtor relation → deduplicate by Account No.
  const filtered      = getFilteredRows();
  const debtorRows    = filtered.filter(r => r.relation.toLowerCase() === 'debtor');
  const debtorUnique  = deduplicateByAccount(debtorRows);

  if (!debtorUnique.length) {
    alert('No Debtor records found for the current filters.');
    return;
  }

  // Build worksheet data — friendly column headers
  const wsData = [
    ['Date', 'Account No.', 'Client', 'Status', 'Relation', 'Call Status',
     'Balance', 'PTP Amount', 'Claim Paid Amount']
  ];

  debtorUnique.forEach(r => {
    wsData.push([
      r.date,
      r.account,
      r.client,
      r.status,
      r.relation,
      r.callStatus,
      r.balance,
      r.ptpAmount,
      r.claimPaid
    ]);
  });

  // Create workbook and style the header row
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // Set column widths for readability
  ws['!cols'] = [
    { wch: 14 }, // Date
    { wch: 18 }, // Account No.
    { wch: 24 }, // Client
    { wch: 16 }, // Status
    { wch: 12 }, // Relation
    { wch: 16 }, // Call Status
    { wch: 16 }, // Balance
    { wch: 16 }, // PTP Amount
    { wch: 20 }, // Claim Paid Amount
  ];

  XLSX.utils.book_append_sheet(wb, ws, 'Debtors');

  // Build a descriptive filename with active filters
  const datePart   = selectedDates.size === uniqueDates.length
    ? 'AllDates'
    : [...selectedDates].sort().join('_').replace(/[\s,]/g, '').slice(0, 40);
  const statusPart = selectedStatuses.size === uniqueStatuses.length
    ? 'AllStatuses'
    : [...selectedStatuses].sort().join('_').replace(/\s/g, '').slice(0, 30);
  const timestamp  = new Date().toISOString().slice(0,10);
  const filename   = `Debtors_${datePart}_${statusPart}_${timestamp}.xlsx`;

  XLSX.writeFile(wb, filename);
}

/* ═══════════════════════════════════════
   DOWNLOAD CONNECTED
═══════════════════════════════════════ */
function downloadConnected() {
  const filtered         = getFilteredRows();
  const connectedRows    = filtered.filter(r => r.callStatus.toUpperCase() === 'CONNECTED');
  const connectedUnique  = deduplicateByAccount(connectedRows);

  if (!connectedUnique.length) {
    alert('No Connected records found for the current filters.');
    return;
  }

  const wsData = [
    ['Date', 'Account No.', 'Client', 'Status', 'Relation', 'Call Status',
     'Balance', 'PTP Amount', 'Claim Paid Amount']
  ];

  connectedUnique.forEach(r => {
    wsData.push([
      r.date,
      r.account,
      r.client,
      r.status,
      r.relation,
      r.callStatus,
      r.balance,
      r.ptpAmount,
      r.claimPaid
    ]);
  });

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);

  ws['!cols'] = [
    { wch: 14 }, // Date
    { wch: 18 }, // Account No.
    { wch: 24 }, // Client
    { wch: 16 }, // Status
    { wch: 12 }, // Relation
    { wch: 16 }, // Call Status
    { wch: 16 }, // Balance
    { wch: 16 }, // PTP Amount
    { wch: 20 }, // Claim Paid Amount
  ];

  XLSX.utils.book_append_sheet(wb, ws, 'Connected');

  const datePart   = selectedDates.size === uniqueDates.length
    ? 'AllDates'
    : [...selectedDates].sort().join('_').replace(/[\s,]/g, '').slice(0, 40);
  const statusPart = selectedStatuses.size === uniqueStatuses.length
    ? 'AllStatuses'
    : [...selectedStatuses].sort().join('_').replace(/\s/g, '').slice(0, 30);
  const timestamp  = new Date().toISOString().slice(0, 10);
  const filename   = `Connected_${datePart}_${statusPart}_${timestamp}.xlsx`;

  XLSX.writeFile(wb, filename);
}

function resetAll() {
  allData = []; uniqueDates = []; uniqueClients = []; uniqueStatuses = [];
  selectedDates.clear(); selectedClients.clear(); selectedStatuses.clear();

  fileInput.value = '';
  fileLoadedRow.style.display = 'none';
  dateBlock.style.display     = 'none';
  clientBlock.style.display   = 'none';
  statusBlock.style.display   = 'none';
  topbarMeta.style.display    = 'none';
  emptyState.style.display    = 'flex';
  dashboard.style.display     = 'none';
  dateList.innerHTML = ''; clientList.innerHTML = ''; statusList.innerHTML = '';
  clientSearch.value = '';
  dateSelCount.textContent = clientSelCount.textContent = statusSelCount.textContent = '0';
}