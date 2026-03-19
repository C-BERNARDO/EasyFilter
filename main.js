/* ═══════════════════════════════════════
   STATE
═══════════════════════════════════════ */
let allData      = [];   // [{date, account, client, balance, ptpAmount, claimPaid, relation, callStatus, status, remarkType}]
let uniqueDates   = [];
let uniqueClients = [];
let uniqueStatuses = [];
let uniqueRemarkTypes = [];
let uniqueRelations   = [];
let selectedDates       = new Set();
let selectedClients     = new Set();
let selectedStatuses    = new Set();
let selectedRemarkTypes = new Set();
let selectedRelations   = new Set();

/* ═══════════════════════════════════════
   CONSTANTS
═══════════════════════════════════════ */
const BLANK_RELATION = '__blank__'; // sentinel for empty Relation values

/* ═══════════════════════════════════════
   DOM REFERENCES
═══════════════════════════════════════ */
const fileInput      = document.getElementById('fileInput');
const uploadZone     = document.getElementById('uploadZone');
const fileLoadedRow  = document.getElementById('fileLoadedRow');
const dateBlock      = document.getElementById('dateBlock');
const clientBlock    = document.getElementById('clientBlock');
const statusBlock    = document.getElementById('statusBlock');
const dateList       = document.getElementById('dateList');
const clientList     = document.getElementById('clientList');
const statusList     = document.getElementById('statusList');
const dateSelCount   = document.getElementById('dateSelCount');
const clientSelCount = document.getElementById('clientSelCount');
const statusSelCount = document.getElementById('statusSelCount');
const remarkTypeBlock    = document.getElementById('remarkTypeBlock');
const remarkTypeList     = document.getElementById('remarkTypeList');
const remarkTypeSelCount = document.getElementById('remarkTypeSelCount');
const clientSearch   = document.getElementById('clientSearch');
const emptyState     = document.getElementById('emptyState');
const dashboard      = document.getElementById('dashboard');
const topbarMeta     = document.getElementById('topbarMeta');
const metaRows       = document.getElementById('metaRows');
const metaDates      = document.getElementById('metaDates');
const metaClients    = document.getElementById('metaClients');
const resetBtn       = document.getElementById('resetBtn');
const kpiAccounts    = document.getElementById('kpiAccounts');
const kpiDials       = document.getElementById('kpiDials');
const kpiBalance     = document.getElementById('kpiBalance');
const kpiPtpSum      = document.getElementById('kpiPtpSum');
const kpiPtpCount    = document.getElementById('kpiPtpCount');
const kpiClaimSum    = document.getElementById('kpiClaimSum');
const kpiClaimCount  = document.getElementById('kpiClaimCount');
const kpiPositiveCount  = document.getElementById('kpiPositiveCount');
const kpiPositiveBal    = document.getElementById('kpiPositiveBal');
const btnDownloadPositive  = document.getElementById('btnDownloadPositive');
const relationBlock    = document.getElementById('relationBlock');
const relationList     = document.getElementById('relationList');
const relationSelCount = document.getElementById('relationSelCount');
const kpiAccountsSub = document.getElementById('kpiAccountsSub');
const kpiBalanceSub  = document.getElementById('kpiBalanceSub');
const filterSummary  = document.getElementById('filterSummary');
const breakdownWrap  = document.getElementById('breakdownWrap');
const tableNote      = document.getElementById('tableNote');

/* ═══════════════════════════════════════
   FILE UPLOAD  — multi-file
═══════════════════════════════════════ */
// Note: the <label for="fileInput"> handles click-to-browse natively (no extra click listener needed).
const loadingOverlay = document.getElementById('loadingOverlay');
const loadSub        = document.getElementById('loadSub');
const fileChips      = document.getElementById('fileChips');

fileInput.addEventListener('change', e => {
  const files = [...e.target.files].filter(f => /\.(xlsx|xls)$/i.test(f.name));
  if (files.length) processFiles(files);
});

uploadZone.addEventListener('dragover',  e => { e.preventDefault(); uploadZone.classList.add('drag-over'); });
uploadZone.addEventListener('dragleave', () => uploadZone.classList.remove('drag-over'));
uploadZone.addEventListener('drop', e => {
  e.preventDefault(); uploadZone.classList.remove('drag-over');
  const files = [...e.dataTransfer.files].filter(f => /\.(xlsx|xls)$/i.test(f.name));
  if (files.length) processFiles(files);
});

resetBtn.addEventListener('click', resetAll);

/* ── Read a single file → array of raw row objects ── */
function readFile(file, index, total) {
  return new Promise((resolve, reject) => {
    loadSub.textContent = `Reading file ${index + 1} of ${total}: ${file.name}`;
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb   = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
        resolve({ rows, fileName: file.name });
      } catch (err) {
        reject(new Error(`${file.name}: ${err.message}`));
      }
    };
    reader.onerror = () => reject(new Error(`Failed to read ${file.name}`));
    reader.readAsArrayBuffer(file);
  });
}

/* ── Process all selected files ── */
async function processFiles(files) {
  loadingOverlay.style.display = 'flex';
  loadSub.textContent = `Preparing ${files.length} file${files.length > 1 ? 's' : ''}…`;

  // Small delay to let the browser render the overlay
  await new Promise(r => setTimeout(r, 50));

  try {
    // Read all files in parallel, but update progress label sequentially via index
    const results = await Promise.all(
      files.map((f, i) => readFile(f, i, files.length))
    );

    loadSub.textContent = 'Validating columns…';
    await new Promise(r => setTimeout(r, 20));

    // Validate columns per file, collect errors
    const errors = [];
    const allRows = [];

    for (const { rows, fileName } of results) {
      if (!rows.length) { errors.push(`"${fileName}" appears to be empty.`); continue; }

      const keys          = Object.keys(rows[0]);
      const findKey       = name => keys.find(k => k.trim().toLowerCase() === name.toLowerCase());
      const dateKey       = findKey('date');
      const accountKey    = findKey('account no.');
      const clientKey     = findKey('client');
      const balanceKey    = findKey('balance');
      const ptpKey        = findKey('ptp amount');
      const claimKey      = findKey('claim paid amount');
      const relationKey   = findKey('relation');
      const callStatusKey = findKey('call status');
      const statusKey     = findKey('status');
      const remarkTypeKey = findKey('remark type');

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
        errors.push(`"${fileName}" missing column(s): ${missing.join(', ')}`);
        continue;
      }

      // Map rows from this file
      const mapped = rows.map(row => ({
        date:       formatDate(row[dateKey]),
        account:    String(row[accountKey]).trim(),
        client:     String(row[clientKey]).trim(),
        balance:    parseBalance(row[balanceKey]),
        ptpAmount:  parseBalance(row[ptpKey]),
        claimPaid:  parseBalance(row[claimKey]),
        relation:   String(row[relationKey]).trim(),
        callStatus: String(row[callStatusKey]).trim(),
        status:     String(row[statusKey]).trim(),
        remarkType: remarkTypeKey ? String(row[remarkTypeKey]).trim() : ''
      })).filter(r => r.date && r.account);

      allRows.push(...mapped);
    }

    if (errors.length) {
      alert('Some files had issues:\n\n' + errors.join('\n'));
    }

    if (!allRows.length) {
      loadingOverlay.style.display = 'none';
      fileInput.value = '';
      return;
    }

    loadSub.textContent = 'Building dashboard…';
    await new Promise(r => setTimeout(r, 20));

    // Replace (don't accumulate) — fresh load each time
    allData = allRows;

    // Render file chips in sidebar
    renderFileChips(files.map(f => f.name));

    buildFilters();
    topbarMeta.style.display = 'flex';
    metaRows.textContent    = `${allData.length} rows`;
    metaDates.textContent   = `${uniqueDates.length} dates`;
    metaClients.textContent = `${uniqueClients.length} clients`;
    emptyState.style.display = 'none';
    dashboard.style.display  = 'flex';
    refreshResults();

  } catch (err) {
    alert('Error: ' + err.message);
  } finally {
    loadingOverlay.style.display = 'none';
    fileInput.value = '';
  }
}

/* ── Render loaded file chips under upload zone ── */
function renderFileChips(names) {
  fileChips.innerHTML = '';
  names.forEach(name => {
    const chip = document.createElement('div');
    chip.className = 'fchip-file';
    chip.innerHTML = `
      <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
        <polyline points="14 2 14 8 20 8"/>
      </svg>
      <span>${name}</span>`;
    fileChips.appendChild(chip);
  });
  fileLoadedRow.style.display = 'block';
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

  // Remark Types
  const rs = new Set(); allData.forEach(r => { if (r.remarkType) rs.add(r.remarkType); });
  uniqueRemarkTypes = [...rs].sort((a,b) => a.localeCompare(b));
  selectedRemarkTypes = new Set(uniqueRemarkTypes);

  // Relations — include a sentinel for blank/empty values
  const rels = new Set();
  allData.forEach(r => { rels.add(r.relation === '' ? BLANK_RELATION : r.relation); });
  uniqueRelations = [...rels].sort((a, b) => {
    if (a === BLANK_RELATION) return 1;   // blanks sort to bottom
    if (b === BLANK_RELATION) return -1;
    return a.localeCompare(b);
  });
  selectedRelations = new Set(uniqueRelations);

  renderDateList();
  renderClientList('');
  renderStatusList();
  renderRemarkTypeList();
  renderRelationList();
  updateCountBadges();

  dateBlock.style.display       = 'block';
  clientBlock.style.display     = 'block';
  statusBlock.style.display     = 'block';
  remarkTypeBlock.style.display = uniqueRemarkTypes.length ? 'block' : 'none';
  relationBlock.style.display   = uniqueRelations.length  ? 'block' : 'none';
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

/* ── Remark Type checkboxes ── */
function renderRemarkTypeList() {
  remarkTypeList.innerHTML = '';
  uniqueRemarkTypes.forEach(rt => {
    const item = makeCheckItem(rt, selectedRemarkTypes.has(rt), rt, (checked) => {
      checked ? selectedRemarkTypes.add(rt) : selectedRemarkTypes.delete(rt);
      updateCountBadges();
      refreshResults();
    });
    remarkTypeList.appendChild(item);
  });
}

document.getElementById('remarkTypeAll').addEventListener('click', () => {
  selectedRemarkTypes = new Set(uniqueRemarkTypes);
  syncCheckboxes(remarkTypeList, selectedRemarkTypes);
  updateCountBadges(); refreshResults();
});
document.getElementById('remarkTypeNone').addEventListener('click', () => {
  selectedRemarkTypes.clear();
  syncCheckboxes(remarkTypeList, selectedRemarkTypes);
  updateCountBadges(); refreshResults();
});

/* ── Relation checkboxes ── */
function renderRelationList() {
  relationList.innerHTML = '';
  uniqueRelations.forEach(rel => {
    const isBlank   = rel === BLANK_RELATION;
    const labelText = isBlank ? '(Blank)' : rel;
    const item = makeCheckItem(rel, selectedRelations.has(rel), labelText, (checked) => {
      checked ? selectedRelations.add(rel) : selectedRelations.delete(rel);
      updateCountBadges();
      refreshResults();
    });
    // Style the blank option distinctly
    if (isBlank) {
      const lbl = item.querySelector('.chk-lbl');
      lbl.style.fontStyle  = 'italic';
      lbl.style.color      = 'var(--muted2)';
    }
    relationList.appendChild(item);
  });
}

document.getElementById('relationAll').addEventListener('click', () => {
  selectedRelations = new Set(uniqueRelations);
  syncCheckboxes(relationList, selectedRelations);
  updateCountBadges(); refreshResults();
});
document.getElementById('relationNone').addEventListener('click', () => {
  selectedRelations.clear();
  syncCheckboxes(relationList, selectedRelations);
  updateCountBadges(); refreshResults();
});
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
  dateSelCount.textContent        = selectedDates.size;
  clientSelCount.textContent      = selectedClients.size;
  statusSelCount.textContent      = selectedStatuses.size;
  remarkTypeSelCount.textContent  = selectedRemarkTypes.size;
  relationSelCount.textContent    = selectedRelations.size;
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
  return allData.filter(r => {
    const rel = r.relation === '' ? BLANK_RELATION : r.relation;
    return (
      selectedDates.has(r.date) &&
      selectedClients.has(r.client) &&
      selectedStatuses.has(r.status) &&
      (uniqueRemarkTypes.length === 0 || selectedRemarkTypes.has(r.remarkType)) &&
      (uniqueRelations.length === 0   || selectedRelations.has(rel))
    );
  });
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

  // Positive: union of accounts where Relation === 'Debtor' (RPC) OR Call Status === 'CONNECTED'
  // Deduplicated by Account No. — each account counted once even if it matches both conditions
  const positiveAccounts = new Map();
  filtered.forEach(r => {
    const isDebtor    = r.relation.toLowerCase() === 'debtor';
    const isConnected = r.callStatus.toUpperCase() === 'CONNECTED';
    if ((isDebtor || isConnected) && !positiveAccounts.has(r.account)) {
      positiveAccounts.set(r.account, r);
    }
  });
  const positiveUnique  = [...positiveAccounts.values()];
  const totalPositive   = positiveUnique.length;
  const totalPositiveBal = positiveUnique.reduce((s, r) => s + r.balance, 0);

  // Dials: total row count with NO deduplication — every dial attempt counted
  const totalDials = filtered.length;

  animateNum(kpiAccounts, totalAccts, false);
  animateNum(kpiDials, totalDials, false);

  // Balance: display exact full number (no abbreviation)
  animateNumExact(kpiBalance, totalBal);
  animateNumExact(kpiPtpSum, ptpSum);
  animateNumExact(kpiClaimSum, claimSum);

  kpiPtpCount.textContent   = ptpRows.length.toLocaleString();
  kpiClaimCount.textContent = claimRows.length.toLocaleString();
  animateNum(kpiPositiveCount, totalPositive,    false);
  animateNumExact(kpiPositiveBal, totalPositiveBal);

  const allDatesSelected        = selectedDates.size === uniqueDates.length;
  const allClientsSelected      = selectedClients.size === uniqueClients.length;
  const allStatusesSelected     = selectedStatuses.size === uniqueStatuses.length;
  const allRemarkTypesSelected  = !uniqueRemarkTypes.length || selectedRemarkTypes.size === uniqueRemarkTypes.length;
  const allRelationsSelected    = !uniqueRelations.length   || selectedRelations.size   === uniqueRelations.length;
  kpiAccountsSub.textContent = `of ${allData.length} total rows`;
  kpiBalanceSub.textContent  = `${totalAccts} unique account${totalAccts !== 1 ? 's' : ''} summed`;

  renderFilterSummary(allDatesSelected, allClientsSelected, allStatusesSelected, allRemarkTypesSelected, allRelationsSelected);
  renderBreakdown(unique, filtered);
  tableNote.textContent = `${totalAccts} unique account${totalAccts!==1?'s':''} · ${formatCurrencyExact(totalBal)}`;
}

/* ═══════════════════════════════════════
   FILTER SUMMARY CHIPS
═══════════════════════════════════════ */
function renderFilterSummary(allDates, allClients, allStatuses, allRemarkTypes, allRelations) {
  filterSummary.innerHTML = '';

  if (allDates && allClients && allStatuses && allRemarkTypes && allRelations) {
    filterSummary.appendChild(chip_('All filters active', 'fchip fchip-all'));
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

  if (uniqueRemarkTypes.length) {
    if (allRemarkTypes) {
      filterSummary.appendChild(chip_('All remark types', 'fchip fchip-remark'));
    } else {
      [...selectedRemarkTypes].sort().slice(0, 5).forEach(rt => {
        filterSummary.appendChild(chip_(rt, 'fchip fchip-remark', 'Remark'));
      });
      if (selectedRemarkTypes.size > 5)
        filterSummary.appendChild(chip_(`+${selectedRemarkTypes.size - 5} more`, 'fchip fchip-remark'));
    }
  }

  if (uniqueRelations.length) {
    if (allRelations) {
      filterSummary.appendChild(chip_('All relations', 'fchip fchip-relation'));
    } else {
      [...selectedRelations].sort().slice(0, 5).forEach(rel => {
        const label = rel === BLANK_RELATION ? '(Blank)' : rel;
        filterSummary.appendChild(chip_(label, 'fchip fchip-relation', 'Relation'));
      });
      if (selectedRelations.size > 5)
        filterSummary.appendChild(chip_(`+${selectedRelations.size - 5} more`, 'fchip fchip-relation'));
    }
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
// Use a plain Map (not WeakMap) — WeakMap only accepts objects as keys and throws
// "Invalid value used as weak map key" if a null or primitive is passed, which can
// happen on GitHub Pages where DOM elements may not resolve at call time.
const _counters = new Map();

function animateNum(el, target, isCurrency) {
  if (!el) return; // null-guard: skip if element not found in DOM
  const prev = _counters.get(el) || 0;
  _counters.set(el, target);
  if (prev === target) return;

  const duration = 500;
  const start    = performance.now();

  function step(now) {
    const t  = Math.min((now - start) / duration, 1);
    const e  = 1 - Math.pow(1 - t, 3);
    const v  = prev + (target - prev) * e;
    el.textContent = isCurrency ? formatCurrencyExact(v) : Math.round(v).toLocaleString();
    if (t < 1) requestAnimationFrame(step);
    else el.textContent = isCurrency ? formatCurrencyExact(target) : target.toLocaleString();
  }
  requestAnimationFrame(step);
}

function animateNumExact(el, target) {
  if (!el) return; // null-guard
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
   DOWNLOAD POSITIVE (Debtor RPC + Connected)
═══════════════════════════════════════ */
btnDownloadPositive.addEventListener('click', downloadPositive);

function downloadPositive() {
  const filtered = getFilteredRows();

  // Union: Debtor (RPC) OR Connected — deduplicated by Account No.
  const positiveMap = new Map();
  filtered.forEach(r => {
    const isDebtor    = r.relation.toLowerCase() === 'debtor';
    const isConnected = r.callStatus.toUpperCase() === 'CONNECTED';
    if ((isDebtor || isConnected) && !positiveMap.has(r.account)) {
      positiveMap.set(r.account, r);
    }
  });
  const positiveUnique = [...positiveMap.values()];

  if (!positiveUnique.length) {
    alert('No Positive records (Debtor/Connected) found for the current filters.');
    return;
  }

  const wsData = [
    ['Date', 'Account No.', 'Client', 'Status', 'Relation', 'Call Status',
     'Balance', 'PTP Amount', 'Claim Paid Amount']
  ];
  positiveUnique.forEach(r => {
    wsData.push([r.date, r.account, r.client, r.status, r.relation,
                 r.callStatus, r.balance, r.ptpAmount, r.claimPaid]);
  });

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols'] = [
    { wch: 14 }, { wch: 18 }, { wch: 24 }, { wch: 16 },
    { wch: 12 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 20 }
  ];
  XLSX.utils.book_append_sheet(wb, ws, 'Positive');

  const datePart   = selectedDates.size === uniqueDates.length ? 'AllDates'
    : [...selectedDates].sort().join('_').replace(/[\s,]/g, '').slice(0, 40);
  const statusPart = selectedStatuses.size === uniqueStatuses.length ? 'AllStatuses'
    : [...selectedStatuses].sort().join('_').replace(/\s/g, '').slice(0, 30);
  const timestamp  = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Positive_${datePart}_${statusPart}_${timestamp}.xlsx`);
}

function resetAll() {
  allData = []; uniqueDates = []; uniqueClients = []; uniqueStatuses = [];
  uniqueRemarkTypes = []; uniqueRelations = [];
  selectedDates.clear(); selectedClients.clear(); selectedStatuses.clear();
  selectedRemarkTypes.clear(); selectedRelations.clear();

  fileInput.value = '';
  fileChips.innerHTML = '';
  fileLoadedRow.style.display   = 'none';
  dateBlock.style.display       = 'none';
  clientBlock.style.display     = 'none';
  statusBlock.style.display     = 'none';
  remarkTypeBlock.style.display = 'none';
  relationBlock.style.display   = 'none';
  topbarMeta.style.display      = 'none';
  emptyState.style.display      = 'flex';
  dashboard.style.display       = 'none';
  dateList.innerHTML = ''; clientList.innerHTML = ''; statusList.innerHTML = '';
  remarkTypeList.innerHTML = ''; relationList.innerHTML = '';
  clientSearch.value = '';
  dateSelCount.textContent = clientSelCount.textContent = statusSelCount.textContent =
    remarkTypeSelCount.textContent = relationSelCount.textContent = '0';
}