/* ===================================================
   Login Activity Analyzer — main.js
   Security Bank UI · Multi-file merge · Time-sorted
=================================================== */

(function () {
  "use strict";

  /* ---- DOM ---- */
  const uploadZone     = document.getElementById("uploadZone");
  const fileInput      = document.getElementById("fileInput");
  const fileListEl     = document.getElementById("fileList");
  const fileListLabel  = document.getElementById("fileListLabel");
  const statusBar      = document.getElementById("statusBar");
  const statusSpinner  = document.getElementById("statusSpinner");
  const statusText     = document.getElementById("statusText");
  const resultsEmpty   = document.getElementById("resultsEmpty");
  const tableScroll    = document.getElementById("tableScroll");
  const resultsMeta    = document.getElementById("resultsMeta");
  const tableBody      = document.getElementById("tableBody");
  const errorSection   = document.getElementById("errorSection");
  const errorMsg       = document.getElementById("errorMsg");
  const retryBtn       = document.getElementById("retryBtn");
  const navDate        = document.getElementById("navDate");

  const TARGET_SHEET  = "Login Logout Activity";
  const COL_COLLECTOR = "Collector";
  const COL_DATETIME  = "Connect/Disconnect Date Time";

  /* ---- Date in nav ---- */
  navDate.textContent = new Date().toLocaleDateString("en-US", {
    weekday: "short", year: "numeric", month: "short", day: "numeric"
  });

  /* ---- Drag & Drop ---- */
  uploadZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    uploadZone.classList.add("drag-over");
  });
  uploadZone.addEventListener("dragleave", () => uploadZone.classList.remove("drag-over"));
  uploadZone.addEventListener("drop", (e) => {
    e.preventDefault();
    uploadZone.classList.remove("drag-over");
    if (e.dataTransfer.files.length > 0) handleFiles(e.dataTransfer.files);
  });
  fileInput.addEventListener("change", () => {
    if (fileInput.files.length > 0) handleFiles(fileInput.files);
  });
  retryBtn.addEventListener("click", resetUI);

  /* ===================================================  HANDLE FILES  */
  function handleFiles(fileList) {
    const files = Array.from(fileList);
    const invalid = files.filter(
      (f) => !["xlsx", "xls", "csv"].includes(f.name.split(".").pop().toLowerCase())
    );
    if (invalid.length) {
      showError(`Unsupported file type: <strong>${invalid.map(f => f.name).join(", ")}</strong>.<br>Only <strong>.xlsx</strong> and <strong>.csv</strong> are supported.`);
      return;
    }

    resetResults();
    renderFileList(files);
    showStatus(`Reading ${files.length} file${files.length !== 1 ? "s" : ""}…`);

    Promise.all(files.map(readFileAsArrayBuffer))
      .then((buffers) => {
        showStatus("Parsing and merging data…");
        return mergeAllFiles(files, buffers);
      })
      .then(({ allRows, totalRows, fileCount }) => {
        showStatus("Analysing merged dataset…");
        processMergedRows(allRows, totalRows, fileCount);
      })
      .catch((err) => showError("Error processing files: " + err.message));
  }

  function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload  = (e) => resolve(e.target.result);
      r.onerror = ()  => reject(new Error(`Could not read "${file.name}"`));
      r.readAsArrayBuffer(file);
    });
  }

  /* ===================================================  MERGE  */
  function mergeAllFiles(files, buffers) {
    const allRows = [], warnings = [];
    files.forEach((file, idx) => {
      const ext = file.name.split(".").pop().toLowerCase();
      try {
        if (ext === "csv") {
          allRows.push(...extractCSVRows(buffers[idx]));
        } else {
          const { rows, warning } = extractXLSXRows(buffers[idx], file.name);
          allRows.push(...rows);
          if (warning) warnings.push(warning);
        }
      } catch (err) {
        warnings.push(`"${file.name}": ${err.message}`);
      }
    });
    if (warnings.length && !allRows.length) throw new Error(warnings.join(" | "));
    return { allRows, totalRows: allRows.length, fileCount: files.length };
  }

  function extractXLSXRows(arrayBuffer, fileName) {
    const wb = XLSX.read(arrayBuffer, { type: "array" });
    const sn = wb.SheetNames.find(n => n.trim().toLowerCase() === TARGET_SHEET.toLowerCase());
    if (!sn) return { rows: [], warning: `"${fileName}": sheet "${TARGET_SHEET}" not found (available: ${wb.SheetNames.join(", ")})` };
    return { rows: XLSX.utils.sheet_to_json(wb.Sheets[sn], { defval: "" }), warning: null };
  }

  function extractCSVRows(arrayBuffer) {
    const text = new TextDecoder("utf-8").decode(arrayBuffer);
    const wb   = XLSX.read(text, { type: "string" });
    return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
  }

  /* ===================================================  PROCESS  */
  function processMergedRows(allRows, totalRows, fileCount) {
    if (!allRows.length) {
      showError(`No data found across ${fileCount} file${fileCount !== 1 ? "s" : ""}.<br>Make sure each file contains a sheet named <strong>"${TARGET_SHEET}"</strong>.`);
      return;
    }

    const headers         = Object.keys(allRows[0]);
    const collectorHeader = findHeader(headers, COL_COLLECTOR);
    const datetimeHeader  = findHeader(headers, COL_DATETIME);

    if (!collectorHeader) {
      showError(`Column <strong>"${COL_COLLECTOR}"</strong> not found.<br>Available: ${headers.map(h => `<em>${h}</em>`).join(", ")}`);
      return;
    }
    if (!datetimeHeader) {
      showError(`Column <strong>"${COL_DATETIME}"</strong> not found.<br>Available: ${headers.map(h => `<em>${h}</em>`).join(", ")}`);
      return;
    }

    const earliestMap = new Map();
    allRows.forEach((row) => {
      const collector = String(row[collectorHeader] || "").trim();
      const rawDT     = String(row[datetimeHeader]  || "").trim();
      if (!collector || !rawDT) return;
      const parsed = parseDateTime(rawDT);
      if (!parsed) return;
      if (!earliestMap.has(collector) || parsed.dateObj < earliestMap.get(collector).dateObj) {
        earliestMap.set(collector, parsed);
      }
    });

    if (!earliestMap.size) {
      showError("No valid rows found. Check that times look like <em>11-03-2026 8:13:28 am</em>.");
      return;
    }

    /* Sort by time ascending */
    const results = Array.from(earliestMap.entries())
      .sort((a, b) => a[1].dateObj - b[1].dateObj);

    renderTable(results, totalRows, fileCount);
  }

  /* ===================================================  PARSER  */
  function parseDateTime(raw) {
    if (!raw) return null;
    if (raw instanceof Date) return { dateObj: raw, time24: dateToTime24(raw) };

    const str = String(raw).trim();

    /* DD-MM-YYYY H:MM:SS am/pm */
    const m1 = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})\s*(am|pm)$/i);
    if (m1) {
      let [, dd, mo, yyyy, hh, mm, ss, ap] = m1;
      hh = parseInt(hh);
      if (ap.toLowerCase() === "am") { if (hh === 12) hh = 0; }
      else { if (hh !== 12) hh += 12; }
      const dateObj = new Date(+yyyy, +mo - 1, +dd, hh, +mm, +ss);
      return { dateObj, time24: pad2(hh) + ":" + pad2(mm) + ":" + pad2(ss) };
    }

    /* DD-MM-YYYY H:MM:SS */
    const m2 = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})$/);
    if (m2) {
      let [, dd, mo, yyyy, hh, mm, ss] = m2;
      const dateObj = new Date(+yyyy, +mo - 1, +dd, +hh, +mm, +ss);
      return { dateObj, time24: pad2(hh) + ":" + pad2(mm) + ":" + pad2(ss) };
    }

    const d = new Date(str);
    if (!isNaN(d.getTime())) return { dateObj: d, time24: dateToTime24(d) };
    return null;
  }

  function dateToTime24(d) { return pad2(d.getHours()) + ":" + pad2(d.getMinutes()) + ":" + pad2(d.getSeconds()); }
  function pad2(n) { return String(n).padStart(2, "0"); }
  function findHeader(headers, target) {
    return headers.find(h => h.trim().toLowerCase() === target.trim().toLowerCase()) || null;
  }

  /* ===================================================  RENDER FILE PILLS  */
  function renderFileList(files) {
    fileListEl.innerHTML = "";
    if (!files.length) return;
    fileListLabel.style.display = "block";
    files.forEach((f) => {
      const ext  = f.name.split(".").pop().toLowerCase();
      const pill = document.createElement("div");
      pill.className = "file-pill";
      pill.innerHTML = `<span class="file-pill-icon">${ext === "csv" ? "📄" : "📊"}</span><span class="file-pill-name">${escapeHTML(f.name)}</span>`;
      fileListEl.appendChild(pill);
    });
  }

  /* ===================================================  RENDER TABLE  */
  function renderTable(results, totalRows, fileCount) {
    tableBody.innerHTML = "";
    results.forEach(([collector, { time24 }]) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${escapeHTML(collector)}</td><td>${escapeHTML(time24)}</td>`;
      tableBody.appendChild(tr);
    });

    resultsMeta.textContent =
      `${results.length} collector${results.length !== 1 ? "s" : ""} · ${totalRows} records · ${fileCount} file${fileCount !== 1 ? "s" : ""}`;

    resultsEmpty.style.display  = "none";
    tableScroll.style.display   = "block";

    updateStatusDone(`Done — ${results.length} collector${results.length !== 1 ? "s" : ""} from ${fileCount} file${fileCount !== 1 ? "s" : ""}.`);
  }

  /* ===================================================  UI HELPERS  */
  function showStatus(msg) {
    statusBar.className    = "status-bar";
    statusSpinner.className = "status-spinner";
    statusText.textContent = msg;
    statusBar.style.display    = "flex";
    errorSection.style.display = "none";
  }

  function updateStatusDone(msg) {
    statusBar.className    = "status-bar done";
    statusSpinner.className = "status-spinner done";
    statusText.textContent = msg;
  }

  function showError(html) {
    statusBar.style.display    = "none";
    tableScroll.style.display  = "none";
    resultsEmpty.style.display = "flex";
    errorMsg.innerHTML = html;
    errorSection.style.display = "block";
  }

  function resetResults() {
    tableScroll.style.display  = "none";
    errorSection.style.display = "none";
    tableBody.innerHTML = "";
    resultsMeta.textContent = "";
  }

  function resetUI() {
    resetResults();
    statusBar.style.display    = "none";
    resultsEmpty.style.display = "flex";
    fileListLabel.style.display = "none";
    fileListEl.innerHTML = "";
    fileInput.value = "";
  }

  function escapeHTML(str) {
    return String(str)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;")
      .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
})();
