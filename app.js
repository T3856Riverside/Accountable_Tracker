(() => {
  "use strict";

  const $ = (id) => document.getElementById(id);
  // Keep the active spreadsheet and selected employee in one shared state object.
  const state = { rows: [], employees: [], selected: null, fileName: "", paycodeFilter: "" };
  const SESSION_KEY = "occurrenceTrackerDataV1";

  const els = {
    fileInput: $("fileInput"), fileStatus: $("fileStatus"), search: $("searchInput"),
    clear: $("clearSearch"), status: $("statusMessage"), empty: $("emptyState"),
    details: $("employeeDetails"), name: $("employeeName"), dcKey: $("dcKey"),
    employeeNumber: $("employeeNumber"), count: $("occurrenceCount"), hours: $("totalHours"),
    historyEmpty: $("historyEmpty"), tableWrap: $("tableWrap"), historyBody: $("historyBody"),
    historyCount: $("historyCount"), matchCard: $("matchCard"), matchList: $("matchList"),
    paycodeFilter: $("paycodeFilter"), exportPdf: $("exportPdf"), reset: $("resetData")
  };

  const clean = (value) => value == null ? "" : String(value).trim();
  const norm = (value) => clean(value).toLocaleLowerCase().replace(/\s+/g, " ");
  const idNorm = (value) => clean(value).replace(/\.0$/, "").replace(/\s+/g, "");
  const firstValue = (obj, names) => {
    const keys = Object.keys(obj);
    for (const wanted of names) {
      const found = keys.find(k => norm(k) === norm(wanted));
      if (found !== undefined) return obj[found];
    }
    return "";
  };
  const num = (v) => {
    if (typeof v === "number" && Number.isFinite(v)) return v;
    const s = clean(v).replace(/[$,]/g, "");
    if (!s || !Number.isFinite(Number(s))) return 0;
    return Number(s);
  };
  const fmtHours = (v) => num(v).toLocaleString(undefined, { maximumFractionDigits: 2 });
  const dateText = (v) => {
    if (v instanceof Date && !Number.isNaN(v.getTime())) return v.toLocaleDateString();
    if (typeof v === "number" && window.XLSX?.SSF) {
      const parsed = XLSX.SSF.parse_date_code(v);
      if (parsed) return `${String(parsed.m).padStart(2,"0")}/${String(parsed.d).padStart(2,"0")}/${parsed.y}`;
    }
    const s = clean(v);
    if (!s) return "";
    const d = new Date(s);
    return !Number.isNaN(d.getTime()) && !/^\d{4,}$/.test(s) ? d.toLocaleDateString() : s;
  };
  const dayText = (row) => {
    const explicit = clean(row.day);
    if (explicit) return explicit;
    const v = row.date;
    const d = v instanceof Date ? v : new Date(v);
    return v && !Number.isNaN(d.getTime()) ? d.toLocaleDateString(undefined, { weekday: "long" }).toUpperCase() : "";
  };

  function normalizeRow(raw) {
    return {
      region: firstValue(raw, ["REGION"]),
      location: firstValue(raw, ["LOCATION"]),
      dcKey: clean(firstValue(raw, ["DC KEY", "DC KEY "])).toUpperCase(),
      lastName: clean(firstValue(raw, ["LAST NAME", "LASTNAME"])),
      firstName: clean(firstValue(raw, ["FIRST NAME", "FIRSTNAME"])),
      employeeNumber: idNorm(firstValue(raw, ["EMPLOYEE NUMBER", "EMPLOYEE #", "EMPLOYEE ID"])),
      date: firstValue(raw, ["DATE OF OCC", "DATE OF OCCURRENCE", "OCCURRENCE DATE"]),
      day: clean(firstValue(raw, ["DAY"])),
      hours: num(firstValue(raw, ["HOURS", "OCCURRENCE HOURS"])),
      paycode: clean(firstValue(raw, ["PAY CODE", "PAYCODE", "PAY_CODE"])),
      type: clean(firstValue(raw, ["TYPE", "OCCURRENCE TYPE"])),
      raw
    };
  }

  function rowIsUseful(r) {
    return !!(r.employeeNumber || r.lastName || r.firstName || r.dcKey) &&
      (r.employeeNumber || r.lastName || r.firstName);
  }

  function groupEmployees(rows) {
    const map = new Map();
    for (const row of rows) {
      const key = row.employeeNumber
        ? `id:${row.employeeNumber}`
        : `name:${norm(row.lastName)}|${norm(row.firstName)}|${norm(row.dcKey)}`;
      if (!map.has(key)) {
        map.set(key, {
          key, employeeNumber: row.employeeNumber, lastName: row.lastName, firstName: row.firstName,
          dcKey: row.dcKey, rows: []
        });
      }
      const employee = map.get(key);
      employee.rows.push(row);
      if (!employee.dcKey && row.dcKey) employee.dcKey = row.dcKey;
      if (!employee.lastName && row.lastName) employee.lastName = row.lastName;
      if (!employee.firstName && row.firstName) employee.firstName = row.firstName;
    }
    return [...map.values()].map(e => ({
      ...e,
      displayName: [e.lastName, e.firstName].filter(Boolean).join(", "),
      occurrenceCount: e.rows.length,
      totalHours: e.rows.reduce((sum, r) => sum + r.hours, 0)
    })).sort((a,b) => a.displayName.localeCompare(b.displayName));
  }

  async function loadFile(file) {
    if (!file) return;
    if (!window.XLSX) {
      setStatus("The Excel-reading library did not load. Check your internet connection and refresh the page.");
      return;
    }
    try {
      setStatus("Reading spreadsheet in this browser…");
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array", cellDates: true });
      const sheetName = workbook.SheetNames.find(n => norm(n) === "input data") || workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) throw new Error("No readable worksheet was found.");
      const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });
      if (!rawRows.length) throw new Error(`The "${sheetName}" sheet has no data rows.`);
      const normalized = rawRows.map(normalizeRow).filter(rowIsUseful);
      if (!normalized.length) throw new Error("I couldn't find employee rows. Check that the sheet has headers such as DC KEY, LAST NAME, FIRST NAME, EMPLOYEE NUMBER, DATE OF OCC, HOURS, and TYPE.");
      state.rows = normalized;
      state.employees = groupEmployees(normalized);
      state.selected = null;
      state.fileName = file.name;
      // Share parsed rows with the separate Top 10 page for this browser session.
      saveSessionData();
      els.fileStatus.textContent = `${file.name} · ${normalized.length.toLocaleString()} rows`;
      els.search.disabled = false;
      els.reset.disabled = false;
      els.empty.classList.remove("hidden");
      els.details.classList.add("hidden");
      clearHistory("Search for an employee to see their occurrence history.");
      resetPaycodeFilter();
      els.exportPdf.disabled = true;
      els.matchCard.classList.add("hidden");
      els.search.value = "";
      setStatus(`Loaded ${normalized.length.toLocaleString()} occurrence rows from "${sheetName}". Search by name, DC Key, or employee number.`);
      els.search.focus();
    } catch (error) {
      console.error(error);
      setStatus(error.message || "Couldn't read that file. Try an .xlsx, .xlsm, .xls, or .csv file.");
    }
  }

  function setStatus(message) { els.status.textContent = message; }

  function clearHistory(message) {
    els.historyBody.replaceChildren();
    els.historyCount.textContent = "";
    els.tableWrap.classList.add("hidden");
    els.historyEmpty.textContent = message;
    els.historyEmpty.classList.remove("hidden");
  }

  function searchEmployees(query) {
    const q = norm(query);
    if (!q) return [];
    const qId = idNorm(query).toLowerCase();
    return state.employees.filter(e => {
      const haystack = norm(`${e.lastName} ${e.firstName} ${e.firstName} ${e.lastName} ${e.displayName} ${e.dcKey} ${e.employeeNumber}`);
      return haystack.includes(q) || (qId && e.employeeNumber.toLowerCase().includes(qId));
    });
  }

  function runSearch() {
    const query = els.search.value.trim();
    els.matchCard.classList.add("hidden");
    if (!state.rows.length) { setStatus("Upload your Input Data spreadsheet first."); return; }
    if (!query) {
      state.selected = null;
      resetPaycodeFilter();
      els.exportPdf.disabled = true;
      els.details.classList.add("hidden");
      els.empty.classList.remove("hidden");
      clearHistory("Search for an employee to see their occurrence history.");
      setStatus(`Loaded ${state.rows.length.toLocaleString()} occurrence rows.`);
      return;
    }
    const matches = searchEmployees(query);
    if (!matches.length) {
      state.selected = null;
      resetPaycodeFilter();
      els.exportPdf.disabled = true;
      els.details.classList.add("hidden");
      els.empty.classList.add("hidden");
      clearHistory("No occurrence history found for this search.");
      setStatus(`No employee matched "${query}". Try their full name or employee number.`);
      return;
    }
    if (matches.length > 1) {
      els.matchCard.classList.remove("hidden");
      els.matchList.replaceChildren();
      for (const employee of matches.slice(0, 30)) {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "match-option";
        const label = document.createElement("span");
        label.textContent = employee.displayName || "(Name not available)";
        const meta = document.createElement("small");
        meta.textContent = `DC Key: ${employee.dcKey || "—"} · Employee #: ${employee.employeeNumber || "—"} · ${employee.occurrenceCount} occurrence(s)`;
        button.append(label, meta);
        button.addEventListener("click", () => selectEmployee(employee));
        els.matchList.append(button);
      }
      setStatus(`${matches.length} employees matched. Choose the correct employee below.`);
      if (matches.length === 1) selectEmployee(matches[0]);
      return;
    }
    selectEmployee(matches[0]);
  }

  function selectEmployee(employee) {
    state.selected = employee;
    resetPaycodeFilter();
    els.exportPdf.disabled = false;
    els.matchCard.classList.add("hidden");
    els.empty.classList.add("hidden");
    els.details.classList.remove("hidden");
    els.name.textContent = employee.displayName || "—";
    els.dcKey.textContent = employee.dcKey || "—";
    els.employeeNumber.textContent = employee.employeeNumber || "—";
    els.count.textContent = employee.occurrenceCount.toLocaleString();
    els.hours.textContent = fmtHours(employee.totalHours);
    populatePaycodeFilter(employee.rows);
    renderHistory(getVisibleRows());
    setStatus(`Showing ${employee.occurrenceCount} occurrence row(s) for ${employee.displayName || employee.employeeNumber}.`);
  }

  function renderHistory(rows) {
    els.historyBody.replaceChildren();
    const sorted = [...rows].sort((a,b) => {
      const da = a.date instanceof Date ? a.date.getTime() : new Date(a.date).getTime();
      const db = b.date instanceof Date ? b.date.getTime() : new Date(b.date).getTime();
      return (Number.isNaN(da) ? 0 : da) - (Number.isNaN(db) ? 0 : db);
    });
    const frag = document.createDocumentFragment();
    for (const row of sorted) {
      const tr = document.createElement("tr");
      const values = [row.dcKey, row.lastName, row.firstName, row.employeeNumber, dateText(row.date), dayText(row), fmtHours(row.hours), row.paycode, row.type];
      values.forEach((value, i) => {
        const td = document.createElement("td");
        td.textContent = value || "—";
        if (i === 6) td.className = "num";
        tr.append(td);
      });
      frag.append(tr);
    }
    els.historyBody.append(frag);
    els.historyEmpty.classList.add("hidden");
    els.tableWrap.classList.remove("hidden");
    els.historyCount.textContent = `(${rows.length.toLocaleString()} rows)`;
  }

  // Store only the parsed spreadsheet rows in session storage so the Top 10 page
  // can use the same upload without asking the user to select the file again.
  function saveSessionData() {
    try {
      sessionStorage.setItem(SESSION_KEY, JSON.stringify({ fileName: state.fileName, rows: state.rows }));
    } catch (error) {
      console.warn("Could not save spreadsheet data for page navigation.", error);
    }
  }

  function restoreSessionData() {
    try {
      const saved = JSON.parse(sessionStorage.getItem(SESSION_KEY) || "null");
      if (!saved || !Array.isArray(saved.rows) || !saved.rows.length) return false;
      state.rows = saved.rows.map(row => ({ ...row, date: row.date ? new Date(row.date) : "" }));
      state.employees = groupEmployees(state.rows);
      state.fileName = saved.fileName || "Previously loaded spreadsheet";
      els.fileStatus.textContent = `${state.fileName} · ${state.rows.length.toLocaleString()} rows`;
      els.search.disabled = false;
      els.reset.disabled = false;
      setStatus(`Restored ${state.rows.length.toLocaleString()} occurrence rows from this browser session.`);
      return true;
    } catch (error) {
      console.warn("Could not restore the spreadsheet session.", error);
      return false;
    }
  }

  // Populate the PAY CODE dropdown using the selected employee's occurrence records.
  function populatePaycodeFilter(rows) {
    const paycodes = [...new Set(rows.map(row => clean(row.paycode)).filter(Boolean))]
      .sort((a, b) => a.localeCompare(b));
    els.paycodeFilter.replaceChildren();
    const allOption = document.createElement("option");
    allOption.value = "";
    allOption.textContent = "All pay codes";
    els.paycodeFilter.append(allOption);
    paycodes.forEach(paycode => {
      const option = document.createElement("option");
      option.value = paycode;
      option.textContent = paycode;
      els.paycodeFilter.append(option);
    });
    els.paycodeFilter.disabled = !state.selected;
    els.paycodeFilter.value = "";
    state.paycodeFilter = "";
  }

  function resetPaycodeFilter() {
    if (!els.paycodeFilter) return;
    els.paycodeFilter.replaceChildren();
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "All pay codes";
    els.paycodeFilter.append(option);
    els.paycodeFilter.value = "";
    els.paycodeFilter.disabled = true;
    state.paycodeFilter = "";
  }

  function getVisibleRows() {
    if (!state.selected) return [];
    return state.paycodeFilter
      ? state.selected.rows.filter(row => clean(row.paycode) === state.paycodeFilter)
      : state.selected.rows;
  }

  function applyPaycodeFilter() {
    if (!state.selected) return;
    state.paycodeFilter = els.paycodeFilter.value;
    const visibleRows = getVisibleRows();
    renderHistory(visibleRows);
    const filterText = state.paycodeFilter ? ` for PAY CODE “${state.paycodeFilter}”` : " across all pay codes";
    setStatus(`Showing ${visibleRows.length} of ${state.selected.occurrenceCount} occurrence row(s)${filterText}.`);
  }

  // Create a downloadable PDF report for the selected team member.
  function exportEmployeePdf() {
    if (!state.selected) return;
    if (!window.jspdf?.jsPDF) {
      setStatus("The PDF library did not load. Check your internet connection and try again.");
      return;
    }
    const employee = state.selected;
    const visibleRows = getVisibleRows();
    const doc = new window.jspdf.jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
    const pageWidth = doc.internal.pageSize.getWidth();
    doc.setFillColor(204, 0, 0);
    doc.rect(0, 0, pageWidth, 66, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(19);
    doc.text("EMPLOYEE OCCURRENCE REPORT", 36, 29);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.text(`Generated ${new Date().toLocaleString()}`, 36, 48);

    doc.setTextColor(29, 41, 56);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(13);
    doc.text(employee.displayName || "Team Member", 36, 91);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    doc.text(`DC Key: ${employee.dcKey || "—"}`, 36, 111);
    doc.text(`Employee Number: ${employee.employeeNumber || "—"}`, 250, 111);
    doc.text(`Total Accountables: ${employee.occurrenceCount}`, 36, 130);
    doc.text(`Total Hours: ${fmtHours(employee.totalHours)}`, 250, 130);
    doc.text(`Report filter: ${state.paycodeFilter || "All pay codes"}`, 470, 130);
    doc.text(`Rows in report: ${visibleRows.length}`, 690, 130);

    const body = [...visibleRows].sort((a, b) => {
      const da = a.date instanceof Date ? a.date.getTime() : new Date(a.date).getTime();
      const db = b.date instanceof Date ? b.date.getTime() : new Date(b.date).getTime();
      return (Number.isNaN(da) ? 0 : da) - (Number.isNaN(db) ? 0 : db);
    }).map(row => [row.dcKey, row.lastName, row.firstName, row.employeeNumber, dateText(row.date), dayText(row), fmtHours(row.hours), row.paycode || "—", row.type || "—"]);

    if (typeof doc.autoTable === "function") {
      doc.autoTable({
        startY: 148,
        head: [["DC Key", "Last Name", "First Name", "Employee Number", "Date of Occ", "Day", "Hours", "Pay Code", "Type"]],
        body,
        theme: "grid",
        styles: { fontSize: 8, cellPadding: 5, textColor: [29, 41, 56] },
        headStyles: { fillColor: [204, 0, 0], textColor: 255, fontStyle: "bold" },
        alternateRowStyles: { fillColor: [255, 245, 246] },
        margin: { left: 36, right: 36, bottom: 35 },
        didDrawPage: () => {
          const pageHeight = doc.internal.pageSize.getHeight();
          doc.setFontSize(8);
          doc.setTextColor(104, 115, 131);
          doc.text("Employee Occurrence Tracker · Generated in browser", 36, pageHeight - 16);
          doc.text(`Page ${doc.internal.getNumberOfPages()}`, pageWidth - 70, pageHeight - 16);
        }
      });
    } else {
      setStatus("The PDF table component did not load. Check your internet connection and try again.");
      return;
    }

    const safeName = (employee.displayName || employee.employeeNumber || "team-member")
      .replace(/[^a-z0-9]+/gi, "_").replace(/^_|_$/g, "");
    doc.save(`TM_Occurrence_Report_${safeName || "team-member"}.pdf`);
  }

  function reset() {
    state.rows = []; state.employees = []; state.selected = null; state.fileName = ""; state.paycodeFilter = "";
    sessionStorage.removeItem(SESSION_KEY);
    els.fileInput.value = ""; els.fileStatus.textContent = "No spreadsheet loaded";
    els.search.value = ""; els.search.disabled = true; els.reset.disabled = true;
    els.details.classList.add("hidden"); els.empty.classList.remove("hidden");
    resetPaycodeFilter();
    els.exportPdf.disabled = true;
    els.matchCard.classList.add("hidden");
    clearHistory("Search for an employee to see their occurrence history.");
    setStatus("Upload your Input Data spreadsheet to get started.");
  }

  els.fileInput.addEventListener("change", event => loadFile(event.target.files?.[0]));
  els.search.addEventListener("input", runSearch);
  els.clear.addEventListener("click", () => { els.search.value = ""; runSearch(); els.search.focus(); });
  els.reset.addEventListener("click", reset);
  els.paycodeFilter.addEventListener("change", applyPaycodeFilter);
  els.exportPdf.addEventListener("click", exportEmployeePdf);

  // Restore the active spreadsheet when returning from the Top 10 page.
  const restored = restoreSessionData();
  if (restored) {
    const requestedEmployee = new URLSearchParams(window.location.search).get("employeeNumber");
    if (requestedEmployee) {
      els.search.value = requestedEmployee;
      runSearch();
    }
  }
})(); 
