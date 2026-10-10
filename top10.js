/*
    Top 10 Accountables page logic
    Reads parsed occurrence rows saved by app.js in this browser session.
*/
(() => {
  "use strict";
  const SESSION_KEY = "occurrenceTrackerDataV1";
  const $ = id => document.getElementById(id);
  const clean = value => value == null ? "" : String(value).trim();
  const norm = value => clean(value).toLocaleLowerCase().replace(/\s+/g, " ");
  const num = value => {
    const parsed = Number(clean(value).replace(/[$,]/g, ""));
    return Number.isFinite(parsed) ? parsed : 0;
  };
  const fmt = value => num(value).toLocaleString(undefined, { maximumFractionDigits: 2 });

  function buildEmployees(rows) {
    const employees = new Map();
    rows.forEach(row => {
      const employeeNumber = clean(row.employeeNumber).replace(/\.0$/, "");
      const key = employeeNumber ? `id:${employeeNumber}` : `name:${norm(row.lastName)}|${norm(row.firstName)}|${norm(row.dcKey)}`;
      if (!employees.has(key)) {
        employees.set(key, {
          key,
          employeeNumber,
          lastName: clean(row.lastName),
          firstName: clean(row.firstName),
          dcKey: clean(row.dcKey),
          rows: [],
          totalHours: 0
        });
      }
      const employee = employees.get(key);
      employee.rows.push(row);
      employee.totalHours += num(row.hours);
      if (!employee.dcKey && row.dcKey) employee.dcKey = clean(row.dcKey);
      if (!employee.lastName && row.lastName) employee.lastName = clean(row.lastName);
      if (!employee.firstName && row.firstName) employee.firstName = clean(row.firstName);
    });
    return [...employees.values()].map(employee => ({
      ...employee,
      displayName: [employee.lastName, employee.firstName].filter(Boolean).join(", "),
      accountableCount: employee.rows.length
    })).sort((a, b) => b.accountableCount - a.accountableCount || a.displayName.localeCompare(b.displayName));
  }

  function render() {
    let saved;
    try { saved = JSON.parse(sessionStorage.getItem(SESSION_KEY) || "null"); }
    catch (error) { saved = null; }

    if (!saved || !Array.isArray(saved.rows) || !saved.rows.length) return;

    $("fileStatus").textContent = `${saved.fileName || "Loaded spreadsheet"} · ${saved.rows.length.toLocaleString()} rows`;
    const topEmployees = buildEmployees(saved.rows).slice(0, 10);
    const body = $("leaderboardBody");
    body.replaceChildren();

    topEmployees.forEach((employee, index) => {
      const row = document.createElement("tr");
      if (index < 3) row.className = "top-rank";
      const values = [
        String(index + 1),
        employee.displayName || "Name unavailable",
        employee.dcKey || "—",
        employee.employeeNumber || "—",
        employee.accountableCount.toLocaleString(),
        fmt(employee.totalHours)
      ];
      values.forEach(value => {
        const cell = document.createElement("td");
        cell.textContent = value;
        row.append(cell);
      });

      const actionCell = document.createElement("td");
      const button = document.createElement("button");
      button.type = "button";
      button.className = "view-tm";
      button.textContent = "View TM";
      button.addEventListener("click", () => {
        const query = employee.employeeNumber || employee.displayName;
        window.location.href = `index.html?employeeNumber=${encodeURIComponent(query)}`;
      });
      actionCell.append(button);
      row.append(actionCell);
      body.append(row);
    });

    $("leaderboardEmpty").classList.add("hidden");
    $("leaderboardWrap").classList.remove("hidden");
  }

  render();
})();
