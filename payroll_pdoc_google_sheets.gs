/**
 * Raman Financial Services - CRA PDOC helper for Google Sheets.
 *
 * What this does:
 * - Adds a "Payroll PDOC" menu.
 * - Creates a "PDOC Results" tab.
 * - Lets you paste text copied from a CRA PDOC result PDF.
 * - Extracts CPP, EI, federal tax, provincial tax, total deductions and net pay.
 * - Updates the matching payroll row by Employee ID and Pay Date.
 *
 * Suggested payroll tab headers:
 * employee_id, pay_date, gross, cpp, cpp2, ei, tax_fed, tax_prov,
 * total_deductions, net, status
 */

const PDOC_RESULTS_SHEET = "PDOC Results";
const DEFAULT_PAYROLL_SHEET = "Payroll";

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Payroll PDOC")
    .addItem("Setup PDOC sheets", "setupPdocSheets")
    .addItem("Paste PDOC result text", "showPdocImportDialog")
    .addItem("Apply selected PDOC row to Payroll", "applySelectedPdocToPayroll")
    .addToUi();
}

function setupPdocSheets() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(PDOC_RESULTS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PDOC_RESULTS_SHEET);
  }
  const headers = [
    "employee_id",
    "employee_name",
    "pay_date",
    "frequency",
    "gross",
    "cpp",
    "cpp2",
    "ei",
    "tax_fed",
    "tax_prov",
    "total_deductions",
    "net",
    "source",
    "imported_at"
  ];
  ensureHeaders_(sheet, headers);
  sheet.setFrozenRows(1);
  SpreadsheetApp.getUi().alert("PDOC Results tab is ready.");
}

function showPdocImportDialog() {
  setupPdocSheets();
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:Arial,sans-serif;padding:12px;">
      <h3 style="margin-top:0;">Paste CRA PDOC Result Text</h3>
      <p>Open the CRA PDOC PDF, select all text, copy it, then paste it below.</p>
      <label>Employee ID</label><br>
      <input id="employeeId" style="width:100%;box-sizing:border-box;margin:4px 0 10px;"><br>
      <label>Employee name</label><br>
      <input id="employeeName" style="width:100%;box-sizing:border-box;margin:4px 0 10px;"><br>
      <label>PDOC text</label><br>
      <textarea id="pdocText" style="width:100%;height:240px;box-sizing:border-box;margin-top:4px;"></textarea>
      <div style="margin-top:12px;">
        <button onclick="submitPdoc()" style="background:#1268df;color:#fff;border:0;padding:8px 14px;border-radius:4px;">Import PDOC</button>
        <button onclick="google.script.host.close()" style="margin-left:8px;padding:8px 14px;">Cancel</button>
      </div>
      <p id="status" style="color:#42526a;"></p>
      <script>
        function submitPdoc() {
          document.getElementById('status').textContent = 'Importing...';
          google.script.run
            .withSuccessHandler(function(message) {
              document.getElementById('status').textContent = message;
            })
            .withFailureHandler(function(error) {
              document.getElementById('status').textContent = error.message || String(error);
            })
            .importPdocTextToSheet(
              document.getElementById('pdocText').value,
              document.getElementById('employeeId').value,
              document.getElementById('employeeName').value
            );
        }
      </script>
    </div>
  `).setWidth(560).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, "CRA PDOC Import");
}

function importPdocTextToSheet(pdocText, employeeId, employeeName) {
  const parsed = parsePdocText_(pdocText || "");
  parsed.employee_id = employeeId || "";
  parsed.employee_name = employeeName || parsed.employee_name || "";
  parsed.source = "CRA PDOC";
  parsed.imported_at = new Date();

  const sheet = SpreadsheetApp.getActive().getSheetByName(PDOC_RESULTS_SHEET);
  const headers = getHeaders_(sheet);
  const row = headers.map(header => parsed[header] === undefined ? "" : parsed[header]);
  sheet.appendRow(row);
  return "Imported PDOC for pay date " + parsed.pay_date + ". Select that row and choose Apply selected PDOC row to Payroll.";
}

function applySelectedPdocToPayroll() {
  const ss = SpreadsheetApp.getActive();
  const pdocSheet = ss.getActiveSheet();
  if (pdocSheet.getName() !== PDOC_RESULTS_SHEET) {
    SpreadsheetApp.getUi().alert("Go to the PDOC Results tab and select the PDOC row first.");
    return;
  }
  const rowNumber = pdocSheet.getActiveCell().getRow();
  if (rowNumber <= 1) {
    SpreadsheetApp.getUi().alert("Select a PDOC result row, not the header.");
    return;
  }

  const pdoc = rowObject_(pdocSheet, rowNumber);
  const payrollSheet = ss.getSheetByName(DEFAULT_PAYROLL_SHEET);
  if (!payrollSheet) {
    SpreadsheetApp.getUi().alert("Payroll tab was not found.");
    return;
  }

  const payrollHeaders = ensureHeaders_(payrollSheet, [
    "employee_id",
    "pay_date",
    "gross",
    "cpp",
    "cpp2",
    "ei",
    "tax_fed",
    "tax_prov",
    "total_deductions",
    "net",
    "status",
    "pdoc_imported_at"
  ]);

  const matchRow = findPayrollRow_(payrollSheet, payrollHeaders, pdoc.employee_id, pdoc.pay_date);
  if (!matchRow) {
    SpreadsheetApp.getUi().alert("No matching Payroll row found for Employee ID " + pdoc.employee_id + " and pay date " + pdoc.pay_date + ".");
    return;
  }

  const updates = {
    cpp: pdoc.cpp,
    cpp2: pdoc.cpp2,
    ei: pdoc.ei,
    tax_fed: pdoc.tax_fed,
    tax_prov: pdoc.tax_prov,
    total_deductions: pdoc.total_deductions,
    net: pdoc.net,
    status: "CRA PDOC",
    pdoc_imported_at: new Date()
  };
  Object.keys(updates).forEach(header => {
    const col = payrollHeaders.indexOf(header) + 1;
    payrollSheet.getRange(matchRow, col).setValue(updates[header]);
  });
  SpreadsheetApp.getUi().alert("Payroll row updated with CRA PDOC amounts.");
}

function parsePdocText_(text) {
  if (text.indexOf("Payroll Deductions Online Calculator") === -1) {
    throw new Error("This does not look like CRA PDOC result text.");
  }
  const result = {
    pay_date: matchText_(text, /Date the employee is paid:\s+(\d{4}-\d{2}-\d{2})/i),
    frequency: matchText_(text, /Pay period frequency:\s+([^\n(]+)/i).trim(),
    gross: amount_(text, "Salary or wages income"),
    tax_fed: amount_(text, "Federal tax deduction"),
    tax_prov: amount_(text, "Provincial tax deduction"),
    cpp: amount_(text, "CPP deductions"),
    cpp2: amount_(text, "CPP2 deductions") || 0,
    ei: amount_(text, "EI deductions"),
    total_deductions: amount_(text, "Total deductions"),
    net: amount_(text, "Net amount")
  };
  const required = ["pay_date", "gross", "tax_fed", "tax_prov", "cpp", "ei", "total_deductions", "net"];
  required.forEach(key => {
    if (result[key] === "" || result[key] === null || result[key] === undefined) {
      throw new Error("Could not find " + key + " in the PDOC text.");
    }
  });
  return result;
}

function amount_(text, label) {
  const escaped = label.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = text.match(new RegExp(escaped + "\\s+(-?\\d[\\d,]*\\.\\d{2})", "i"));
  return match ? Number(match[1].replace(/,/g, "")) : null;
}

function matchText_(text, regex) {
  const match = text.match(regex);
  return match ? match[1] : "";
}

function getHeaders_(sheet) {
  if (sheet.getLastColumn() === 0) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
}

function ensureHeaders_(sheet, wantedHeaders) {
  let headers = getHeaders_(sheet).filter(Boolean);
  if (headers.length === 0) {
    sheet.getRange(1, 1, 1, wantedHeaders.length).setValues([wantedHeaders]);
    return wantedHeaders.slice();
  }
  const missing = wantedHeaders.filter(header => headers.indexOf(header) === -1);
  if (missing.length) {
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
    headers = headers.concat(missing);
  }
  return headers;
}

function rowObject_(sheet, rowNumber) {
  const headers = getHeaders_(sheet);
  const values = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  const result = {};
  headers.forEach((header, index) => {
    result[header] = values[index];
  });
  if (result.pay_date instanceof Date) {
    result.pay_date = Utilities.formatDate(result.pay_date, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return result;
}

function findPayrollRow_(sheet, headers, employeeId, payDate) {
  const employeeCol = headers.indexOf("employee_id");
  const payDateCol = headers.indexOf("pay_date");
  if (employeeCol === -1 || payDateCol === -1 || sheet.getLastRow() < 2) return 0;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  const wantedEmployee = String(employeeId || "").trim();
  const wantedDate = String(payDate || "").trim();
  for (let i = 0; i < values.length; i++) {
    const rowEmployee = String(values[i][employeeCol] || "").trim();
    let rowDate = values[i][payDateCol];
    if (rowDate instanceof Date) {
      rowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else {
      rowDate = String(rowDate || "").trim();
    }
    if (rowEmployee === wantedEmployee && rowDate === wantedDate) {
      return i + 2;
    }
  }
  return 0;
}
