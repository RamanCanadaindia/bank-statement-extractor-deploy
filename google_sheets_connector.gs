/**
 * Raman Financial Services transaction receiver.
 *
 * Setup:
 * 1. In Apps Script, open Project Settings and add a Script Property named
 *    RFS_SHARED_SECRET with a long private value.
 * 2. Deploy as a Web App that executes as you. Choose the access setting that
 *    permits the Streamlit server to call it.
 * 3. Copy the deployment URL ending in /exec into the website.
 */

const LOG_SHEET_NAME = "_RFS Upload Log";

function jsonResponse_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function safeCell_(value) {
  if (typeof value === "string" && /^[=+\-@]/.test(value)) {
    return "'" + value;
  }
  return value === null || value === undefined ? "" : value;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const payload = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const expectedSecret = PropertiesService.getScriptProperties()
      .getProperty("RFS_SHARED_SECRET");

    if (!expectedSecret || payload.secret !== expectedSecret) {
      return jsonResponse_({ok: false, error: "Unauthorized connector request."});
    }
    if (payload.action !== "append_transactions") {
      return jsonResponse_({ok: false, error: "Unsupported action."});
    }
    if (!/^[A-Za-z0-9_-]{20,}$/.test(payload.spreadsheet_id || "")) {
      return jsonResponse_({ok: false, error: "Invalid spreadsheet ID."});
    }
    if (!payload.sheet_name || !payload.batch_id || !Array.isArray(payload.rows)) {
      return jsonResponse_({ok: false, error: "Missing destination or transaction data."});
    }
    if (payload.rows.length === 0) {
      return jsonResponse_({ok: false, error: "No transaction rows were supplied."});
    }

    const spreadsheet = SpreadsheetApp.openById(payload.spreadsheet_id);
    let logSheet = spreadsheet.getSheetByName(LOG_SHEET_NAME);
    if (!logSheet) {
      logSheet = spreadsheet.insertSheet(LOG_SHEET_NAME);
      logSheet.appendRow(["Batch ID", "Uploaded At", "Source File", "Bank", "Rows", "Tab"]);
      logSheet.hideSheet();
    }

    const duplicate = logSheet
      .getRange(1, 1, Math.max(logSheet.getLastRow(), 1), 1)
      .createTextFinder(payload.batch_id)
      .matchEntireCell(true)
      .findNext();
    if (duplicate) {
      return jsonResponse_({ok: true, duplicate: true, rows_added: 0});
    }

    let target = spreadsheet.getSheetByName(payload.sheet_name);
    if (!target) {
      target = spreadsheet.insertSheet(payload.sheet_name);
    }

    const incomingHeaders = [];
    payload.rows.forEach(row => {
      Object.keys(row).forEach(key => {
        if (!incomingHeaders.includes(key)) incomingHeaders.push(key);
      });
    });
    ["Upload Batch ID", "Uploaded At"].forEach(key => {
      if (!incomingHeaders.includes(key)) incomingHeaders.push(key);
    });

    let headers = [];
    if (target.getLastRow() > 0 && target.getLastColumn() > 0) {
      headers = target.getRange(1, 1, 1, target.getLastColumn()).getValues()[0]
        .map(String)
        .filter(Boolean);
    }
    if (headers.length === 0) {
      headers = incomingHeaders.slice();
      target.getRange(1, 1, 1, headers.length).setValues([headers]);
      target.setFrozenRows(1);
    } else {
      const missingHeaders = incomingHeaders.filter(header => !headers.includes(header));
      if (missingHeaders.length) {
        target.getRange(1, headers.length + 1, 1, missingHeaders.length)
          .setValues([missingHeaders]);
        headers = headers.concat(missingHeaders);
      }
    }

    const uploadedAt = new Date();
    const values = payload.rows.map(row => headers.map(header => {
      if (header === "Upload Batch ID") return payload.batch_id;
      if (header === "Uploaded At") return uploadedAt;
      return safeCell_(row[header]);
    }));
    target.getRange(target.getLastRow() + 1, 1, values.length, headers.length)
      .setValues(values);

    logSheet.appendRow([
      payload.batch_id,
      uploadedAt,
      payload.source_file || "",
      payload.bank || "",
      values.length,
      payload.sheet_name
    ]);

    return jsonResponse_({ok: true, duplicate: false, rows_added: values.length});
  } catch (error) {
    return jsonResponse_({ok: false, error: String(error.message || error)});
  } finally {
    try {
      lock.releaseLock();
    } catch (ignored) {}
  }
}
