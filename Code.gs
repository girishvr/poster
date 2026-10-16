function doGet() {
  return jsonResponse({ ok: true, message: "Slogan poster counter endpoint is ready." });
}

function doPost(event) {
  const payload = JSON.parse(event.postData.contents);
  const rowNumber = Number(payload.row);
  const columnName = payload.metric === "views"
    ? "View count"
    : payload.metric === "downloads"
      ? "Downloads"
      : "";

  if (!Number.isInteger(rowNumber) || rowNumber < 2 || !columnName) {
    return jsonResponse({ ok: false, error: "Invalid row or metric." });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheets().find((candidate) => {
      const rowCount = candidate.getLastRow();
      return rowCount > 0 && candidate.getRange(1, 2, rowCount, 1)
        .getDisplayValues()
        .some((row) => row[0].trim().toLowerCase() === "slogan");
    });

    if (!sheet) throw new Error('No sheet with a "Slogan" heading in column B was found.');

    const values = sheet.getDataRange().getDisplayValues();
    const headerRowIndex = values.findIndex((row) => row[1]?.trim().toLowerCase() === "slogan");
    if (headerRowIndex < 0 || rowNumber <= headerRowIndex + 1 || rowNumber > sheet.getMaxRows()) {
      throw new Error("The requested poster row is outside the slogan data.");
    }
    if (!values[rowNumber - 1]?.[1]?.trim()) {
      throw new Error("The requested row does not contain a slogan.");
    }

    const headers = values[headerRowIndex].map((value) => value.trim().toLowerCase());
    let columnIndex = headers.indexOf(columnName.toLowerCase());
    if (columnIndex < 0) {
      columnIndex = Math.max(sheet.getLastColumn(), 2);
      sheet.getRange(headerRowIndex + 1, columnIndex + 1).setValue(columnName);
    }

    const counterCell = sheet.getRange(rowNumber, columnIndex + 1);
    const currentCount = Math.max(0, Math.trunc(Number(counterCell.getValue()) || 0));
    counterCell.setValue(currentCount + 1);
    return jsonResponse({ ok: true, row: rowNumber, metric: columnName, count: currentCount + 1 });
  } finally {
    lock.releaseLock();
  }
}

function jsonResponse(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
