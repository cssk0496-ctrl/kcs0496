const SHEET_NAME = "예산내역";
const HEADERS = ["ID", "연월일", "팀명", "팀원", "항목", "금액"];

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || "list";
    if (action !== "list") {
      throw new Error("지원하지 않는 GET 요청입니다.");
    }
    return jsonResponse({ success: true, data: readAllRows() });
  } catch (error) {
    return errorResponse(error);
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const payload = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const action = payload.action;

    if (action === "append") {
      appendRow(payload);
    } else if (action === "delete") {
      deleteRows(payload.ids || []);
    } else if (action === "clear") {
      replaceAllRows([]);
    } else if (action === "replaceAll") {
      replaceAllRows(payload.data || []);
    } else {
      throw new Error("지원하지 않는 요청입니다: " + action);
    }

    SpreadsheetApp.flush();
    return jsonResponse({ success: true });
  } catch (error) {
    return errorResponse(error);
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

function getSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error("Apps Script를 사용할 스프레드시트에서 직접 만들어 주세요.");
  }

  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  } else {
    const currentHeaders = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(String);
    const oldHeaders = ["ID", "연월", "팀원", "항목", "금액"];
    if (currentHeaders.join("|") === oldHeaders.join("|")) {
      sheet.insertColumnBefore(3);
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      if (sheet.getLastRow() > 1) {
        sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).setValue("미지정");
      }
    } else if (
      currentHeaders.join("|") ===
      ["ID", "연월", "팀명", "팀원", "항목", "금액"].join("|")
    ) {
      sheet.getRange(1, 2).setValue("연월일");
    }
  }
  return sheet;
}

function readAllRows() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  return sheet
    .getRange(2, 1, lastRow - 1, HEADERS.length)
    .getValues()
    .filter(row => row.some(value => value !== ""))
    .map(row => ({
      ID: String(row[0]),
      "연월일": normalizeDate(row[1]),
      "팀명": String(row[2]),
      "팀원": String(row[3]),
      "항목": String(row[4]),
      "금액": Number(row[5]) || 0,
    }));
}

function normalizeDate(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }

  const text = String(value).trim();
  const match = text.match(/^(\d{4})-(\d{2})(?:-(\d{2}))?/);
  if (match) {
    return match[1] + "-" + match[2] + "-" + (match[3] || "01");
  }

  const parsed = new Date(value);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return text;
}

function appendRow(payload) {
  validateRow(payload);
  getSheet().appendRow([
    String(payload.ID),
    String(payload["연월일"]),
    String(payload["팀명"]),
    String(payload["팀원"]),
    String(payload["항목"]),
    Number(payload["금액"]),
  ]);
}

function deleteRows(ids) {
  const idSet = new Set(ids.map(String));
  if (idSet.size === 0) {
    return;
  }

  const remainingRows = readAllRows().filter(row => !idSet.has(String(row.ID)));
  replaceAllRows(remainingRows);
}

function replaceAllRows(rows) {
  rows.forEach(validateRow);
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, HEADERS.length).clearContent();
  }
  if (rows.length > 0) {
    const values = rows.map(row => [
      String(row.ID),
      String(row["연월일"]),
      String(row["팀명"]),
      String(row["팀원"]),
      String(row["항목"]),
      Number(row["금액"]),
    ]);
    sheet.getRange(2, 1, values.length, HEADERS.length).setValues(values);
  }
}

function validateRow(row) {
  const required = ["ID", "연월일", "팀명", "팀원", "항목", "금액"];
  required.forEach(key => {
    if (row[key] === undefined || row[key] === null || row[key] === "") {
      throw new Error("필수 값이 없습니다: " + key);
    }
  });

  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(row["연월일"]))) {
    throw new Error("연월일은 YYYY-MM-DD 형식이어야 합니다.");
  }
  if (!Number.isFinite(Number(row["금액"])) || Number(row["금액"]) < 0) {
    throw new Error("금액은 0 이상의 숫자여야 합니다.");
  }
}

function jsonResponse(body) {
  return ContentService
    .createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON);
}

function errorResponse(error) {
  return jsonResponse({
    success: false,
    error: error && error.message ? error.message : String(error),
  });
}
