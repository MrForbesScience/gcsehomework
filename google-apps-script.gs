const SPREADSHEET_ID = "1Yq3bcdnPVQTw5hT_rid0zD0iA5PO2pTMRAQCs2puIkI";
const SHEET_NAME = "Results";
const HEADERS = ["year", "subject", "week", "class", "name", "weeklyScore", "total"];
const TEACHER_PASSWORD = "Ed3xcel!";

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const data = { ok: true, results: listResults() };
  const text = params.callback ? params.callback + "(" + JSON.stringify(data) + ")" : JSON.stringify(data);
  return ContentService
    .createTextOutput(text)
    .setMimeType(params.callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function doPost(e) {
  const payload = parsePayload(e);
  if (payload.action === "resetLeaderboard") {
    if (payload.password !== TEACHER_PASSWORD) return json({ ok: false, error: "Incorrect password" });
    return json({ ok: true, removed: resetLeaderboard() });
  }
  if (payload.action === "deleteStudents") {
    return json({ ok: true, removed: deleteStudents(payload.students || []) });
  }
  appendResult(payload.result || payload);
  return json({ ok: true });
}

function parsePayload(e) {
  if (!e || !e.postData || !e.postData.contents) return {};
  try {
    return JSON.parse(e.postData.contents);
  } catch (err) {
    return {};
  }
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  const firstRow = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (firstRow.join("|") !== HEADERS.join("|")) {
    sh.clearContents();
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  return sh;
}

function clean(value) {
  return String(value || "").trim().replace(/\s+/g, " ");
}

function appendResult(raw) {
  const r = raw || {};
  const name = clean(r.name || [r.first, r.last].filter(Boolean).join(" "));
  sheet().appendRow([
    clean(r.year),
    clean(r.subject).toLowerCase(),
    Number(r.week || 0),
    clean(r.cls || r.class),
    name,
    Number(r.weeklyScore || r.score || 0),
    Number(r.total || 10)
  ]);
}

function listResults() {
  const sh = sheet();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1)
    .filter(row => row.join("") !== "")
    .map(row => {
      const item = {};
      headers.forEach((header, index) => item[header] = row[index]);
      return item;
    });
}

function deleteStudents(students) {
  const sh = sheet();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return 0;

  const headers = values[0].map(String);
  const nameIndex = headers.indexOf("name");
  const classIndex = headers.indexOf("class");
  const targets = new Set(students.map(s => [
    clean(s.name || [s.firstKey, s.lastKey].filter(Boolean).join(" ")).toLowerCase(),
    clean(s.cls || s.class)
  ].join("|")));

  let removed = 0;
  for (let row = values.length; row >= 2; row--) {
    const data = values[row - 1];
    const rowKey = [
      clean(data[nameIndex]).toLowerCase(),
      clean(data[classIndex])
    ].join("|");
    if (targets.has(rowKey)) {
      sh.deleteRow(row);
      removed++;
    }
  }
  return removed;
}

function resetLeaderboard() {
  const sh = sheet();
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return 0;
  sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).clearContent();
  return lastRow - 1;
}

