/**********************************************************************
 * MOHANPUR CONSTRUCTION — Google Apps Script backend
 *
 * SETUP:
 *  1. Create a Google Sheet. Rename tab to "hisab-mohanpur-construction"
 *     Row 1 headers: id | date | member | flow | category | amount | note
 *     (A "History" tab is created automatically for activity logs.)
 *  2. Extensions > Apps Script → paste this file → Save
 *  3. Deploy > Web app → Execute as Me → Anyone → copy URL
 *  4. Paste URL into CONFIG.APPS_SCRIPT_URL in index.html
 **********************************************************************/

var SHEET_NAME = "hisab-mohanpur-construction";
var HISTORY_SHEET = "History";
var HEADERS = ["id", "date", "member", "flow", "category", "amount", "note"];
var HISTORY_HEADERS = ["id", "ts", "action", "by", "entry", "prev"];
var VERSION = "1.1-history-sync";

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  return sh;
}

function getHistorySheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HISTORY_SHEET);
  if (!sh) {
    sh = ss.insertSheet(HISTORY_SHEET);
    sh.getRange(1, 1, 1, HISTORY_HEADERS.length).setValues([HISTORY_HEADERS]);
  } else if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HISTORY_HEADERS.length).setValues([HISTORY_HEADERS]);
  }
  return sh;
}

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === "ping") {
      return json_({ ok: true, version: VERSION });
    }
    if (e && e.parameter && e.parameter.action === "history") {
      return json_({ ok: true, data: readHistory_() });
    }
    var sh = getSheet_();
    var values = sh.getDataRange().getValues();
    var out = [];
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      if (!row[0]) continue;
      out.push({
        id: String(row[0]),
        date: formatDate_(row[1]),
        member: row[2],
        flow: row[3],
        category: row[4],
        amount: Number(row[5]) || 0,
        note: row[6] || ""
      });
    }
    return json_({ ok: true, data: out });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function readHistory_() {
  var sh = getHistorySheet_();
  if (sh.getLastRow() <= 1) return [];
  var values = sh.getDataRange().getValues();
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (!row[0]) continue;
    out.push({
      id: String(row[0]),
      ts: String(row[1] || ""),
      action: String(row[2] || ""),
      by: String(row[3] || ""),
      entry: parseJson_(row[4]),
      prev: row[5] ? parseJson_(row[5]) : null
    });
  }
  out.sort(function(a, b) { return String(b.ts).localeCompare(String(a.ts)); });
  return out;
}

function parseJson_(val) {
  try {
    return JSON.parse(String(val || "{}"));
  } catch (e) {
    return {};
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var sh = getSheet_();

    if (body.action === "add" && body.entry) {
      sh.appendRow(rowFromEntry_(body.entry));
      return json_({ ok: true });
    }

    if (body.action === "update" && body.entry && body.entry.id) {
      var r = findRowById_(sh, body.entry.id);
      if (r === -1) return json_({ ok: false, error: "id not found" });
      sh.getRange(r, 1, 1, HEADERS.length).setValues([rowFromEntry_(body.entry)]);
      return json_({ ok: true });
    }

    if (body.action === "delete" && body.id) {
      var rd = findRowById_(sh, body.id);
      if (rd === -1) return json_({ ok: false, error: "id not found" });
      sh.deleteRow(rd);
      return json_({ ok: true });
    }

    if (body.action === "logHistory" && body.item) {
      var hi = body.item;
      var hsh = getHistorySheet_();
      hsh.appendRow([
        hi.id || String(Date.now()),
        hi.ts || new Date().toISOString(),
        hi.action || "",
        hi.by || "",
        JSON.stringify(hi.entry || {}),
        hi.prev ? JSON.stringify(hi.prev) : ""
      ]);
      return json_({ ok: true });
    }

    return json_({ ok: false, error: "Unknown action" });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function rowFromEntry_(en) {
  return [
    en.id || String(Date.now()),
    en.date || "",
    en.member || "",
    en.flow || "",
    en.category || "",
    Number(en.amount) || 0,
    en.note || ""
  ];
}

function findRowById_(sh, id) {
  var ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (var i = 1; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function formatDate_(d) {
  if (d instanceof Date && !isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  if (typeof d === "number" && d > 1000) {
    var serial = new Date(Math.round((d - 25569) * 86400 * 1000));
    if (!isNaN(serial.getTime())) {
      return Utilities.formatDate(serial, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
  }
  var s = String(d || "").trim();
  if (!s) return "";
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[1] + "-" + iso[2] + "-" + iso[3];
  var dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (dmy) {
    var yr = Number(dmy[3]);
    if (yr < 100) yr += 2000;
    return yr + "-" + ("0" + dmy[2]).slice(-2) + "-" + ("0" + dmy[1]).slice(-2);
  }
  return s;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
