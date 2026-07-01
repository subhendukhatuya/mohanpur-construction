/**********************************************************************
 * MOHANPUR CONSTRUCTION — Google Apps Script backend
 * Connects the GitHub website to your Google Sheet (acts as the database)
 *
 * SETUP (one time):
 *  1. Create a Google Sheet. Rename the tab to "hisab-mohanpur-construction"
 *     (or create a sheet with that exact name). In row 1 put these headers
 *     (exactly, in this order):
 *        id | date | member | flow | category | amount | note
 *  2. In the Sheet menu: Extensions > Apps Script. Delete any code,
 *     paste THIS whole file, and Save.
 *  3. Click Deploy > New deployment > type "Web app".
 *        - Execute as:  Me
 *        - Who has access:  Anyone
 *     Deploy, authorise, and COPY the Web app URL.
 *  4. Paste that URL into CONFIG.APPS_SCRIPT_URL in index.html.
 *
 * Re-deploy (Deploy > Manage deployments > edit > Deploy) after any
 * code change so the new version goes live.
 *
 * flow values: "Credit" (money in) or "Expense" (debit / construction spend)
 * categories: Rod, Cement, Chips, Sand, Bricks, Labour, Soil, Paper,
 *             Tiles, Pipe, Electricity, Other (for Expense)
 **********************************************************************/

var SHEET_NAME = "hisab-mohanpur-construction";
var HEADERS = ["id", "date", "member", "flow", "category", "amount", "note"];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  return sh;
}

var VERSION = "1.0-mohanpur";

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === "ping") {
      return json_({ ok: true, version: VERSION });
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
  if (d instanceof Date) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(d || "");
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
