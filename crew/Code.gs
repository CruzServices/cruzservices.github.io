/**
 * CRUZ Crew Check-In — Google Apps Script
 * Paste this whole file into Extensions > Apps Script in your Google Sheet,
 * run setup() once, then Deploy > New deployment > Web app
 * (Execute as: Me, Who has access: Anyone). Copy the URL into crew/index.html.
 */

var LOG = "Log";
var HEADERS = ["Visit ID", "Date", "Site", "Name", "Station", "Station Name",
               "Clock In", "Clock Out", "Minutes", "Items Done", "Items Total", "Notes", "Status"];

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = ss.getSheetByName(LOG) || ss.insertSheet(LOG);
  if (log.getLastRow() === 0) {
    log.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
    log.setFrozenRows(1);
    log.getRange("B:B").setNumberFormat("yyyy-mm-dd");
    log.getRange("G:H").setNumberFormat("h:mm am/pm");
  }
  buildReports(ss);
}

function buildReports(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var d = ss.getSheetByName("Daily") || ss.insertSheet("Daily");
  d.clear();
  d.getRange("A1").setValue("Time to service the gym, by day").setFontWeight("bold").setFontSize(13);
  d.getRange("A2").setValue("Total minutes = everyone's minutes added together. Span = first clock-in to last clock-out.");
  d.getRange("A4").setFormula(
    '=QUERY(Log!B2:M, "select B, count(D), sum(I), min(G), max(H) where M=\'complete\' group by B order by B desc label B \'Date\', count(D) \'People\', sum(I) \'Total minutes\', min(G) \'First in\', max(H) \'Last out\'", 0)');
  d.getRange("G4").setValue("Avg minutes per person").setFontWeight("bold");
  d.getRange("G5").setFormula('=ARRAYFORMULA(IF(A5:A="","",ROUND(C5:C/B5:B)))');
  d.getRange("D5:E").setNumberFormat("h:mm am/pm");
  d.setFrozenRows(4);

  var p = ss.getSheetByName("By Person") || ss.insertSheet("By Person");
  p.clear();
  p.getRange("A1").setValue("Minutes per person per visit").setFontWeight("bold").setFontSize(13);
  p.getRange("A3").setFormula(
    '=QUERY(Log!B2:M, "select D, F, count(B), avg(I), min(I), max(I) where M=\'complete\' group by D, F order by D label D \'Name\', F \'Station\', count(B) \'Visits\', avg(I) \'Avg min\', min(I) \'Fastest\', max(I) \'Slowest\'", 0)');
  p.getRange("D4:D").setNumberFormat("0");

  var w = ss.getSheetByName("Weekly") || ss.insertSheet("Weekly");
  w.clear();
  w.getRange("A1").setValue("Weekly rollup (weeks start Monday)").setFontWeight("bold").setFontSize(13);
  w.getRange("A3").setFormula(
    '=QUERY({ARRAYFORMULA(IF(Log!B2:B="","",Log!B2:B-WEEKDAY(Log!B2:B,2)+1)), Log!I2:I, Log!M2:M, Log!B2:B}, "select Col1, count(Col4), sum(Col2) where Col3=\'complete\' group by Col1 order by Col1 desc label Col1 \'Week of\', count(Col4) \'Visits\', sum(Col2) \'Total minutes\'", 0)');
  w.getRange("A4:A").setNumberFormat("yyyy-mm-dd");
}

/* ---------- web app ---------- */

function doGet(e) {
  var action = (e.parameter.action || "");
  if (action === "today") return json(todayClaims(e.parameter.site));
  return json({ ok: true, msg: "CRUZ crew endpoint is live" });
}

function doPost(e) {
  var p = JSON.parse(e.postData.contents);
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG);
    if (p.action === "clockin") {
      var start = new Date(p.startedAt);
      sheet.appendRow([p.visitId, dateOnly(start), p.site, p.name, p.station, p.stationName, start, "", "", "", "", "", "in progress"]);
      return json({ ok: true });
    }
    if (p.action === "complete") {
      var start2 = new Date(p.startedAt), end = new Date(p.endedAt);
      var row = findRow(sheet, p.visitId);
      var values = [p.visitId, dateOnly(start2), p.site, p.name, p.station, p.stationName, start2, end,
                    p.minutes, p.itemsDone, p.itemsTotal, p.notes || "", "complete"];
      if (row) sheet.getRange(row, 1, 1, values.length).setValues([values]);
      else sheet.appendRow(values);
      return json({ ok: true });
    }
    return json({ ok: false, error: "unknown action" });
  } finally { lock.releaseLock(); }
}

function todayClaims(site) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG);
  var out = {};
  if (sheet.getLastRow() < 2) return out;
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
  rows.forEach(function (r) {
    var d = r[1] instanceof Date ? Utilities.formatDate(r[1], tz, "yyyy-MM-dd") : String(r[1]);
    if (d === today && (!site || r[2] === site)) out[r[4]] = r[3];
  });
  return out;
}

function findRow(sheet, visitId) {
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) if (ids[i][0] === visitId) return i + 2;
  return 0;
}

function dateOnly(d) {
  var tz = Session.getScriptTimeZone();
  return new Date(Utilities.formatDate(d, tz, "yyyy-MM-dd") + "T00:00:00");
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- optional: Monday email summary ----------
 * In Apps Script: Triggers (clock icon) > Add trigger > weeklyEmail, Time-driven, Week timer, Every Monday, 7-8am.
 */
function weeklyEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = ss.getSheetByName(LOG);
  if (log.getLastRow() < 2) return;
  var tz = Session.getScriptTimeZone();
  var since = new Date(); since.setDate(since.getDate() - 7);
  var rows = log.getRange(2, 1, log.getLastRow() - 1, HEADERS.length).getValues()
    .filter(function (r) { return r[12] === "complete" && r[1] >= since; });
  var byDay = {};
  rows.forEach(function (r) {
    var k = Utilities.formatDate(r[1], tz, "EEE MMM d");
    byDay[k] = byDay[k] || { people: 0, minutes: 0, lines: [] };
    byDay[k].people++; byDay[k].minutes += Number(r[8]);
    byDay[k].lines.push("   " + r[3] + " — " + r[5] + ": " + r[8] + " min" + (r[11] ? " (" + r[11] + ")" : ""));
  });
  var body = "CRUZ crew — last 7 days at " + (rows[0] ? rows[0][2] : "the gym") + "\n\n";
  Object.keys(byDay).forEach(function (k) {
    body += k + ": " + byDay[k].people + " people, " + byDay[k].minutes + " total minutes\n" + byDay[k].lines.join("\n") + "\n\n";
  });
  body += "Sheet: " + ss.getUrl();
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), "CRUZ weekly crew summary", body);
}
