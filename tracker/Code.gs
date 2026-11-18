/**
 * Bobcat Course Match — traffic sheet
 *
 * This Google Apps Script receives anonymous events from the website
 * and adds each one as a row in a Google Sheet.
 *
 * Set up (one time):
 *   1. Go to script.google.com > New project. Paste this whole file in.
 *   2. Pick "setup" in the function menu at the top and click Run.
 *      Allow the permissions it asks for. It creates the sheet and
 *      shows its link in the log.
 *   3. Deploy > New deployment > type "Web app".
 *      Execute as: Me.   Who has access: Anyone.   Click Deploy.
 *   4. Copy the Web app URL and paste it into TRACK_URL in index.html.
 *
 * What gets saved: time, event, detail (course name, interest...),
 * role picked in the quiz (student / staff), and phone vs computer.
 * No names, emails or IP addresses.
 */

const TITLE = "Bobcat Course Match — Traffic";
const LOG = "📝 All events";
const SUM = "📊 Summary";
const MAROON = "#6b1a1f", GOLD = "#e8b44c", CREAM = "#fff6e3", PINK = "#fbe3e1";

/* ---------- receives events from the website ---------- */
function doPost(e) {
  let ev = {};
  try { ev = JSON.parse((e && e.postData && e.postData.contents) || "{}"); } catch (err) {}
  return save_(ev);
}
function doGet(e) {               // also accepts ?event=...&detail=... (handy for testing)
  return save_((e && e.parameter) || {});
}
function save_(ev) {
  const clean = s => String(s || "").replace(/^[=+\-@]/, "'").slice(0, 120);   // keep cells safe
  const event = clean(ev.event);
  if (!event) return out_("ignored");
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    log_().appendRow([new Date(), event, clean(ev.detail), clean(ev.role), clean(ev.device)]);
  } finally { lock.releaseLock(); }
  return out_("ok");
}
function out_(msg) { return ContentService.createTextOutput(msg); }

function book_() {
  const id = PropertiesService.getScriptProperties().getProperty("SHEET_ID");
  if (!id) throw new Error("Run setup() first.");
  return SpreadsheetApp.openById(id);
}
function log_() { return book_().getSheetByName(LOG); }

/* ---------- builds the pretty sheet (run once) ---------- */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let ss = props.getProperty("SHEET_ID") ? SpreadsheetApp.openById(props.getProperty("SHEET_ID")) : null;
  if (!ss) {
    ss = SpreadsheetApp.create(TITLE);
    props.setProperty("SHEET_ID", ss.getId());
  }

  // --- All events (the raw log) ---
  let log = ss.getSheetByName(LOG) || ss.insertSheet(LOG);
  log.clear();
  log.getRange("A1:E1").setValues([["When", "Event", "Detail", "Role", "Device"]])
     .setBackground(MAROON).setFontColor("white").setFontWeight("bold");
  log.setFrozenRows(1);
  log.setColumnWidths(1, 1, 160); log.setColumnWidth(2, 110); log.setColumnWidth(3, 380);
  log.setColumnWidth(4, 90); log.setColumnWidth(5, 90);
  log.getRange("A2:A").setNumberFormat("mmm d, h:mm am/pm");
  log.getBandings().forEach(b => b.remove());
  log.getRange("A1:E1000").applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  log.getRange("A1:E1").setBackground(MAROON).setFontColor("white");

  // --- Summary ---
  let sum = ss.getSheetByName(SUM) || ss.insertSheet(SUM, 0);
  sum.clear(); sum.getCharts().forEach(c => sum.removeChart(c));
  sum.setHiddenGridlines(true);
  sum.getRange("A1:H200").setFontFamily("Nunito").setFontSize(11);
  sum.getRange("A1:H1").merge().setValue("🐾 Bobcat Course Match — Traffic")
     .setFontSize(20).setFontWeight("bold").setFontColor(MAROON).setBackground(CREAM);
  sum.getRange("A2:H2").merge().setValue("Updates live as students use the site. No names, emails or IP addresses are collected.")
     .setFontColor("#7a6c66").setBackground(CREAM);
  sum.setRowHeight(1, 44);

  const L = `'${LOG}'`;
  const cards = [
    ["👀 Visits",              `=COUNTIF(${L}!B:B,"visit")`],
    ["🚀 Started quiz",         `=COUNTIF(${L}!B:B,"start")`],
    ["🎉 Finished quiz",        `=COUNTIF(${L}!B:B,"finished")`],
    ["🎒 Finished — students",  `=COUNTIFS(${L}!B:B,"finished",${L}!C:C,"student")`],
    ["🏫 Finished — faculty/staff", `=COUNTIFS(${L}!B:B,"finished",${L}!C:C,"staff")`],
    ["🖱️ Course clicks",        `=COUNTIF(${L}!B:B,"open")`],
    ["📄 PDFs made",            `=COUNTIF(${L}!B:B,"pdf")`],
    ["🗂️ Opened All courses",   `=COUNTIF(${L}!B:B,"view")`],
    ["📱 Phone visits",         `=COUNTIFS(${L}!B:B,"visit",${L}!E:E,"phone")`],
    ["✅ Finish rate",          `=IFERROR(B7/B6,0)`]
  ];
  sum.getRange("A4:B4").setValues([["At a glance", ""]]);
  sum.getRange(5, 1, cards.length, 2).setValues(cards);
  sum.getRange(5 + cards.length - 1, 2).setNumberFormat("0%");
  styleTable_(sum, 4, 1, cards.length, 2);

  const tables = [
    ["🖱️ Most-clicked courses", "open",     "Course",   "Clicks"],
    ["✨ What students want to learn", "interest", "Topic", "Picks"],
    ["🤷 Asked for, but no match yet", "no-match", "Topic", "Picks"]
  ];
  let col = 4;
  tables.forEach(([title, ev, head, count]) => {
    sum.getRange(4, col).setValue(title);
    sum.getRange(5, col).setFormula(
      `=IFERROR(QUERY(${L}!B:C,"select C, count(C) where B = '${ev}' group by C order by count(C) desc label C '${head}', count(C) '${count}'",0),"—")`);
    styleTable_(sum, 4, col, 15, 2, true);
    col += 3;
  });
  // Visits per day
  sum.getRange(20, 1).setValue("📅 Visits per day");
  sum.getRange(21, 1).setFormula(
    `=IFERROR(QUERY(${L}!A:B,"select toDate(A), count(B) where B = 'visit' group by toDate(A) order by toDate(A) label toDate(A) 'Day', count(B) 'Visits'",1),"—")`);
  sum.getRange("A22:A60").setNumberFormat("mmm d");
  styleTable_(sum, 20, 1, 30, 2, true);

  sum.setColumnWidth(1, 210); sum.setColumnWidth(2, 80); sum.setColumnWidth(3, 20);
  sum.setColumnWidth(4, 330); sum.setColumnWidth(5, 70); sum.setColumnWidth(6, 20);
  sum.setColumnWidth(7, 260); sum.setColumnWidth(8, 70);
  // the third table sits under the second one so the sheet stays easy to read
  sum.getRange(4, 10, 17, 2).moveTo(sum.getRange(20, 4));

  // Chart: most-clicked courses
  const chart = sum.newChart().asBarChart()
    .addRange(sum.getRange("D6:E15"))
    .setOption("title", "Most-clicked courses")
    .setOption("colors", [MAROON])
    .setOption("legend", { position: "none" })
    .setPosition(20, 7, 0, 0).build();
  sum.insertChart(chart);

  // clean up the default empty tab
  const extra = ss.getSheetByName("Sheet1"); if (extra && ss.getSheets().length > 2) ss.deleteSheet(extra);
  ss.setActiveSheet(sum);

  Logger.log("Sheet ready: " + ss.getUrl());
  return ss.getUrl();
}

function styleTable_(sh, row, col, rows, cols, isQuery) {
  sh.getRange(row, col, 1, cols).merge().setFontWeight("bold").setFontSize(13)
    .setFontColor(MAROON).setBackground(PINK);
  if (isQuery) sh.getRange(row + 1, col, 1, cols).setFontWeight("bold").setFontColor("white").setBackground(MAROON);
  const body = sh.getRange(row + (isQuery ? 2 : 1), col, rows - (isQuery ? 1 : 0), cols);
  body.setBorder(false, false, true, false, false, true, "#f1e4da", SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(row + 1, col + cols - 1, rows, 1).setHorizontalAlignment("center").setFontWeight("bold");
}
