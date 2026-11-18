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
 * What gets saved: date, event, detail (course name, interest...),
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

/* ---------- builds the pretty sheet ----------
   Safe to run again: it rebuilds the Summary and By day tabs but never
   deletes the rows already saved in "All events". */
const DAY = "📅 By day";
function setup() {
  const props = PropertiesService.getScriptProperties();
  let ss = props.getProperty("SHEET_ID") ? SpreadsheetApp.openById(props.getProperty("SHEET_ID")) : null;
  if (!ss) {
    ss = SpreadsheetApp.create(TITLE);
    props.setProperty("SHEET_ID", ss.getId());
  }
  const L = `'${LOG}'`, D = `'${DAY}'`;

  // --- All events (the raw log). Only the header and formats are touched. ---
  let log = ss.getSheetByName(LOG) || ss.insertSheet(LOG);
  log.getRange("A1:F1").setValues([["Date", "Event", "Detail", "Role", "Device", "Day"]])
     .setBackground(MAROON).setFontColor("white").setFontWeight("bold");
  log.getRange("F1").setFormula('={"Day";ARRAYFORMULA(IF(A2:A="",,INT(A2:A)))}');
  log.setFrozenRows(1);
  log.setColumnWidth(1, 130); log.setColumnWidth(2, 110); log.setColumnWidth(3, 380);
  log.setColumnWidth(4, 90); log.setColumnWidth(5, 90);
  log.getRange("A2:A").setNumberFormat("ddd, mmm d, yyyy");      // days, no hours
  log.getRange("F2:F").setNumberFormat("yyyy-mm-dd");
  log.hideColumns(6);                                            // helper column for the day counts
  log.getBandings().forEach(b => b.remove());
  log.getRange("A1:E2000").applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  log.getRange("A1:F1").setBackground(MAROON).setFontColor("white");

  // --- By day: one row per day, counts per column ---
  let day = ss.getSheetByName(DAY) || ss.insertSheet(DAY, 1);
  day.clear();
  day.getRange("A1:F1").setValues([["Day", "Visits", "Started quiz", "Finished quiz", "Course clicks", "PDFs + copies"]])
     .setBackground(MAROON).setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  day.getRange("A2").setFormula(`=IFERROR(SORT(UNIQUE(FILTER(${L}!F2:F,${L}!F2:F<>""))),)`);
  const count = (col, crit) => day.getRange(col + "2").setFormula(
    `=ARRAYFORMULA(IF(A2:A="",,COUNTIFS(${L}!F:F,A2:A,${L}!B:B,"${crit}")))`);
  count("B", "visit"); count("C", "start"); count("D", "finished"); count("E", "open");
  day.getRange("F2").setFormula(
    `=ARRAYFORMULA(IF(A2:A="",,COUNTIFS(${L}!F:F,A2:A,${L}!B:B,"pdf")+COUNTIFS(${L}!F:F,A2:A,${L}!B:B,"copy")))`);
  day.getRange("A2:A").setNumberFormat("ddd, mmm d");
  day.getRange("B2:F").setHorizontalAlignment("center");
  day.setFrozenRows(1);
  day.setColumnWidth(1, 140); for (let c = 2; c <= 6; c++) day.setColumnWidth(c, 120);
  day.getBandings().forEach(b => b.remove());
  day.getRange("A1:F400").applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  day.getRange("A1:F1").setBackground(MAROON).setFontColor("white");

  // --- Summary ---
  let sum = ss.getSheetByName(SUM) || ss.insertSheet(SUM, 0);
  sum.getCharts().forEach(c => sum.removeChart(c));
  sum.getRange("A1:K200").breakApart();
  sum.clear();
  sum.setHiddenGridlines(true);
  sum.getRange("A1:H200").setFontFamily("Nunito").setFontSize(11);
  sum.getRange("A1:H1").merge().setValue("🐾 Bobcat Course Match — Traffic")
     .setFontSize(20).setFontWeight("bold").setFontColor(MAROON).setBackground(CREAM);
  sum.getRange("A2:H2").merge().setValue("Updates live as students use the site. No names, emails or IP addresses are collected.")
     .setFontColor("#7a6c66").setBackground(CREAM);
  sum.setRowHeight(1, 44);

  const cards = [
    ["👀 Visits",              `=COUNTIF(${L}!B:B,"visit")`],
    ["🚀 Started quiz",         `=COUNTIF(${L}!B:B,"start")`],
    ["🎉 Finished quiz",        `=COUNTIF(${L}!B:B,"finished")`],
    ["🎒 Finished — students",  `=COUNTIFS(${L}!B:B,"finished",${L}!C:C,"student")`],
    ["🏫 Finished — faculty/staff", `=COUNTIFS(${L}!B:B,"finished",${L}!C:C,"staff")`],
    ["🖱️ Course clicks",        `=COUNTIF(${L}!B:B,"open")`],
    ["📄 Saved (PDF or copy)",  `=COUNTIF(${L}!B:B,"pdf")+COUNTIF(${L}!B:B,"copy")`],
    ["🗂️ Opened All courses",   `=COUNTIF(${L}!B:B,"view")`],
    ["📱 Phone visits",         `=COUNTIFS(${L}!B:B,"visit",${L}!E:E,"phone")`],
    ["✅ Finish rate",          `=IFERROR(B7/B6,0)`]
  ];
  sum.getRange("A4:B4").setValues([["At a glance", ""]]);
  sum.getRange(5, 1, cards.length, 2).setValues(cards);
  sum.getRange(5 + cards.length - 1, 2).setNumberFormat("0%");
  styleTable_(sum, 4, 1, cards.length, 2);

  const query = (ev, head, cnt) =>
    `=IFERROR(QUERY(${L}!B:C,"select C, count(C) where B = '${ev}' group by C order by count(C) desc label C '${head}', count(C) '${cnt}'",0),"—")`;
  sum.getRange(4, 4).setValue("🖱️ Most-clicked courses");
  sum.getRange(5, 4).setFormula(query("open", "Course", "Clicks"));
  styleTable_(sum, 4, 4, 15, 2, true);
  sum.getRange(4, 7).setValue("✨ What students want to learn");
  sum.getRange(5, 7).setFormula(query("interest", "Topic", "Picks"));
  styleTable_(sum, 4, 7, 15, 2, true);
  sum.getRange(21, 1).setValue("🤷 Asked for, but no match yet");
  sum.getRange(22, 1).setFormula(query("no-match", "Topic", "Picks"));
  styleTable_(sum, 21, 1, 15, 2, true);

  sum.setColumnWidth(1, 210); sum.setColumnWidth(2, 80); sum.setColumnWidth(3, 20);
  sum.setColumnWidth(4, 330); sum.setColumnWidth(5, 70); sum.setColumnWidth(6, 20);
  sum.setColumnWidth(7, 260); sum.setColumnWidth(8, 70);

  // Trend: visits, started, finished and course clicks per day
  const trend = sum.newChart().asLineChart()
    .addRange(day.getRange("A1:E400"))
    .setNumHeaders(1)
    .setOption("title", "📈 Day by day")
    .setOption("colors", [MAROON, GOLD, "#3b8a64", "#2f6fe0"])
    .setOption("legend", { position: "bottom" })
    .setOption("pointSize", 6)
    .setOption("curveType", "function")
    .setOption("vAxis", { minValue: 0, format: "0" })
    .setOption("hAxis", { format: "MMM d" })
    .setOption("width", 760).setOption("height", 330)
    .setPosition(21, 4, 0, 0).build();
  sum.insertChart(trend);

  // Bar chart: most-clicked courses
  const bars = sum.newChart().asBarChart()
    .addRange(sum.getRange("D6:E15"))
    .setOption("title", "🖱️ Most-clicked courses")
    .setOption("colors", [MAROON])
    .setOption("legend", { position: "none" })
    .setOption("width", 760).setOption("height", 300)
    .setPosition(40, 4, 0, 0).build();
  sum.insertChart(bars);

  // clean up the default empty tab
  const extra = ss.getSheetByName("Sheet1"); if (extra && ss.getSheets().length > 3) ss.deleteSheet(extra);
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
