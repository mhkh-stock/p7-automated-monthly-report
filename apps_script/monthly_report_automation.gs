/**
 * Automated Monthly Report — Google Sheets version of this repo's Excel build.
 * Setup: paste the month's order lines into a sheet named "Data" with headers
 * order_id | date | product | category | qty | unit_price_pkr | amount_pkr
 * Then run buildMonthlyReport() from the "Monthly Report" menu (or set the
 * monthly trigger with setupMonthlyTrigger(), run once). It rebuilds the
 * Summary sheet (KPIs + category chart) from whatever is in Data — no
 * re-typing, no manual chart rebuilds. NOTE: unlike the Excel build, this
 * script rebuilds Summary only; it does not build the Monthly Trend sheet.
 * The Summary values are computed fresh on each run (not live formulas),
 * so re-run it after pasting a new month's lines.
 * This is the automation a client gets: same logic as src/build_report.py,
 * in the tool they already use (Google Sheets).
 */
// Month string ("yyyy-MM") for a Data date cell. Dates pasted into Sheets
// usually arrive as Date objects, for which String(date).slice(0, 7) gives
// "Mon Sep", not "2026-09" — so format real Dates properly, and accept
// plain "yyyy-MM-dd" text as-is.
function monthOf(v) {
  if (v instanceof Date) return Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM');
  return String(v).slice(0, 7);
}
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Monthly Report').addItem('Build this month\u2019s report', 'buildMonthlyReport').addToUi();
}
function buildMonthlyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = ss.getSheetByName('Data');
  if (!data) throw new Error('No "Data" sheet found. Paste order lines there first.');
  const values = data.getDataRange().getValues();
  const rows = values.slice(1).filter(r => r[0]);
  if (!rows.length) throw new Error('No data rows in "Data". Paste order lines there first.');
  const months = [...new Set(rows.map(r => monthOf(r[1])))].sort();
  const latest = months[months.length - 1], prev = months[months.length - 2];
  const sumMonth = m => rows.filter(r => monthOf(r[1]) === m).reduce((a, r) => a + Number(r[6]), 0);
  const countMonth = m => rows.filter(r => monthOf(r[1]) === m).length;
  let summary = ss.getSheetByName('Summary') || ss.insertSheet('Summary');
  summary.clear();
  summary.getCharts().forEach(c => summary.removeChart(c)); // clear() leaves charts behind; drop them so re-runs don't stack duplicates
  summary.getRange('A1').setValue('Automated Monthly Report — report month: ' + latest);
  summary.getRange('A3:C3').setValues([['KPI', 'Value', 'Note']]);
  summary.getRange('A4:C8').setValues([
    ['Report-month revenue (PKR)', sumMonth(latest), 'recalculated on each run'],
    ['Previous-month revenue (PKR)', sumMonth(prev), 'recalculated on each run'],
    ['Month-on-month growth', prev ? sumMonth(latest) / sumMonth(prev) - 1 : 'n/a', 'vs previous month'],
    ['Report-month orders', countMonth(latest), 'recalculated on each run'],
    ['Average order value (PKR)', Math.round(sumMonth(latest) / countMonth(latest)), 'revenue / orders'],
  ]);
  summary.getRange('B6').setNumberFormat('0.00%');
  const cats = [...new Set(rows.map(r => r[3]))].sort();
  summary.getRange('A10:B10').setValues([['Category', 'Revenue (PKR) — ' + latest]]);
  cats.forEach((c, i) => summary.getRange(11 + i, 1, 1, 2).setValues([[c, rows.filter(r => monthOf(r[1]) === latest && r[3] === c).reduce((a, r) => a + Number(r[6]), 0)]]));
  const chart = summary.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(summary.getRange(10, 1, cats.length + 1, 2)).setPosition(11 + cats.length + 1, 1, 0, 0)
    .setOption('title', 'Revenue by category — ' + latest).build();
  summary.insertChart(chart);
}

// Monthly trigger: run setupMonthlyTrigger() ONCE and the report rebuilds
// on the 1st of every month at 09:00. Kept separate from
// buildMonthlyReport() on purpose — creating the trigger inside the build
// would add a duplicate trigger on every run. Safe to re-run: it skips
// creating a second trigger if one already exists.
function setupMonthlyTrigger() {
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'buildMonthlyReport');
  if (!exists) ScriptApp.newTrigger('buildMonthlyReport').timeBased().onMonthDay(1).atHour(9).create();
}
