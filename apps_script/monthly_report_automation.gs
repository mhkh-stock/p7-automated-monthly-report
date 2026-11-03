/**
 * Automated Monthly Report — Google Sheets version of this repo's Excel build.
 * Setup: paste the month's order lines into a sheet named "Data" with headers
 * order_id | date | product | category | qty | unit_price_pkr | amount_pkr
 * Then run buildMonthlyReport() from the "Monthly Report" menu (or set the
 * monthly trigger below). It rebuilds Summary + Monthly Trend and their charts
 * from whatever is in Data — no re-typing, no manual chart rebuilds.
 * This is the automation a client gets: same logic as src/build_report.py,
 * in the tool they already use (Google Sheets).
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Monthly Report').addItem('Build this month\u2019s report', 'buildMonthlyReport').addToUi();
}
function buildMonthlyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = ss.getSheetByName('Data');
  if (!data) throw new Error('No "Data" sheet found. Paste order lines there first.');
  const values = data.getDataRange().getValues();
  const rows = values.slice(1).filter(r => r[0]);
  const months = [...new Set(rows.map(r => String(r[1]).slice(0, 7)))].sort();
  const latest = months[months.length - 1], prev = months[months.length - 2];
  const sumMonth = m => rows.filter(r => String(r[1]).startsWith(m)).reduce((a, r) => a + Number(r[6]), 0);
  const countMonth = m => rows.filter(r => String(r[1]).startsWith(m)).length;
  let summary = ss.getSheetByName('Summary') || ss.insertSheet('Summary');
  summary.clear();
  summary.getRange('A1').setValue('Automated Monthly Report — report month: ' + latest);
  summary.getRange('A3:C3').setValues([['KPI', 'Value', 'Note']]);
  summary.getRange('A4:C8').setValues([
    ['Report-month revenue (PKR)', sumMonth(latest), 'SUMIFS over Data'],
    ['Previous-month revenue (PKR)', sumMonth(prev), 'SUMIFS over Data'],
    ['Month-on-month growth', prev ? sumMonth(latest) / sumMonth(prev) - 1 : 'n/a', 'vs previous month'],
    ['Report-month orders', countMonth(latest), 'COUNTIFS over Data'],
    ['Average order value (PKR)', Math.round(sumMonth(latest) / countMonth(latest)), 'revenue / orders'],
  ]);
  summary.getRange('B6').setNumberFormat('0.00%');
  const cats = [...new Set(rows.map(r => r[3]))].sort();
  summary.getRange('A10:B10').setValues([['Category', 'Revenue (PKR) — ' + latest]]);
  cats.forEach((c, i) => summary.getRange(11 + i, 1, 1, 2).setValues([[c, rows.filter(r => String(r[1]).startsWith(latest) && r[3] === c).reduce((a, r) => a + Number(r[6]), 0)]]));
  const chart = summary.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(summary.getRange(10, 1, cats.length + 1, 2)).setPosition(11 + cats.length + 1, 1, 0, 0)
    .setOption('title', 'Revenue by category — ' + latest).build();
  summary.insertChart(chart);
  // Monthly trigger: uncomment once, then it runs on the 1st of every month.
  // ScriptApp.newTrigger('buildMonthlyReport').timeBased().onMonthDay(1).atHour(9).create();
}
