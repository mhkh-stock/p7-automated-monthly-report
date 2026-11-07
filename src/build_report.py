"""Build the automated monthly report workbook from data/raw/sales_SAMPLE.csv.
Data sheet = raw lines (Excel Table). Summary + Monthly Trend sheets use
LIVE formulas (SUMPRODUCT) over Data, so pasting a new
month's lines into Data and extending the month row rebuilds the report
with no re-typing. Charts are native Excel charts over the summary tables.
Report month = latest month in the data (Sep 2026 in the sample). NOTE:
the live formulas are built for the months present at build time (fixed
row range and month literals) — a new month needs a rebuild via this
script, it does not move Summary on its own."""
import csv
from pathlib import Path
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font, PatternFill
BASE = Path(__file__).resolve().parents[1]
rows = list(csv.DictReader(open(BASE / "data" / "raw" / "sales_SAMPLE.csv")))
months = sorted({r["date"][:7] for r in rows})
latest, prev = months[-1], months[-2]
wb = Workbook()
ws = wb.active; ws.title = "Data"
ws.append(["order_id", "date", "product", "category", "qty", "unit_price_pkr", "amount_pkr"])
for r in rows:
    ws.append([r["order_id"], r["date"], r["product"], r["category"], int(r["qty"]), int(r["unit_price_pkr"]), int(r["amount_pkr"])])
ws.add_table(Table(displayName="SalesData", ref=f"A1:G{len(rows)+1}"))
ws.column_dimensions["A"].width = 12; ws.column_dimensions["B"].width = 12
ws.column_dimensions["C"].width = 24; ws.column_dimensions["D"].width = 14
for col in "EFG": ws.column_dimensions[col].width = 16
for c in ws[1]: c.font = Font(bold=True)
# Summary sheet: KPIs for latest month, all live formulas over Data
s = wb.create_sheet("Summary")
s["A1"] = "Automated Monthly Report — CityMart (SAMPLE)"; s["A1"].font = Font(bold=True, size=14)
s["A2"] = f"Report month: {latest}  |  Previous month: {prev}  |  To add a new month, paste its lines into Data and re-run src/build_report.py — this sheet is rebuilt from Data."
s["A4"] = "KPI"; s["B4"] = "Value"; s["C4"] = "How it is calculated (live formula)"
for c in s[4]: c.font = Font(bold=True)
# Data rows 2..561; date col B, amount G, order A. Month match via LEFT(B,7)=month text in A col of trend? Use SUMPRODUCT for month filtering.
def sum_month(month_cell):
    return f'=SUMPRODUCT((LEFT(Data!$B$2:$B${len(rows)+1},7)={month_cell})*Data!$G$2:$G${len(rows)+1})'
def count_month(month_cell):
    return f'=SUMPRODUCT((LEFT(Data!$B$2:$B${len(rows)+1},7)={month_cell})*1)'
s["A5"] = "Report-month revenue (PKR)"; s["B5"] = sum_month(f'"{latest}"')
s["A6"] = "Previous-month revenue (PKR)"; s["B6"] = sum_month(f'"{prev}"')
s["A7"] = "Month-on-month growth"; s["B7"] = "=IF(B6=0,\"n/a\",B5/B6-1)"; s["B7"].number_format = '0.00%'
s["A8"] = "Report-month orders"; s["B8"] = count_month(f'"{latest}"')
s["A9"] = "Average order value (PKR)"; s["B9"] = "=IF(B8=0,0,B5/B8)"; s["B9"].number_format = '#,##0'
for r in (5, 6): s[f"B{r}"].number_format = '#,##0'
s["A11"] = "Revenue by category — report month (PKR)"; s["A11"].font = Font(bold=True)
s["A12"] = "Category"; s["B12"] = "Revenue"; s["C12"] = "Share"
for c in s[12]: c.font = Font(bold=True)
cats = sorted({r["category"] for r in rows})
for i, cat in enumerate(cats):
    row = 13 + i
    s[f"A{row}"] = cat
    s[f"B{row}"] = f'=SUMPRODUCT((LEFT(Data!$B$2:$B${len(rows)+1},7)="{latest}")*(Data!$D$2:$D${len(rows)+1}="{cat}")*Data!$G$2:$G${len(rows)+1})'
    s[f"B{row}"].number_format = '#,##0'
    s[f"C{row}"] = f"=IF($B$5=0,0,B{row}/$B$5)"; s[f"C{row}"].number_format = '0.0%'
bar = BarChart(); bar.title = f"Revenue by category — {latest}"; bar.y_axis.title = "PKR"; bar.x_axis.title = "Category"
bar.add_data(Reference(s, min_col=2, min_row=12, max_row=12 + len(cats)), titles_from_data=True)
bar.set_categories(Reference(s, min_col=1, min_row=13, max_row=12 + len(cats))); bar.height = 7; bar.width = 13
s.add_chart(bar, f"A{14 + len(cats)}")
# Monthly trend sheet
t = wb.create_sheet("Monthly Trend")
t["A1"] = "Month"; t["B1"] = "Revenue (PKR)"; t["C1"] = "Orders"; t["D1"] = "Avg order value (PKR)"
for c in t[1]: c.font = Font(bold=True)
for i, m in enumerate(months):
    row = 2 + i
    t[f"A{row}"] = m
    t[f"B{row}"] = sum_month(f"A{row}")
    t[f"C{row}"] = count_month(f"A{row}")
    t[f"D{row}"] = f"=IF(C{row}=0,0,B{row}/C{row})"; t[f"D{row}"].number_format = '#,##0'; t[f"B{row}"].number_format = '#,##0'
line = LineChart(); line.title = "Monthly revenue trend"; line.y_axis.title = "PKR"
line.add_data(Reference(t, min_col=2, min_row=1, max_row=1 + len(months)), titles_from_data=True)
line.set_categories(Reference(t, min_col=1, min_row=2, max_row=1 + len(months))); line.height = 7; line.width = 13
t.add_chart(line, "A6")
t["A10"] = "How to run next month: paste the new month's lines into Data and re-run src/build_report.py — it re-derives the report month and rebuilds Summary and Monthly Trend from Data. The Google Apps Script in apps_script/ rebuilds Summary inside Google Sheets (run buildMonthlyReport() from the 'Monthly Report' menu)."
for sh in (s, t):
    sh.column_dimensions["A"].width = 42; sh.column_dimensions["B"].width = 20; sh.column_dimensions["C"].width = 22; sh.column_dimensions["D"].width = 20
wb.save(BASE / "outputs" / "monthly_report_SAMPLE.xlsx")
print(f"saved workbook: {len(rows)} data rows, months {months}, latest {latest}")
