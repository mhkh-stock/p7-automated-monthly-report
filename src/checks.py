"""Checks as code for P7. Recomputes every headline number from the raw CSV
and verifies the workbook structure (live formulas, tables, charts)."""
import csv, sys
from pathlib import Path
from openpyxl import load_workbook
BASE = Path(__file__).resolve().parents[1]
rows = list(csv.DictReader(open(BASE / "data" / "raw" / "sales_SAMPLE.csv")))
fails = []
def check(name, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + name + (f" — {detail}" if detail else ""))
    if not cond: fails.append(name)
def tot(month): return sum(int(r["amount_pkr"]) for r in rows if r["date"].startswith(month))
check("row count = 560", len(rows) == 560, str(len(rows)))
check("Jul revenue = 1,987,250", tot("2026-07") == 1987250, str(tot("2026-07")))
check("Aug revenue = 2,471,500", tot("2026-08") == 2471500, str(tot("2026-08")))
check("Sep revenue = 2,862,500", tot("2026-09") == 2862500, str(tot("2026-09")))
check("grand total = 7,321,250", sum(tot(m) for m in ["2026-07","2026-08","2026-09"]) == 7321250)
check("Sep orders = 205", sum(1 for r in rows if r["date"].startswith("2026-09")) == 205)
mom = tot("2026-09")/tot("2026-08")-1
check("Sep MoM growth ~ +15.8%", abs(mom - 0.1582) < 0.001, f"{mom:.2%}")
check("Sep category sums = Sep total", sum(sum(int(r["amount_pkr"]) for r in rows if r["date"].startswith("2026-09") and r["category"]==c) for c in sorted({r["category"] for r in rows})) == tot("2026-09"))
check("Sep top category = Furniture", max(((c, sum(int(r["amount_pkr"]) for r in rows if r["date"].startswith("2026-09") and r["category"]==c)) for c in sorted({r["category"] for r in rows})), key=lambda x: x[1])[0] == "Furniture")
wb = load_workbook(BASE / "outputs" / "monthly_report_SAMPLE.xlsx")
check("sheets = Data, Summary, Monthly Trend", wb.sheetnames == ["Data", "Summary", "Monthly Trend"], str(wb.sheetnames))
check("Data rows = 561 incl header", wb["Data"].max_row == 561, str(wb["Data"].max_row))
check("Summary B5 is a live formula", isinstance(wb["Summary"]["B5"].value, str) and wb["Summary"]["B5"].value.startswith("="), str(wb["Summary"]["B5"].value))
check("Summary B7 is a growth formula", isinstance(wb["Summary"]["B7"].value, str) and "B5/B6" in wb["Summary"]["B7"].value)
check("Summary has a chart", len(wb["Summary"]._charts) >= 1)
check("Monthly Trend has a chart", len(wb["Monthly Trend"]._charts) >= 1)
check("Apps Script file exists", (BASE / "apps_script" / "monthly_report_automation.gs").exists())
print(f"\n{len(fails)} failures" if fails else "\nALL CHECKS PASS")
sys.exit(1 if fails else 0)
