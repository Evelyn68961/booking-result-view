"""Generate test xlsx files exercising every rule in the manager view.

Settings to use in the manager panel:
  Gate Day = 2026-12-05
  每日通過上限人數 = 2
  單筆預假最少天數 = 4
  單筆預假最多天數 = 10
  個人年度核准次數上限（點數）= 12

The window is 2026-12-05 ~ 2027-06-06 (Sunday of the week that contains Gate+6mo).
The historical data is dense pre-2027-01-09 and empty after, so we use clean
post-2027-01-09 dates plus seeded earlier-in-batch entries to trigger quota.
"""
import io, sys
from datetime import datetime, timedelta
from openpyxl import Workbook
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

HEADERS = ["你的名字", "預假【起日】", "預假【迄日】", "預假天數", "送出時間"]


def fmt(d):
    return d.isoformat() if d else ""


def submit_ts(i):
    return (datetime(2026, 11, 25, 9, 0) + timedelta(minutes=i)).strftime("%Y-%m-%d %H:%M")


def D(y, m, d):
    return datetime(y, m, d).date()


# (Label, Name, Start, End, Category, Expected)
cases = [
    # --- Day-count rule ---
    ("A. 通過 — 5 天、全空檔",
        "測試甲", D(2026, 12, 6),  D(2026, 12, 10), "4-10天", "通過"),
    ("B. 未通過 — 3 天 < 最少 4 天",
        "測試乙", D(2026, 12, 12), D(2026, 12, 14), "1-3天",  "未通過 - 預假天數錯誤"),
    ("C. 未通過 — 11 天 > 最多 10 天",
        "測試丙", D(2026, 12, 16), D(2026, 12, 26), ">10天",  "未通過 - 預假天數錯誤"),
    ("D. 未通過 — 迄日早於起日",
        "測試丁", D(2026, 12, 30), D(2026, 12, 26), "4-10天", "未通過 - 預假天數錯誤"),

    # --- Per-day quota (build up via the batch itself) ---
    ("E1. 通過 — 第 1 人 (該日佔用 1/2)",
        "測試戊一", D(2027, 1, 12), D(2027, 1, 16), "4-10天", "通過"),
    ("E2. 通過 — 第 2 人，相同日期 (該日佔用 2/2)",
        "測試戊二", D(2027, 1, 12), D(2027, 1, 16), "4-10天", "通過"),
    ("E3. 未通過 — 第 3 人，相同日期 (已超過上限人數)",
        "測試戊三", D(2027, 1, 12), D(2027, 1, 16), "4-10天", "未通過 - 已超過上限人數"),
    ("E4. 未通過 — 與 E1/E2 部分重疊 (重疊那幾天已 2/2)",
        "測試戊四", D(2027, 1, 14), D(2027, 1, 19), "4-10天", "未通過 - 已超過上限人數"),

    # --- Yearly 12-point cap (測試庚 submits 13 successful applications in 2027) ---
    ("F01. 通過 — 測試庚 第 1 次 (累計 1/12)",
        "測試庚", D(2027, 2, 1),  D(2027, 2, 4),  "4-10天", "通過"),
    ("F02. 通過 — 測試庚 第 2 次 (累計 2/12)",
        "測試庚", D(2027, 2, 8),  D(2027, 2, 11), "4-10天", "通過"),
    ("F03. 通過 — 測試庚 第 3 次 (累計 3/12)",
        "測試庚", D(2027, 2, 15), D(2027, 2, 18), "4-10天", "通過"),
    ("F04. 通過 — 測試庚 第 4 次 (累計 4/12)",
        "測試庚", D(2027, 2, 22), D(2027, 2, 25), "4-10天", "通過"),
    ("F05. 通過 — 測試庚 第 5 次 (累計 5/12)",
        "測試庚", D(2027, 3, 1),  D(2027, 3, 4),  "4-10天", "通過"),
    ("F06. 通過 — 測試庚 第 6 次 (累計 6/12)",
        "測試庚", D(2027, 3, 8),  D(2027, 3, 11), "4-10天", "通過"),
    ("F07. 通過 — 測試庚 第 7 次 (累計 7/12)",
        "測試庚", D(2027, 3, 15), D(2027, 3, 18), "4-10天", "通過"),
    ("F08. 通過 — 測試庚 第 8 次 (累計 8/12)",
        "測試庚", D(2027, 3, 22), D(2027, 3, 25), "4-10天", "通過"),
    ("F09. 通過 — 測試庚 第 9 次 (累計 9/12)",
        "測試庚", D(2027, 3, 29), D(2027, 4, 1),  "4-10天", "通過"),
    ("F10. 通過 — 測試庚 第 10 次 (累計 10/12)",
        "測試庚", D(2027, 4, 5),  D(2027, 4, 8),  "4-10天", "通過"),
    ("F11. 通過 — 測試庚 第 11 次 (累計 11/12)",
        "測試庚", D(2027, 4, 12), D(2027, 4, 15), "4-10天", "通過"),
    ("F12. 通過 — 測試庚 第 12 次 (累計 12/12，剛好用完)",
        "測試庚", D(2027, 4, 19), D(2027, 4, 22), "4-10天", "通過"),
    ("F13. 未通過 — 測試庚 第 13 次 (年度點數不足)",
        "測試庚", D(2027, 4, 26), D(2027, 4, 29), "4-10天", "未通過 - 年度點數不足"),

    # --- Bookable window ---
    ("G. 未通過 — 起日早於 Gate Day (2026-12-05)",
        "測試辛", D(2026, 11, 25), D(2026, 11, 29), "4-10天", "未通過 - 超出可預約範圍"),
    ("H. 未通過 — 迄日晚於本輪結束 (2027-06-06)",
        "測試壬", D(2027, 6, 4),  D(2027, 6, 10), "4-10天", "未通過 - 超出可預約範圍"),
]


def write_book(path, rows_):
    wb = Workbook()
    sh = wb.active
    sh.title = "新申請"
    sh.append(HEADERS)
    for i, (_label, name, s, e, cat, _exp) in enumerate(rows_):
        sh.append([name, fmt(s), fmt(e), cat, submit_ts(i)])
    doc = wb.create_sheet("測試說明")
    doc.append(["#", "情境", "姓名", "起日", "迄日", "天數", "預期結果"])
    for i, (label, name, s, e, _cat, expected) in enumerate(rows_, 1):
        days = (e - s).days + 1 if (s and e) else ""
        doc.append([i, label, name, fmt(s), fmt(e), days, expected])
    wb.save(path)
    print(f"Wrote {path} ({len(rows_)} rows)")


# A second small file with only the rows expected to pass on a clean state.
happy = [c for c in cases if c[5] == "通過"]
write_book("test-mixed.xlsx", cases)
write_book("test-pass-only.xlsx", happy)

print()
print("=" * 64)
print("Set the manager panel to:")
print("  Gate Day = 2026-12-05")
print("  Quota=2  Min=4  Max=10  YearlyPoints=12")
print("=" * 64)
print()
print("Cases (in order — predictor processes top-to-bottom):")
for i, (label, name, s, e, cat, exp) in enumerate(cases, 1):
    days = (e - s).days + 1
    print(f"  {i:2d}. [{exp:<24s}] {label}")
    print(f"        {name:<8s} {fmt(s)} → {fmt(e)} ({days} 天, 類別 {cat})")
