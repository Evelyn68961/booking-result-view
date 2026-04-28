"""Generate one xlsx per month, simulating monthly申請 batches the manager
processes via the 新申請審核 tab.

Manager settings to use throughout:
  Gate Day = 2026-12-05
  每日通過上限人數 = 2
  單筆預假最少天數 = 4
  單筆預假最多天數 = 10
  個人年度核准次數上限（點數）= 12

Window: 2026-12-05 ~ 2027-08-01 (first Sunday of the month after Gate+7mo).

Process the files in chronological order (filename order). After each batch,
commit 通過 rows so they accumulate as historical baseline for the next month.
Each file's first sheet is the upload payload; the second sheet documents the
expected verdicts for that month.
"""
import io, sys
from datetime import date, datetime, timedelta
from openpyxl import Workbook

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

HEADERS = ["你的名字", "預假【起日】", "預假【迄日】", "預假天數", "送出時間"]


def D(y, m, d):
    return date(y, m, d)


def fmt(d):
    return d.isoformat()


def days(s, e):
    return (e - s).days + 1


def category(s, e):
    n = days(s, e)
    if n < 4:
        return "1-3天"
    if n > 10:
        return ">10天"
    return "4-10天"


# Each batch is keyed by (year, month) of submission. Submissions in a batch
# are timestamped sequentially within that month so the manager tab evaluates
# them top-to-bottom (priority = submission order).
#
# Row schema: (label, name, start_date, end_date, expected_verdict)
#   expected_verdict assumes batches are processed in order and 通過 rows from
#   earlier batches are committed before the next batch is uploaded.

batches = {
    # First batch — clean state, exercises gate-day boundary.
    (2026, 11): [
        ("A1. 通過 — 5 天，視窗開頭",          "測試甲", D(2026, 12,  7), D(2026, 12, 11), "通過"),
        ("A2. 未通過 — 起日早於 Gate Day",     "測試乙", D(2026, 11, 25), D(2026, 11, 29), "未通過 - 超出可預約範圍"),
        ("A3. 通過 — 5 天，月中",              "測試丙", D(2026, 12, 14), D(2026, 12, 18), "通過"),
    ],

    # Day-count edge cases.
    (2026, 12): [
        ("B1. 通過 — 4 天剛好下限",            "測試丁", D(2026, 12, 28), D(2026, 12, 31), "通過"),
        ("B2. 未通過 — 11 天 > 上限 10",       "測試戊", D(2027,  1,  5), D(2027,  1, 15), "未通過 - 預假天數錯誤"),
        ("B3. 未通過 — 3 天 < 下限 4",         "測試己", D(2027,  1, 18), D(2027,  1, 20), "未通過 - 預假天數錯誤"),
        ("B4. 未通過 — 迄日早於起日",          "測試庚", D(2027,  1, 25), D(2027,  1, 21), "未通過 - 預假天數錯誤"),
    ],

    # Per-day quota — three people on the same window, fourth partial overlap.
    (2027, 1): [
        ("C1. 通過 — 第 1 人 (該日佔 1/2)",    "測試辛", D(2027,  2,  8), D(2027,  2, 12), "通過"),
        ("C2. 通過 — 第 2 人，相同日期 (2/2)", "測試壬", D(2027,  2,  8), D(2027,  2, 12), "通過"),
        ("C3. 未通過 — 第 3 人，相同日期",      "測試癸", D(2027,  2,  8), D(2027,  2, 12), "未通過 - 已超過上限人數"),
        ("C4. 未通過 — 與 C1/C2 部分重疊",      "測試地", D(2027,  2, 10), D(2027,  2, 15), "未通過 - 已超過上限人數"),
        ("C5. 通過 — 不重疊的隔週",            "測試玄", D(2027,  2, 15), D(2027,  2, 19), "通過"),
    ],

    # 測試年 starts a yearly-12-points marathon.
    (2027, 2): [
        ("D1. 通過 — 測試年 1/12",             "測試年", D(2027,  3,  1), D(2027,  3,  4), "通過"),
        ("D2. 通過 — 測試年 2/12",             "測試年", D(2027,  3,  8), D(2027,  3, 11), "通過"),
        ("D3. 通過 — 測試年 3/12",             "測試年", D(2027,  3, 15), D(2027,  3, 18), "通過"),
        ("D4. 通過 — 測試年 4/12",             "測試年", D(2027,  3, 22), D(2027,  3, 25), "通過"),
    ],

    (2027, 3): [
        ("E1. 通過 — 測試年 5/12",             "測試年", D(2027,  4,  5), D(2027,  4,  8), "通過"),
        ("E2. 通過 — 測試年 6/12",             "測試年", D(2027,  4, 12), D(2027,  4, 15), "通過"),
        ("E3. 通過 — 測試年 7/12",             "測試年", D(2027,  4, 19), D(2027,  4, 22), "通過"),
        ("E4. 通過 — 測試年 8/12",             "測試年", D(2027,  4, 26), D(2027,  4, 29), "通過"),
    ],

    # Fill remaining points; trip yearly cap; trip end-of-window.
    (2027, 4): [
        ("F1. 通過 — 測試年 9/12",             "測試年", D(2027,  5,  3), D(2027,  5,  6), "通過"),
        ("F2. 通過 — 測試年 10/12",            "測試年", D(2027,  5, 10), D(2027,  5, 13), "通過"),
        ("F3. 通過 — 測試年 11/12",            "測試年", D(2027,  5, 17), D(2027,  5, 20), "通過"),
        ("F4. 通過 — 測試年 12/12 剛好用完",    "測試年", D(2027,  5, 24), D(2027,  5, 27), "通過"),
        ("F5. 未通過 — 測試年 13 (年度點數不足)", "測試年", D(2027,  5, 31), D(2027,  6,  3), "未通過 - 年度點數不足"),
        ("F6. 未通過 — 迄日晚於本輪結束",        "測試宙", D(2027,  8,  2), D(2027,  8,  6), "未通過 - 超出可預約範圍"),
    ],
}


def submit_ts(year, month, idx):
    # Spread submissions across the month, preserving order.
    base = datetime(year, month, 5, 9, 0)
    return (base + timedelta(hours=idx)).strftime("%Y-%m-%d %H:%M")


def write_batch(year, month, rows):
    wb = Workbook()
    sh = wb.active
    sh.title = "新申請"
    sh.append(HEADERS)
    for i, (_label, name, s, e, _exp) in enumerate(rows):
        sh.append([name, fmt(s), fmt(e), category(s, e), submit_ts(year, month, i)])

    doc = wb.create_sheet("測試說明")
    doc.append(["#", "情境", "姓名", "起日", "迄日", "天數", "類別", "預期結果"])
    for i, (label, name, s, e, exp) in enumerate(rows, 1):
        doc.append([i, label, name, fmt(s), fmt(e), days(s, e), category(s, e), exp])

    path = f"batch-{year:04d}-{month:02d}.xlsx"
    wb.save(path)
    return path


written = []
for (y, m), rows in batches.items():
    written.append((y, m, write_batch(y, m, rows), rows))

print()
print("=" * 64)
print("Manager panel settings:")
print("  Gate Day = 2026-12-05   Quota=2   Min=4   Max=10   YearlyPoints=12")
print("=" * 64)
print()
print("Process these in order, committing 通過 rows between batches:")
print()
for y, m, path, rows in written:
    pass_n = sum(1 for r in rows if r[4] == "通過")
    fail_n = len(rows) - pass_n
    print(f"  {path}  ({len(rows)} rows: {pass_n} 通過 / {fail_n} 未通過)")
    for i, (label, name, s, e, exp) in enumerate(rows, 1):
        print(f"    {i}. [{exp:<24s}] {label}")
        print(f"         {name}  {fmt(s)} → {fmt(e)}  ({days(s, e)} 天)")
    print()
