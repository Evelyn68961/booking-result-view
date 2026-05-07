"""Generate one xlsx per month, simulating monthly申請 batches the manager
processes via the 新申請審核 tab.

Manager settings to use throughout:
  Gate Day = 2026-05-02
  平日上限（週一~週五）= 2
  假日上限（週六、週日）= 4
  單筆預假最少天數 = 4
  單筆預假最多天數 = 10
  個人年度核准次數上限（點數）= 12
  上限例外 (overrides) = empty unless a batch's 測試說明 says otherwise

Window: 2026-05-02 ~ 2027-01-03 (first Sunday of the month after Gate+7mo).

Process the files in chronological order (filename order). After each batch,
commit 通過 rows so they accumulate as historical baseline for the next month.
Each file's first sheet is the upload payload; the second sheet documents the
expected verdicts for that month.

Most batches' rows are timestamped sequentially within their month so row
order == 送出時間 order == evaluation priority. The 2026-11 batch is the
exception — it deliberately ships rows in reverse 送出時間 order to verify
that recomputeBatch() sorts by submittedAt before evaluating.
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
# them top-to-bottom (priority = submission order). One batch (2026-11)
# overrides this default — see the inline note for that month.
#
# Row schema: (label, name, start_date, end_date, expected_verdict[, ts_override])
#   expected_verdict assumes batches are processed in order and 通過 rows from
#   earlier batches are committed before the next batch is uploaded.
#   ts_override (optional) is a "%Y-%m-%d %H:%M" string that bypasses the
#   default sequential-by-row submit_ts() — used by the 2026-11 priority test.

batches = {
    # First batch — clean state, exercises gate-day boundary.
    (2026, 4): [
        ("A1. 通過 — 5 天，視窗開頭",          "測試甲", D(2026,  5,  4), D(2026,  5,  8), "通過"),
        ("A2. 未通過 — 起日早於 Gate Day",     "測試乙", D(2026,  4, 25), D(2026,  4, 29), "未通過 - 超出可預約範圍"),
        ("A3. 通過 — 5 天，月中",              "測試丙", D(2026,  5, 11), D(2026,  5, 15), "通過"),
    ],

    # Day-count edge cases.
    (2026, 5): [
        ("B1. 通過 — 4 天剛好下限",            "測試丁", D(2026,  5, 25), D(2026,  5, 28), "通過"),
        ("B2. 未通過 — 11 天 > 上限 10",       "測試戊", D(2026,  6,  1), D(2026,  6, 11), "未通過 - 預假天數錯誤"),
        ("B3. 未通過 — 3 天 < 下限 4",         "測試己", D(2026,  6, 15), D(2026,  6, 17), "未通過 - 預假天數錯誤"),
        ("B4. 未通過 — 迄日早於起日",          "測試庚", D(2026,  6, 22), D(2026,  6, 18), "未通過 - 預假天數錯誤"),
    ],

    # Per-day quota — three people on the same window, fourth partial overlap.
    # All dates here are Mon-Fri, so 平日上限=2 binds.
    (2026, 6): [
        ("C1. 通過 — 第 1 人 (該日佔 1/2)",    "測試辛", D(2026,  7,  6), D(2026,  7, 10), "通過"),
        ("C2. 通過 — 第 2 人，相同日期 (2/2)", "測試壬", D(2026,  7,  6), D(2026,  7, 10), "通過"),
        ("C3. 未通過 — 第 3 人，相同日期",      "測試癸", D(2026,  7,  6), D(2026,  7, 10), "未通過 - 已超過上限人數"),
        ("C4. 未通過 — 與 C1/C2 部分重疊",      "測試地", D(2026,  7,  8), D(2026,  7, 13), "未通過 - 已超過上限人數"),
        ("C5. 通過 — 不重疊的隔週",            "測試玄", D(2026,  7, 13), D(2026,  7, 17), "通過"),
    ],

    # 測試年 starts a yearly-12-points marathon (all 2026 starts).
    (2026, 7): [
        ("D1. 通過 — 測試年 1/12",             "測試年", D(2026,  8,  3), D(2026,  8,  6), "通過"),
        ("D2. 通過 — 測試年 2/12",             "測試年", D(2026,  8, 10), D(2026,  8, 13), "通過"),
        ("D3. 通過 — 測試年 3/12",             "測試年", D(2026,  8, 17), D(2026,  8, 20), "通過"),
        ("D4. 通過 — 測試年 4/12",             "測試年", D(2026,  8, 24), D(2026,  8, 27), "通過"),
    ],

    (2026, 8): [
        ("E1. 通過 — 測試年 5/12",             "測試年", D(2026,  9,  7), D(2026,  9, 10), "通過"),
        ("E2. 通過 — 測試年 6/12",             "測試年", D(2026,  9, 14), D(2026,  9, 17), "通過"),
        ("E3. 通過 — 測試年 7/12",             "測試年", D(2026,  9, 21), D(2026,  9, 24), "通過"),
        ("E4. 通過 — 測試年 8/12",             "測試年", D(2026,  9, 28), D(2026, 10,  1), "通過"),
    ],

    # Fill remaining points; trip yearly cap.
    (2026, 9): [
        ("F1. 通過 — 測試年 9/12",             "測試年", D(2026, 10, 12), D(2026, 10, 15), "通過"),
        ("F2. 通過 — 測試年 10/12",            "測試年", D(2026, 10, 19), D(2026, 10, 22), "通過"),
        ("F3. 通過 — 測試年 11/12",            "測試年", D(2026, 10, 26), D(2026, 10, 29), "通過"),
        ("F4. 通過 — 測試年 12/12 剛好用完",    "測試年", D(2026, 11,  2), D(2026, 11,  5), "通過"),
        ("F5. 未通過 — 測試年 13 (年度點數不足)", "測試年", D(2026, 11,  9), D(2026, 11, 12), "未通過 - 年度點數不足"),
    ],

    # December rush — most pharmacists aim for year-end. Tests quota contention
    # on prime weeks, Christmas overlap, and the 2027-01-03 window boundary.
    # All G-rows are Mon-Fri, so 平日上限=2 binds (same as old single-quota=2 behaviour).
    (2026, 10): [
        ("G1. 通過 — 12月第1週 第1人 (1/2)",    "測試卯", D(2026, 12,  7), D(2026, 12, 11), "通過"),
        ("G2. 通過 — 12月第1週 第2人 (2/2)",    "測試辰", D(2026, 12,  7), D(2026, 12, 11), "通過"),
        ("G3. 未通過 — 12月第1週 第3人 (額滿)",  "測試巳", D(2026, 12,  7), D(2026, 12, 11), "未通過 - 已超過上限人數"),
        ("G4. 通過 — 聖誕週 第1人 (1/2)",       "測試午", D(2026, 12, 21), D(2026, 12, 25), "通過"),
        ("G5. 通過 — 聖誕週 第2人 (2/2)",       "測試未", D(2026, 12, 21), D(2026, 12, 25), "通過"),
        ("G6. 未通過 — 與聖誕週部分重疊 (額滿)",  "測試申", D(2026, 12, 23), D(2026, 12, 27), "未通過 - 已超過上限人數"),
        ("G7. 通過 — 跨年至視窗末端 2027-01-03", "測試酉", D(2026, 12, 30), D(2027,  1,  3), "通過"),
        ("G8. 未通過 — 迄日 2027-01-04 超出視窗", "測試戌", D(2026, 12, 31), D(2027,  1,  4), "未通過 - 超出可預約範圍"),
    ],

    # Priority-by-送出時間 — three people contend for a Mon-Thu (cap=2) window.
    # Rows ship in REVERSE submission order (row 1 = latest, row 3 = earliest)
    # so the verdicts only come out right if recomputeBatch() sorts by
    # submittedAt before evaluating. Under a broken sort that just walked
    # state.batch in array order, H1 would pass and H3 would fail.
    (2026, 11): [
        ("H1. 未通過 — 最晚送出，輸掉名額",     "測試零", D(2026, 11, 23), D(2026, 11, 26), "未通過 - 已超過上限人數",
         "2026-11-25 18:00"),
        ("H2. 通過 — 次早送出，拿到第 2 名額",   "測試壹", D(2026, 11, 23), D(2026, 11, 26), "通過",
         "2026-11-25 12:00"),
        ("H3. 通過 — 最早送出，拿到第 1 名額",   "測試貳", D(2026, 11, 23), D(2026, 11, 26), "通過",
         "2026-11-25 06:00"),
    ],

    # Weekend cap=4 — four bookings all overlap on Sat 2026-12-12.
    # Two come in via Wed-Sat (12/9-12/12), two via Sat-Tue (12/12-12/15), so
    # weekday slots don't compete (Wed-Fri vs Mon-Tue). Saturday accumulates
    # 1→2→3→4, all under 假日上限=4. Under the old single-quota=2 behaviour
    # the 3rd and 4th (I3, I4) would have failed at Sat=3/2.
    #
    # Manual override walkthrough (after this batch is processed):
    #   1. Open 上限例外, add 2026-12-12 ~ 2026-12-12 with 上限人數=2.
    #   2. Clear the batch and re-upload this xlsx.
    #   3. Expect I3 and I4 to flip from 通過 to 未通過 - 已超過上限人數
    #      (Sat would be 3/2 then 4/2 with the override active).
    (2026, 12): [
        ("I1. 通過 — Wed-Sat 第 1 人 (Sat 1/4)",  "測試三", D(2026, 12,  9), D(2026, 12, 12), "通過"),
        ("I2. 通過 — Wed-Sat 第 2 人 (Sat 2/4)",  "測試四", D(2026, 12,  9), D(2026, 12, 12), "通過"),
        ("I3. 通過 — Sat-Tue 第 3 人 (Sat 3/4)",  "測試五", D(2026, 12, 12), D(2026, 12, 15), "通過"),
        ("I4. 通過 — Sat-Tue 第 4 人 (Sat 4/4)",  "測試六", D(2026, 12, 12), D(2026, 12, 15), "通過"),
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
    for i, row in enumerate(rows):
        name, s, e = row[1], row[2], row[3]
        ts = row[5] if len(row) > 5 else submit_ts(year, month, i)
        sh.append([name, fmt(s), fmt(e), category(s, e), ts])

    doc = wb.create_sheet("測試說明")
    doc.append(["#", "情境", "姓名", "起日", "迄日", "天數", "類別", "送出時間", "預期結果"])
    for i, row in enumerate(rows, 1):
        label, name, s, e, exp = row[0], row[1], row[2], row[3], row[4]
        ts = row[5] if len(row) > 5 else submit_ts(year, month, i - 1)
        doc.append([i, label, name, fmt(s), fmt(e), days(s, e), category(s, e), ts, exp])

    path = f"batch-{year:04d}-{month:02d}.xlsx"
    wb.save(path)
    return path


written = []
for (y, m), rows in batches.items():
    written.append((y, m, write_batch(y, m, rows), rows))

print()
print("=" * 64)
print("Manager panel settings:")
print("  Gate Day = 2026-05-02   平日=2 / 假日=4   Min=4   Max=10   Year=12")
print("  上限例外: empty (one batch documents an optional override walkthrough)")
print("=" * 64)
print()
print("Process these in order, committing 通過 rows between batches:")
print()
for y, m, path, rows in written:
    pass_n = sum(1 for r in rows if r[4] == "通過")
    fail_n = len(rows) - pass_n
    print(f"  {path}  ({len(rows)} rows: {pass_n} 通過 / {fail_n} 未通過)")
    for i, row in enumerate(rows, 1):
        label, name, s, e, exp = row[0], row[1], row[2], row[3], row[4]
        ts = row[5] if len(row) > 5 else submit_ts(y, m, i - 1)
        print(f"    {i}. [{exp:<24s}] {label}")
        print(f"         {name}  {fmt(s)} → {fmt(e)}  ({days(s, e)} 天)  送出={ts}")
    print()
