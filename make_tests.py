"""Generate a single test batch xlsx covering the rules that can be exercised
in December 2026 against the real baked history (`202401-202604預假紀錄.xlsx`).

The baked file already contains 通過 records through 2026-12-05 — most days in
May–November are at or near the daily cap, so test scenarios that book those
months don't actually pass. December 6 onwards is clean, so this batch packs
all the testable rules into 12/06–2027/01/04.

Manager settings to use (defaults — do NOT tweak):
  Gate Day = 2026-05-02
  平日上限（週一~週五）= 2
  假日上限（週六、週日）= 4
  單筆預假最少天數 = 4
  單筆預假最多天數 = 10
  個人年度核准次數上限（點數）= 12
  上限例外 = empty (one optional walkthrough at the bottom of 測試說明)

Window: 2026-05-02 ~ 2027-01-03 (first Sunday after Gate + 7 months).

The yearly-points rule (每人每年 12 點) is enforced by the code but is NOT
tested by an xlsx scenario in this batch — packing 12 non-overlapping 4-day
bookings for one person doesn't fit into the clean December window. Verify
that rule manually if needed; it is documented in booking_rules.md.
"""
import io, sys
from datetime import date
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


# Row schema: (label, name, start_date, end_date, expected_verdict, submitted_at)
# The submit timestamps are explicit so the priority test (P1-P3) can ship
# rows in REVERSE submission order. Other rows are sequenced earlier so they
# evaluate first; only the relative submit-time order matters, not the values.
ROWS = [
    # --- Gate-day boundary -------------------------------------------------
    ("A1. 未通過 — 起日早於 Gate Day（2026-05-02）",
     "測試甲", D(2026, 4, 25), D(2026, 4, 29),
     "未通過 - 超出可預約範圍",   "2026-12-01 09:00"),

    # --- Day-count edges ---------------------------------------------------
    ("B1. 通過 — 4 天剛好下限",
     "測試乙", D(2026, 12,  8), D(2026, 12, 11),
     "通過",                       "2026-12-01 09:10"),
    ("B2. 未通過 — 11 天 > 上限 10",
     "測試丙", D(2026, 12,  8), D(2026, 12, 18),
     "未通過 - 預假天數錯誤",       "2026-12-01 09:20"),
    ("B3. 未通過 — 3 天 < 下限 4",
     "測試丁", D(2026, 12,  8), D(2026, 12, 10),
     "未通過 - 預假天數錯誤",       "2026-12-01 09:30"),
    ("B4. 未通過 — 迄日早於起日",
     "測試戊", D(2026, 12, 12), D(2026, 12,  8),
     "未通過 - 預假天數錯誤",       "2026-12-01 09:40"),

    # --- Weekday quota (cap=2 binds at 3rd) -------------------------------
    ("C1. 通過 — 12/14-12/17 第 1 人 (Mon-Thu, 1/2)",
     "測試己", D(2026, 12, 14), D(2026, 12, 17),
     "通過",                       "2026-12-02 09:00"),
    ("C2. 通過 — 12/14-12/17 第 2 人 (2/2)",
     "測試庚", D(2026, 12, 14), D(2026, 12, 17),
     "通過",                       "2026-12-02 09:10"),
    ("C3. 未通過 — 12/14-12/17 第 3 人 (額滿)",
     "測試辛", D(2026, 12, 14), D(2026, 12, 17),
     "未通過 - 已超過上限人數",      "2026-12-02 09:20"),

    # --- Priority by 送出時間 (rows in REVERSE submit order) --------------
    # Three people contend for 12/22-12/25 (Tue-Fri, clean). The row at the
    # top has the LATEST 送出時間, so if recomputeBatch() really sorts by
    # submittedAt before evaluating, P3 wins slot 1 and P2 slot 2; P1 fails.
    # Under a broken sort that walked array order, P1 would pass.
    ("P1. 未通過 — 最晚送出，輸掉名額",
     "測試壬", D(2026, 12, 22), D(2026, 12, 25),
     "未通過 - 已超過上限人數",      "2026-12-15 18:00"),
    ("P2. 通過 — 次早送出，拿到第 2 名額",
     "測試癸", D(2026, 12, 22), D(2026, 12, 25),
     "通過",                       "2026-12-15 12:00"),
    ("P3. 通過 — 最早送出，拿到第 1 名額",
     "測試子", D(2026, 12, 22), D(2026, 12, 25),
     "通過",                       "2026-12-15 06:00"),

    # --- Weekend cap (Sat=4, Sun=4 vs. weekday=2) -------------------------
    # Four people all book Sat-Tue 12/26-12/29 (clean). The 3rd and 4th fail
    # because Mon 12/28 / Tue 12/29 hit weekday cap=2 — but their 已滿日 line
    # only lists Mon and Tue, NOT Sat/Sun. Under the previous single-quota=2,
    # Sat 12/26 would have appeared in the list at 3/2.
    ("W1. 通過 — Sat-Tue 第 1 人 (Sat 1/4)",
     "測試丑", D(2026, 12, 26), D(2026, 12, 29),
     "通過",                       "2026-12-20 09:00"),
    ("W2. 通過 — Sat-Tue 第 2 人 (Sat 2/4)",
     "測試寅", D(2026, 12, 26), D(2026, 12, 29),
     "通過",                       "2026-12-20 09:10"),
    ("W3. 未通過 — Sat 3/4 ✓ 但 Mon-Tue 3/2 ✗",
     "測試卯", D(2026, 12, 26), D(2026, 12, 29),
     "未通過 - 已超過上限人數",      "2026-12-20 09:20"),
    ("W4. 未通過 — 同 W3 (週末未滿，平日已滿)",
     "測試辰", D(2026, 12, 26), D(2026, 12, 29),
     "未通過 - 已超過上限人數",      "2026-12-20 09:30"),

    # --- Window-end boundary (window ends 2027-01-03 inclusive) -----------
    ("E1. 通過 — 跨年至視窗末端 2027-01-03",
     "測試巳", D(2026, 12, 30), D(2027,  1,  3),
     "通過",                       "2026-12-25 09:00"),
    ("E2. 未通過 — 迄日 2027-01-04 超出視窗",
     "測試午", D(2026, 12, 31), D(2027,  1,  4),
     "未通過 - 超出可預約範圍",      "2026-12-25 09:10"),
]


def write_batch():
    wb = Workbook()
    sh = wb.active
    sh.title = "新申請"
    sh.append(HEADERS)
    for _label, name, s, e, _exp, ts in ROWS:
        sh.append([name, fmt(s), fmt(e), category(s, e), ts])

    doc = wb.create_sheet("測試說明")
    doc.append(["#", "情境", "姓名", "起日", "迄日", "天數", "類別", "送出時間", "預期結果"])
    for i, (label, name, s, e, exp, ts) in enumerate(ROWS, 1):
        doc.append([i, label, name, fmt(s), fmt(e), days(s, e), category(s, e), ts, exp])

    # Optional override walkthrough — separate block at the bottom.
    doc.append([])
    doc.append(["", "選用：上限例外 walkthrough"])
    doc.append(["", "1. 在「上限例外」面板新增 2026-12-26 ~ 2026-12-26，上限人數 = 1。"])
    doc.append(["", "2. 清空批次，重新上傳本檔。"])
    doc.append(["", "3. 預期：W1 仍 通過 (Sat 1/1)；W2/W3/W4 改為 未通過 - 已超過上限人數，"])
    doc.append(["", "        而且已滿日改成只列 12/26 (Sat) 而不是 Mon-Tue。"])

    path = "batch-2026-12.xlsx"
    wb.save(path)
    return path


path = write_batch()

print()
print("=" * 64)
print("Manager panel settings (do NOT tweak):")
print("  Gate Day = 2026-05-02   平日=2 / 假日=4   Min=4   Max=10   Year=12")
print("  上限例外: empty (one optional walkthrough documented in 測試說明)")
print("=" * 64)
print()
print(f"Single batch: {path}  ({len(ROWS)} rows)")
print()
pass_n = sum(1 for r in ROWS if r[4] == "通過")
fail_n = len(ROWS) - pass_n
print(f"  ({pass_n} 通過 / {fail_n} 未通過)")
for i, (label, name, s, e, exp, ts) in enumerate(ROWS, 1):
    print(f"    {i:>2}. [{exp:<24s}] {label}")
    print(f"          {name}  {fmt(s)} → {fmt(e)}  ({days(s, e)} 天)  送出={ts}")
