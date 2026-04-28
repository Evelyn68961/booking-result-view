# 預假審核規則

This file is the source-of-truth for the approval rules applied by 預假紀錄管理系統 (`index.html`). It is **not** the spec for a backend booking system — it only documents the validation rules the manager tab evaluates against uploaded申請.

---

## Business Rules

- Bookable window: Gate Day → first Sunday of the month after (Gate Day + 7 months), inclusive. If Gate Day is left blank in the UI, the range check is skipped.
- Each submission: 1 consecutive block of **4–10 days** (configurable: `minDays` / `maxDays`).
- Multiple blocks per person allowed.
- Max **2 people per day** per calendar date, counted across all approved blocks (configurable: `quota`).
- **每人每年 12 點**：each approved submission consumes 1 point regardless of length, counted in the calendar year of the booking's **start date**. When a person's points for that year reach 12, further submissions for the same year are rejected (configurable: `yearlyPoints`).

### Reject reasons (zh-TW, surfaced verbatim in the預測欄)

- `預假天數錯誤` — day count outside `minDays`–`maxDays`, or 迄日早於起日
- `已超過上限人數` — at least one date in the requested range is already at the daily quota of approved bookings
- `年度點數不足` — submitter has already used all yearly points for that calendar year
- `超出可預約範圍` — start before Gate Day, or end after the round-window Sunday

### Points-vs-days clarification

Yearly points are counted per *approved submission*, not per day. A 4-day approval and a 10-day approval each consume exactly 1 of the year's 12 points. The bucket is the calendar year of `start_date`, not the submission round — a January booking submitted in the prior December still consumes a point from the January year.

### Priority

Within a single contention window, priority follows submission order (`送出時間` / server timestamp). The manager tab does not re-order — it evaluates申請 in the order they appear in the uploaded sheet, so the existing approved set + the申請 above the current row act as the baseline for capacity checks.
