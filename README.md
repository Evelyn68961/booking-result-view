# 預假紀錄管理系統

A single-file HTML viewer + manager-approval helper for the existing 預假 Excel workflow. Not a replacement for the booking system — a thin tool that bakes the historical xlsx into a static page so the manager can review new申請 against existing approvals without retyping anything.

## What it does

**檢視紀錄 tab** — search, filter, sort, and paginate every row in the baked xlsx (姓名 / 審核結果 / 起日 / 迄日 / 天數 / 送出時間 / 特休 / 時數). Manager-committed rows are marked `手動加入`.

**預約日曆 tab** — month grid showing per-day occupancy of all 通過 entries (baked history + manager-committed rows). Cells are colour-coded — green = 空, yellow = 1 人, red = 2 人以上. Click a day to see who's booked (姓名 / 起日 / 迄日).

**新申請審核 tab (password-protected)** — drop in a new申請 xlsx/csv (or add rows manually), and the page predicts 通過 / 未通過 for each entry by replaying the rules in [booking_rules.md](booking_rules.md) against the baked history + already-approved entries earlier in the batch. Manager can override (強制通過 / 強制未通過), commit decisions to localStorage, or export the batch as xlsx.

Day-by-day occupancy for the touched dates is shown so it's obvious which days are full.

**Per-record edit / delete** — when the manager tab is unlocked, every row in the View tab gets 編輯 / 刪除 buttons. Edits to baked records become a localStorage "patches" overlay (the original xlsx is never mutated). Manager-added records are edited or removed in place. A separate panel in the manager tab lists the manual additions for inline tweaking, plus a bulk-wipe button.

**Optional Google Sheet sync** — the manager tab has a 🔗 Google Sheet 同步 panel. Pasting an Apps Script Web App URL ([apps-script.gs](apps-script.gs) is the backend) gives you a shared sheet your managers can view/edit directly. The app auto-pushes local changes every 2 minutes and auto-pulls when the browser tab regains focus; manual 上傳 / 載入 buttons are also provided.

## Files

| File | Purpose |
|---|---|
| [index.html](index.html) | Built output. Open directly in a browser — no server needed (sync requires http(s)). |
| [build.py](build.py) | Reads the xlsx, encrypts the manager block, and writes `index.html`. |
| [booking_rules.md](booking_rules.md) | The approval rules the manager tab enforces. |
| [apps-script.gs](apps-script.gs) | Google Apps Script backend for the optional Sheet sync. Paste into the script editor of your shared sheet. |
| [spec.md](spec.md) | Architecture / data model / sync semantics — read this when changing internals. |
| [make_tests.py](make_tests.py) | Generates `batch-YYYY-MM.xlsx` files — one per month — that simulate the申請 batches the manager would receive. |
| `202401-202604預假紀錄.xlsx` | Source data baked into the page. |
| `batch-*.xlsx` | Monthly申請 batches for trying the manager tab. Process in filename order, committing 通過 rows between batches so history accumulates. |

## Build

```bash
pip install openpyxl cryptography
python build.py
```

This regenerates `index.html` with the current xlsx contents. To swap data sources, change `XLSX` near the top of [build.py](build.py).

## Manager password

The manager tab's UI markup is encrypted at build time with PBKDF2-SHA256 (200k iters) + AES-GCM. Without the password, viewing source only shows an opaque base64 blob. To rotate, edit `MANAGER_PASSWORD` in [build.py](build.py) and rerun.

This is obfuscation of the controls, not protection of the data — the baked records are always visible in the 檢視紀錄 tab.

## Storage

Five localStorage keys (all per-browser, per-device):

| Key | Holds |
|---|---|
| `booking-extra-records-v1` | Manager-added records (post-commit). |
| `booking-batch-v1` | In-progress batch (pre-commit scratch). |
| `booking-baked-patches-v1` | Edits / tombstones for the baked xlsx records. |
| `booking-sheet-url-v1` | Apps Script Web App URL (if Sheet sync configured). |
| `booking-sheet-last-sync-v1` | ISO timestamp of last successful sync. |

The 「移除所有已儲存的新增紀錄」button clears `booking-extra-records-v1`; the original baked xlsx is never modified. To survive a browser-cache wipe or to share state across devices, configure the Google Sheet sync — see [spec.md](spec.md) and [apps-script.gs](apps-script.gs).

## Test scenarios

Run `python make_tests.py` to regenerate the monthly batch files
(`batch-2026-04.xlsx` … `batch-2026-10.xlsx`). Recommended manager-tab settings:

- Gate Day = `2026-05-02` → bookable window 2026-05-02 ~ 2027-01-03
- 每日上限 = 2 / 單筆 4–10 天 / 年度 12 點

Process the batches in filename order. After each batch, commit 通過 rows
before uploading the next month — this lets historical state accumulate the
way it would in real use, so cross-month rules (yearly point cap, day-quota
across earlier approvals) are exercised end-to-end.

What each batch exercises:

| Batch | Coverage |
|---|---|
| `batch-2026-04.xlsx` | Gate-day boundary (start before Gate Day fails) |
| `batch-2026-05.xlsx` | Day-count edges (4 days OK, 11 too long, 3 too short, reversed dates) |
| `batch-2026-06.xlsx` | Per-day quota (3rd person on same dates fails, partial overlap fails) |
| `batch-2026-07.xlsx` … `batch-2026-09.xlsx` | 測試年's 12-points marathon (1/12 → 12/12 → 13/12 fails year cap) |
| `batch-2026-10.xlsx` | **December rush** — quota contention on prime weeks, Christmas overlap, and the 2027-01-03 window-end boundary (G7 ends exactly on the boundary, G8 one day past) |

Each xlsx has the upload payload on the first sheet (`新申請`) and the
expected verdict for every row documented on the second sheet (`測試說明`).
Running `make_tests.py` also prints the same summary to stdout.
