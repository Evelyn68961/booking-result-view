# 預假紀錄管理系統

A single-file HTML viewer + manager-approval helper for the existing 預假 Excel workflow. Not a replacement for the booking system — a thin tool that bakes the historical xlsx into a static page so the manager can review new申請 against existing approvals without retyping anything.

## What it does

**檢視紀錄 tab** — search, filter, sort, and paginate every row in the baked xlsx (姓名 / 審核結果 / 起日 / 迄日 / 天數 / 送出時間 / 特休 / 時數). Manager-committed rows are marked `手動加入`.

**新申請審核 tab (password-protected)** — drop in a new申請 xlsx/csv (or add rows manually), and the page predicts 通過 / 未通過 for each entry by replaying the rules in [booking_rules.md](booking_rules.md) against the baked history + already-approved entries earlier in the batch. Manager can override (強制通過 / 強制未通過), commit decisions to localStorage, or export the batch as xlsx.

Day-by-day occupancy for the touched dates is shown so it's obvious which days are full.

## Files

| File | Purpose |
|---|---|
| [index.html](index.html) | Built output. Open directly in a browser — no server needed. |
| [build.py](build.py) | Reads the xlsx, encrypts the manager block, and writes `index.html`. |
| [booking_rules.md](booking_rules.md) | The approval rules the manager tab enforces. |
| [make_tests.py](make_tests.py) | Generates `test-mixed.xlsx` + `test-pass-only.xlsx` covering every rule. |
| `202401-202604預假紀錄.xlsx` | Source data baked into the page. |
| `test-*.xlsx` | Sample申請 batches for trying the manager tab. |

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

Manager-committed rows and in-progress batches live in `localStorage` (`booking-extra-records-v1`, `booking-batch-v1`). They survive reloads but are per-browser. The 「移除所有已儲存的新增紀錄」button in the manager tab clears them; the original baked xlsx is never modified.

## Test scenarios

Run `python make_tests.py` to regenerate the test files. Recommended manager-tab settings to exercise them:

- Gate Day = `2026-12-05`
- 每日上限 = 2 / 單筆 4–10 天 / 年度 12 點

[make_tests.py](make_tests.py) documents the expected verdict for every row.
