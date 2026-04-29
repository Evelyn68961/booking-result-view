# Specification

Internal reference for the 預假紀錄管理 system. Read after [README.md](README.md) — this assumes you know what the app does and want to understand how it does it.

## Architecture at a glance

Single static HTML file. No server, no framework, no build step at runtime. Three pieces of state, three views, one optional sync layer.

```
                ┌──────────────────────── index.html ────────────────────────┐
                │                                                            │
   xlsx ──build.py──▶  BAKED  (const, embedded JS array, ~600 records)       │
                │       │                                                    │
                │       ▼                                                    │
                │  effectiveBaked()  ◀── booking-baked-patches-v1            │
                │       │                                                    │
                │       ▼                                                    │
                │   allRecords() ──▶  View / Calendar / Batch predictions    │
                │       ▲                                                    │
                │       │                                                    │
                │  loadStored()  ◀── booking-extra-records-v1                │
                │                                                            │
                │  state.batch  ◀── booking-batch-v1   (pre-commit only)     │
                └────────────────────────────────────────────────────────────┘
                                            │
                                            ▼  (optional)
                              ┌────────── Apps Script ─────────┐
                              │   Web App: doGet / doPost      │
                              │   reads/writes "Bookings" tab  │
                              └────────────────────────────────┘
```

Build-time: `build.py` reads the source xlsx, encrypts the manager-tab HTML block with PBKDF2 + AES-GCM, and serializes the data into `index.html`. Runtime: pure browser JS reads `BAKED`, overlays `localStorage` patches, and renders.

## Data model

Three layers, merged in this order:

1. **BAKED** (const) — array of objects coming straight from the source xlsx. Read-only at runtime. Never mutated.
2. **Baked patches** (`booking-baked-patches-v1`) — sparse map keyed by index into `BAKED`, holding partial overrides or `_deleted: true` tombstones.
3. **Manager-stored records** (`booking-extra-records-v1`) — full records the manager added through the batch flow. Each carries a stable `_id` (uuid-ish) backfilled on first load if missing.

The merge is implemented in `effectiveBaked()`:

```js
function effectiveBaked() {
  const patches = loadBakedPatches();
  const out = [];
  for (let i = 0; i < BAKED.length; i++) {
    const p = patches[i];
    if (p && p._deleted) continue;
    out.push(p ? Object.assign({}, BAKED[i], p, { _baked_idx: i })
               :     Object.assign({}, BAKED[i], { _baked_idx: i }));
  }
  return out;
}
function allRecords() { return effectiveBaked().concat(loadStored()); }
```

`allRecords()` is the single source of truth for filtering, sorting, the calendar, the day-quota map, the yearly-points cap, and the batch-prediction logic. Every consumer reads through it.

### Record fields

Schema is whatever the xlsx had, plus a few derived/synthetic fields. The original Chinese keys are kept verbatim so the View tab can map columns 1:1 without a translation layer.

| Field | Source | Notes |
|---|---|---|
| `你的名字` | xlsx | Display name. |
| `審核結果` | xlsx | String, e.g. `通過` or `未通過 - 已超過上限`. Parsed by `classifyStatus()` and `reasonOf()`. |
| `預假【起日】`, `預假【迄日】` | xlsx | Original (often Excel-serial) date values. Not used by render code. |
| `_start_iso`, `_end_iso` | derived in build | ISO strings (`YYYY-MM-DD`). What everything else reads. |
| `預假天數` | xlsx | Category string like `4-10天`, used as-is. |
| `送出時間` | xlsx | Submission timestamp string. |
| `_baked_idx` | runtime | Set by `effectiveBaked()`. Identifies which `BAKED` slot a row came from. |
| `_id` | runtime | Set on stored manager records (backfilled by `loadStored()`). Stable identifier. |
| `_source` | runtime | `'manager'` for stored manager records, absent for baked. |

### Record handles

The Edit / Delete UI and the Google Sheet `id` column use a unified handle string:

| Handle | Meaning | Where it points |
|---|---|---|
| `b:N` | Baked record at index N | `BAKED[N]`, with patches[N] overlaid |
| `m:<id>` | Manager record with `_id` = `<id>` | `loadStored().find(x => x._id === id)` |

`recordHandle(r)` produces the right form; `findRecordByHandle(h)` resolves it back.

### localStorage keys

| Key | Shape | Cleared by |
|---|---|---|
| `booking-extra-records-v1` | `Record[]` (manager rows, full record objects with `_id`) | "移除所有已儲存的新增紀錄" button, manual record delete |
| `booking-batch-v1` | `BatchEntry[]` (uploaded申請 rows pre-commit) | "清空批次" button, batch commit |
| `booking-baked-patches-v1` | `{ [idx]: PartialRecord & { _deleted?: true } }` | Edit-modal-saving the original values back, Pull-from-sheet |
| `booking-sheet-url-v1` | string (Web App `/exec` URL) | Manually clearing the input |
| `booking-sheet-last-sync-v1` | ISO timestamp | Never automatically |

`booking-batch-v1` is **not** synced to Google Sheets — it's pre-decision scratch space, only the post-commit results matter to other managers.

## Manager-tab encryption

The 新申請審核 tab's UI markup is not in the page until the password is provided.

- `build.py` defines `MANAGER_HTML_BLOCK` (the inputs, file dropzone, batch table headers, etc.) and encrypts it via PBKDF2-SHA256 (200k iterations) + AES-GCM, keyed on `MANAGER_PASSWORD`.
- The encrypted blob is embedded as `ENC = {salt, iv, ct, iters}` in the HTML.
- `unlockManager(password)` derives the key in-browser, decrypts, and `innerHTML`'s the result into `#managerContent`.

This is **obfuscation of the manager controls**, not protection of the data. The baked records, patches, and manager-stored records are all visible in the View tab regardless. Treat the password as "don't accidentally let anyone press the approve button," not "data is secret."

To rotate: edit `MANAGER_PASSWORD` in `build.py`, rerun. The encrypted blob in `index.html` regenerates.

`MANAGER_UNLOCKED` (boolean) gates whether:
- The action column appears in the View tab.
- `renderCommittedRecords()` renders.
- `renderSyncPanel()` renders.
- Auto-sync still runs regardless — sync is a function of `getSheetUrl()`, not of the unlock state — so a non-manager visitor still sees the latest data on focus, they just can't edit.

## Editing surfaces

There are **three** ways a record gets edited; they all flow through the same storage:

1. **View-tab modal** (any record, baked or manager). Edit/Delete buttons appear in an extra column when `MANAGER_UNLOCKED`. The modal collects 姓名 / 結果 / 拒絕原因 / 起日 / 迄日 / 類別 and writes through `saveBakedPatches()` or `saveStored()`.
2. **Manager-tab committed-records panel** (manager records only). Inline-editable cells; saves on `change`. Useful for triaging "what did I add?" without the modal.
3. **Manager-tab batch table** (pre-commit only). Inline-editable name/start/end/category for batch entries. Commit moves them into `loadStored()`.

After every mutation, the appropriate `render*` chain runs to refresh the View, Manager panel, and (if visible) the day-grid / batch predictions.

## Sync layer

Optional. If `getSheetUrl()` returns a non-empty Apps Script `/exec` URL, the app two-way-syncs against a Google Sheet.

### Wire format

The sheet has nine columns:

```
id | name | status | start | end | daysCat | submittedAt | source | deleted
```

- `id` is the record handle (`b:N` or `m:<id>`). Treated as opaque by everything except the pull logic.
- `deleted` accepts `TRUE` / `FALSE` / blank. `TRUE` on a `b:N` row becomes a tombstone patch; `TRUE` on an `m:<id>` row drops the record.
- `source` is informational (`baked` or `manager`). Not load-bearing.

### Push (`pushToSheet`)

Idempotent full-overwrite:

1. `buildSheetRecords()` — emit one row per `BAKED[i]` (overlaid with patches) plus one row per stored manager record. Deleted bakeds are emitted with `deleted=TRUE` (so the sheet reflects "this was tombstoned" rather than dropping the row).
2. `POST` to the Web App with `{action: 'overwrite', records}`. `Content-Type: text/plain;charset=utf-8` — important: this keeps the request a "simple" CORS request and avoids the preflight (Apps Script doesn't reply to OPTIONS).
3. Apps Script `doPost` clears the `Bookings` sheet and rewrites it.
4. On success: `setLastSync()`, `clearDirty()`.

### Pull (`pullFromSheet`)

Reconstructs local state from the sheet:

1. `GET` the Web App URL. Apps Script `doGet` returns `{ok: true, records: [...]}`.
2. Iterate rows. For each `b:N`: compare every field to `BAKED[N]`; build a sparse patch with only the differences. For `_deleted` rows, the patch is just `{_deleted: true}`. For each `m:<id>`: build a full manager record (`_source: 'manager'`).
3. `saveBakedPatches(newPatches)` and `saveStored(newStored)` replace local state. Both call `markDirty()` internally; immediately after, `clearDirty()` overrides — pulled state matches the sheet exactly, so it's by definition not dirty.

### Auto-sync state machine

```
                ┌─ user edit ─▶ markDirty() ─▶ state.dirty = true
                │
   any edit ────┤                          ┌─ every 2 min: if dirty, push, clearDirty
                │                          │
                └──── auto-sync timers ────┤
                                           └─ on focus / visibilitychange:
                                              if dirty: push first, clearDirty
                                              then pull, clearDirty
```

A mutex (`autoSyncBusy`) wraps each tick — overlapping pushes/pulls can't interleave. Errors are logged, the dirty flag stays set so the next tick retries.

`startAutoSync()` is called once at init. It:
- Sets the 2-minute interval (`setInterval`).
- Wires `visibilitychange` and `window.focus` to a focus-style tick.
- Fires one bootstrap pull 200ms after init if a URL is configured.

### Conflict semantics

Last-writer-wins. There's no row-level versioning. The push-first-if-dirty rule on focus handles the common case ("I edited locally, switched to the sheet tab, edited there, came back") by ensuring the local edit lands in the sheet before the pull overwrites local state — but if the same row was edited in both places, the last push wins. For ≤3-manager monthly cadence this is fine; it would not be fine for hot multi-user editing.

`source` and `id` columns are not versioned and not protected at the API layer — protect them in the Sheet UI via Data → Protect range if you want a hard guard against accidental edits.

## Apps Script

[apps-script.gs](apps-script.gs) is the entire backend. Two endpoints:

- `doGet(e)` — returns `{ok: true, records: [...]}`. Date cells are formatted as ISO strings before serialization to keep round-trips lossless.
- `doPost(e)` — only `{action: 'overwrite', records}` is implemented. Clears the `Bookings` sheet and rewrites it. Other actions return `{ok: false, error: ...}`.

The sheet is auto-created on first push. Header row is bold and frozen. Data cells are formatted as plain text (`@`) to prevent Sheets from auto-converting `2026-12-01` into a Date object.

Deployment must be **Web app**, **Execute as: Me**, **Who has access: Anyone** — anything stricter triggers a login redirect that strips CORS headers and produces "blocked by CORS policy: No 'Access-Control-Allow-Origin' header" in the browser console. This is an Apps Script behavior, not something the client can fix.

## State and mutations summary

| Trigger | What changes | What re-renders |
|---|---|---|
| Filter / sort / paginate | `state.q/name/status/from/to/sortKey/sortDir/page` | `renderView` |
| Tab switch | active class on `.tab` | the destination tab's render |
| Drop / upload xlsx | `state.batch` | `renderBatch` (predictions recompute) |
| Inline-edit batch cell | one `state.batch[i]` field | `renderBatch` |
| Override decision | `state.batch[i].decision` | `renderBatch` |
| Commit batch | `loadStored()` += batch, `state.batch = []` | `renderManager`, `renderView` |
| Modal save (baked) | `loadBakedPatches()[idx]` updated | `renderView`, `renderManager` |
| Modal save (manager) | `loadStored()` record updated | `renderView`, `renderManager` |
| Modal delete (baked) | `loadBakedPatches()[idx] = {_deleted: true}` | `renderView`, `renderManager` |
| Modal delete (manager) | `loadStored()` filtered | `renderView`, `renderManager` |
| Manager-panel inline edit | `loadStored()` updated | `renderView`, `renderCommittedRecords`, `renderBatch` |
| 移除所有已儲存的新增紀錄 | `loadStored() = []` | `renderManager`, `renderView` |
| Push success | `setLastSync()`, `clearDirty()` | `updateSyncStatus` |
| Pull success | `loadBakedPatches()`, `loadStored()` replaced; `clearDirty()` | `renderView`, `renderManager`, `updateSyncStatus` |

Anything that calls `saveStored()` or `saveBakedPatches()` automatically marks dirty. Pull explicitly clears dirty after the saves. Push clears dirty on success.

## Non-goals

- Row-level conflict detection / OCC.
- Real-time collaboration. Auto-sync cadence is 2 min; sub-minute updates require a real DB.
- Authentication. Manager password is local obfuscation; sheet auth is whatever Google decides for "Anyone with link."
- Server-side validation. The browser enforces the booking rules (`predict()`); a malicious POST to the Apps Script can write arbitrary rows. Acceptable because the URL is unguessable and the audience is trusted.
- Schema migrations across versions. Changing xlsx columns means rebuilding `index.html`; changing the sheet schema means coordinated edits to `apps-script.gs` + `buildSheetRecords()` + the pull parser.

## Pointers

- [README.md](README.md) — getting started, build/run, manager password rotation.
- [booking_rules.md](booking_rules.md) — the rules `predict()` enforces.
- [build.py](build.py) — source of truth for both the HTML template and the `MANAGER_HTML_BLOCK`. Edits to either show up only after a rebuild — except for code in the outer `<script>` block, which is also mirrored in `index.html` so localhost edits don't strictly require running `build.py`.
- [apps-script.gs](apps-script.gs) — paste into the linked Google Sheet's Apps Script editor.
- [index.html](index.html) — the built artifact. Both the source-of-truth runtime AND a regeneratable file. Keep them in sync.
