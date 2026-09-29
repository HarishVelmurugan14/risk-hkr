# Risk HKR — setup

Two files. `index.html` is the whole app. `Code.gs` turns a Google Sheet into its database.

## 1. Sheet + Apps Script

1. New Google Sheet. Name it whatever — the script creates the `calls` and `config` tabs itself on first write.
2. **Extensions → Apps Script**. Delete the placeholder, paste in `Code.gs`.
3. In the editor, run `TEST_setApiToken('your-long-random-secret')` **once**. This stores the key in Script Properties. Keep a copy — you'll paste it into each browser. Running `TEST_setApiToken('')` removes the gate.
4. **Deploy → New deployment → Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Authorise when prompted. Copy the `/exec` URL.

"Anyone" is required because GitHub Pages is a different origin and the page calls it unauthenticated. The key you set in step 3 is what actually gates it — anyone with both the URL and the key can read and write the sheet. That's acceptable for tickers, prices and notes. Never put account numbers, credentials or anything sensitive in this sheet.

**If you skip step 3 the gate is open** and the URL alone grants full access. The app detects this and shows a warning banner on every load.

## 2. Wire up the page

Open `index.html` and set one line at the top of the `<script>` block — the URL only:

```js
const API = 'https://script.google.com/macros/s/AKfy.../exec';
```

**The key does not go in the file.** On first load the page prompts you to paste it once; it is kept in that browser's `localStorage` under `hkr_api_token` and reused silently. If the backend ever answers `unauthorized` — wrong key, or you rotated it — the prompt returns. There is also a *Work offline* option that keeps the app usable with local-only storage.

Because no secret is committed, **this repo can be public.** Each browser you use needs the key pasted once.

## 3. Publish

`index.html` and `Code.gs` live in this repo (`risk-hkr`). Turn on GitHub Pages for it and the page is served as-is — no build step. Same Sheets + Apps Script + Pages pattern as VaultZero, so the trade record reads the same way even though it sits in its own repo; pull it into VaultZero's main workbook with `IMPORTRANGE` off the `calls` sheet when you want it alongside the rest.

The only thing in the source is the web-app URL, which is useless without the key — so the repo does not need to be private. Rotate by re-running `TEST_setApiToken()` with a new value; every browser will prompt again on its next call.

## How it behaves

- Loads instantly from `localStorage`, then reconciles with the Sheet.
- Every save writes locally first, then to the Sheet. If the Sheet is unreachable, a banner appears and you keep working — nothing is lost, it just doesn't sync.
- If `API` is still the placeholder, it runs browser-only with a banner saying so.
- `LockService` serialises writes, so two tabs open at once won't clobber each other.

## Data model

`calls` sheet, one row per mentor **event** — a buy or a sell, copied or skipped:

| field | meaning |
|---|---|
| `id`, `createdAt`, `date`, `ticker`, `kind` | identity; `kind` is `stock` or `etf` |
| `side` | `buy` or `sell` |
| `hisQty`, `hisPrice` | what he did |
| `myPrice`, `myShares`, `myValue` | mine at the core-to-core ratio. Shares are computed even for skipped calls, which is what makes the missed/dodged numbers possible. `myValue` is 0 when skipped, so the tank stays accurate. On a sell, `myShares` is capped at what I actually hold |
| `hisThesis` | his reasoning |
| `myView` | my read on it |
| `decision`, `skipReason` | `copied` or `skipped` |

Rows are **events, not positions.** The app folds them into positions at read time: buys push a lot, sells consume lots **FIFO** — oldest first, spilling into the next lot only for the excess. Units held, average cost and realised P&L are all derived, never stored, so deleting an event recomputes the position cleanly.

`config` sheet holds `tank`, `mine`, `his` — the fuel tank and the two core bases the sizing ratio is derived from — plus one `m:TICKER` row per marked price (the current price used for unrealised P&L).

Because it's a plain sheet, you can pivot it, chart it, or pull it into VaultZero's main workbook with `IMPORTRANGE` without touching the app.

## Note on the journal design

This version logs **his thesis first, then your view on it** — the beginner-appropriate rung. It measures whether you can *evaluate* a thesis. The mandate's actual success test ("can I justify a trade on my own") needs the harder version: your own read written **before** you see his. Worth revisiting at the month-3 review.
