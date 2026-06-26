# form4api-sheets-addin

Google Sheets add-on for [Form4API](https://www.form4api.com) — query SEC Form 4 insider trading data with custom spreadsheet functions.

> **Status: v0.3.0 — 8 functions live.** Phase 1 spike PASSED 2026-06-01 (real data spilled into a real sheet). Phase 2 added 4 more functions + 1-hour caching + plan-gating. Phase 3 added `FORM4API_RETURNS`, apostrophe title-case, cluster-flag window depth fix, and editor-shares-key warning. v0.3.0 adds `FORM4API_SCORECARD` (insider track record, Pro+) and `FORM4API_HOLDINGS` (13F-HR institutional holders, Business+). Next: Google Workspace Marketplace listing per [PLAN_SHEETS_ADDIN.md](https://github.com/theodor90/insiderapi/blob/main/PLAN_SHEETS_ADDIN.md) Phase 5.

## Functions

| Formula | Returns | Plan |
|---|---|---|
| `=FORM4API_TX(ticker, [limit], [code])` | Spill of up to 100 recent transactions for the ticker. `code` filter: `P` purchase, `S` sale, `A` award, `M` option exercise, `F` tax withholding, `D` disposition, `G` gift. | Free |
| `=FORM4API_TX_LATEST(ticker)` | Single most-recent transaction (single-row spill) | Free |
| `=FORM4API_INSIDER_TX(cik, [limit])` | All transactions for one insider across companies. CIK is auto-padded to the 10-digit canonical form. | Free |
| `=FORM4API_RETURNS(ticker, [horizon])` | Average post-trade return (1d / 1w / 1m / 3m / 6m, default 3m) across recent open-market P/S transactions. Returns a decimal — `0.0523` = +5.23%. | Pro+ |
| `=FORM4API_SCORECARD(cik)` | 2-column labelled table: hit rate, avg/median 3m & 6m return, scored-buy count, sample-sufficiency flag, last trade date, methodology. Returns are decimals — format as % if preferred. Shows `Sample Sufficient: No` when fewer than 5 matured buys exist. | Pro+ |
| `=FORM4API_SENTIMENT(ticker, [months])` | Average MSPR-style sentiment score (-100 to +100). 10b5-1 plan trades excluded. | Business+ |
| `=FORM4API_CLUSTER_FLAG(ticker)` | `BUY` / `SELL` / `BOTH` / empty — direction of recent cluster signals | Business+ |
| `=FORM4API_HOLDINGS(ticker)` | Top 10 institutional holders from the latest 13F-HR filings. 4-column spill: Manager \| Shares \| Value (USD) \| Quarter. | Business+ |

Transaction spills return 6 columns: `Date | Insider | Code | Shares | Price | Value`. Dates are parsed to real Date objects so Sheets formats them natively (`5/8/2026`). Insider names are title-cased from EDGAR's ALL-CAPS source.

## Examples

```
=FORM4API_TX("AAPL")                Last 20 Apple insider transactions
=FORM4API_TX("AAPL", 50, "S")       Last 50 Apple sales only
=FORM4API_TX_LATEST("NVDA")         Most recent NVDA insider filing
=FORM4API_INSIDER_TX("1214156")     Tim Cook's career trading activity
=FORM4API_RETURNS("NVDA", "3m")     Average 3-month post-trade return on NVDA insider trades
=FORM4API_SENTIMENT("AAPL", 6)      6-month average sentiment for Apple
=FORM4API_CLUSTER_FLAG("CRT")       "BUY" if multiple insiders accumulating
```

## Caching

Every result is cached in the sheet's `CacheService` with a 1-hour TTL. Repeated formula reads in the same hour serve from cache (no API call, no quota consumption). To force a refresh: **Form4API → Refresh All** in the menu, then click into any cell and press Enter.

## Setup (one-time, ~10 minutes)

### 1. Install dev deps

```bash
npm install
```

### 2. Log into Google with clasp

```bash
npx clasp login
```

Browser opens → sign in with the Google account that will own the bound sheet. Credentials land in `~/.clasprc.json` (gitignored).

### 3. Enable the Apps Script API (one-time, per account)

Visit https://script.google.com/home/usersettings → flip **Google Apps Script API** toggle to ON.

### 4. Create the bound Apps Script project

```bash
npx clasp create --type sheets --title "Form4API Sheets Add-on"
```

Creates a new Google Sheet in your Drive AND an Apps Script project bound to it. Writes `.clasp.json` (gitignored) with the script ID.

### 5. Push the code

```bash
npx clasp push
```

### 6. Open the bound sheet

```bash
npx clasp open --addon
```

Reload the tab — a **Form4API** menu appears in the top bar.

### 7. Set your API key

In the sheet: **Form4API → Set API Key…** → paste a key from your [Form4API dashboard](https://www.form4api.com/dashboard). Stored in `PropertiesService.getDocumentProperties()` — sheet-scoped, encrypted at rest by Google.

### 8. First-cell test

```
=FORM4API_TX("AAPL")
```

Google will ask you to authorise the script's OAuth scopes on first call. Click through → grant. From then on, custom functions run silently.

## Plan-gating UX

Plans: **Free** (limited daily quota), **Starter $19/mo** (7,500 req/day), **Pro** (scorecard + returns), **Business** (all functions). When a key's plan is too low, the cell shows `#ERROR!` with a hover tooltip, e.g.:

```
This function requires the Business plan. Upgrade at https://www.form4api.com/dashboard/billing
```

- `=FORM4API_SCORECARD` and `=FORM4API_RETURNS` → require Pro or higher
- `=FORM4API_SENTIMENT`, `=FORM4API_CLUSTER_FLAG`, `=FORM4API_HOLDINGS` → require Business or higher

This is the intended behaviour — error rendering in Sheets is structured enough that users can read the upgrade path inline. Free/Starter-tier functions degrade gracefully when a key's daily quota is exhausted (`429 → Rate limited. Retry after N seconds`).

## Why `PropertiesService` for auth

Sheet-scoped, encrypted at rest by Google, never sent off-Drive. No OAuth handshake required (Apps Script OAuth requires a Cloud Project + verification process; for an API-key auth pattern that the backend already supports, the property approach is dramatically simpler with no UX cost).

**Caveat:** anyone with edit access to the bound sheet can read or replace the stored API key. The Set API Key prompt surfaces this risk explicitly so the choice is informed at the moment of decision.

## Phase 4 backlog

- Google Workspace Marketplace listing (Phase 5 per the main plan): privacy policy, ToS link, icon, screenshots, listing copy, brand verification, review wait (~1-3 days)
- `=FORM4API_INST_OWNERSHIP(ticker)` — current quarter AUM + QoQ trend per ticker (the `institutionalOwnership` block we noticed is embedded in every transaction)
- Excel parallel via Office Add-ins — defer until Sheets shows ≥50 installs

## What shipped in v0.3.0

- `=FORM4API_SCORECARD(cik)` — insider track-record scorecard. Calls `GET /v1/insiders/{cik}/scorecard` (Pro+). Spills a 14-row × 2-column labelled table with hit rate, avg/median 3m & 6m returns, scored-buy counts, sample-sufficiency flags, last trade date, and methodology. Returns kept as fractions (format cells as % in Sheets).
- `=FORM4API_HOLDINGS(ticker)` — top 10 institutional holders from 13F-HR. Calls `GET /v1/holdings?ticker=...` (Business+). Spills Manager | Shares | Value (USD) | Quarter (up to 10 rows).
- All URLs updated to canonical `www.form4api.com`.
- Plan-gating section updated to mention Starter ($19/mo) tier.

## What shipped in Phase 3

- `=FORM4API_RETURNS(ticker, horizon)` added — uses the embedded `return1d/1w/1m/3m/6m` fields in the `/v1/transactions` response (no backend change needed, just had to probe to confirm)
- Apostrophe title-case so `O'BRIEN DEIRDRE` renders as `O'Brien Deirdre`
- `FORM4API_CLUSTER_FLAG` window depth raised from `per_page: 10` to `per_page: 50` — small-cap tickers with one-off cluster signals no longer get paged out by subsequent non-cluster activity
- Set API Key prompt now warns that anyone with edit access to the sheet can use the key

## License

MIT
