# form4api-sheets-addin

Google Sheets add-on for [Form4API](https://form4api.com) — query SEC Form 4 insider trading data with custom spreadsheet functions.

> **Status: Phase 2 — 5 functions live and tested end-to-end.** Phase 1 spike PASSED 2026-06-01 (real data spilled into a real sheet). Phase 2 added 4 more functions + 1-hour caching + plan-gating error path. Next: Phase 3 polish + Google Workspace Marketplace listing per [PLAN_SHEETS_ADDIN.md](https://github.com/theodor90/insiderapi/blob/main/PLAN_SHEETS_ADDIN.md).

## Functions

| Formula | Returns | Plan |
|---|---|---|
| `=FORM4API_TX(ticker, [limit], [code])` | Spill of up to 100 recent transactions for the ticker. `code` filter: `P` purchase, `S` sale, `A` award, `M` option exercise, `F` tax withholding, `D` disposition, `G` gift. | Free |
| `=FORM4API_TX_LATEST(ticker)` | Single most-recent transaction (single-row spill) | Free |
| `=FORM4API_INSIDER_TX(cik, [limit])` | All transactions for one insider across companies. CIK is auto-padded to the 10-digit canonical form. | Free |
| `=FORM4API_SENTIMENT(ticker, [months])` | Average MSPR-style sentiment score (-100 to +100). 10b5-1 plan trades excluded. | Business+ |
| `=FORM4API_CLUSTER_FLAG(ticker)` | `BUY` / `SELL` / `BOTH` / empty — direction of recent cluster signals | Business+ |

Transaction spills return 6 columns: `Date | Insider | Code | Shares | Price | Value`. Dates are parsed to real Date objects so Sheets formats them natively (`5/8/2026`). Insider names are title-cased from EDGAR's ALL-CAPS source.

## Examples

```
=FORM4API_TX("AAPL")                Last 20 Apple insider transactions
=FORM4API_TX("AAPL", 50, "S")       Last 50 Apple sales only
=FORM4API_TX_LATEST("NVDA")         Most recent NVDA insider filing
=FORM4API_INSIDER_TX("1214156")     Tim Cook's career trading activity
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

In the sheet: **Form4API → Set API Key…** → paste a key from your [Form4API dashboard](https://form4api.com/dashboard). Stored in `PropertiesService.getDocumentProperties()` — sheet-scoped, encrypted at rest by Google.

### 8. First-cell test

```
=FORM4API_TX("AAPL")
```

Google will ask you to authorise the script's OAuth scopes on first call. Click through → grant. From then on, custom functions run silently.

## Plan-gating UX

When a Free-tier key calls a Business+ function (`=FORM4API_SENTIMENT`, `=FORM4API_CLUSTER_FLAG`), the cell shows `#ERROR!` with a hover tooltip:

```
This function requires the Business plan. Upgrade at https://form4api.com/dashboard/billing
```

This is the intended behaviour — error rendering in Sheets is structured enough that users can read the upgrade path inline. Free-tier functions degrade gracefully when called from a key that's run out of daily quota (`429 → Rate limited. Retry after N seconds`).

## Why `PropertiesService` for auth

Sheet-scoped, encrypted at rest by Google, never sent off-Drive. No OAuth handshake required (Apps Script OAuth requires a Cloud Project + verification process; for an API-key auth pattern that the backend already supports, the property approach is dramatically simpler with no UX cost).

**Caveat:** anyone with edit access to the bound sheet can read or replace the stored API key. Phase 3 will add a "Editors can use your API key — share carefully" banner on first set.

## Phase 3 backlog

- `FORM4API_RETURNS(ticker, horizon)` — backend doesn't yet expose an aggregate-returns endpoint; would require either a new endpoint or a per-tx walk + average
- Apostrophe title-case (`O'brien Deirdre` → `O'Brien Deirdre`) — ~3 line tweak to `titleCaseInsider_`
- `FORM4API_CLUSTER_FLAG` per-ticker window — raise from 10 to 50 rows so older clusters (like UBCP's 2026-05-28 BUY) don't get paged out
- Banner on the bound sheet warning editors that they share the API key
- Google Workspace Marketplace listing (Phase 5 per the main plan)

## License

MIT
