# form4api-sheets-addin

Google Sheets add-on for [Form4API](https://form4api.com) — query SEC Form 4 insider trading data with custom spreadsheet functions.

> **Status: Phase 1 spike.** Single function (`=FORM4API_TX`) + API-key handling. If the spike works end-to-end (`=FORM4API_TX("AAPL")` returns real insider transactions in a real Google Sheet), expand to the full 6-function surface per [PLAN_SHEETS_ADDIN.md](https://github.com/theodor90/insiderapi/blob/main/PLAN_SHEETS_ADDIN.md).

## What you get (Phase 1)

```
=FORM4API_TX("AAPL")           returns the last 20 Apple insider transactions
=FORM4API_TX("AAPL", 50)       returns the last 50
```

Output is a 6-column spill: `Date | Insider | Code | Shares | Price | Value`.

## Setup (one-time, ~10 minutes)

### 1. Install dev deps

```bash
npm install
```

### 2. Log into Google with clasp

```bash
npx clasp login
```

Browser opens → sign in with the same Google account you'll use the add-on from. Stores credentials in `~/.clasprc.json` (already gitignored).

### 3. Create the bound Apps Script project

```bash
npx clasp create --type sheets --title "Form4API Sheets Add-on"
```

This creates **both** a new Google Sheet and an Apps Script project bound to it. Writes `.clasp.json` (gitignored) with the script ID. Open the new sheet with:

```bash
npx clasp open --addon
```

(or visit the URL clasp prints).

### 4. Push the code

```bash
npx clasp push
```

The script is now live in the Apps Script project. Reload the bound Google Sheet — a **Form4API** menu should appear in the top bar.

### 5. Set your API key

In the sheet: **Form4API → Set API Key…** → paste your key (get one free at [form4api.com](https://form4api.com) → Sign in → Dashboard).

### 6. Test it

In any cell:

```
=FORM4API_TX("AAPL")
```

Should spill 6 columns × ~20 rows of recent insider transactions at Apple.

## What this proves

- Apps Script can `UrlFetchApp.fetch()` `api.form4api.com` with `X-Api-Key` headers (CORS isn't an issue server-side)
- `PropertiesService.getDocumentProperties()` is a workable home for the API key (encrypted at rest, sheet-scoped)
- Custom-function array spill maps the JSON response into a usable spreadsheet layout
- The 6-min Apps Script execution timeout isn't tripped by a single API call

If all 4 hold, Phase 2 productionizes 5 more functions (`FORM4API_TX_LATEST`, `FORM4API_SENTIMENT`, `FORM4API_CLUSTER_FLAG`, `FORM4API_INSIDER_TX`, `FORM4API_RETURNS`), adds caching (`CacheService`), and prepares the Workspace Marketplace listing.

## Why DocumentProperties for auth

`PropertiesService.getDocumentProperties()` is scoped to the bound sheet — accessible to every editor of that sheet but invisible to other sheets and to the wider Google account. Encrypted at rest by Google. No OAuth flow needed (Apps Script OAuth requires a Cloud Project + verification process; for an API-key pattern we already support, this is dramatically simpler with no UX cost).

**Caveat:** anyone with edit access to the bound sheet can use the stored API key (via the menu's "Set API Key" prompt they could overwrite, or by writing their own Apps Script that calls `getProperty`). A future Phase 2 menu addition should display a "Editors of this sheet can use your API key — share carefully" banner.

## License

MIT
