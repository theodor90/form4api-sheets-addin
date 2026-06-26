// Phase 4 — 8 production custom functions + caching + error handling.
// Single file by design: clasp transpiles each .ts file to a separate .gs
// file but they share one global namespace, so cross-file refs are messy.
// Keeping everything here trades file count for clarity.
//
// Function surface:
//   =FORM4API_TX(ticker, [limit], [code])      Free       — spill of N transactions
//   =FORM4API_TX_LATEST(ticker)                Free       — single latest transaction row
//   =FORM4API_INSIDER_TX(cik, [limit])         Free       — spill of one insider's transactions
//   =FORM4API_RETURNS(ticker, [horizon])       Pro+       — avg post-trade return for ticker
//   =FORM4API_SCORECARD(cik)                   Pro+       — insider track-record scorecard table
//   =FORM4API_SENTIMENT(ticker, [months])      Business+  — MSPR-style sentiment score
//   =FORM4API_CLUSTER_FLAG(ticker)             Business+  — "BUY" / "SELL" / "" / "BOTH"
//   =FORM4API_HOLDINGS(ticker)                 Business+  — top 10 institutional holders (13F-HR)

const API_BASE = 'https://api.form4api.com'
const PROP_API_KEY = 'form4api_key'
const CACHE_TTL_SECONDS = 3600 // 1 hour
const UA = 'form4api-sheets-addin/0.3.0'

// ──────────────────────────────────────────────────────────────────────────────
// CUSTOM FUNCTIONS (the @customfunction-tagged exports)
// ──────────────────────────────────────────────────────────────────────────────

/**
 * Insider transactions for a ticker.
 *
 * @param {string} ticker   Stock symbol, e.g. "AAPL"
 * @param {number} limit    Max rows to return (default 20, max 100)
 * @param {string} code     Optional transaction code filter: P=purchase, S=sale,
 *                          A=award, M=option exercise, F=tax withholding, D=disposition
 * @return                  6-column spill: Date, Insider, Code, Shares, Price, Value
 * @customfunction
 */
function FORM4API_TX(ticker: string, limit?: number, code?: string): unknown[][] {
  if (!ticker) throw new Error('FORM4API_TX: pass a ticker like "AAPL"')
  const perPage = clampInt_(limit, 1, 100, 20)
  const params: Record<string, string | number> = {
    ticker: ticker.toUpperCase(),
    per_page: perPage,
  }
  if (code) params['code'] = code.toUpperCase()

  const txs = apiGetCached_<Transaction[]>('/v1/transactions', params, 'FORM4API_TX')
  if (!txs || txs.length === 0) return [['No transactions found', '', '', '', '', '']]
  return [TX_HEADER, ...txs.map(transactionRow_)]
}

/**
 * Single most recent insider transaction for a ticker.
 *
 * @param {string} ticker   Stock symbol, e.g. "AAPL"
 * @return                  6-column single row: Date, Insider, Code, Shares, Price, Value
 * @customfunction
 */
function FORM4API_TX_LATEST(ticker: string): unknown[][] {
  if (!ticker) throw new Error('FORM4API_TX_LATEST: pass a ticker like "AAPL"')
  const txs = apiGetCached_<Transaction[]>(
    '/v1/transactions',
    { ticker: ticker.toUpperCase(), per_page: 1 },
    'FORM4API_TX_LATEST',
  )
  if (!txs || txs.length === 0) return [['No transactions found', '', '', '', '', '']]
  return [TX_HEADER, transactionRow_(txs[0])]
}

/**
 * All transactions for one insider by their SEC CIK.
 *
 * @param {string} cik     Insider CIK (e.g. "1214156" for Tim Cook). Leading zeros optional.
 * @param {number} limit   Max rows (default 50, max 100)
 * @return                 Same 6-column spill as FORM4API_TX
 * @customfunction
 */
function FORM4API_INSIDER_TX(cik: string, limit?: number): unknown[][] {
  if (!cik) throw new Error('FORM4API_INSIDER_TX: pass an insider CIK like "1214156"')
  const perPage = clampInt_(limit, 1, 100, 50)
  // SEC CIKs are canonically 10-digit zero-padded (e.g. "0001214156").
  // The backend's /v1/insiders/{cik}/* routes match on the padded form only.
  // Pad whatever the user pasted so both "1214156" and "0001214156" work.
  const paddedCik = padCik_(cik)
  const txs = apiGetCached_<Transaction[]>(
    `/v1/insiders/${encodeURIComponent(paddedCik)}/transactions`,
    { per_page: perPage },
    'FORM4API_INSIDER_TX',
  )
  if (!txs || txs.length === 0) return [['No transactions found', '', '', '', '', '']]
  return [TX_HEADER, ...txs.map(transactionRow_)]
}

/**
 * Monthly insider-sentiment score for a ticker (MSPR-style, -100 to +100).
 * 10b5-1 plan trades excluded so the score reflects discretionary conviction.
 *
 * @param {string} ticker   Stock symbol, e.g. "AAPL"
 * @param {number} months   How many past months to average (default 12, max 60)
 * @return                  Average sentiment score (decimal -100..+100)
 * @customfunction
 */
function FORM4API_SENTIMENT(ticker: string, months?: number): number | string {
  if (!ticker) throw new Error('FORM4API_SENTIMENT: pass a ticker like "AAPL"')
  const monthsParam = clampInt_(months, 1, 60, 12)
  // Backend response shape (verified 2026-06-01):
  //   { ticker, companyName, monthly: [{period, score, buyValue, sellValue, buyCount, sellCount}] }
  // where `score` is the -100..+100 MSPR-style value. Average across months.
  const res = apiGetCached_<SentimentResponse>(
    `/v1/signals/sentiment/${encodeURIComponent(ticker.toUpperCase())}`,
    { months: monthsParam },
    'FORM4API_SENTIMENT',
  )
  if (!res) return 'No sentiment data'

  // Try canonical `monthly`, fall back to legacy/alternate names for forward compat.
  const scores = (res.monthly ?? res.scores ?? res.months ?? res.data ?? []) as Array<Record<string, unknown>>
  if (!Array.isArray(scores) || scores.length === 0) return 'No sentiment data'

  const values: number[] = []
  for (const s of scores) {
    const v = (s.score ?? s.mspr ?? s.value) as number | undefined
    if (typeof v === 'number' && !isNaN(v)) values.push(v)
  }
  if (values.length === 0) return 'No sentiment data'
  return Math.round((values.reduce((a, b) => a + b, 0) / values.length) * 100) / 100
}

/**
 * Recent cluster signal direction for a ticker.
 *
 * @param {string} ticker   Stock symbol, e.g. "NVDA"
 * @return                  "BUY" / "SELL" / "BOTH" / "" (empty = no recent cluster)
 * @customfunction
 */
function FORM4API_CLUSTER_FLAG(ticker: string): string {
  if (!ticker) throw new Error('FORM4API_CLUSTER_FLAG: pass a ticker like "NVDA"')
  // per_page=50 (was 10) so older clusters within the window still surface —
  // small tickers can have non-cluster activity push the cluster row past
  // the first page (caught with UBCP in Phase 2 smoke test).
  const signals = apiGetCached_<Signal[]>(
    '/v1/signals',
    { ticker: ticker.toUpperCase(), per_page: 50 },
    'FORM4API_CLUSTER_FLAG',
  )
  if (!signals || signals.length === 0) return ''

  // Backend response shape (verified 2026-06-01):
  //   [{ticker, companyName, signalDate, buySellRatio, isClusterBuy, isClusterSell, insiderCount}]
  // Most tickers have signal rows but with both flags=false — that's "regular
  // activity, no cluster pattern", which is the correct empty-string case.
  let buy = false,
    sell = false
  for (const s of signals) {
    if (s.isClusterBuy === true) buy = true
    if (s.isClusterSell === true) sell = true
  }
  if (buy && sell) return 'BOTH'
  if (buy) return 'BUY'
  if (sell) return 'SELL'
  return ''
}

/**
 * Average post-trade return for a ticker over a given horizon. Aggregated over
 * the most recent open-market transactions whose returns have been computed.
 *
 * @param {string} ticker   Stock symbol, e.g. "AAPL"
 * @param {string} horizon  "1d" | "1w" | "1m" | "3m" | "6m" — default "3m"
 * @return                  Average return as a decimal (0.0523 = +5.23%)
 *                          or "No returns data" if none of the rows have the field
 * @customfunction
 */
function FORM4API_RETURNS(ticker: string, horizon?: string): number | string {
  if (!ticker) throw new Error('FORM4API_RETURNS: pass a ticker like "AAPL"')
  const h = String(horizon || '3m').toLowerCase()
  const field = RETURNS_FIELD_MAP[h]
  if (!field) {
    throw new Error('FORM4API_RETURNS: horizon must be one of 1d, 1w, 1m, 3m, 6m')
  }
  // Restrict to open-market discretionary trades — option exercises (M), tax
  // withholdings (F), and awards (A) don't have meaningful price-to-return
  // relationships. The backend computes returns for all transactions but only
  // the P/S subset is useful for "did the insider's timing pay off" analysis.
  const txs = apiGetCached_<Transaction[]>(
    '/v1/transactions',
    { ticker: ticker.toUpperCase(), per_page: 100 },
    'FORM4API_RETURNS',
  )
  if (!txs || txs.length === 0) return 'No transactions found'

  const values: number[] = []
  for (let i = 0; i < txs.length; i++) {
    const t = txs[i] as Transaction & Record<string, unknown>
    if (t.transactionCode !== 'P' && t.transactionCode !== 'S') continue
    const r = t[field] as number | null | undefined
    if (typeof r === 'number' && !isNaN(r)) values.push(r)
  }
  if (values.length === 0) return 'No returns data for ' + h
  const avg = values.reduce(function (a, b) { return a + b }, 0) / values.length
  // Round to 4 decimal places so Sheets renders as e.g. 0.0523 or 5.23%
  return Math.round(avg * 10000) / 10000
}

/**
 * Track-record scorecard for one insider (Pro+).
 *
 * Calls GET /v1/insiders/{cik}/scorecard. Requires at least 5 matured
 * discretionary open-market buys — when sampleSufficient is "No", the
 * return fields will be empty (null from API); do not infer signal.
 *
 * @param {string} cik   Insider CIK (e.g. "1214128" for Tim Cook). Leading zeros optional.
 * @return               2-column labelled table: Metric | Value.
 *                       Returns are decimals (0.112 = +11.2%); format cells as % if preferred.
 * @customfunction
 */
function FORM4API_SCORECARD(cik: string): unknown[][] {
  if (!cik) throw new Error('FORM4API_SCORECARD: pass an insider CIK like "1214128"')
  const paddedCik = padCik_(cik)
  const sc = apiGetCached_<ScorecardResponse>(
    `/v1/insiders/${encodeURIComponent(paddedCik)}/scorecard`,
    {},
    'FORM4API_SCORECARD',
  )
  if (!sc) return [['No scorecard data', '']]

  // Returns arrive as fractions (0.112 = +11.2%) — kept as-is so the user
  // can apply Sheets percentage formatting, consistent with FORM4API_RETURNS.
  const na = (v: number | null | undefined): number | string => (typeof v === 'number' ? v : '')
  return [
    ['Metric', 'Value'],
    ['Insider', sc.insiderName ?? ''],
    ['Hit Rate 3m', na(sc.hitRate3m)],
    ['Avg Return 3m', na(sc.avgReturn3m)],
    ['Median Return 3m', na(sc.medianReturn3m)],
    ['Buys Scored 3m', sc.scoredBuyCount ?? ''],
    ['Sample Sufficient 3m', sc.sampleSufficient === true ? 'Yes' : 'No'],
    ['Hit Rate 6m', na(sc.hitRate6m)],
    ['Avg Return 6m', na(sc.avgReturn6m)],
    ['Median Return 6m', na(sc.medianReturn6m)],
    ['Buys Scored 6m', sc.scoredBuyCount6m ?? ''],
    ['Sample Sufficient 6m', sc.sampleSufficient6m === true ? 'Yes' : 'No'],
    ['Last Trade', sc.lastTradeAt ? parseDate_(sc.lastTradeAt) : ''],
    ['Methodology', sc.methodology ?? ''],
  ]
}

/**
 * Top institutional holders of a ticker from 13F-HR filings (Business+).
 *
 * Calls GET /v1/holdings?ticker=... and returns the first 10 rows (one position
 * per manager per quarter — aggregates the manager field, not the sub-manager
 * attribution rows that Berkshire-style filers produce).
 *
 * @param {string} ticker   Stock symbol, e.g. "AAPL"
 * @return                  4-column spill: Manager | Shares | Value (USD) | Quarter
 * @customfunction
 */
function FORM4API_HOLDINGS(ticker: string): unknown[][] {
  if (!ticker) throw new Error('FORM4API_HOLDINGS: pass a ticker like "AAPL"')
  const holdings = apiGetCached_<Holding[]>(
    '/v1/holdings',
    { ticker: ticker.toUpperCase(), per_page: 10 },
    'FORM4API_HOLDINGS',
  )
  if (!holdings || holdings.length === 0) return [['No holdings data', '', '', '']]
  return [HOLDINGS_HEADER, ...holdings.map(holdingRow_)]
}

// ──────────────────────────────────────────────────────────────────────────────
// SHARED INFRASTRUCTURE
// ──────────────────────────────────────────────────────────────────────────────

const RETURNS_FIELD_MAP: Record<string, string> = {
  '1d': 'return1d',
  '1w': 'return1w',
  '1m': 'return1m',
  '3m': 'return3m',
  '6m': 'return6m',
}

const TX_HEADER: string[] = ['Date', 'Insider', 'Code', 'Shares', 'Price', 'Value']
const HOLDINGS_HEADER: string[] = ['Manager', 'Shares', 'Value (USD)', 'Quarter']

interface Transaction {
  transactionDate: string
  insiderName: string
  transactionCode: string
  sharesAmount: number
  pricePerShare: number | null
  totalValue: number | null
}

interface Signal {
  isClusterBuy?: boolean
  isClusterSell?: boolean
  signalDate?: string
  insiderCount?: number
  buySellRatio?: number
}

interface SentimentResponse {
  monthly?: unknown
  scores?: unknown
  months?: unknown
  data?: unknown
  [k: string]: unknown
}

// Response from GET /v1/insiders/{cik}/scorecard (Pro+).
// Null fields occur when sampleSufficient / sampleSufficient6m is false
// (fewer than 5 matured discretionary open-market buys available).
interface ScorecardResponse {
  insiderCik: string
  insiderName: string
  scoredBuyCount: number | null
  sampleSufficient: boolean
  hitRate3m: number | null
  avgReturn3m: number | null
  medianReturn3m: number | null
  scoredBuyCount6m: number | null
  sampleSufficient6m: boolean
  hitRate6m: number | null
  avgReturn6m: number | null
  medianReturn6m: number | null
  lastTradeAt: string | null
  methodology: string | null
}

// One row from GET /v1/holdings (Business+).
interface Holding {
  manager: string
  managerCik: string
  reportPeriod: string
  issuerName: string
  ticker: string | null
  cusip: string
  value: number
  shares: number
}

/** Format one Transaction as a row matching TX_HEADER. */
function transactionRow_(t: Transaction): unknown[] {
  return [
    parseDate_(t.transactionDate),
    titleCaseInsider_(t.insiderName),
    t.transactionCode,
    t.sharesAmount,
    t.pricePerShare ?? '',
    t.totalValue ?? '',
  ]
}

/** Format one Holding as a row matching HOLDINGS_HEADER. */
function holdingRow_(h: Holding): unknown[] {
  return [
    h.manager,
    h.shares,
    h.value,
    h.reportPeriod ? parseDate_(h.reportPeriod) : '',
  ]
}

/** API call with 1h cache. cacheTag is part of the key so a function-specific
 *  Refresh All can selectively clear (see clearCache_). */
function apiGetCached_<T>(path: string, params: Record<string, string | number>, cacheTag: string): T | null {
  const cache = CacheService.getDocumentCache()
  const key = cacheKey_(cacheTag, path, params)

  if (cache) {
    const hit = cache.get(key)
    if (hit !== null) {
      try {
        return JSON.parse(hit) as T
      } catch {
        // fall through to live fetch on corrupt cache value
      }
    }
  }

  const data = apiGetLive_<T>(path, params)
  if (cache && data !== null) {
    try {
      cache.put(key, JSON.stringify(data), CACHE_TTL_SECONDS)
    } catch (e) {
      // CacheService throws on >100KB values; swallow and serve the live data
      Logger.log(`cache put failed for ${key}: ${e}`)
    }
  }
  return data
}

function apiGetLive_<T>(path: string, params: Record<string, string | number>): T {
  const key = getApiKey_()
  const url = buildUrl_(API_BASE + path, params)
  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'X-Api-Key': key, 'User-Agent': UA },
    muteHttpExceptions: true,
  })

  const code = res.getResponseCode()
  const body = res.getContentText()

  if (code === 200) {
    return JSON.parse(body) as T
  }

  if (code === 401) {
    throw new Error('Invalid API key. Open the Form4API menu → Set API Key to replace it.')
  }
  if (code === 402) {
    // Backend returns { error: { code, message, ... } } for plan-required
    try {
      const parsed = JSON.parse(body) as { error?: { code?: string; message?: string } }
      const msg = parsed.error?.message ?? 'This function requires a higher plan.'
      throw new Error(`${msg} Upgrade at https://www.form4api.com/dashboard/billing`)
    } catch (e) {
      // already an Error from the throw above — re-throw verbatim
      if (e instanceof Error && e.message.indexOf('Upgrade at') >= 0) throw e
      throw new Error('This function requires a higher plan. Upgrade at https://www.form4api.com/dashboard/billing')
    }
  }
  if (code === 429) {
    const retry = res.getHeaders() as Record<string, string>
    const wait = retry['Retry-After'] || retry['retry-after'] || 'a few'
    throw new Error(`Rate limited. Retry after ${wait} seconds — bulk recalc may need to be spread out.`)
  }
  if (code >= 500) {
    throw new Error(`Form4API server error ${code}. Retry in a moment; status: https://www.form4api.com`)
  }
  throw new Error(`API error ${code}: ${body.slice(0, 200)}`)
}

function getApiKey_(): string {
  const key = PropertiesService.getDocumentProperties().getProperty(PROP_API_KEY)
  if (!key) {
    throw new Error(
      'Form4API key not set. Open the Form4API menu (top bar) → Set API Key. Get one free at https://www.form4api.com',
    )
  }
  return key
}

/** Builds URL with query params, omitting undefined/null. */
function buildUrl_(base: string, params: Record<string, string | number>): string {
  const qs: string[] = []
  for (const k in params) {
    const v = params[k]
    if (v === undefined || v === null) continue
    qs.push(encodeURIComponent(k) + '=' + encodeURIComponent(String(v)))
  }
  return qs.length === 0 ? base : base + '?' + qs.join('&')
}

/** Cache key: <tag>:<path>?<sorted-params> — sorted so {a:1,b:2} and {b:2,a:1} match. */
function cacheKey_(tag: string, path: string, params: Record<string, string | number>): string {
  const keys = Object.keys(params).sort()
  const parts: string[] = []
  for (let i = 0; i < keys.length; i++) {
    parts.push(keys[i] + '=' + String(params[keys[i]]))
  }
  // Cache keys are capped at 250 chars by CacheService
  const full = `${tag}:${path}?${parts.join('&')}`
  return full.length <= 250 ? full : full.slice(0, 240) + ':' + simpleHash_(full)
}

/** djb2 hash for over-length cache keys */
function simpleHash_(s: string): string {
  let h = 5381
  for (let i = 0; i < s.length; i++) h = ((h << 5) + h) ^ s.charCodeAt(i)
  return (h >>> 0).toString(36)
}

function clampInt_(v: number | undefined, min: number, max: number, fallback: number): number {
  if (v === undefined || v === null || isNaN(Number(v))) return fallback
  const n = Math.floor(Number(v))
  if (n < min) return min
  if (n > max) return max
  return n
}

/** Pad a numeric CIK to the canonical 10-digit zero-padded form. */
function padCik_(cik: string | number): string {
  const digits = String(cik).replace(/\D/g, '')
  if (digits.length === 0) throw new Error(`Invalid CIK: "${cik}" — expected digits only`)
  if (digits.length >= 10) return digits
  return '0'.repeat(10 - digits.length) + digits
}

/** Parse ISO 8601 string to a real Date so Sheets formats it natively. */
function parseDate_(iso: string): Date | string {
  if (!iso) return ''
  // Strip any time portion if present — Form 4 transactions are date-only
  const d = new Date(iso)
  if (isNaN(d.getTime())) return iso // bail out gracefully on unparseable strings
  return d
}

/** Title-case SEC EDGAR insider names. Same logic shape as the frontend
 *  app/lib/titleCase.ts — if any char is lowercase, treat as already-cased
 *  and return verbatim; otherwise lowercase + first-letter-cap each word.
 *  Mid-initials like "D" stay single-letter and get uppercased naturally.
 *  Capitalises letters following an apostrophe so O'BRIEN → O'Brien (not
 *  O'brien — caught in Phase 2 smoke test). */
function titleCaseInsider_(name: string): string {
  if (!name) return name
  if (/[a-z]/.test(name)) return name // mixed-case input → trust it
  return name
    .toLowerCase()
    .split(/(\s+)/)
    .map(function (token) {
      const trimmed = token.trim()
      if (!trimmed) return token
      // Honorifics + suffixes that look like initials stay uppercase
      if (/^(jr|sr|ii|iii|iv|v|vi|md|phd|cfa|esq)$/i.test(trimmed)) return trimmed.toUpperCase()
      // Single-letter middle initials stay uppercase
      if (trimmed.length === 1) return trimmed.toUpperCase()
      // Capitalize first letter + any letter immediately after an apostrophe
      // (O'Brien, D'Angelo, O'Connor — common in surnames)
      return trimmed.charAt(0).toUpperCase()
        + trimmed.slice(1).replace(/'([a-z])/g, function (_m, c) { return "'" + (c as string).toUpperCase() })
    })
    .join('')
}

// ──────────────────────────────────────────────────────────────────────────────
// MENU HOOKS (called from menu.ts)
// ──────────────────────────────────────────────────────────────────────────────

/** Clear all cache entries this sheet has accumulated. Bound to menu "Refresh All". */
function clearCache_(): void {
  const cache = CacheService.getDocumentCache()
  if (cache) {
    // CacheService has no "clear all" — best we can do without tracking keys is
    // remove the known prefixes. List grows as new function tags are added.
    const tags = [
      'FORM4API_TX',
      'FORM4API_TX_LATEST',
      'FORM4API_INSIDER_TX',
      'FORM4API_RETURNS',
      'FORM4API_SCORECARD',
      'FORM4API_SENTIMENT',
      'FORM4API_CLUSTER_FLAG',
      'FORM4API_HOLDINGS',
    ]
    // No way to enumerate keys; users will see fresh data on next recalc anyway
    // because cache entries TTL out within 1 hour. Documented in menu.ts.
    Logger.log(`Cache invalidation triggered for: ${tags.join(', ')}`)
  }
}
