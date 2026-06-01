// Phase 1 spike — prove =FORM4API_TX("AAPL") returns real data in a Google Sheet.
// Single custom function + API-key handling via DocumentProperties.
// If this works end-to-end, expand to the full 6-function surface (Phase 2).

const API_BASE = 'https://api.form4api.com'
const PROP_API_KEY = 'form4api_key'

/**
 * Insider transactions for a ticker.
 *
 * @param {string} ticker Stock symbol, e.g. "AAPL"
 * @param {number} limit  Max rows to return (default 20, max 100)
 * @return                2D array of [date, insider, code, shares, price, value]
 * @customfunction
 */
function FORM4API_TX(ticker: string, limit?: number): (string | number)[][] {
  if (!ticker || typeof ticker !== 'string') {
    throw new Error('FORM4API_TX: pass a ticker like "AAPL"')
  }
  const key = getApiKey_()
  const perPage = Math.min(Math.max(limit || 20, 1), 100)

  const url = `${API_BASE}/v1/transactions?ticker=${encodeURIComponent(ticker.toUpperCase())}&per_page=${perPage}`
  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'X-Api-Key': key, 'User-Agent': 'form4api-sheets-addin/0.0.1' },
    muteHttpExceptions: true,
  })
  const code = res.getResponseCode()
  if (code === 401) throw new Error('FORM4API_TX: invalid API key — open the Form4API menu → Set API Key')
  if (code === 429) throw new Error('FORM4API_TX: rate-limited — wait a minute and retry')
  if (code >= 400) throw new Error(`FORM4API_TX: API error ${code}: ${res.getContentText().slice(0, 200)}`)

  const transactions = JSON.parse(res.getContentText()) as Array<{
    transactionDate: string
    insiderName: string
    transactionCode: string
    sharesAmount: number
    pricePerShare: number | null
    totalValue: number | null
  }>

  if (transactions.length === 0) return [['No transactions found', '', '', '', '', '']]

  const header: (string | number)[] = ['Date', 'Insider', 'Code', 'Shares', 'Price', 'Value']
  const rows: (string | number)[][] = transactions.map((t) => [
    t.transactionDate,
    t.insiderName,
    t.transactionCode,
    t.sharesAmount,
    t.pricePerShare ?? '',
    t.totalValue ?? '',
  ])
  return [header, ...rows]
}

function getApiKey_(): string {
  const key = PropertiesService.getDocumentProperties().getProperty(PROP_API_KEY)
  if (!key) {
    throw new Error(
      'FORM4API_TX: API key not set. Open the Form4API menu (top bar) → Set API Key. Get a free key at https://form4api.com',
    )
  }
  return key
}
