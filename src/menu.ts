// "Form4API" menu added to the spreadsheet's top bar on open.
// Phase 1 has only the essentials: Set API Key + About. Phase 2 adds Refresh All.

const PROP_API_KEY_M = 'form4api_key'

function onOpen(): void {
  SpreadsheetApp.getUi()
    .createMenu('Form4API')
    .addItem('Set API Key…', 'setApiKey_')
    .addItem('Clear API Key', 'clearApiKey_')
    .addSeparator()
    .addItem('About', 'showAbout_')
    .addToUi()
}

function setApiKey_(): void {
  const ui = SpreadsheetApp.getUi()
  const existing = PropertiesService.getDocumentProperties().getProperty(PROP_API_KEY_M)
  const prompt = existing
    ? 'API key is already set. Paste a new one to replace it, or Cancel to keep the current key.'
    : 'Paste your Form4API key (starts with fapi_ or smk_). Get one free at https://form4api.com'

  const res = ui.prompt('Form4API — Set API Key', prompt, ui.ButtonSet.OK_CANCEL)
  if (res.getSelectedButton() !== ui.Button.OK) return

  const key = res.getResponseText().trim()
  if (!key) {
    ui.alert('No key entered. Try again from the Form4API menu.')
    return
  }
  PropertiesService.getDocumentProperties().setProperty(PROP_API_KEY_M, key)
  ui.alert('API key saved. Try =FORM4API_TX("AAPL") in any cell.')
}

function clearApiKey_(): void {
  PropertiesService.getDocumentProperties().deleteProperty(PROP_API_KEY_M)
  SpreadsheetApp.getUi().alert('API key cleared.')
}

function showAbout_(): void {
  SpreadsheetApp.getUi().alert(
    'Form4API Sheets Add-on (v0.0.1 spike)\n\n' +
      'Custom functions: =FORM4API_TX(ticker, [limit])\n\n' +
      'Docs: https://form4api.com/docs\n' +
      'Source: https://github.com/theodor90/form4api-sheets-addin',
  )
}
