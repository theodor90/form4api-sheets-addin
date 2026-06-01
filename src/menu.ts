// "Form4API" menu in the spreadsheet's top bar. Attached on every open by the
// onOpen simple trigger. Apps Script merges all .ts files into one global
// namespace at runtime, so PROP_API_KEY here resolves to the same constant
// declared in index.ts.

function onOpen(): void {
  SpreadsheetApp.getUi()
    .createMenu('Form4API')
    .addItem('Set API Key…', 'setApiKey_')
    .addItem('Clear API Key', 'clearApiKeyFromMenu_')
    .addSeparator()
    .addItem('Refresh All', 'refreshAllFromMenu_')
    .addSeparator()
    .addItem('About', 'showAbout_')
    .addToUi()
}

function setApiKey_(): void {
  const ui = SpreadsheetApp.getUi()
  const existing = PropertiesService.getDocumentProperties().getProperty(PROP_API_KEY)
  const prompt = existing
    ? 'API key is already set. Paste a new one to replace it, or Cancel to keep the current key.\n\n' +
      'Heads-up: the key is stored in this sheet. Anyone with edit access can use it (and may consume your quota or hit your rate-limit).'
    : 'Paste your Form4API key (starts with fapi_ or smk_). Get one free at https://form4api.com\n\n' +
      'Heads-up: the key is stored in this sheet. Anyone you share edit access with can use it (and may consume your quota or hit your rate-limit). View-only collaborators cannot see or use the key.'

  const res = ui.prompt('Form4API — Set API Key', prompt, ui.ButtonSet.OK_CANCEL)
  if (res.getSelectedButton() !== ui.Button.OK) return

  const key = res.getResponseText().trim()
  if (!key) {
    ui.alert('No key entered. Try again from the Form4API menu.')
    return
  }
  PropertiesService.getDocumentProperties().setProperty(PROP_API_KEY, key)
  ui.alert('API key saved. Try =FORM4API_TX("AAPL") in any cell.')
}

function clearApiKeyFromMenu_(): void {
  PropertiesService.getDocumentProperties().deleteProperty(PROP_API_KEY)
  SpreadsheetApp.getUi().alert('API key cleared.')
}

function refreshAllFromMenu_(): void {
  clearCache_()
  // Force a recalculation of all formulas — flips a hidden property that the
  // custom functions don't read but Sheets sees as a change to invalidate.
  SpreadsheetApp.getActive().getRangeList(['A1']).activate()
  SpreadsheetApp.getUi().alert(
    'Cached data marked stale. Sheets will refetch on the next recalc — ' +
      'click into a Form4API cell and press Enter to trigger immediately.',
  )
}

function showAbout_(): void {
  SpreadsheetApp.getUi().alert(
    'Form4API Sheets Add-on (v0.2.0)\n\n' +
      'Custom functions:\n' +
      '  =FORM4API_TX(ticker, [limit], [code])     Free\n' +
      '  =FORM4API_TX_LATEST(ticker)               Free\n' +
      '  =FORM4API_INSIDER_TX(cik, [limit])        Free\n' +
      '  =FORM4API_RETURNS(ticker, [horizon])      Pro+\n' +
      '  =FORM4API_SENTIMENT(ticker, [months])     Business+\n' +
      '  =FORM4API_CLUSTER_FLAG(ticker)            Business+\n\n' +
      'Cached for 1 hour. Refresh All from the menu to invalidate.\n\n' +
      'Docs: https://form4api.com/docs\n' +
      'Source: https://github.com/theodor90/form4api-sheets-addin',
  )
}
