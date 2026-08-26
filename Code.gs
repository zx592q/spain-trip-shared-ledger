const LEDGER_SHEET = '系統資料';

function doGet() {
  return jsonResponse_(readLedger_());
}

function doPost(event) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const payload = JSON.parse(event.postData.contents || '{}');
    const current = readLedger_();
    if (!payload.state || !Array.isArray(payload.state.people) || !Array.isArray(payload.state.expenses) || !Array.isArray(payload.state.incomes)) {
      return jsonResponse_({ error: '帳本資料格式不正確' });
    }
    if (!Number.isInteger(payload.revision) || payload.revision !== current.revision) {
      return jsonResponse_({ error: '帳本已由其他旅伴更新，請重新整理後再試一次', conflict: true, revision: current.revision });
    }
    const nextRevision = current.revision + 1;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LEDGER_SHEET);
    sheet.getRange('B2:D2').setValues([[nextRevision, new Date().toISOString(), JSON.stringify(payload.state)]]);
    SpreadsheetApp.flush();
    return jsonResponse_({ revision: nextRevision });
  } finally {
    lock.releaseLock();
  }
}

function readLedger_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LEDGER_SHEET);
  if (!sheet) throw new Error('找不到「系統資料」工作表');
  const values = sheet.getRange('B2:D2').getValues()[0];
  return { revision: Number(values[0] || 0), updatedAt: values[1] || '', state: JSON.parse(values[2] || '{"people":[],"expenses":[],"incomes":[],"categories":[]}') };
}

function jsonResponse_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
