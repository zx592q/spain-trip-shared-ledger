const LEDGER_SHEET = '系統資料';
const BACKUP_SHEET = '自動備份';
const BACKUP_INTERVAL_MS = 12 * 60 * 60 * 1000;
const MAX_BACKUPS = 30;

function doGet() {
  createBackupIfDue_();
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
    createBackupIfDue_();
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

function createBackupIfDue_() {
  const properties = PropertiesService.getScriptProperties();
  const lastBackup = Number(properties.getProperty('LAST_BACKUP_AT') || 0);
  if (Date.now() - lastBackup < BACKUP_INTERVAL_MS) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    const latest = Number(properties.getProperty('LAST_BACKUP_AT') || 0);
    if (Date.now() - latest < BACKUP_INTERVAL_MS) return;
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let backupSheet = spreadsheet.getSheetByName(BACKUP_SHEET);
    if (!backupSheet) {
      backupSheet = spreadsheet.insertSheet(BACKUP_SHEET);
      backupSheet.getRange('A1:D1').setValues([['備份時間', '資料版本', '最後更新時間', '完整帳本 JSON']]);
      backupSheet.setFrozenRows(1);
    }
    const ledger = readLedger_();
    backupSheet.appendRow([new Date(), ledger.revision, ledger.updatedAt, JSON.stringify(ledger.state)]);
    const backupCount = backupSheet.getLastRow() - 1;
    if (backupCount > MAX_BACKUPS) backupSheet.deleteRows(2, backupCount - MAX_BACKUPS);
    properties.setProperty('LAST_BACKUP_AT', String(Date.now()));
  } finally {
    lock.releaseLock();
  }
}

function createBackupNow() {
  PropertiesService.getScriptProperties().deleteProperty('LAST_BACKUP_AT');
  createBackupIfDue_();
}

function jsonResponse_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
