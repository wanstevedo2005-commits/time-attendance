/**
 * Google Apps Script backend for the LIFF attendance page.
 * Store FOLDER_ID, TELEGRAM_TOKEN, TELEGRAM_CHAT_IDS, and ADMIN_API_KEY
 * in Project Settings > Script properties. Do not commit credentials.
 */

function doPost(e) {
  try {
    const data = parsePostData(e);
    if (data.action === 'admin_logs') {
      assertValidAdminKey(data.apiKey);
      return jsonResponse(getAttendanceRows(SpreadsheetApp.getActiveSpreadsheet()));
    }
    if (data.action === 'update_name') {
      assertValidAdminKey(data.apiKey);
      return updateEmployeeName(SpreadsheetApp.getActiveSpreadsheet(), data.userId, data.newName);
    }
    data.userId = verifyLineAccessToken(data.accessToken).userId;
    return jsonResponse(processUploadData(data));
  }
  catch (error) {
    console.error(error && error.stack ? error.stack : error);
    return jsonResponse({ status: 'error', message: error.message || String(error) });
  }
}

/** Compatibility entry point for an HTML file embedded in Apps Script. */
function uploadData(data) {
  try { return processUploadData(data); }
  catch (error) { return { status: 'error', message: error.message || String(error) }; }
}

function processUploadData(data) {
  validateUploadData(data);
  const userId = String(data.userId).trim();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = spreadsheet.getSheetByName('Log');
  if (!logSheet) throw new Error('Sheet "Log" was not found.');
  const now = new Date();
  const name = findEmployeeName(spreadsheet, logSheet, userId);
  const time = Utilities.formatDate(now, getTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  const lateStatus = calculateLateStatus(now);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  let imageFile;
  try {
    if (hasAttendanceForDate(logSheet, userId, now)) {
      throw new Error('วันนี้คุณลงเวลาเข้างานแล้ว ไม่สามารถบันทึกซ้ำได้');
    }
    imageFile = saveAttendancePhoto(data.image, name, now);
    logSheet.appendRow([time, userId, name, imageFile.getUrl(), lateStatus]);
  } finally {
    lock.releaseLock();
  }
  const telegramResults = sendTelegramNotification(name, time, imageFile, userId, lateStatus);
  return {
    status: 'success',
    name: name,
    isLate: lateStatus === 1,
    telegramOk: telegramResults.length > 0 && telegramResults.every(function (result) { return result.ok; })
  };
}

function doGet(e) {
  if (e && e.parameter && e.parameter.adminBridge === '1') return getAdminBridgePage();
  return jsonResponse({ status: 'error', message: 'Admin API accepts POST requests only.' });
}

function getAdminBridgePage() {
  const allowedOrigin = 'https://wanstevedo2005-commits.github.io';
  const html = '<!doctype html><html><body><script>' +
    'const ALLOWED_ORIGIN=' + JSON.stringify(allowedOrigin) + ';' +
    'window.addEventListener("message",function(event){' +
      'if(event.origin!==ALLOWED_ORIGIN||!event.data||event.data.type!=="admin-request")return;' +
      'const source=event.source,id=event.data.id;' +
      'google.script.run.withSuccessHandler(function(data){source.postMessage({type:"admin-response",id:id,data:data},event.origin);})' +
      '.withFailureHandler(function(error){source.postMessage({type:"admin-response",id:id,error:error&&error.message?error.message:String(error)},event.origin);})' +
      '.adminBridgeRequest(event.data.request);' +
    '});' +
    'window.top.postMessage({type:"admin-bridge-ready"},ALLOWED_ORIGIN);' +
    '<\/script></body></html>';
  return HtmlService.createHtmlOutput(html)
    .setTitle('Admin API Bridge')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function adminBridgeRequest(request) {
  const data = request || {};
  assertValidAdminKey(data.apiKey);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (data.action === 'admin_logs') return getAttendanceRows(ss);
  if (data.action === 'update_name') return updateEmployeeNameData(ss, data.userId, data.newName);
  throw new Error('Unsupported admin action.');
}

function updateEmployeeName(ss, userId, newName) {
  return jsonResponse(updateEmployeeNameData(ss, userId, newName));
}

function updateEmployeeNameData(ss, userId, newName) {
  const id = String(userId || '').trim();
  const name = validateEmployeeName(newName);
  if (!id || !name) return { status: 'error', message: 'userId and newName are required.' };
  const employeeSheet = ss.getSheetByName('Employee');
  if (employeeSheet) {
    const employees = employeeSheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < employees.length; i++) {
      if (String(employees[i][0]).trim() === id) {
        employeeSheet.getRange(i + 1, 2).setValue(name); found = true; break;
      }
    }
    if (!found) employeeSheet.appendRow([id, name]);
  }
  const logSheet = ss.getSheetByName('Log');
  if (logSheet) {
    const rows = logSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]).trim() === id) logSheet.getRange(i + 1, 3).setValue(name);
    }
  }
  return { status: 'success' };
}

function getAttendanceLog(ss) {
  return jsonResponse(getAttendanceRows(ss));
}

function getAttendanceRows(ss) {
  const logSheet = ss.getSheetByName('Log');
  if (!logSheet) return [];
  return logSheet.getDataRange().getValues().slice(1).map(function (row) {
    return { time: formatLogDate(row[0]), userId: String(row[1]), name: String(row[2]),
      image: String(row[3]), isLate: row[4] === 1 || String(row[4]).toLowerCase() === 'true' };
  }).reverse();
}

/** Sends independently to every configured chat; a failure does not stop later recipients. */
function sendTelegramNotification(name, time, imageFile, userId, lateStatus) {
  const token = getRequiredProperty('TELEGRAM_TOKEN');
  const message = ['📢 มีการลงเวลาเข้างาน!', '👤 ชื่อ: ' + name, '🆔 LINE ID: ' + userId,
    '⏰ เวลา: ' + time, lateStatus === 1 ? '⚠️ มาสาย' : '✅ ตรงเวลา'].join('\n');
  const url = 'https://api.telegram.org/bot' + token + '/sendPhoto';
  return getTelegramChatIds().map(function (chatId) {
    try {
      const response = UrlFetchApp.fetch(url, { method: 'post', muteHttpExceptions: true,
        payload: { chat_id: chatId, caption: message, photo: imageFile.getBlob() } });
      const body = response.getContentText();
      const ok = response.getResponseCode() >= 200 && response.getResponseCode() < 300;
      console.log('Telegram response for ' + chatId + ': ' + response.getResponseCode() + ' ' + body);
      return { chatId: chatId, ok: ok, statusCode: response.getResponseCode(), response: body };
    } catch (error) {
      console.error('Telegram error for ' + chatId + ': ' + error);
      return { chatId: chatId, ok: false, error: String(error) };
    }
  });
}

function testTelegramDirect() {
  const token = getRequiredProperty('TELEGRAM_TOKEN');
  const url = 'https://api.telegram.org/bot' + token + '/sendMessage';
  return getTelegramChatIds().map(function (chatId) {
    const response = UrlFetchApp.fetch(url, { method: 'post', muteHttpExceptions: true,
      payload: { chat_id: chatId, text: 'TEST MESSAGE FROM TIME ATTENDANCE' } });
    Logger.log('Telegram response for ' + chatId + ': ' + response.getResponseCode() + ' ' + response.getContentText());
    return response.getResponseCode();
  });
}

/** Existing dashboard generator retained with the current Log columns. */
function createAutoDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Dashboard_Auto');
  if (!sheet) sheet = ss.insertSheet('Dashboard_Auto'); else sheet.clear();
  sheet.getRange('A1:D1').setValues([['ชื่อพนักงาน', 'จำนวนวันทำงาน', 'มาสาย (ครั้ง)', 'เปอร์เซ็นต์การสาย']])
    .setFontWeight('bold').setBackground('#f3f3f3');
  sheet.getRange('A2').setFormula('=UNIQUE(Log!C2:C)');
  sheet.getRange('B2').setFormula('=BYROW(A2:A,LAMBDA(emp,IF(emp="","",COUNTIF(Log!C2:C,emp))))');
  sheet.getRange('C2').setFormula('=BYROW(A2:A,LAMBDA(emp,IF(emp="","",SUMPRODUCT((Log!C2:C=emp)*(Log!E2:E=1)))))');
  sheet.getRange('D2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(C2:C/B2:B,0)))');
  sheet.getRange('D2:D').setNumberFormat('0.00%'); sheet.autoResizeColumns(1, 4);
}

function parsePostData(e) {
  if (!e || !e.postData || !e.postData.contents) throw new Error('Request body is missing.');
  try { return JSON.parse(e.postData.contents); } catch (error) { throw new Error('Request body must be valid JSON.'); }
}
function validateUploadData(data) {
  if (!data || !data.userId) throw new Error('LINE user ID is required.');
  if (!data.image || !/^data:image\/(jpeg|jpg|png);base64,/i.test(data.image)) throw new Error('A JPEG or PNG image is required.');
}
function verifyLineAccessToken(accessToken) {
  if (!accessToken) throw new Error('LINE session has expired. Please reopen this page in LINE.');
  const response = UrlFetchApp.fetch('https://api.line.me/v2/profile', {
    headers: { Authorization: 'Bearer ' + accessToken },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) throw new Error('Unable to verify your LINE account. Please reopen this page in LINE.');
  return JSON.parse(response.getContentText());
}
function findEmployeeName(spreadsheet, logSheet, userId) {
  const id = userId.toLowerCase();
  const employeeSheet = spreadsheet.getSheetByName('Employee');
  if (employeeSheet) {
    const employees = employeeSheet.getDataRange().getValues();
    for (let i = 1; i < employees.length; i++) {
      const employeeId = String(employees[i][0]).trim().toLowerCase();
      const employeeName = String(employees[i][1]).trim();
      if (employeeId === id && employeeName) return employeeName;
    }
  }
  const rows = logSheet.getDataRange().getValues();
  for (let i = rows.length - 1; i > 0; i--) {
    const foundId = String(rows[i][1]).trim().toLowerCase(); const name = String(rows[i][2]).trim();
    if (foundId === id && name && name !== 'พนักงานใหม่ (ยังไม่ระบุชื่อ)') return name;
  }
  return 'พนักงานใหม่ (ยังไม่ระบุชื่อ)';
}
function hasAttendanceForDate(logSheet, userId, now) {
  const targetId = String(userId).trim().toLowerCase();
  const targetDate = Utilities.formatDate(now, getTimeZone(), 'dd/MM/yyyy');
  const rows = logSheet.getDataRange().getValues();
  for (let i = rows.length - 1; i > 0; i--) {
    const rowId = String(rows[i][1]).trim().toLowerCase();
    if (rowId !== targetId) continue;
    const rowDate = rows[i][0] instanceof Date
      ? Utilities.formatDate(rows[i][0], getTimeZone(), 'dd/MM/yyyy')
      : String(rows[i][0]).trim().split(' ')[0];
    if (rowDate === targetDate) return true;
  }
  return false;
}
function saveAttendancePhoto(imageData, name, now) {
  const contentType = imageData.split(',')[0].split(':')[1].split(';')[0];
  const bytes = Utilities.base64Decode(imageData.split(',')[1]);
  const filename = 'CheckIn_' + sanitizeFilename(name) + '_' + now.getTime() + '.jpg';
  const file = DriveApp.getFolderById(getRequiredProperty('FOLDER_ID')).createFile(Utilities.newBlob(bytes, contentType, filename));
  // Files remain private; grant managers access to the folder in Google Drive if needed.
  return file;
}
function calculateLateStatus(now) {
  const time = Utilities.formatDate(now, getTimeZone(), 'HH:mm').split(':');
  return Number(time[0]) > 8 || (Number(time[0]) === 8 && Number(time[1]) > 40) ? 1 : 0;
}
function getTelegramChatIds() {
  const ids = getRequiredProperty('TELEGRAM_CHAT_IDS').split(',').map(function (id) { return id.trim(); }).filter(Boolean);
  if (!ids.length) throw new Error('TELEGRAM_CHAT_IDS has no recipient IDs.'); return ids;
}
function assertValidAdminKey(apiKey) {
  const expected = getRequiredProperty('ADMIN_API_KEY');
  if (!apiKey || String(apiKey) !== expected) throw new Error('Unauthorized.');
}
function validateEmployeeName(value) {
  const name = String(value || '').trim();
  if (!name) return '';
  if (name.length > 100) throw new Error('Employee name must not exceed 100 characters.');
  if (/^[=+\-@]/.test(name) || /[\u0000-\u001F\u007F]/.test(name)) {
    throw new Error('Employee name contains unsupported characters.');
  }
  return name;
}
function getRequiredProperty(key) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) throw new Error('Missing Script Property: ' + key); return value;
}
function getTimeZone() { return PropertiesService.getScriptProperties().getProperty('TIMEZONE') || 'Asia/Bangkok'; }
function formatLogDate(value) {
  if (value instanceof Date) return Utilities.formatDate(value, getTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const parsed = new Date(String(value)); return isNaN(parsed.getTime()) ? String(value) : Utilities.formatDate(parsed, getTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}
function sanitizeFilename(value) { return String(value).replace(/[\\/:*?"<>|]/g, '_').slice(0, 100); }
function jsonResponse(value) { return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); }
