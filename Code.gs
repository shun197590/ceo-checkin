// ============================================================
// 115年度高雄軟體園區CEO聯誼會
// 完整報名＋QR報到系統
// 架構：Apps Script 控制頁 + GitHub 連續相機掃描頁
// 通訊：window.postMessage + google.script.run
// 不使用 doPost()
// ============================================================
 
const CONFIG = {
  EVENT_NAME: '115年度高雄軟體園區CEO聯誼會',
  MAX_PEOPLE: 80,
  TIMEZONE: 'Asia/Taipei',
  DAILY_SUMMARY_HOUR: 8,
 
  FORM_URL:
    'https://docs.google.com/forms/d/1x6tuN-Sf3ijR3z8I-qeo5Rh7KECyDp3hHnbFGQg93dw/edit',
 
  SPREADSHEET_ID:
    '1-EFuZGivOEESxtwmNdjy8krUBSCNhrsFWylJjm90CAo',
 
  SHEET_NAME: '表單回覆 1',
  // 主要管理員（系統錯誤、QR 測試等仍以此信箱為主要收件人）
  ADMIN_EMAIL: 'shun@bip.gov.tw',
 
  // 其他管理成員：每日報名摘要＋報名階段通知會一併寄送。
  // 尚未設定的欄位請保留空字串即可；日後只需填入 Email，不必再改寄信函式。
  ADDITIONAL_ADMIN_EMAILS: [
    '',
    '',
    ''
  ],
 
  SENDER_NAME: '高雄軟體園區辦公室',
 
  // ★ 必填：請改成您實際發布的 GitHub Pages HTTPS 網址。
  // 例如：https://your-account.github.io/ceo-checkin/
  SCANNER_URL: 'https://shun197590.github.io/ceo-checkin/scanner/',
 
  // Google 表單題目名稱（必須與表單完全一致）
  FIELD_COMPANY: '公司名稱',
  FIELD_NAME: '姓名',
  FIELD_TITLE: '職稱',
  FIELD_EMAIL: 'E-mail',
  FIELD_MOBILE: '手機號碼',
  FIELD_MEAL: '餐食需求',
  FIELD_PRIVACY: '個資同意聲明',
 
  // 系統自訂欄位名稱
  HEADER_REG_STATUS: '報名狀態',
  HEADER_CHECKIN_CODE: '報到碼',
  HEADER_QR_CODE: 'QR Code',
  HEADER_CHECKIN_STATUS: '報到狀態',
  HEADER_CHECKIN_TIME: '報到時間',
  HEADER_QR_SEND_STATUS: 'QR寄送狀態',
  HEADER_QR_SEND_TIME: 'QR寄送時間',
  HEADER_BADGE_NUMBER: '名牌編號'
};
 
 
// ============================================================
// 一次性設定
// 1. 補齊系統欄位
// 2. 建立「提交表單時」觸發器
// 3. 建立每日摘要觸發器
// ============================================================
function setupSystem() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = getRegistrationSheet_();
 
  ensureCustomHeaders_(sheet);
 
  const managed = new Set([
    'sendRegistrationConfirmation',
    'sendDailyRegistrationSummary'
  ]);
 
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (managed.has(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
 
  ScriptApp
    .newTrigger('sendRegistrationConfirmation')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
 
  ScriptApp
    .newTrigger('sendDailyRegistrationSummary')
    .timeBased()
    .atHour(CONFIG.DAILY_SUMMARY_HOUR)
    .everyDays(1)
    .inTimezone(CONFIG.TIMEZONE)
    .create();
 
  Logger.log('系統設定完成。');
  Logger.log('Apps Script 控制頁：' + (ScriptApp.getService().getUrl() || '尚未部署 Web App'));
  Logger.log('GitHub Scanner：' + CONFIG.SCANNER_URL);
}
 
 
// ============================================================
// Google 表單提交主程式
// 只能由「從試算表 → 提交表單時」觸發器執行
// ============================================================
function sendRegistrationConfirmation(e) {
  if (!e || !e.range || !e.namedValues) {
    Logger.log('此函式必須由「提交表單時」觸發，請勿手動執行。');
    return;
  }
 
  const lock = LockService.getScriptLock();
 
  try {
    lock.waitLock(30000);
 
    const sheet = e.range.getSheet();
    const row = e.range.getRow();
    const namedValues = e.namedValues;
 
    ensureCustomHeaders_(sheet);
    const headers = getHeaderMap_(sheet);
 
    const statusCol = getColumn_(headers, CONFIG.HEADER_REG_STATUS);
    const codeCol = getColumn_(headers, CONFIG.HEADER_CHECKIN_CODE);
    const checkinStatusCol = getColumn_(headers, CONFIG.HEADER_CHECKIN_STATUS);
    const badgeCol = getColumn_(headers, CONFIG.HEADER_BADGE_NUMBER);
 
    const existingStatus = String(
      sheet.getRange(row, statusCol).getDisplayValue() || ''
    ).trim();
 
    if (existingStatus !== '') {
      Logger.log('第 ' + row + ' 列已處理：' + existingStatus);
      return;
    }
 
    const company = getNamedValue_(namedValues, CONFIG.FIELD_COMPANY).trim();
    const name = getNamedValue_(namedValues, CONFIG.FIELD_NAME).trim();
    const title = getNamedValue_(namedValues, CONFIG.FIELD_TITLE).trim();
    const email = getNamedValue_(namedValues, CONFIG.FIELD_EMAIL).trim();
    const mobile = getNamedValue_(namedValues, CONFIG.FIELD_MOBILE).trim();
    const meal = getNamedValue_(namedValues, CONFIG.FIELD_MEAL).trim();
    const privacyConsent = getNamedValue_(namedValues, CONFIG.FIELD_PRIVACY).trim();
 
    if (!privacyConsent) {
      sheet.getRange(row, statusCol).setValue('個資未同意');
      SpreadsheetApp.flush();
      return;
    }
 
    // 每家公司限報1名：只要前面已有「報名成功」的同公司即視為重複。
    if (hasPriorSuccessfulCompany_(sheet, row, company, headers)) {
      sheet.getRange(row, statusCol).setValue('重複報名');
      SpreadsheetApp.flush();
 
      safeSendEmail_({
        to: email,
        subject: '【報名未成立】' + CONFIG.EVENT_NAME,
        body:
          name + ' ' + title + ' 您好：\n\n' +
          '感謝您填寫「' + CONFIG.EVENT_NAME + '」報名資料。\n\n' +
          '經系統檢查，「' + company + '」已有正式報名紀錄。\n' +
          '因本活動每家公司限報1名，本次重複填寫不列入正式報名名額。\n\n' +
          '如需更換出席人員或修正資料，請與高雄軟體園區辦公室聯繫。\n\n' +
          '經濟部產業園區管理局\n高雄軟體園區辦公室'
      });
      return;
    }
 
    const successCountBefore = countStatus_(sheet, CONFIG.HEADER_REG_STATUS, '報名成功');
 
    if (successCountBefore >= CONFIG.MAX_PEOPLE) {
      sheet.getRange(row, statusCol).setValue('候補／額滿');
      SpreadsheetApp.flush();
      safeCloseForm_();
 
      safeSendEmail_({
        to: email,
        subject: '【報名額滿】' + CONFIG.EVENT_NAME,
        body:
          name + ' ' + title + ' 您好：\n\n' +
          '感謝您填寫「' + CONFIG.EVENT_NAME + '」報名資料。\n\n' +
          '本活動正式報名名額80名目前已額滿，本次資料暫不列入正式報名名單。\n' +
          '如後續有名額釋出，將再另行通知。\n\n' +
          '經濟部產業園區管理局\n高雄軟體園區辦公室'
      });
      return;
    }
 
    // 正式報名成功
    sheet.getRange(row, statusCol).setValue('報名成功');
 
    // 名牌編號依「報名成功」順序固定編 001～080。
    const badgeNumber = String(successCountBefore + 1).padStart(3, '0');
    sheet.getRange(row, badgeCol).setNumberFormat('@').setValue(badgeNumber);
 
    // 建立唯一報到碼。
    let checkinCode = String(sheet.getRange(row, codeCol).getDisplayValue() || '').trim();
    if (checkinCode === '') {
      checkinCode = generateCheckinCode_(row);
      sheet.getRange(row, codeCol).setNumberFormat('@').setValue(checkinCode);
    }
 
    // 預設未報到。
    sheet.getRange(row, checkinStatusCol).setValue('未報到');
 
    SpreadsheetApp.flush();
 
    // 試算表內先建立 QR Code；正式 Email 等報名截止後再統一寄。
    try {
      createQRCodeInSheet_(sheet, row, checkinCode, headers);
    } catch (qrError) {
      notifyAdminError_('試算表 QR Code 建立失敗', qrError, row);
    }
 
    // 報名當下只寄「報名成功通知」，不寄 QR Code。
    sendRegistrationReceivedEmail_(
      email,
      company,
      name,
      title,
      mobile,
      meal
    );
 
    const successCountAfter = successCountBefore + 1;
 
    if (successCountAfter >= CONFIG.MAX_PEOPLE) {
      safeCloseForm_();
    }
 
    sendAdminNotification_(
      successCountAfter,
      company,
      name,
      title,
      meal
    );
 
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}
 
 
// ============================================================
// 報名成功通知（當下不寄 QR Code）
// ============================================================
function sendRegistrationReceivedEmail_(email, company, name, title, mobile, meal) {
  const subject = '【報名成功】' + CONFIG.EVENT_NAME;
 
  const body =
    name + ' ' + title + ' 您好：\n\n' +
    '感謝您報名「' + CONFIG.EVENT_NAME + '」，您的報名已成功完成。\n\n' +
    '【報名資料】\n' +
    '公司名稱：' + company + '\n' +
    '姓名：' + name + '\n' +
    '職稱：' + title + '\n' +
    '手機號碼：' + mobile + '\n' +
    '餐食需求：' + meal + '\n\n' +
    '您的報名資料已列入正式名單。\n\n' +
    '個人專屬報到 QR Code 將於報名截止、名單確認完成後另行寄送，敬請留意電子郵件。\n\n' +
    '經濟部產業園區管理局\n高雄軟體園區辦公室';
 
  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: body,
    name: CONFIG.SENDER_NAME
  });
}
 
 
// ============================================================
// 報名截止後：統一寄送所有正式參加者 QR Code
// 已寄送者不會重複寄。
// ============================================================
function sendAllQRCodesAfterDeadline() {
  const sheet = getRegistrationSheet_();
  ensureCustomHeaders_(sheet);
  const headers = getHeaderMap_(sheet);
  const lastRow = sheet.getLastRow();
 
  if (lastRow < 2) {
    Logger.log('目前沒有報名資料。');
    return;
  }
 
  const regCol = getColumn_(headers, CONFIG.HEADER_REG_STATUS);
  const codeCol = getColumn_(headers, CONFIG.HEADER_CHECKIN_CODE);
  const qrSendStatusCol = getColumn_(headers, CONFIG.HEADER_QR_SEND_STATUS);
  const qrSendTimeCol = getColumn_(headers, CONFIG.HEADER_QR_SEND_TIME);
  const badgeCol = getColumn_(headers, CONFIG.HEADER_BADGE_NUMBER);
 
  const companyCol = getColumn_(headers, CONFIG.FIELD_COMPANY);
  const nameCol = getColumn_(headers, CONFIG.FIELD_NAME);
  const titleCol = getColumn_(headers, CONFIG.FIELD_TITLE);
  const emailCol = getColumn_(headers, CONFIG.FIELD_EMAIL);
  const mobileCol = getColumn_(headers, CONFIG.FIELD_MOBILE);
  const mealCol = getColumn_(headers, CONFIG.FIELD_MEAL);
 
  let needSendCount = 0;
 
  for (let row = 2; row <= lastRow; row++) {
    const regStatus = String(sheet.getRange(row, regCol).getDisplayValue() || '').trim();
    const sendStatus = String(sheet.getRange(row, qrSendStatusCol).getDisplayValue() || '').trim();
    if (regStatus === '報名成功' && sendStatus !== '已寄送') {
      needSendCount++;
    }
  }
 
  if (needSendCount === 0) {
    Logger.log('目前沒有尚待寄送的 QR Code。');
    return;
  }
 
  const remainingQuota = MailApp.getRemainingDailyQuota();
  if (remainingQuota < needSendCount) {
    throw new Error(
      '今日 MailApp 剩餘寄信額度不足。需要寄送 ' +
      needSendCount + ' 人，但目前只剩 ' + remainingQuota + ' 人額度。'
    );
  }
 
  let sent = 0;
  let failed = 0;
 
  for (let row = 2; row <= lastRow; row++) {
    const regStatus = String(sheet.getRange(row, regCol).getDisplayValue() || '').trim();
    const oldSendStatus = String(sheet.getRange(row, qrSendStatusCol).getDisplayValue() || '').trim();
 
    if (regStatus !== '報名成功' || oldSendStatus === '已寄送') {
      continue;
    }
 
    try {
      const company = String(sheet.getRange(row, companyCol).getDisplayValue() || '').trim();
      const name = String(sheet.getRange(row, nameCol).getDisplayValue() || '').trim();
      const title = String(sheet.getRange(row, titleCol).getDisplayValue() || '').trim();
      const email = String(sheet.getRange(row, emailCol).getDisplayValue() || '').trim();
      const mobile = String(sheet.getRange(row, mobileCol).getDisplayValue() || '').trim();
      const meal = String(sheet.getRange(row, mealCol).getDisplayValue() || '').trim();
      const badgeNumber = String(sheet.getRange(row, badgeCol).getDisplayValue() || '').trim();
 
      if (!email) {
        throw new Error('E-mail 空白。');
      }
 
      let checkinCode = String(sheet.getRange(row, codeCol).getDisplayValue() || '').trim();
      if (!checkinCode) {
        checkinCode = generateCheckinCode_(row);
        sheet.getRange(row, codeCol).setNumberFormat('@').setValue(checkinCode);
      }
 
      createQRCodeInSheet_(sheet, row, checkinCode, headers);
 
      sendQRCodeEmail_(
        email,
        company,
        name,
        title,
        mobile,
        meal,
        badgeNumber,
        checkinCode
      );
 
      sheet.getRange(row, qrSendStatusCol).setValue('已寄送');
      sheet.getRange(row, qrSendTimeCol).setValue(new Date());
      SpreadsheetApp.flush();
 
      sent++;
      Utilities.sleep(250);
 
    } catch (error) {
      failed++;
      sheet.getRange(row, qrSendStatusCol).setValue('寄送失敗');
      Logger.log('第 ' + row + ' 列 QR 寄送失敗：' + error.message);
    }
  }
 
  safeSendEmail_({
    to: CONFIG.ADMIN_EMAIL,
    subject: '【QR Code寄送完成】' + CONFIG.EVENT_NAME,
    body:
      'QR Code 統一寄送完成。\n\n' +
      '成功寄送：' + sent + ' 人\n' +
      '寄送失敗：' + failed + ' 人\n\n' +
      '如有失敗，請查看試算表「QR寄送狀態」欄。'
  });
}
 
 
// ============================================================
// 寄送個人專屬 QR Code Email
// QR 只編碼「報到碼」，不含姓名、Email 等個資。
// ============================================================
function sendQRCodeEmail_(email, company, name, title, mobile, meal, badgeNumber, checkinCode) {
  const qrBlob = fetchQrBlob_(checkinCode);
  const subject = '【專屬報到QR Code】' + CONFIG.EVENT_NAME;
 
  const body =
    name + ' ' + title + ' 您好：\n\n' +
    '「' + CONFIG.EVENT_NAME + '」個人專屬報到 QR Code 如下。\n\n' +
    '【報名資料】\n' +
    '公司名稱：' + company + '\n' +
    '姓名：' + name + '\n' +
    '職稱：' + title + '\n' +
    '手機號碼：' + mobile + '\n' +
    '餐食需求：' + meal + '\n' +
    '名牌編號：' + badgeNumber + '\n\n' +
    '活動當日請開啟本信件，向現場工作人員出示 QR Code。\n' +
    '此 QR Code 為個人專屬報到憑證，請勿任意轉傳。\n\n' +
    '經濟部產業園區管理局\n高雄軟體園區辦公室';
 
  const htmlBody =
    '<div style="max-width:640px;margin:0 auto;padding:12px;font-family:Arial,\'Noto Sans TC\',\'Microsoft JhengHei\',sans-serif;color:#222;line-height:1.8;font-size:16px;">' +
      '<p>' + escapeHtml_(name) + ' ' + escapeHtml_(title) + ' 您好：</p>' +
      '<p>以下為 <strong>「' + escapeHtml_(CONFIG.EVENT_NAME) + '」</strong> 個人專屬報到 QR Code。</p>' +
      '<div style="background:#f5f7fa;border-radius:10px;padding:18px 20px;margin:22px 0;">' +
        '<strong>公司名稱：</strong>' + escapeHtml_(company) + '<br>' +
        '<strong>姓名：</strong>' + escapeHtml_(name) + '<br>' +
        '<strong>職稱：</strong>' + escapeHtml_(title) + '<br>' +
        '<strong>手機號碼：</strong>' + escapeHtml_(mobile) + '<br>' +
        '<strong>餐食需求：</strong>' + escapeHtml_(meal) + '<br>' +
        '<strong>名牌編號：</strong>' + escapeHtml_(badgeNumber) +
      '</div>' +
      '<div style="border:2px solid #1a73e8;border-radius:14px;padding:22px;text-align:center;">' +
        '<div style="font-size:22px;font-weight:bold;">專屬活動報到 QR Code</div>' +
        '<p style="color:#555;">活動當日請向現場工作人員出示。</p>' +
        '<img src="cid:registrationQRCode" width="300" height="300" style="display:block;margin:12px auto;max-width:100%;">' +
        '<div style="color:#d93025;font-weight:bold;">※ 個人專屬報到憑證，請勿任意轉傳。</div>' +
      '</div>' +
      '<p>若郵件系統無法顯示 QR Code，現場仍可由工作人員以公司名稱或姓名搜尋報到。</p>' +
      '<p>經濟部產業園區管理局<br>高雄軟體園區辦公室</p>' +
    '</div>';
 
  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: body,
    htmlBody: htmlBody,
    inlineImages: {
      registrationQRCode: qrBlob
    },
    name: CONFIG.SENDER_NAME
  });
}
 
 
// ============================================================
// Web App 控制頁
// 工作人員先開 Apps Script /exec，再按「啟動連續掃碼器」。
// GitHub 頁不直接存取 Google Sheets。
// ============================================================
function doGet() {
  return HtmlService
    .createHtmlOutput(buildControlPage_())
    .setTitle(CONFIG.EVENT_NAME + '－報到控制頁')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
 
 
function buildControlPage_() {
  const eventNameJson = safeJsonForHtml_(CONFIG.EVENT_NAME);
  const scannerUrlJson = safeJsonForHtml_(CONFIG.SCANNER_URL);
 
  return `<!DOCTYPE html>
<html lang="zh-Hant">
<head>
  <base target="_top">
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    *{box-sizing:border-box}
    body{margin:0;padding:18px;background:#f4f6f8;font-family:Arial,"Noto Sans TC","Microsoft JhengHei",sans-serif;color:#222}
    .wrap{max-width:680px;margin:0 auto}
    .card{background:#fff;border-radius:16px;padding:20px;margin-bottom:16px;box-shadow:0 2px 12px rgba(0,0,0,.08)}
    h1{font-size:24px;line-height:1.4;margin:0 0 6px}
    .sub{color:#666;margin-bottom:18px}
    button{width:100%;padding:15px;border:0;border-radius:10px;font-size:18px;font-weight:700;background:#1a73e8;color:#fff;cursor:pointer;margin-top:10px}
    button.secondary{background:#5f6368}
    .status{padding:14px;border-radius:10px;background:#eef3fd;line-height:1.7;margin-top:12px;word-break:break-word}
    .ok{background:#e6f4ea}.warn{background:#fef7e0}.err{background:#fce8e6}
    .big{font-size:26px;font-weight:800}
    .note{color:#666;line-height:1.7}
  </style>
</head>
<body>
<div class="wrap">
  <div class="card">
    <h1>${escapeHtml_(CONFIG.EVENT_NAME)}</h1>
    <div class="sub">Apps Script 報到控制頁（請勿關閉此頁）</div>
    <p class="note">請先在此頁按「啟動 GitHub 連續掃碼器」。掃碼器會開在另一個分頁，真正的 Google Sheets 讀寫仍由本頁透過 <strong>google.script.run</strong> 執行。</p>
    <button id="launchBtn" onclick="launchScanner()">📷 啟動連續掃碼器</button>
    <button class="secondary" onclick="testBackend()">✅ 測試 Apps Script 後端</button>
    <div id="status" class="status">等待操作。</div>
  </div>
 
  <div class="card">
    <div class="big">使用方式</div>
    <ol class="note">
      <li>本頁先保持開啟。</li>
      <li>按「啟動連續掃碼器」。</li>
      <li>切到新開的 GitHub 掃碼分頁。</li>
      <li>在掃碼頁按「啟用相機＋語音」。</li>
      <li>之後同一頁連續掃所有來賓 QR Code。</li>
    </ol>
  </div>
</div>
 
<script>
  const EVENT_NAME = ${eventNameJson};
  const SCANNER_BASE_URL = ${scannerUrlJson};
  let scannerWindow = null;
  let scannerOrigin = '';
  let sessionToken = '';
 
  function randomToken(){
    if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
    return 'S-' + Date.now() + '-' + Math.random().toString(36).slice(2);
  }
 
  function setStatus(text, type){
    const el = document.getElementById('status');
    el.textContent = text;
    el.className = 'status ' + (type || '');
  }
 
  function validateScannerUrl(){
    if (!SCANNER_BASE_URL || SCANNER_BASE_URL.indexOf('請改成') !== -1) {
      throw new Error('請先在 Code.gs 的 CONFIG.SCANNER_URL 貼上 GitHub Pages HTTPS 網址。');
    }
    const url = new URL(SCANNER_BASE_URL);
    if (url.protocol !== 'https:') {
      throw new Error('SCANNER_URL 必須是 https:// 網址。');
    }
    return url;
  }
 
  function launchScanner(){
    try{
      const url = validateScannerUrl();
      sessionToken = randomToken();
      scannerOrigin = url.origin;
      url.searchParams.set('controlOrigin', window.location.origin);
      url.searchParams.set('sessionToken', sessionToken);
      url.searchParams.set('v', String(Date.now()));
 
      scannerWindow = window.open(url.toString(), '_blank');
      if (!scannerWindow) {
        throw new Error('瀏覽器阻擋新視窗，請允許此網站開啟彈出式視窗。');
      }
 
      setStatus('掃碼器已開啟。請切到新分頁；本控制頁請勿關閉。', 'ok');
    }catch(err){
      setStatus(err.message || String(err), 'err');
    }
  }
 
  function sendToScanner(message){
    if (!scannerWindow || scannerWindow.closed) {
      setStatus('掃碼器分頁已關閉，請重新按「啟動連續掃碼器」。', 'err');
      return;
    }
    scannerWindow.postMessage(message, scannerOrigin);
  }
 
  window.addEventListener('message', function(event){
    if (!scannerWindow || event.source !== scannerWindow) return;
    if (!scannerOrigin || event.origin !== scannerOrigin) return;
 
    const data = event.data || {};
    if (data.source !== 'CEO_SCANNER') return;
    if (!sessionToken || data.sessionToken !== sessionToken) return;
 
    if (data.type === 'HELLO') {
      sendToScanner({
        source:'CEO_CONTROL',
        type:'HELLO_ACK',
        sessionToken:sessionToken,
        eventName:EVENT_NAME
      });
      setStatus('掃碼器與控制頁已連線。', 'ok');
      return;
    }
 
    if (data.type === 'PING') {
      sendToScanner({source:'CEO_CONTROL',type:'PONG',sessionToken:sessionToken,ts:Date.now()});
      return;
    }
 
    if (data.type === 'DASHBOARD_REQUEST') {
      google.script.run
        .withSuccessHandler(function(result){
          sendToScanner({
            source:'CEO_CONTROL',
            type:'DASHBOARD_RESULT',
            sessionToken:sessionToken,
            requestId:data.requestId || '',
            payload:result
          });
        })
        .withFailureHandler(function(error){
          sendToScanner({
            source:'CEO_CONTROL',
            type:'DASHBOARD_RESULT',
            sessionToken:sessionToken,
            requestId:data.requestId || '',
            payload:{error:true,message:error && error.message ? error.message : String(error)}
          });
        })
        .getCheckinDashboard();
      return;
    }
 
    if (data.type === 'CHECKIN_REQUEST') {
      setStatus('正在處理報到碼：' + String(data.code || ''), 'warn');
      google.script.run
        .withSuccessHandler(function(result){
          sendToScanner({
            source:'CEO_CONTROL',
            type:'CHECKIN_RESULT',
            sessionToken:sessionToken,
            requestId:data.requestId,
            payload:result
          });
          setStatus('最近一次報到處理完成。', 'ok');
        })
        .withFailureHandler(function(error){
          sendToScanner({
            source:'CEO_CONTROL',
            type:'CHECKIN_RESULT',
            sessionToken:sessionToken,
            requestId:data.requestId,
            payload:{type:'error',message:'後端錯誤：' + (error && error.message ? error.message : String(error))}
          });
          setStatus('報到後端發生錯誤。', 'err');
        })
        .checkInByCode(data.code);
      return;
    }
 
    if (data.type === 'ROW_CHECKIN_REQUEST') {
      google.script.run
        .withSuccessHandler(function(result){
          sendToScanner({
            source:'CEO_CONTROL',
            type:'CHECKIN_RESULT',
            sessionToken:sessionToken,
            requestId:data.requestId,
            payload:result
          });
          setStatus('人工搜尋報到處理完成。', 'ok');
        })
        .withFailureHandler(function(error){
          sendToScanner({
            source:'CEO_CONTROL',
            type:'CHECKIN_RESULT',
            sessionToken:sessionToken,
            requestId:data.requestId,
            payload:{type:'error',message:'後端錯誤：' + (error && error.message ? error.message : String(error))}
          });
          setStatus('人工搜尋報到後端發生錯誤。', 'err');
        })
        .checkInByRow(Number(data.row));
      return;
    }
 
    if (data.type === 'SEARCH_REQUEST') {
      google.script.run
        .withSuccessHandler(function(result){
          sendToScanner({
            source:'CEO_CONTROL',
            type:'SEARCH_RESULT',
            sessionToken:sessionToken,
            requestId:data.requestId,
            payload:result
          });
        })
        .withFailureHandler(function(error){
          sendToScanner({
            source:'CEO_CONTROL',
            type:'SEARCH_RESULT',
            sessionToken:sessionToken,
            requestId:data.requestId,
            payload:[],
            errorMessage:error && error.message ? error.message : String(error)
          });
        })
        .searchParticipants(data.keyword);
    }
  });
 
  function testBackend(){
    setStatus('正在測試 Apps Script 後端…', 'warn');
    google.script.run
      .withSuccessHandler(function(data){
        setStatus('後端正常。正式報名 ' + data.registered + ' 人，已報到 ' + data.checkedIn + ' 人。', 'ok');
      })
      .withFailureHandler(function(error){
        setStatus('後端測試失敗：' + (error && error.message ? error.message : String(error)), 'err');
      })
      .getCheckinDashboard();
  }
</script>
</body>
</html>`;
}
 
 
// ============================================================
// 掃碼／手動輸入報到碼：公開函式，供 google.script.run 呼叫
// ============================================================
function checkInByCode(rawCode) {
  const code = normalizeCheckinCode_(rawCode);
 
  if (!code) {
    return {
      ok: false,
      type: 'error',
      message: '報到碼為空白。'
    };
  }
 
  const lock = LockService.getScriptLock();
 
  try {
    lock.waitLock(30000);
 
    const sheet = getRegistrationSheet_();
    ensureCustomHeaders_(sheet);
    const headers = getHeaderMap_(sheet);
    const lastRow = sheet.getLastRow();
    const codeCol = getColumn_(headers, CONFIG.HEADER_CHECKIN_CODE);
 
    if (lastRow < 2) {
      return {ok:false,type:'error',message:'目前沒有報名資料。'};
    }
 
    const codes = sheet.getRange(2, codeCol, lastRow - 1, 1).getDisplayValues();
    let targetRow = -1;
 
    for (let i = 0; i < codes.length; i++) {
      if (normalizeCheckinCode_(codes[i][0]) === code) {
        targetRow = i + 2;
        break;
      }
    }
 
    if (targetRow === -1) {
      return {
        ok: false,
        type: 'error',
        message: '查無此報到碼：' + code + '。請確認是否為最新 QR Code。'
      };
    }
 
    return processCheckIn_(sheet, targetRow, headers);
 
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}
 
 
// ============================================================
// 姓名／公司搜尋：公開函式，供 google.script.run 呼叫
// ============================================================
function searchParticipants(keyword) {
  keyword = normalizeSearchText_(keyword);
  if (!keyword) return [];
 
  const sheet = getRegistrationSheet_();
  ensureCustomHeaders_(sheet);
  const headers = getHeaderMap_(sheet);
  const lastRow = sheet.getLastRow();
 
  if (lastRow < 2) return [];
 
  const companyCol = getColumn_(headers, CONFIG.FIELD_COMPANY);
  const nameCol = getColumn_(headers, CONFIG.FIELD_NAME);
  const titleCol = getColumn_(headers, CONFIG.FIELD_TITLE);
  const mealCol = getColumn_(headers, CONFIG.FIELD_MEAL);
  const regCol = getColumn_(headers, CONFIG.HEADER_REG_STATUS);
  const checkinStatusCol = getColumn_(headers, CONFIG.HEADER_CHECKIN_STATUS);
  const checkinTimeCol = getColumn_(headers, CONFIG.HEADER_CHECKIN_TIME);
  const badgeCol = getColumn_(headers, CONFIG.HEADER_BADGE_NUMBER);
 
  const results = [];
 
  for (let row = 2; row <= lastRow; row++) {
    const company = String(sheet.getRange(row, companyCol).getDisplayValue() || '').trim();
    const name = String(sheet.getRange(row, nameCol).getDisplayValue() || '').trim();
    const searchable = normalizeSearchText_(company + ' ' + name);
 
    if (searchable.indexOf(keyword) === -1) continue;
 
    results.push({
      row: row,
      company: company,
      name: name,
      title: String(sheet.getRange(row, titleCol).getDisplayValue() || '').trim(),
      meal: String(sheet.getRange(row, mealCol).getDisplayValue() || '').trim(),
      badgeNumber: String(sheet.getRange(row, badgeCol).getDisplayValue() || '').trim(),
      registrationStatus: String(sheet.getRange(row, regCol).getDisplayValue() || '').trim(),
      checkinStatus: String(sheet.getRange(row, checkinStatusCol).getDisplayValue() || '').trim() || '未報到',
      checkinTime: formatCheckinTime_(sheet.getRange(row, checkinTimeCol).getValue())
    });
 
    if (results.length >= 20) break;
  }
 
  return results;
}
 
 
// ============================================================
// 人工搜尋結果一鍵報到：公開函式
// ============================================================
function checkInByRow(row) {
  row = Number(row);
 
  if (!Number.isInteger(row) || row < 2) {
    return {ok:false,type:'error',message:'資料列無效。'};
  }
 
  const lock = LockService.getScriptLock();
 
  try {
    lock.waitLock(30000);
    const sheet = getRegistrationSheet_();
    ensureCustomHeaders_(sheet);
    const headers = getHeaderMap_(sheet);
 
    if (row > sheet.getLastRow()) {
      return {ok:false,type:'error',message:'查無此報名資料。'};
    }
 
    return processCheckIn_(sheet, row, headers);
 
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}
 
 
// ============================================================
// 報到核心
// ============================================================
function processCheckIn_(sheet, row, headers) {
  const companyCol = getColumn_(headers, CONFIG.FIELD_COMPANY);
  const nameCol = getColumn_(headers, CONFIG.FIELD_NAME);
  const titleCol = getColumn_(headers, CONFIG.FIELD_TITLE);
  const mealCol = getColumn_(headers, CONFIG.FIELD_MEAL);
  const regCol = getColumn_(headers, CONFIG.HEADER_REG_STATUS);
  const checkinStatusCol = getColumn_(headers, CONFIG.HEADER_CHECKIN_STATUS);
  const checkinTimeCol = getColumn_(headers, CONFIG.HEADER_CHECKIN_TIME);
  const badgeCol = getColumn_(headers, CONFIG.HEADER_BADGE_NUMBER);
 
  // 一次讀取整列，降低活動現場每筆報到的 Sheets API 往返次數。
  const displayRow = sheet
    .getRange(row, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];
 
  const company = String(displayRow[companyCol - 1] || '').trim();
  const name = String(displayRow[nameCol - 1] || '').trim();
  const title = String(displayRow[titleCol - 1] || '').trim();
  const meal = String(displayRow[mealCol - 1] || '').trim();
  const badgeNumber = String(displayRow[badgeCol - 1] || '').trim();
  const registrationStatus = String(displayRow[regCol - 1] || '').trim();
  const checkinStatus = String(displayRow[checkinStatusCol - 1] || '').trim();
  const oldTime = sheet.getRange(row, checkinTimeCol).getValue();
 
  if (registrationStatus !== '報名成功') {
    return {
      ok: false,
      type: 'error',
      message: '此筆資料狀態為「' + registrationStatus + '」，不可報到。',
      company: company,
      name: name,
      title: title,
      meal: meal,
      badgeNumber: badgeNumber
    };
  }
 
  if (checkinStatus === '已報到') {
    return {
      ok: true,
      type: 'warning',
      message: '此參加者已完成報到。',
      company: company,
      name: name,
      title: title,
      meal: meal,
      badgeNumber: badgeNumber,
      checkinTime: formatCheckinTime_(oldTime)
    };
  }
 
  const now = new Date();
  sheet.getRange(row, checkinStatusCol).setValue('已報到');
  sheet.getRange(row, checkinTimeCol).setValue(now);
  SpreadsheetApp.flush();
 
  return {
    ok: true,
    type: 'success',
    message: '報到成功',
    company: company,
    name: name,
    title: title,
    meal: meal,
    badgeNumber: badgeNumber,
    checkinTime: formatCheckinTime_(now)
  };
}
 
 
// ============================================================
// 報到儀表板：公開函式
// ============================================================
function getCheckinDashboard() {
  const sheet = getRegistrationSheet_();
  ensureCustomHeaders_(sheet);
  const headers = getHeaderMap_(sheet);
  const lastRow = sheet.getLastRow();
 
  if (lastRow < 2) {
    return {registered:0,checkedIn:0,notCheckedIn:0,rate:0};
  }
 
  const regCol = getColumn_(headers, CONFIG.HEADER_REG_STATUS);
  const checkinStatusCol = getColumn_(headers, CONFIG.HEADER_CHECKIN_STATUS);
 
  let registered = 0;
  let checkedIn = 0;
 
  for (let row = 2; row <= lastRow; row++) {
    const regStatus = String(sheet.getRange(row, regCol).getDisplayValue() || '').trim();
    if (regStatus !== '報名成功') continue;
 
    registered++;
    const checkinStatus = String(sheet.getRange(row, checkinStatusCol).getDisplayValue() || '').trim();
    if (checkinStatus === '已報到') checkedIn++;
  }
 
  return {
    registered: registered,
    checkedIn: checkedIn,
    notCheckedIn: Math.max(registered - checkedIn, 0),
    rate: registered > 0 ? Math.round(checkedIn / registered * 100) : 0
  };
}
 
 
// ============================================================
// 重新依正式報名順序建立名牌編號（截止後、印名牌前才使用）
// ============================================================
function rebuildBadgeNumbers() {
  const sheet = getRegistrationSheet_();
  ensureCustomHeaders_(sheet);
  const headers = getHeaderMap_(sheet);
  const regCol = getColumn_(headers, CONFIG.HEADER_REG_STATUS);
  const badgeCol = getColumn_(headers, CONFIG.HEADER_BADGE_NUMBER);
  const lastRow = sheet.getLastRow();
 
  let counter = 0;
 
  for (let row = 2; row <= lastRow; row++) {
    const status = String(sheet.getRange(row, regCol).getDisplayValue() || '').trim();
    if (status === '報名成功') {
      counter++;
      sheet.getRange(row, badgeCol).setNumberFormat('@').setValue(String(counter).padStart(3, '0'));
    } else {
      sheet.getRange(row, badgeCol).clearContent();
    }
  }
 
  SpreadsheetApp.flush();
  Logger.log('名牌編號重建完成，共 ' + counter + ' 人。');
}
 
 
// ============================================================
// 重建所有正式報名者試算表 QR Code
// QR Code 僅含報到碼，與 Web App 部署網址完全解耦。
// ============================================================
function rebuildAllQRCodes() {
  const sheet = getRegistrationSheet_();
  ensureCustomHeaders_(sheet);
  const headers = getHeaderMap_(sheet);
  const regCol = getColumn_(headers, CONFIG.HEADER_REG_STATUS);
  const codeCol = getColumn_(headers, CONFIG.HEADER_CHECKIN_CODE);
  const lastRow = sheet.getLastRow();
 
  for (let row = 2; row <= lastRow; row++) {
    const status = String(sheet.getRange(row, regCol).getDisplayValue() || '').trim();
    if (status !== '報名成功') continue;
 
    let code = String(sheet.getRange(row, codeCol).getDisplayValue() || '').trim();
    if (!code) {
      code = generateCheckinCode_(row);
      sheet.getRange(row, codeCol).setNumberFormat('@').setValue(code);
    }
 
    createQRCodeInSheet_(sheet, row, code, headers);
  }
 
  SpreadsheetApp.flush();
  Logger.log('所有正式報名者 QR Code 已重建完成。');
}
 
 
// ============================================================
// QR Email 單筆測試：寄到管理員信箱，不修改正式資料。
// ============================================================
function testQRCodeEmail() {
  sendQRCodeEmail_(
    CONFIG.ADMIN_EMAIL,
    '測試科技股份有限公司',
    '王小明',
    '總經理',
    '0912345678',
    '葷食',
    '025',
    'CEO2026-EMAILTEST-' + Date.now()
  );
  Logger.log('QR Code 測試信已寄到：' + CONFIG.ADMIN_EMAIL);
}
 
 
// ============================================================
// 每日報名摘要
// ============================================================
function sendDailyRegistrationSummary() {
  const sheet = getRegistrationSheet_();
  ensureCustomHeaders_(sheet);
  const headers = getHeaderMap_(sheet);
  const lastRow = sheet.getLastRow();
 
  const regCol = getColumn_(headers, CONFIG.HEADER_REG_STATUS);
  const checkinStatusCol = getColumn_(headers, CONFIG.HEADER_CHECKIN_STATUS);
  const mealCol = getColumn_(headers, CONFIG.FIELD_MEAL);
 
  let success = 0;
  let meat = 0;
  let vegetarian = 0;
  let duplicate = 0;
  let waiting = 0;
  let checkedIn = 0;
 
  for (let row = 2; row <= lastRow; row++) {
    const status = String(sheet.getRange(row, regCol).getDisplayValue() || '').trim();
    const meal = String(sheet.getRange(row, mealCol).getDisplayValue() || '').trim();
    const checkinStatus = String(sheet.getRange(row, checkinStatusCol).getDisplayValue() || '').trim();
 
    if (status === '報名成功') {
      success++;
      if (meal === '葷食') meat++;
      if (meal === '素食') vegetarian++;
      if (checkinStatus === '已報到') checkedIn++;
    } else if (status === '重複報名') {
      duplicate++;
    } else if (status === '候補／額滿') {
      waiting++;
    }
  }
 
  const today = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy/MM/dd');
  const remaining = Math.max(CONFIG.MAX_PEOPLE - success, 0);
  const rate = Math.round(success / CONFIG.MAX_PEOPLE * 100);
 
  safeSendEmail_({
    to: getManagementNotificationRecipients_(),
    subject: '【每日報名摘要】' + CONFIG.EVENT_NAME + '－' + today,
    body:
      '承辦人您好：\n\n' +
      '正式報名：' + success + ' / ' + CONFIG.MAX_PEOPLE + ' 人\n' +
      '剩餘名額：' + remaining + ' 人\n' +
      '報名率：' + rate + '%\n\n' +
      '葷食：' + meat + ' 人\n' +
      '素食：' + vegetarian + ' 人\n\n' +
      '重複報名：' + duplicate + ' 筆\n' +
      '候補／額滿：' + waiting + ' 筆\n\n' +
      '已報到：' + checkedIn + ' 人\n\n' +
      '試算表：\n' + SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getUrl()
  });
}
 
 
// ============================================================
// 管理員階段通知
// ============================================================
function sendAdminNotification_(successCount, company, name, title, meal) {
  const points = new Set([10,20,30,40,50,60,70,75,80]);
  if (!points.has(successCount)) return;
 
  safeSendEmail_({
    to: getManagementNotificationRecipients_(),
    subject: '【報名進度】' + CONFIG.EVENT_NAME + '－' + successCount + '/' + CONFIG.MAX_PEOPLE,
    body:
      '目前正式報名：' + successCount + ' 人\n' +
      '剩餘名額：' + Math.max(CONFIG.MAX_PEOPLE - successCount, 0) + ' 人\n\n' +
      '最新正式報名：\n' +
      company + '\n' + name + ' ' + title + '\n餐食：' + meal
  });
}
 
 
// ============================================================
// 共用：公司重複判斷
// ============================================================
function hasPriorSuccessfulCompany_(sheet, currentRow, company, headers) {
  if (currentRow <= 2) return false;
 
  const companyCol = getColumn_(headers, CONFIG.FIELD_COMPANY);
  const regCol = getColumn_(headers, CONFIG.HEADER_REG_STATUS);
  const target = normalizeCompanyName_(company);
 
  for (let row = 2; row < currentRow; row++) {
    const previousCompany = normalizeCompanyName_(sheet.getRange(row, companyCol).getDisplayValue());
    const previousStatus = String(sheet.getRange(row, regCol).getDisplayValue() || '').trim();
    if (previousCompany && previousCompany === target && previousStatus === '報名成功') {
      return true;
    }
  }
 
  return false;
}
 
 
function normalizeCompanyName_(value) {
  return String(value == null ? '' : value)
    .normalize('NFKC')
    .replace(/\s+/g, '')
    .toLowerCase()
    .trim();
}
 
 
// ============================================================
// 共用：報到碼標準化
// 可接受：
// 1. 純 CEO2026-... 報到碼
// 2. 舊版含 ?code=CEO2026-... 的完整網址
// ============================================================
function normalizeCheckinCode_(value) {
  let text = String(value == null ? '' : value).trim();
 
  const queryMatch = text.match(/[?&]code=([^&#]+)/i);
  if (queryMatch && queryMatch[1]) {
    try {
      text = decodeURIComponent(queryMatch[1]);
    } catch (e) {
      text = queryMatch[1];
    }
  }
 
  const codeMatch = text.match(/CEO2026-[A-Z0-9-]+/i);
  if (codeMatch) {
    text = codeMatch[0];
  }
 
  return text
    .normalize('NFKC')
    .replace(/\s+/g, '')
    .trim()
    .toUpperCase();
}
 
 
function normalizeSearchText_(value) {
  return String(value == null ? '' : value)
    .normalize('NFKC')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}
 
 
function generateCheckinCode_(row) {
  const timestamp = Utilities.formatDate(
    new Date(),
    CONFIG.TIMEZONE,
    'yyyyMMddHHmmss'
  );
 
  const random = Utilities
    .getUuid()
    .replace(/-/g, '')
    .substring(0, 10)
    .toUpperCase();
 
  return 'CEO2026-' + String(row).padStart(4, '0') + '-' + random + '-' + timestamp;
}
 
 
// ============================================================
// QR Code
// ============================================================
function buildQrImageUrl_(content, size) {
  return (
    'https://quickchart.io/qr' +
    '?format=png' +
    '&size=' + encodeURIComponent(String(size || 400)) +
    '&margin=4' +
    '&ecLevel=M' +
    '&text=' + encodeURIComponent(content)
  );
}
 
 
function createQRCodeInSheet_(sheet, row, checkinCode, headers) {
  const qrCol = getColumn_(headers, CONFIG.HEADER_QR_CODE);
  const qrUrl = buildQrImageUrl_(normalizeCheckinCode_(checkinCode), 260);
  const formula = '=IMAGE("' + qrUrl.replace(/"/g, '""') + '",4,180,180)';
  sheet.getRange(row, qrCol).setFormula(formula);
  sheet.setRowHeight(row, 190);
}
 
 
function fetchQrBlob_(checkinCode) {
  const qrUrl = buildQrImageUrl_(normalizeCheckinCode_(checkinCode), 420);
  let lastError = null;
 
  for (let attempt = 1; attempt <= 2; attempt++) {
    try {
      const response = UrlFetchApp.fetch(qrUrl, {
        method: 'get',
        followRedirects: true,
        muteHttpExceptions: true
      });
 
      const status = response.getResponseCode();
      if (status >= 200 && status < 300) {
        return response
          .getBlob()
          .setName('115年度高雄軟體園區CEO聯誼會_報到QRCode.png');
      }
 
      lastError = new Error('QR Code 圖片服務回傳 HTTP ' + status);
    } catch (error) {
      lastError = error;
    }
 
    Utilities.sleep(500);
  }
 
  throw lastError || new Error('QR Code 圖片產生失敗。');
}
 
 
// ============================================================
// 工作表欄位：依第1列欄位名稱自動定位，不依賴固定欄號。
// ============================================================
function ensureCustomHeaders_(sheet) {
  const required = [
    CONFIG.HEADER_REG_STATUS,
    CONFIG.HEADER_CHECKIN_CODE,
    CONFIG.HEADER_QR_CODE,
    CONFIG.HEADER_CHECKIN_STATUS,
    CONFIG.HEADER_CHECKIN_TIME,
    CONFIG.HEADER_QR_SEND_STATUS,
    CONFIG.HEADER_QR_SEND_TIME,
    CONFIG.HEADER_BADGE_NUMBER
  ];
 
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  let headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
 
  required.forEach(function (header) {
    if (headers.indexOf(header) === -1) {
      const newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(header);
      headers.push(header);
    }
  });
}
 
 
function getHeaderMap_(sheet) {
  const lastColumn = sheet.getLastColumn();
  const row = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const map = {};
 
  row.forEach(function (header, index) {
    const key = String(header || '').trim();
    if (key) map[key] = index + 1;
  });
 
  return map;
}
 
 
function getColumn_(headers, headerName) {
  const col = headers[headerName];
  if (!col) {
    throw new Error('找不到欄位：「' + headerName + '」。請確認 Google 表單／試算表第1列欄位名稱。');
  }
  return col;
}
 
 
function getNamedValue_(namedValues, fieldName) {
  if (!namedValues[fieldName] || namedValues[fieldName].length === 0) {
    throw new Error('找不到表單欄位：「' + fieldName + '」。請確認 Google 表單題目名稱完全一致。');
  }
  return String(namedValues[fieldName][0]);
}
 
 
function countStatus_(sheet, statusHeader, targetStatus) {
  ensureCustomHeaders_(sheet);
  const headers = getHeaderMap_(sheet);
  const col = getColumn_(headers, statusHeader);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
 
  const values = sheet.getRange(2, col, lastRow - 1, 1).getDisplayValues();
  return values.reduce(function (count, row) {
    return count + (String(row[0] || '').trim() === targetStatus ? 1 : 0);
  }, 0);
}
 
 
function getRegistrationSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    throw new Error('找不到工作表：「' + CONFIG.SHEET_NAME + '」。請確認左下角分頁名稱。');
  }
  return sheet;
}
 
 
function formatCheckinTime_(value) {
  if (!value) return '';
  return Utilities.formatDate(new Date(value), CONFIG.TIMEZONE, 'yyyy/MM/dd HH:mm:ss');
}
 
 
function closeForm_() {
  const form = FormApp.openByUrl(CONFIG.FORM_URL);
  form.setCustomClosedFormMessage(
    '感謝您的關注，' + CONFIG.EVENT_NAME + '報名名額已額滿，目前已停止接受報名。'
  );
  form.setAcceptingResponses(false);
}
 
 
function safeCloseForm_() {
  try {
    closeForm_();
  } catch (error) {
    notifyAdminError_('關閉 Google 表單失敗', error, '');
  }
}
 
 
// ============================================================
// 共用：每日摘要／階段通知的管理成員收件人
// 主要管理員＋其他管理成員會合併、去空白、去重複。
// ============================================================
function getManagementNotificationRecipients_() {
  const emails = [CONFIG.ADMIN_EMAIL].concat(
    Array.isArray(CONFIG.ADDITIONAL_ADMIN_EMAILS)
      ? CONFIG.ADDITIONAL_ADMIN_EMAILS
      : []
  );
 
  const unique = [];
  const seen = new Set();
 
  emails.forEach(function (email) {
    const value = String(email || '').trim();
    if (!value) return;
 
    const key = value.toLowerCase();
    if (seen.has(key)) return;
 
    seen.add(key);
    unique.push(value);
  });
 
  if (unique.length === 0) {
    throw new Error('尚未設定管理員通知 Email。');
  }
 
  return unique.join(',');
}
 
 
function safeSendEmail_(message) {
  try {
    const payload = {
      to: message.to,
      subject: message.subject,
      body: message.body,
      name: CONFIG.SENDER_NAME
    };
 
    if (message.htmlBody) {
      payload.htmlBody = message.htmlBody;
    }
 
    MailApp.sendEmail(payload);
  } catch (error) {
    console.error('寄信失敗：', message.to, error);
  }
}
 
 
function notifyAdminError_(title, error, row) {
  const details = error && error.stack ? error.stack : String(error);
  safeSendEmail_({
    to: CONFIG.ADMIN_EMAIL,
    subject: '【系統錯誤】' + CONFIG.EVENT_NAME + '－' + title,
    body:
      '報名／報到系統發生錯誤。\n\n' +
      (row ? '資料列：' + row + '\n' : '') +
      '項目：' + title + '\n\n' + details
  });
}
 
 
function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
 
 
function safeJsonForHtml_(value) {
  return JSON.stringify(value)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026');
}
 
 
// ============================================================
// 系統檢查
// ============================================================
function checkSystem() {
  const sheet = getRegistrationSheet_();
  ensureCustomHeaders_(sheet);
 
  Logger.log('工作表：' + sheet.getName());
  Logger.log('Apps Script 控制頁：' + (ScriptApp.getService().getUrl() || '尚未部署')); 
  Logger.log('GitHub Scanner：' + CONFIG.SCANNER_URL);
  Logger.log('管理員：' + CONFIG.ADMIN_EMAIL);
  Logger.log('今日剩餘 MailApp 收件者額度：' + MailApp.getRemainingDailyQuota());
}
