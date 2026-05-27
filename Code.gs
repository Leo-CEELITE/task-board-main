/**
 * CE Task API - v1
 *
 * 基於 Xcity Task API v14 遷移至 CE 版本：
 * - 全面品牌重塑：XCITY → CE
 * - 推播頻率調整：由每日兩次（早晨 + 晚間）改為每日一次
 * - 推播時間固定為台灣時間（Asia/Taipei）每天 10:00 AM
 *
 * v1 核心設計（繼承自 Xcity v14）：
 * - doPost 路徑「絕對不會發 push」——
 *     ‣ 所有 webhook 觸發的訊息都只走 Reply API（免費）
 *     ‣ reply 失敗就記 log 然後丟掉，使用者那邊就是收不到回覆
 * - push 函式只有 sendReminder（排程入口）會呼叫
 * - doPost 只能呼叫 replyTextOnly_ / replyFlexOnly_ 兩個函式
 *
 * 部署提醒：
 * 1. 編輯器右上「部署」→「管理部署作業」→ 編輯（鉛筆）→ 版本選「新版本」→ 部署
 *    （只按儲存不會生效）
 * 2. LINE Developers Console → Messaging API → Webhook URL 維持原樣，
 *    Use webhook 開關打開
 * 3. 貼上 v1 之後，先在群組打 /今日任務 測試一次 → 應該會收到 Flex 訊息
 *    然後去 spreadsheet 看 debug_log 分頁，應該有 reply-ok 紀錄
 * 4. 執行 setupDailyTrigger() 建立每天 10:00 AM (Asia/Taipei) 的推播排程
 *
 * 預期 push 額度：每天 1 次（10:00 AM 排程），其餘指令都走 reply 不算額度
 *
 * ── LINE Token 安全儲存（Script Properties）────────────────────────────────
 * LINE_CHANNEL_ACCESS_TOKEN 改存在 GAS 指令碼屬性，不放在 Google Sheet config 分頁。
 *
 * 首次設定步驟：
 *   1. 在 Apps Script 編輯器左側點齒輪圖示 ⚙️「專案設定」
 *      (Project Settings)
 *   2. 捲動到「指令碼屬性」(Script Properties) 區塊
 *   3. 點「新增指令碼屬性」
 *      屬性 (Property)：LINE_CHANNEL_ACCESS_TOKEN
 *      值    (Value)  ：貼上你的 LINE Channel Access Token（Long-lived）
 *   4. 點「儲存指令碼屬性」
 *
 * 驗證：執行 testGetConfig() → Logger 應顯示「CHANNEL_ACCESS_TOKEN: 已設定」
 * ────────────────────────────────────────────────────────────────────────────
 */


const SHEET_NAME_TASKS = 'Master Task Board';
const SHEET_NAME_CONFIG = 'config';
const SHEET_NAME_LOG = 'debug_log';
const TIMEZONE = 'Asia/Taipei';
const DASHBOARD_URL = 'https://YOUR_GITHUB_USERNAME.github.io/ce-task-board/';  // ⚠️ 請更新為 CE 的 dashboard 網址
const TASK_URL = 'https://YOUR_CE_TASK_SHORTLINK';                              // ⚠️ 請更新為 CE 的任務更新表短連結

const MAX_MESSAGE_LEN = 4800;
const MY_USER_ID = 'YOUR_LINE_USER_ID';  // ⚠️ 請更新為 CE 負責人的 LINE User ID
const LOG_MAX_ROWS = 2001;

const OWNER_COLORS = {
  'Ryota':   '#4A90D9',
  'Charlie': '#27AE60',
  'Leo':     '#9B59B6',
  '培培':    '#E67E22',
  '溫':      '#E74C3C',
  // ⚠️ 請依 CE 團隊成員更新此對應表
};


function getOwnerColor_(owner) {
  return OWNER_COLORS[owner] || '#888888';
}


// ════════════════════════════════════════════════════════════════════════════
//  LINE Token（從 Script Properties 安全讀取，不經過 Google Sheet）
// ════════════════════════════════════════════════════════════════════════════

/**
 * 從 GAS 指令碼屬性讀取 LINE Channel Access Token。
 * 所有需要 token 的函式都呼叫這裡——往後只需改這一處。
 *
 * 尚未設定時會拋出明確錯誤，避免靜默失敗。
 */
function getLineToken_() {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) throw new Error(
    'LINE_CHANNEL_ACCESS_TOKEN 未設定。' +
    '請至 Apps Script → ⚙️ 專案設定 → 指令碼屬性 新增該 key。'
  );
  return token;
}


// ════════════════════════════════════════════════════════════════════════════
//  Sheet log
// ════════════════════════════════════════════════════════════════════════════

function logToSheet_(category, message) {
  try {
    const ss = SpreadsheetApp.getActive();
    let sheet = ss.getSheetByName(SHEET_NAME_LOG);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME_LOG);
      sheet.getRange(1, 1, 1, 3).setValues([['timestamp', 'category', 'message']]);
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 160);
      sheet.setColumnWidth(2, 140);
      sheet.setColumnWidth(3, 800);
    }
    const ts = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    const msgStr = (typeof message === 'string') ? message : JSON.stringify(message);
    sheet.appendRow([ts, String(category), msgStr]);

    const lastRow = sheet.getLastRow();
    if (lastRow > LOG_MAX_ROWS) {
      sheet.deleteRows(2, lastRow - LOG_MAX_ROWS);
    }
  } catch (err) {
    Logger.log('logToSheet_ 失敗: ' + err);
  }
}


// ════════════════════════════════════════════════════════════════════════════
//  Config
// ════════════════════════════════════════════════════════════════════════════

function getConfig_(key) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_CONFIG);
  if (!sheet) throw new Error('找不到 config 分頁');
  const values = sheet.getDataRange().getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) return String(values[i][1]).trim();
  }
  return '';
}


function setConfig_(key, value) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_CONFIG);
  if (!sheet) throw new Error('找不到 config 分頁');
  const values = sheet.getDataRange().getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}


function isDate_(value) {
  return value !== null
    && value !== undefined
    && typeof value === 'object'
    && typeof value.getTime === 'function'
    && !isNaN(value.getTime());
}


function formatDateString_(value) {
  if (value === null || value === undefined || value === '') return '';
  if (isDate_(value)) return Utilities.formatDate(value, TIMEZONE, 'yyyy-MM-dd');
  const str = String(value).trim();
  if (!str) return '';
  const parsed = new Date(str);
  if (isDate_(parsed)) return Utilities.formatDate(parsed, TIMEZONE, 'yyyy-MM-dd');
  return str;
}


function getTodayString_() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
}


function getTomorrowString_() {
  const tomorrow = new Date(Date.now() + 24 * 60 * 60 * 1000);
  return Utilities.formatDate(tomorrow, TIMEZONE, 'yyyy-MM-dd');
}


function getAllTasks_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME_TASKS);
  if (!sheet) throw new Error('找不到 Master Task Board 分頁');

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => String(h).trim());

  const headerMap = {
    '週次': 'week',
    '任務性質 (Type)': 'type',
    '負責人': 'owner',
    '優先級': 'priority',
    '執行狀態': 'status',
    '任務名稱': 'name',
    'Due Day': 'dueDay',
    '備註': 'note',
    'Date': 'date'
  };

  const dateFields = ['dueDay', 'date'];

  const tasks = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const hasContent = row.some(cell => String(cell).trim() !== '');
    if (!hasContent) continue;

    const task = { _rowIndex: i + 1 };
    headers.forEach((header, idx) => {
      const field = headerMap[header];
      if (!field) return;
      const raw = row[idx];
      if (dateFields.indexOf(field) !== -1) {
        task[field] = formatDateString_(raw);
      } else if (isDate_(raw)) {
        task[field] = Utilities.formatDate(raw, TIMEZONE, 'yyyy-MM-dd');
      } else {
        task[field] = String(raw).trim();
      }
    });

    tasks.push(task);
  }

  return tasks;
}


function getSummary_(tasks) {
  const today = getTodayString_();

  const isDone      = t => t.status && t.status.indexOf('Done')     !== -1;
  const isCancelled = t => t.status && t.status.indexOf('Canceled') !== -1;
  const isActive    = t => !isDone(t) && !isCancelled(t);

  const doneCount = tasks.filter(isDone).length;
  const total     = tasks.length;

  const overdue = tasks.filter(t => t.dueDay && isActive(t) && t.dueDay < today);
  const dueToday = tasks.filter(t => t.dueDay && isActive(t) && t.dueDay === today);

  const byStatus = {};
  tasks.forEach(t => {
    const s = t.status || '未定義';
    byStatus[s] = (byStatus[s] || 0) + 1;
  });

  const byOwner = {};
  tasks.forEach(t => {
    const o = t.owner || '未指派';
    if (!byOwner[o]) byOwner[o] = { total: 0, done: 0 };
    byOwner[o].total++;
    if (isDone(t)) byOwner[o].done++;
  });

  return {
    total,
    doneCount,
    completionRate: total > 0 ? Math.round((doneCount / total) * 10000) / 100 : 0,
    overdueCount:   overdue.length,
    dueTodayCount:  dueToday.length,
    byStatus,
    byOwner,
    today
  };
}


// ════════════════════════════════════════════════════════════════════════════
//  doGet
// ════════════════════════════════════════════════════════════════════════════

function doGet(e) {
  try {
    const expectedToken = getConfig_('API_TOKEN');
    const providedToken = (e && e.parameter && e.parameter.token) || '';

    if (!expectedToken) return jsonResponse_({ error: 'API_TOKEN 未設定在 config 分頁' });
    if (providedToken !== expectedToken) return jsonResponse_({ error: 'Invalid token' });

    const action = (e && e.parameter && e.parameter.action) || 'all';
    const tasks  = getAllTasks_();

    let payload;
    if (action === 'tasks') {
      payload = { tasks };
    } else if (action === 'summary') {
      payload = { summary: getSummary_(tasks) };
    } else {
      payload = { tasks, summary: getSummary_(tasks), updatedAt: new Date().toISOString() };
    }

    return jsonResponse_(payload);
  } catch (err) {
    return jsonResponse_({ error: String(err) });
  }
}


// ════════════════════════════════════════════════════════════════════════════
//  doPost（reply-only，永遠不發 push）
// ════════════════════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    const expectedToken = getConfig_('WEBHOOK_TOKEN');
    const providedToken = (e && e.parameter && e.parameter.webhook_token) || '';
    if (expectedToken) {
      if (providedToken !== expectedToken) {
        logToSheet_('auth-fail', 'token 驗證失敗，已忽略');
        return jsonResponse_({ error: 'unauthorized' });
      }
    } else {
      logToSheet_('warn', 'WEBHOOK_TOKEN 尚未設定，webhook 未受保護');
    }

    if (!e || !e.postData || !e.postData.contents) {
      logToSheet_('empty-body', '收到沒有 body 的請求');
      return jsonResponse_({ status: 'no body' });
    }

    const body   = JSON.parse(e.postData.contents);
    const events = body.events || [];

    logToSheet_('webhook-in', '收到 ' + events.length + ' 個事件');

    events.forEach((event, idx) => {
      const source     = event.source || {};
      const replyToken = event.replyToken || null;

      const eventInfo = {
        idx: idx,
        type: event.type,
        sourceType: source.type || null,
        groupId: source.groupId || null,
        userId: source.userId || null,
        messageType: (event.message && event.message.type) || null,
        text: (event.message && event.message.text) || null,
        hasReplyToken: !!replyToken
      };
      logToSheet_('event', JSON.stringify(eventInfo));

      // ── handler dispatch ──────────────────────────────────────────────

      // Bot 加入群組 → 記錄 Group ID + 歡迎訊息
      if (event.type === 'join' && source.type === 'group' && source.groupId) {
        setConfig_('LINE_GROUP_ID', source.groupId);
        logToSheet_('handler', 'join → 記錄 Group ID ' + source.groupId);
        const welcomeMsg = '✅ CE 任務提醒機器人已加入群組\n' +
          '每天上午 10:00 推播任務提醒\n' +
          '群組內輸入 /今日任務 可立即查詢\n\n' +
          '給開發者：Group ID 已自動記錄';
        replyTextOnly_(replyToken, welcomeMsg, 'join-welcome');
      }

      // 「抓 ID」→ 覆寫 Group ID
      else if (event.type === 'message'
        && event.message && event.message.type === 'text'
        && event.message.text === '抓 ID'
        && source.groupId) {
        setConfig_('LINE_GROUP_ID', source.groupId);
        logToSheet_('handler', '抓 ID → ' + source.groupId);
        replyTextOnly_(replyToken, '✅ Group ID 已記錄：' + source.groupId, '抓ID');
      }

      // /今日任務 → Flex 即時提醒
      else if (event.type === 'message'
        && event.message && event.message.type === 'text'
        && (event.message.text === '/今日任務'
          || event.message.text === '/今日'
          || event.message.text === '/task')) {
        logToSheet_('handler', '/今日任務');
        const { altText, flexContents } = buildFlexReminderMessage_();
        replyFlexOnly_(replyToken, altText, flexContents, '/今日任務');
      }

      // 私訊「我的 ID」
      else if (event.type === 'message'
        && event.message && event.message.type === 'text'
        && event.message.text === '我的 ID'
        && source.userId) {
        logToSheet_('handler', '我的 ID → ' + source.userId);
        replyTextOnly_(replyToken, '你的 userId：\n' + source.userId, '我的ID');
      }

      // Bot 被移出群組
      else if (event.type === 'leave' && source.type === 'group' && source.groupId) {
        const currentGroupId = getConfig_('LINE_GROUP_ID');
        if (currentGroupId === source.groupId) {
          setConfig_('LINE_GROUP_ID', '');
          logToSheet_('handler', 'leave → 清除 LINE_GROUP_ID');
        } else {
          logToSheet_('handler', 'leave → 非當前群組，不處理');
        }
      }

      else {
        logToSheet_('skip', '事件 #' + idx + ' 無對應 handler');
      }
    });

    return jsonResponse_({ status: 'ok' });
  } catch (err) {
    logToSheet_('error', 'doPost 異常: ' + err);
    return jsonResponse_({ error: String(err) });
  }
}


function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ════════════════════════════════════════════════════════════════════════════
//  doPost 專用：Reply Only
//  ⚠️ 這兩個函式絕對不能呼叫任何 push API。
//  ⚠️ reply 失敗就只記 log，使用者收不到回覆，但 push 額度不會被扣。
// ════════════════════════════════════════════════════════════════════════════

function replyTextOnly_(replyToken, text, label) {
  const tag = label || 'unknown';
  if (!replyToken) {
    logToSheet_('reply-skip', '[' + tag + '] 無 replyToken，放棄（不發 push）');
    return;
  }
  try {
    callLineReplyText_(replyToken, text);
    logToSheet_('reply-ok', '[' + tag + '] reply 成功（純文字，免費）');
  } catch (err) {
    logToSheet_('reply-fail', '[' + tag + '] reply 失敗（放棄，不發 push）: ' + err);
  }
}


function replyFlexOnly_(replyToken, altText, flexContents, label) {
  const tag = label || 'unknown';
  if (!replyToken) {
    logToSheet_('reply-skip', '[' + tag + '] 無 replyToken，放棄（不發 push）');
    return;
  }
  try {
    callLineReplyFlex_(replyToken, altText, flexContents);
    logToSheet_('reply-ok', '[' + tag + '] reply 成功（Flex，免費）');
  } catch (err) {
    logToSheet_('reply-fail', '[' + tag + '] reply 失敗（放棄，不發 push）: ' + err);
  }
}


// ════════════════════════════════════════════════════════════════════════════
//  LINE API 低階呼叫（reply 給 doPost 用，push 給 sendReminder 用）
// ════════════════════════════════════════════════════════════════════════════

function callLineReplyText_(replyToken, text) {
  const token = getLineToken_();

  let safeText = text;
  if (safeText.length > MAX_MESSAGE_LEN) {
    safeText = safeText.slice(0, MAX_MESSAGE_LEN) + '\n\n...訊息過長已截斷';
  }
  const payload = {
    replyToken: replyToken,
    messages: [{ type: 'text', text: safeText }]
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  const res  = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', options);
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code !== 200) {
    throw new Error('LINE Reply Text ' + code + ': ' + body);
  }
  return body;
}


function callLineReplyFlex_(replyToken, altText, flexContents) {
  const token = getLineToken_();

  const payload = {
    replyToken: replyToken,
    messages: [{ type: 'flex', altText: altText, contents: flexContents }]
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  const res  = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', options);
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code !== 200) {
    throw new Error('LINE Reply Flex ' + code + ': ' + body);
  }
  return body;
}


// ⚠️ 這個 push 函式只給排程入口（sendReminder）用。
//    doPost 程式碼路徑「不允許」呼叫此函式。
function callLinePushFlex_(to, altText, flexContents) {
  const token = getLineToken_();
  const payload = {
    to: to,
    messages: [{ type: 'flex', altText: altText, contents: flexContents }]
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  const res  = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', options);
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code !== 200) {
    logToSheet_('line-api-fail', 'Push Flex ' + code + ': ' + body);
    throw new Error('LINE 推送失敗：' + body);
  }
  return body;
}


// ════════════════════════════════════════════════════════════════════════════
//  Flex 訊息組裝
// ════════════════════════════════════════════════════════════════════════════

function buildFlexReminderMessage_() {
  const tasks    = getAllTasks_();
  const today    = getTodayString_();
  const tomorrow = getTomorrowString_();

  const isDone      = t => t.status && t.status.indexOf('Done')     !== -1;
  const isCancelled = t => t.status && t.status.indexOf('Canceled') !== -1;
  const isActive    = t => !isDone(t) && !isCancelled(t);

  const overdue = tasks
    .filter(t => t.dueDay && isActive(t) && t.dueDay < today)
    .sort((a, b) => a.dueDay.localeCompare(b.dueDay));
  const dueToday    = tasks.filter(t => t.dueDay && isActive(t) && t.dueDay === today);
  const dueTomorrow = tasks.filter(t => t.dueDay && isActive(t) && t.dueDay === tomorrow);

  const dateLabel = Utilities.formatDate(new Date(), TIMEZONE, 'M月d日');
  const weekLabel = getWeekLabel_();

  const headerEmoji = '📋';
  const headerText  = 'CE 任務提醒';
  const altText = headerEmoji + ' ' + headerText + '｜' + weekLabel + '｜' + dateLabel;

  function taskBox(t, overdayCount) {
    const color  = getOwnerColor_(t.owner || '');
    const owner  = t.owner  || '?';
    const name   = t.name   || '（無名稱）';
    const week   = t.week   || '-';
    const status = t.status || '-';
    const meta   = overdayCount != null
      ? week + '｜逾期 ' + overdayCount + ' 天（' + t.dueDay + '）'
      : week + '｜' + status;
    return {
      type: 'box',
      layout: 'horizontal',
      spacing: 'sm',
      margin: 'sm',
      contents: [
        {
          type: 'box',
          layout: 'vertical',
          width: '52px',
          backgroundColor: color,
          cornerRadius: '20px',
          paddingAll: '4px',
          justifyContent: 'center',
          alignItems: 'center',
          contents: [{
            type: 'text', text: owner, size: 'xs', weight: 'bold',
            color: '#ffffff', align: 'center', wrap: false, adjustMode: 'shrink-to-fit'
          }]
        },
        {
          type: 'box',
          layout: 'vertical',
          flex: 1,
          spacing: 'none',
          contents: [
            { type: 'text', text: name, size: 'sm', color: '#111111', wrap: true, weight: 'bold' },
            { type: 'text', text: meta, size: 'xs', color: '#888888', wrap: true, margin: 'xs' }
          ]
        }
      ]
    };
  }

  function sectionHeader(emoji, label, count) {
    return {
      type: 'text',
      text: emoji + ' ' + label + ' ' + count + ' 筆',
      size: 'sm', weight: 'bold', color: '#333333', margin: 'md'
    };
  }

  const separator = { type: 'separator', margin: 'md' };
  const bodyContents = [];

  if (overdue.length === 0 && dueToday.length === 0 && dueTomorrow.length === 0) {
    bodyContents.push({
      type: 'text',
      text: '🎉 目前沒有逾期或即將到期的任務，大家辛苦了！',
      size: 'sm', color: '#555555', wrap: true, margin: 'md'
    });
  } else {
    if (overdue.length > 0) {
      bodyContents.push(sectionHeader('⚠️', '逾期任務', overdue.length));
      overdue.forEach(t => bodyContents.push(taskBox(t, daysBetween_(t.dueDay, today))));
    }
    if (dueToday.length > 0) {
      if (bodyContents.length > 0) bodyContents.push(separator);
      bodyContents.push(sectionHeader('📌', '今日到期', dueToday.length));
      dueToday.forEach(t => bodyContents.push(taskBox(t, null)));
    }
    if (dueTomorrow.length > 0) {
      if (bodyContents.length > 0) bodyContents.push(separator);
      bodyContents.push(sectionHeader('⏰', '明日到期', dueTomorrow.length));
      dueTomorrow.forEach(t => bodyContents.push(taskBox(t, null)));
    }
  }

  const flexContents = {
    type: 'bubble',
    size: 'mega',
    header: {
      type: 'box',
      layout: 'vertical',
      backgroundColor: '#1A1A2E',
      paddingAll: '16px',
      contents: [
        { type: 'text', text: headerEmoji + ' ' + headerText, color: '#ffffff', size: 'md', weight: 'bold' },
        { type: 'text', text: weekLabel + '｜' + dateLabel, color: '#aaaaaa', size: 'xs', margin: 'xs' }
      ]
    },
    body: {
      type: 'box',
      layout: 'vertical',
      paddingAll: '16px',
      spacing: 'none',
      contents: bodyContents
    },
    footer: {
      type: 'box',
      layout: 'vertical',
      paddingAll: '12px',
      backgroundColor: '#f5f5f5',
      contents: [
        {
          type: 'button',
          action: { type: 'uri', label: '查看完整看板', uri: DASHBOARD_URL },
          style: 'primary', color: '#1A1A2E', height: 'sm', margin: 'none'
        },
        {
          type: 'button',
          action: { type: 'uri', label: '更新任務狀態', uri: TASK_URL },
          style: 'secondary', height: 'sm', margin: 'sm'
        }
      ]
    }
  };

  return { altText, flexContents };
}


// ════════════════════════════════════════════════════════════════════════════
//  工具
// ════════════════════════════════════════════════════════════════════════════

function daysBetween_(dateStr1, dateStr2) {
  const d1 = new Date(dateStr1);
  const d2 = new Date(dateStr2);
  return Math.round(Math.abs((d2 - d1) / (1000 * 60 * 60 * 24)));
}


function getWeekLabel_() {
  const now   = new Date();
  const tzStr = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd');
  const [y, m, d] = tzStr.split('-').map(Number);
  const date  = new Date(Date.UTC(y, m - 1, d));
  const dayNum = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  const weekNo    = Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
  return 'W' + String(weekNo).padStart(2, '0');
}


// ════════════════════════════════════════════════════════════════════════════
//  排程入口（唯一會呼叫 push API 的地方）
//  CE 版本：每日一次，固定 10:00 AM 台灣時間
// ════════════════════════════════════════════════════════════════════════════

function sendReminder() {
  const groupId = getConfig_('LINE_GROUP_ID');
  if (!groupId) {
    logToSheet_('scheduled', 'sendReminder 跳過：LINE_GROUP_ID 未設定');
    return;
  }
  try {
    const { altText, flexContents } = buildFlexReminderMessage_();
    callLinePushFlex_(groupId, altText, flexContents);
    logToSheet_('scheduled', '💸 sendReminder 推播完成（push 額度 -1）→ ' + groupId);
  } catch (err) {
    logToSheet_('scheduled-fail', 'sendReminder 失敗: ' + err);
    throw err;
  }
}

// 向後相容別名
function sendDailyReminder() { sendReminder(); }


// ════════════════════════════════════════════════════════════════════════════
//  Trigger 檢視
// ════════════════════════════════════════════════════════════════════════════

function listTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  Logger.log('目前共有 ' + triggers.length + ' 個觸發器：');
  triggers.forEach((t, i) => {
    Logger.log((i + 1) + '. ' + t.getHandlerFunction() + ' | ' + t.getEventType() + ' | uid=' + t.getUniqueId());
  });
}


// ════════════════════════════════════════════════════════════════════════════
//  測試
// ════════════════════════════════════════════════════════════════════════════

function testSheetLog() {
  logToSheet_('test', '這是一筆測試 log，時間：' + new Date());
  Logger.log('已寫入 debug_log 分頁');
}


function testSendReminderDryRun() {
  const { altText, flexContents } = buildFlexReminderMessage_();
  Logger.log('altText: ' + altText);
  Logger.log(JSON.stringify(flexContents, null, 2));
}


function testSendReminder() {
  const groupId = getConfig_('LINE_GROUP_ID');
  Logger.log('LINE_GROUP_ID: ' + (groupId || '未設定'));
  if (groupId) {
    sendReminder();
    Logger.log('已推送到群組');
  } else {
    Logger.log('沒有 Group ID，未推送');
  }
}


function testSendReminderToMyself() {
  if (!MY_USER_ID || MY_USER_ID === 'YOUR_LINE_USER_ID') {
    Logger.log('MY_USER_ID 尚未設定，請先更新程式碼頂部的 MY_USER_ID 常數');
    return;
  }
  const { altText, flexContents } = buildFlexReminderMessage_();
  callLinePushFlex_(MY_USER_ID, altText, flexContents);
  logToSheet_('manual-test', '💸 手動推送到自己 → ' + MY_USER_ID);
}


function testGetTasks() {
  const tasks = getAllTasks_();
  Logger.log('共 ' + tasks.length + ' 筆任務');
  Logger.log(JSON.stringify(tasks.slice(0, 3), null, 2));
}


function testGetSummary() {
  const tasks = getAllTasks_();
  Logger.log(JSON.stringify(getSummary_(tasks), null, 2));
}


function testGetConfig() {
  Logger.log('API_TOKEN: '              + (getConfig_('API_TOKEN')              ? '已設定' : '未設定'));
  Logger.log('CHANNEL_SECRET: '         + (getConfig_('CHANNEL_SECRET')         ? '已設定' : '未設定'));
  const _tok = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  Logger.log('CHANNEL_ACCESS_TOKEN: '   + (_tok ? '已設定（Script Properties）' : '⚠️ 未設定'));
  Logger.log('WEBHOOK_TOKEN: '          + (getConfig_('WEBHOOK_TOKEN')          ? '已設定' : '未設定'));
  Logger.log('LINE_GROUP_ID: '          + (getConfig_('LINE_GROUP_ID')          || '未設定'));
  Logger.log('MY_USER_ID（程式碼常數）: ' + (MY_USER_ID                          || '未設定'));
  Logger.log('今天（台灣時區）：'        + getTodayString_());
  Logger.log('明天（台灣時區）：'        + getTomorrowString_());
  Logger.log('本週：'                   + getWeekLabel_());
}


/**
 * CE 版本：建立每天 10:00 AM（台北時區）的單一觸發器
 *
 * 操作方式：從上方函式下拉選 setupDailyTrigger → 點 ▶ 執行
 * 執行後可用 listTriggers() 確認觸發器已正確建立。
 *
 * 注意：此函式會先清除所有舊的 reminder 觸發器（包含從 Xcity 遷移過來的
 * sendMorningReminder / sendEveningReminder），再建立 CE 版本的單一每日觸發器。
 */
function setupDailyTrigger() {
  // 清除所有舊的 reminder 觸發器（含 Xcity 遺留的早晚排程）
  ScriptApp.getProjectTriggers().forEach(t => {
    const handler = t.getHandlerFunction();
    if (handler === 'sendReminder'
      || handler === 'sendDailyReminder'
      || handler === 'sendMorningReminder'   // Xcity 遺留
      || handler === 'sendEveningReminder') { // Xcity 遺留
      ScriptApp.deleteTrigger(t);
      Logger.log('已刪除舊觸發器：' + handler);
    }
  });

  // 建立每天 10:00 AM 台北時間的觸發器
  ScriptApp.newTrigger('sendReminder')
    .timeBased()
    .everyDays(1)
    .atHour(10)
    .nearMinute(0)
    .inTimezone('Asia/Taipei')
    .create();

  Logger.log('✅ 已建立：每天 10:00 AM (Asia/Taipei) sendReminder');
}
