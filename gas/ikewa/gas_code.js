// ============================================================
// コワーキングスペース 入退館管理システム - Google Apps Script
// ============================================================
// 【設定】以下の値をご自身の環境に合わせて変更してください

const CONFIG = {
  SHEET_NAME_LOG: '入退館ログ',
  SHEET_NAME_MEMBERS: '会員マスタ',
  SHEET_NAME_STATS: '統計',
  ADMIN_EMAIL: 'info@handanotane.com',
  FROM_EMAIL:   'info@handanotane.com',
  SPACE_NAME: 'cococorin',
  DROPIN_HOURLY: 400,
  DROPIN_DAILY_MAX: 1000,
};

// 会員種別と月額料金
const MEMBER_TYPES = {
  'monthly_general':  { label: '月額会員（一般）',         monthly: 8000, isMonthly: true },
  'monthly_student':  { label: '月額会員（学生）',         monthly: 4000, isMonthly: true },
  'monthly_weekend':  { label: '月額プラン（土日祝）',     monthly: 4000, isMonthly: true },
  'monthly_weekday':  { label: '月額プラン（平日）',       monthly: 6000, isMonthly: true },
  'rental_office':    { label: 'レンタルオフィス入居者',   monthly: 0,    isMonthly: true },
  'dropin':           { label: 'ドロップイン',             monthly: null, isMonthly: false },
};

// ============================================================
// メインエントリーポイント（GETリクエスト処理）
// ============================================================
function doGet(e) {
  const params = e.parameter || {};
  // 受信停止ページ
  if (params.action === 'unsubscribe') {
    const result = unsubscribe(params.memberId);
    const html = result.success
      ? `<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>メール受信停止</title>
<style>body{font-family:sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;background:#f8f8f6;}
.card{background:#fff;border-radius:16px;padding:2.5rem 2rem;max-width:420px;width:90%;text-align:center;box-shadow:0 2px 16px rgba(0,0,0,0.08);}
h2{color:#2e2826;font-size:18px;margin-bottom:1rem;}
p{color:#555;font-size:14px;line-height:1.8;margin:0 0 1rem;}
a{color:#00a3af;}</style></head>
<body><div class="card">
<h2>メールの受信を停止しました</h2>
<p>今後、このメールアドレス宛のメールはお届けしません。</p>
<p>再度メールを受け取りたい場合は、<a href="mailto:info@handanotane.com">info@handanotane.com</a> までご連絡ください。</p>
<p style="color:#aaa;font-size:12px;">cococorin</p>
</div></body></html>`
      : `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>エラー</title>
<style>body{font-family:sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;background:#f8f8f6;}
.card{background:#fff;border-radius:16px;padding:2.5rem 2rem;max-width:420px;width:90%;text-align:center;}
h2{color:#ce5242;}p{color:#555;font-size:14px;}</style></head>
<body><div class="card"><h2>エラーが発生しました</h2><p>${result.error || '会員情報が見つかりませんでした'}</p></div></body></html>`;
    return HtmlService.createHtmlOutput(html).setTitle('メール受信停止');
  }
  return handleRequest(e);
}

// POSTリクエスト処理
function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const params = e.parameter || {};
  const action = params.action;

  let result;
  try {
    switch (action) {
      case 'checkin':
        result = doCheckIn(params.memberId);
        break;
      case 'checkout':
        result = doCheckOut(params.memberId);
        break;
      case 'getMember':
        result = getMemberInfo(params.memberId);
        break;
      case 'getLog':
        result = getLog(params.date);
        break;
      case 'getStats':
        result = getStats();
        break;
      case 'searchByEmail':
        result = searchByEmail(params.email);
        break;
      case 'getActiveUsers':
        result = getActiveUsers();
        break;
      case 'getMailSettings':
        result = getMailSettings();
        break;
      case 'saveMailSettings':
        result = saveMailSettings(params);
        break;
      case 'unsubscribe':
        result = unsubscribe(params.memberId);
        break;
      default:
        result = { success: false, error: '不明なアクションです' };
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// 入館処理
// ============================================================
function doCheckIn(memberId) {
  if (!memberId) return { success: false, error: '会員番号が入力されていません' };

  const member = findMember(memberId);
  if (!member) return { success: false, error: '会員番号が見つかりません: ' + memberId };

  // 月額プランの曜日チェック
  const dayCheck = checkDayRestriction(member.type);
  if (!dayCheck.ok) return { success: false, error: dayCheck.message };

  // 既に入館中かチェック
  const existing = findActiveSession(memberId);
  if (existing) return { success: false, error: 'すでに入館中です' };

  // 平日プランが土日に来た場合はdropin扱いで入館
  const effectiveType = dayCheck.dropinFallback ? 'dropin' : member.type;
  const fallbackLabel = member.type === 'monthly_weekend'
    ? 'ドロップイン（休日プラン・平日利用）'
    : 'ドロップイン（平日プラン・土日利用）';
  const effectiveLabel = dayCheck.dropinFallback ? fallbackLabel : (MEMBER_TYPES[member.type]?.label || member.type);

  const now = new Date();
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAME_LOG);

  ensureLogHeader(sheet);

  sheet.appendRow([
    Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd'),
    Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm:ss'),
    '',
    memberId,
    member.name,
    effectiveType,
    effectiveLabel,
    '',
    '',
    '利用中',
  ]);

  return {
    success: true,
    action: 'checkin',
    memberId: memberId,
    name: member.name,
    type: effectiveType,
    typeLabel: effectiveLabel,
    checkinTime: Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm'),
    isMonthly: dayCheck.dropinFallback ? false : (MEMBER_TYPES[member.type]?.isMonthly || false),
    dropinFallback: dayCheck.dropinFallback || false,
    fallbackMessage: dayCheck.message || '',
  };
}


// ============================================================
// 退館処理
// ============================================================
function doCheckOut(memberId) {
  if (!memberId) return { success: false, error: '会員番号が入力されていません' };

  const member = findMember(memberId);
  if (!member) return { success: false, error: '会員番号が見つかりません: ' + memberId };

  const session = findActiveSession(memberId);
  if (!session) return { success: false, error: '入館記録が見つかりません' };

  const now = new Date();
  const checkinDate = session.date + ' ' + session.checkinTime;
  const checkinDt = new Date(checkinDate.replace(/\//g, '-').replace(' ', 'T') + '+09:00');
  const diffMs = now - checkinDt;
  const diffMin = Math.max(1, Math.floor(diffMs / 60000));
  const diffHours = Math.floor(diffMin / 60);
  const diffRemain = diffMin % 60;
  const durationStr = diffHours > 0 ? `${diffHours}時間${diffRemain}分` : `${diffMin}分`;

  // 料金計算（入館時の実効プランで判定）
  const effectiveType = session.effectiveType || member.type;
  let fee = 0;
  let feeLabel = '';
  if (!MEMBER_TYPES[effectiveType]?.isMonthly) {
    const hours = Math.ceil((diffMin - 9) / 60);  // 9分のバッファ（例：1時間9分→1時間として計算）
    fee = Math.min(Math.max(hours, 1) * CONFIG.DROPIN_HOURLY, CONFIG.DROPIN_DAILY_MAX);
    feeLabel = `¥${fee.toLocaleString()}`;
  } else {
    feeLabel = '月額会員（追加料金なし）';
  }

  // ログ行を更新
  updateSessionRow(session.rowIndex, now, durationStr, fee);

  // 最終利用日を会員マスタに記録
  updateLastVisit(memberId, now);

  // 初回利用チェック＆メール送信
  if (member.email && isFirstVisit(memberId)) {
    const mailOpt = getMemberMailOption(memberId);
    const globalMailEnabled = getMailEnabled();
    if (mailOpt !== '0' && globalMailEnabled) {
      sendFirstVisitEmail(member, session.checkinTime, Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm'), durationStr, fee);
    }
  }

  return {
    success: true,
    action: 'checkout',
    memberId: memberId,
    name: member.name,
    type: effectiveType,
    typeLabel: MEMBER_TYPES[effectiveType]?.label || member.type,
    checkinTime: session.checkinTime.substring(0, 5),
    checkoutTime: Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm'),
    duration: durationStr,
    fee: fee,
    feeLabel: feeLabel,
    isMonthly: MEMBER_TYPES[effectiveType]?.isMonthly || false,
  };
}

// ============================================================
// ヘルパー：会員マスタ検索
// ============================================================
function findMember(memberId) {
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAME_MEMBERS);
  const data = sheet.getDataRange().getValues();
  const inputNum = parseInt(memberId, 10);
  for (let i = 1; i < data.length; i++) {
    const rowNum = parseInt(String(data[i][0]).trim(), 10);
    if (!isNaN(rowNum) && rowNum === inputNum) {
      return {
        id:    String(data[i][0]),
        name:  data[i][1],
        type:  data[i][2],
        email: data[i][3] || '',
        note:  data[i][4] || '',
      };
    }
  }
  return null;
}

// ============================================================
// ヘルパー：アクティブセッション検索（入館中レコード）
// ============================================================
function findActiveSession(memberId) {
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAME_LOG);
  const data = sheet.getDataRange().getValues();
  // 最新行から逆順で検索
  const inputNum = parseInt(memberId, 10);
  for (let i = data.length - 1; i >= 1; i--) {
    const rowNum = parseInt(String(data[i][3]).trim(), 10);
    if (!isNaN(rowNum) && rowNum === inputNum && data[i][9] === '利用中') {
      const rawDate = data[i][0];
      const rawTime = data[i][1];
      const checkinDate = (rawDate instanceof Date)
        ? Utilities.formatDate(rawDate, 'Asia/Tokyo', 'yyyy/MM/dd')
        : String(rawDate);
      const checkinTime = (rawTime instanceof Date)
        ? Utilities.formatDate(rawTime, 'Asia/Tokyo', 'HH:mm:ss')
        : String(rawTime);
      return {
        rowIndex:    i + 1,
        date:        checkinDate,
        checkinTime: checkinTime,
        effectiveType: String(data[i][5]).trim(), // 入館時に記録した会員種別コード
      };
    }
  }
  return null;
}

// ============================================================
// ヘルパー：ログ行の退館情報を更新
// ============================================================
function updateSessionRow(rowIndex, checkoutTime, duration, fee) {
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAME_LOG);
  sheet.getRange(rowIndex, 3).setValue(Utilities.formatDate(checkoutTime, 'Asia/Tokyo', 'HH:mm:ss'));
  sheet.getRange(rowIndex, 8).setValue(duration);
  sheet.getRange(rowIndex, 9).setValue(fee > 0 ? fee : '');
  sheet.getRange(rowIndex, 10).setValue('退館済');
}

// ============================================================
// ヘルパー：祝日判定（Googleカレンダー参照）
// ============================================================
function isJapaneseHoliday(date) {
  try {
    const calendar = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
    if (!calendar) return false;
    const startOfDay = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0);
    const endOfDay   = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59);
    const events = calendar.getEvents(startOfDay, endOfDay);
    return events.length > 0;
  } catch (e) {
    console.log('祝日カレンダー取得エラー:', e.message);
    return false;
  }
}

// ============================================================
// ヘルパー：曜日制限チェック（祝日対応）
// ============================================================
function checkDayRestriction(type) {
  const now = new Date();
  const day = now.getDay(); // 0=日, 1=月..., 6=土
  const isSatOrSun = (day === 0 || day === 6);
  const isHoliday  = isJapaneseHoliday(now);
  const isWeekend  = isSatOrSun || isHoliday; // 土日 または 祝日

  if (type === 'monthly_weekend' && !isWeekend) {
    // エラーにせずドロップイン扱いで入館させる
    return { ok: true, dropinFallback: true, message: '休日プランのため、本日はドロップインでのご利用となります（400円/時間・上限1,000円/日）' };
  }
  if (type === 'monthly_weekday' && isWeekend) {
    // エラーにせずドロップイン扱いで入館させる
    return { ok: true, dropinFallback: true, message: '平日プランのため、本日はドロップインでのご利用となります（400円/時間・上限1,000円/日）' };
  }
  return { ok: true, dropinFallback: false };
}

// ============================================================
// 【検証用】祝日判定とプラン制限の動作確認
//   引数なし → 今日
//   引数あり → 'YYYY/MM/DD' または 'YYYY-MM-DD' を渡す
// 例： debugHolidayCheck('2026/05/04')
// ============================================================
function debugHolidayCheck(dateStr) {
  const date = dateStr ? new Date(String(dateStr).replace(/\//g, '-') + 'T12:00:00+09:00') : new Date();
  const day = date.getDay();
  const isSatOrSun = (day === 0 || day === 6);
  const isHoliday = isJapaneseHoliday(date);
  const isWeekend = isSatOrSun || isHoliday;
  console.log('日付: ' + Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd (EEE)'));
  console.log('  getDay=' + day + ' isSatOrSun=' + isSatOrSun + ' isHoliday=' + isHoliday + ' → isWeekend=' + isWeekend);
  if (isWeekend) {
    console.log('  monthly_weekend (土日祝プラン) : ✓ 月額利用OK（追加料金なし）');
    console.log('  monthly_weekday (平日プラン)   : ❌ ドロップイン扱い（400円/時間）');
  } else {
    console.log('  monthly_weekend (土日祝プラン) : ❌ ドロップイン扱い（400円/時間）');
    console.log('  monthly_weekday (平日プラン)   : ✓ 月額利用OK（追加料金なし）');
  }
}

// ============================================================
// 【検証用】指定会員番号の checkDayRestriction を本日でシミュレーション
// 例： debugDayRestrictionForMember('10001')
// ============================================================
function debugDayRestrictionForMember(memberId) {
  const member = findMember(memberId);
  if (!member) { console.log('会員が見つかりません: ' + memberId); return; }
  console.log('会員番号=' + member.id + ' 氏名=' + member.name + ' 種別コード=' + member.type
    + ' 種別ラベル=' + (MEMBER_TYPES[member.type] && MEMBER_TYPES[member.type].label));
  const r = checkDayRestriction(member.type);
  console.log('checkDayRestriction結果: ok=' + r.ok + ' dropinFallback=' + (r.dropinFallback || false)
    + ' message=' + (r.message || '(なし)'));
  debugHolidayCheck();
}

// ============================================================
// ヘルパー：メールアドレスで会員検索
// ============================================================
function searchByEmail(email) {
  if (!email) return { success: false, error: 'メールアドレスが入力されていません' };
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAME_MEMBERS);
  const data = sheet.getDataRange().getValues();
  const target = email.trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    const rowEmail = String(data[i][3]).trim().toLowerCase();
    if (rowEmail === target) {
      const type = data[i][2];
      return {
        success:   true,
        memberId:  String(data[i][0]),
        name:      data[i][1],
        type:      type,
        typeLabel: MEMBER_TYPES[type]?.label || type,
      };
    }
  }
  return { success: false, error: '該当する会員が見つかりませんでした' };
}


function getMemberInfo(memberId) {
  const member = findMember(memberId);
  if (!member) return { success: false, error: '会員番号が見つかりません' };
  return {
    success: true,
    ...member,
    typeLabel: MEMBER_TYPES[member.type]?.label || member.type,
    isMonthly: MEMBER_TYPES[member.type]?.isMonthly || false,
  };
}

// ============================================================
// ヘルパー：ログ取得（管理者画面用）
// ============================================================
function formatLogDate(val) {
  if (val instanceof Date) return Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy/MM/dd');
  return String(val).trim();
}

function formatLogTime(val) {
  if (val instanceof Date) return Utilities.formatDate(val, 'Asia/Tokyo', 'HH:mm');
  return String(val).trim().substring(0, 5);
}

function getLog(dateStr) {
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAME_LOG);
  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const rowDate = formatLogDate(data[i][0]);
    if (!rowDate) continue;
    rows.push({
      date:         rowDate,
      checkinTime:  formatLogTime(data[i][1]),
      checkoutTime: data[i][2] ? formatLogTime(data[i][2]) : '',
      memberId:     data[i][3],
      name:         data[i][4],
      typeLabel:    data[i][6],
      duration:     data[i][7] || '',
      fee:          data[i][8] || '',
      status:       data[i][9],
    });
  }
  return { success: true, logs: rows.reverse() };
}

// ============================================================
// ヘルパー：現在の利用中ユーザー取得
// ============================================================
function getActiveUsers() {
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAME_LOG);
  const data = sheet.getDataRange().getValues();
  const active = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][9] === '利用中') {
      active.push({
        memberId:    data[i][3],
        name:        data[i][4],
        typeLabel:   data[i][6],
        checkinTime: formatLogTime(data[i][1]),
      });
    }
  }
  return { success: true, users: active };
}

// ============================================================
// ヘルパー：統計データ取得
// ============================================================
function getStats() {
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAME_LOG);
  const data = sheet.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  const thisMonth = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM');

  let todayCount = 0, todayRevenue = 0;
  let monthCount = 0, monthRevenue = 0;
  const dailyMap = {};

  for (let i = 1; i < data.length; i++) {
    const rowDate = formatLogDate(data[i][0]);
    if (!rowDate) continue;
    const fee = Number(data[i][8]) || 0;

    if (rowDate === today) { todayCount++; todayRevenue += fee; }
    if (rowDate.startsWith(thisMonth)) { monthCount++; monthRevenue += fee; }

    dailyMap[rowDate] = (dailyMap[rowDate] || { count: 0, revenue: 0 });
    dailyMap[rowDate].count++;
    dailyMap[rowDate].revenue += fee;
  }

  const daily = Object.entries(dailyMap)
    .sort((a, b) => a[0] < b[0] ? 1 : -1)
    .slice(0, 30)
    .map(([date, v]) => ({ date, ...v }));

  return {
    success: true,
    today: { count: todayCount, revenue: todayRevenue },
    month: { count: monthCount, revenue: monthRevenue },
    daily,
  };
}


// ============================================================
// 初回利用判定（2026/04/18以降のログに記録があるか）
// ============================================================
function isFirstVisit(memberId) {
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAME_LOG);
  const data = sheet.getDataRange().getValues();
  const inputNum = parseInt(memberId, 10);
  const startDate = '2026/04/18';
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const rowNum = parseInt(String(data[i][3]).trim(), 10);
    if (!isNaN(rowNum) && rowNum === inputNum) {
      const rowDate = (data[i][0] instanceof Date)
        ? Utilities.formatDate(data[i][0], 'Asia/Tokyo', 'yyyy/MM/dd')
        : String(data[i][0]).trim();
      if (rowDate >= startDate) count++;
    }
  }
  // 今回の退館分を含めて1件のみなら初回
  return count <= 1;
}

// ============================================================
// 最終利用日を会員マスタに記録（G列）
// ============================================================
function updateLastVisit(memberId, date) {
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAME_MEMBERS);
  const data = sheet.getDataRange().getValues();
  const inputNum = parseInt(memberId, 10);
  for (let i = 1; i < data.length; i++) {
    const rowNum = parseInt(String(data[i][0]).trim(), 10);
    if (!isNaN(rowNum) && rowNum === inputNum) {
      const dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
      sheet.getRange(i + 1, 7).setValue(dateStr); // G列
      return;
    }
  }
}

// ============================================================
// 会員のメール受信設定を取得（F列）
// ============================================================
function getMemberMailOption(memberId) {
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAME_MEMBERS);
  const data = sheet.getDataRange().getValues();
  const inputNum = parseInt(memberId, 10);
  for (let i = 1; i < data.length; i++) {
    const rowNum = parseInt(String(data[i][0]).trim(), 10);
    if (!isNaN(rowNum) && rowNum === inputNum) {
      const val = String(data[i][5]).trim(); // F列
      return val === '' ? '1' : val; // 未設定は受け取る扱い
    }
  }
  return '1';
}

// ============================================================
// メール受信拒否（unsubscribe）
// ============================================================
function unsubscribe(memberId) {
  if (!memberId) return { success: false, error: '会員番号が必要です' };
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAME_MEMBERS);
  const data = sheet.getDataRange().getValues();
  const inputNum = parseInt(memberId, 10);
  for (let i = 1; i < data.length; i++) {
    const rowNum = parseInt(String(data[i][0]).trim(), 10);
    if (!isNaN(rowNum) && rowNum === inputNum) {
      sheet.getRange(i + 1, 6).setValue('0'); // F列を0に
      return { success: true, message: 'メール受信を停止しました' };
    }
  }
  return { success: false, error: '会員番号が見つかりません' };
}

// ============================================================
// メール配信のグローバルON/OFFを取得
// ============================================================
function getMailEnabled() {
  const settings = getMailSettings().settings || {};
  return settings['mail_enabled'] !== '0';
}

// ============================================================
// メール設定をスプレッドシートから取得
// ============================================================
function getMailSettings() {
  const ss = getDestSpreadsheet();
  let sheet = ss.getSheetByName('メール設定');
  if (!sheet) return { success: false, error: 'メール設定シートが見つかりません' };
  const data = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 0; i < data.length; i++) {
    if (data[i][0]) settings[String(data[i][0]).trim()] = String(data[i][1] !== undefined && data[i][1] !== null ? data[i][1] : '');
  }
  return { success: true, settings };
}

// ============================================================
// メール設定をスプレッドシートに保存
// ============================================================
function saveMailSettings(params) {
  const ss = getDestSpreadsheet();
  let sheet = ss.getSheetByName('メール設定');
  if (!sheet) sheet = ss.insertSheet('メール設定');
  const data = sheet.getDataRange().getValues();
  const updates = {
    first_visit_subject: params.first_visit_subject || '',
    first_visit_body:    params.first_visit_body    || '',
    mail_enabled:        params.mail_enabled !== undefined ? params.mail_enabled : '1',
  };
  for (const [key, val] of Object.entries(updates)) {
    let found = false;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === key) {
        sheet.getRange(i + 1, 2).setValue(val);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, val]);
    }
  }
  return { success: true, message: '保存しました' };
}

// ============================================================
// 初回利用メール送信
// ============================================================
function sendFirstVisitEmail(member, checkinTime, checkoutTime, duration, fee) {
  const subject = `【cococorin】はじめてのご利用ありがとうございました`;
  const unsubscribeUrl = `${ScriptApp.getService().getUrl()}?action=unsubscribe&memberId=${member.id}`;
  const isMonthly = MEMBER_TYPES[member.type]?.isMonthly || false;
  const feeText = isMonthly ? '月額会員のため無料' : `¥${fee.toLocaleString()}`;

  // 共通パーツ
  const header = `
    <div style="background:#2e2826;padding:18px 24px;border-radius:8px 8px 0 0;">
      <p style="color:#fff;font-size:17px;font-weight:bold;margin:0;">cococorin</p>
      <p style="color:#aaa;font-size:11px;margin:3px 0 0;">半田市創造・連携・実践センター</p>
    </div>`;

  const greeting = `
    <p style="font-size:14px;margin:0 0 14px;">${member.name} 様</p>
    <p style="font-size:13px;line-height:1.85;margin:0 0 14px;">はじめてcococorinをご利用いただきありがとうございました。</p>
    <p style="font-size:13px;line-height:1.85;margin:0 0 14px;">cococorinは、静かに集中するだけじゃない場所です。隣接するカフェの心地よいざわめきのなかで、自分のペースで仕事や勉強に向き合える、そんな「ちょうどいい」空間を目指しています。そして、ここに集まる人同士がゆるやかにつながり、あたたかいコミュニティに育っていけたらと思っています。またぜひ、顔を出しにきてください。</p>`;

  const usageInfo = `
    <div style="background:#f8f8f6;border-radius:8px;padding:14px 18px;margin:18px 0;">
      <p style="font-size:11px;color:#888;font-weight:bold;margin:0 0 10px;">本日のご利用内容</p>
      <table style="font-size:12px;width:100%;border-collapse:collapse;">
        <tr><td style="color:#888;padding:3px 0;width:80px;">入館時刻</td><td>${checkinTime.substring(0,5)}</td></tr>
        <tr><td style="color:#888;padding:3px 0;">退館時刻</td><td>${checkoutTime}</td></tr>
        <tr><td style="color:#888;padding:3px 0;">利用時間</td><td>${duration}</td></tr>
        <tr><td style="color:#888;padding:3px 0;">ご利用料金</td><td style="font-weight:bold;">${feeText}</td></tr>
      </table>
    </div>`;

  const survey = `
    <div style="border-left:3px solid #00a3af;padding:10px 14px;margin:18px 0;background:#f0fafb;border-radius:0 8px 8px 0;">
      <p style="font-size:12px;color:#085041;line-height:1.75;margin:0;">ご利用の感想をぜひ聞かせてください。アンケートにお答えいただいた方には、次回ご来館時に<strong>コワーキング1時間無料券 または cococorinオリジナルステッカー</strong>をプレゼントしています。<br><br>
      <a href="https://forms.gle/tvYFJqgT1HFQ2LEaA" style="color:#00a3af;font-weight:bold;">アンケートに答える（1〜2分）</a><br><br>
      ご回答後、次回ご来館時にスタッフへお声がけください。</p>
    </div>`;

  const planInfo = `
    <div style="border:1px solid #c8ead9;border-radius:8px;padding:14px 18px;margin:18px 0;background:#f0faf7;">
      <p style="font-size:11px;color:#0F6E56;font-weight:bold;margin:0 0 8px;">よく来るなら、月額プランがおトクかも</p>
      <table style="font-size:12px;width:100%;border-collapse:collapse;">
        <tr><td style="color:#0F6E56;padding:3px 0;width:140px;">月額プラン（平日）</td><td>月6,000円〜 平日使い放題</td></tr>
        <tr><td style="color:#0F6E56;padding:3px 0;">月額プラン（土日祝）</td><td>月4,000円〜 土日祝使い放題</td></tr>
        <tr><td style="color:#0F6E56;padding:3px 0;">月額会員（一般）</td><td>月8,000円〜 毎日使い放題</td></tr>
        <tr><td style="color:#0F6E56;padding:3px 0;">月額会員（学生）</td><td>月4,000円〜 毎日使い放題</td></tr>
      </table>
      <p style="font-size:11px;color:#0F6E56;margin:10px 0 0;">詳しくは<a href="https://handanotane.com/news/howtousecoworkingspace/" style="color:#0F6E56;">こちらのページ</a>をご覧ください。</p>
    </div>`;

  const footer_contact = `<p style="font-size:12px;color:#888;margin:14px 0 0;">ご不明な点は <a href="mailto:info@handanotane.com" style="color:#00a3af;">info@handanotane.com</a> までお気軽にどうぞ。</p>`;

  const footer_unsub = `
    <div style="background:#f4f4f0;padding:12px 24px;border-radius:0 0 8px 8px;border:1px solid #e0e0e0;border-top:none;">
      <p style="font-size:11px;color:#aaa;margin:0;">メールの受信を停止する場合は<a href="${unsubscribeUrl}" style="color:#00a3af;">こちら</a>からお手続きください。</p>
    </div>`;

  // ドロップイン向け（月額プラン案内あり）
  // 月額会員向け（月額プラン案内なし）
  const bodyContent = isMonthly
    ? `${greeting}${usageInfo}${survey}${footer_contact}`
    : `${greeting}${usageInfo}${planInfo}${survey}${footer_contact}`;

  const htmlBody = `<div style="font-family:sans-serif;max-width:560px;margin:0 auto;color:#333;">
    ${header}
    <div style="background:#fff;padding:24px;border:1px solid #e0e0e0;">
      ${bodyContent}
    </div>
    ${footer_unsub}
  </div>`;

  const plainBody = isMonthly
    ? `${member.name} 様\n\nはじめてcococorinをご利用いただきありがとうございました。\n\n入館時刻：${checkinTime.substring(0,5)}\n退館時刻：${checkoutTime}\n利用時間：${duration}\n\nアンケートにご協力ください：https://forms.gle/tvYFJqgT1HFQ2LEaA\n（回答後、次回来館時にスタッフへお声がけください）\n\ncococorin\ninfo@handanotane.com`
    : `${member.name} 様\n\nはじめてcococorinをご利用いただきありがとうございました。\n\n入館時刻：${checkinTime.substring(0,5)}\n退館時刻：${checkoutTime}\n利用時間：${duration}\nご利用料金：${feeText}\n\nよく来るなら月額プランもご検討ください。詳細は info@handanotane.com まで。\n\nアンケートにご協力ください：https://forms.gle/tvYFJqgT1HFQ2LEaA\n（回答後、次回来館時にスタッフへお声がけください）\n\ncococorin\ninfo@handanotane.com`;

  try {
    GmailApp.sendEmail(member.email, subject, plainBody, {
      name: 'cococorin',
      from: CONFIG.FROM_EMAIL,
      replyTo: CONFIG.ADMIN_EMAIL,
      htmlBody: htmlBody,
    });
    console.log('初回利用メール送信完了:', member.name, isMonthly ? '月額会員向け' : 'ドロップイン向け');
  } catch (e) {
    console.log('メール送信エラー:', e.message);
  }
}

// ============================================================
// ヘルパー：サンキューメール送信
// ============================================================
function sendThankYouEmail(member, checkinTime, checkoutTime, duration, fee) {
  const subject = `【${CONFIG.SPACE_NAME}】ご利用ありがとうございました`;
  const feeText = fee > 0 ? `ご利用料金：¥${fee.toLocaleString()}` : '月額会員のため追加料金はありません';
  const body = `${member.name} 様

本日も${CONFIG.SPACE_NAME}をご利用いただきありがとうございました。

■ ご利用内容
　入館時刻：${checkinTime.substring(0, 5)}
　退館時刻：${checkoutTime}
　利用時間：${duration}
　${feeText}

またのご利用をお待ちしております。

${CONFIG.SPACE_NAME}`;

  try {
    GmailApp.sendEmail(member.email, subject, body, {
      name: CONFIG.SPACE_NAME,
      from: CONFIG.FROM_EMAIL,
      replyTo: CONFIG.ADMIN_EMAIL,
    });
  } catch (e) {
    console.log('メール送信エラー:', e.message);
  }
}

// ============================================================
// ヘルパー：シート取得または作成
// ============================================================

// ★「コワーキング入退館管理」スプレッドシートのID（書き込み先を明示固定）
const DEST_SS_ID = '1knYE9NMyYkVAWQqqNb5DsUoUHLAF4RFVQ3k6MOzTgCU';

function getDestSpreadsheet() {
  return SpreadsheetApp.openById(DEST_SS_ID);
}

function getOrCreateSheet(name) {
  const ss = getDestSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === CONFIG.SHEET_NAME_LOG) ensureLogHeader(sheet);
    if (name === CONFIG.SHEET_NAME_MEMBERS) ensureMemberHeader(sheet);
  }
  return sheet;
}

function ensureLogHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['日付', '入館時刻', '退館時刻', '会員番号', '氏名', '会員種別コード', '会員種別', '利用時間', '料金', '状態']);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#E1F5EE');
  }
}

function ensureMemberHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['会員番号', '氏名', '会員種別', 'メールアドレス', '備考']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#E1F5EE');
    // サンプルデータ
    sheet.appendRow(['10001', '田中 太郎', 'monthly_general', 'tanaka@example.com', '']);
    sheet.appendRow(['10002', '鈴木 花子', 'monthly_student', 'suzuki@example.com', '']);
    sheet.appendRow(['20001', 'ゲスト',   'dropin',          '', '都度利用']);
  }
}

// ============================================================
// リマインドメール送信（時間トリガー設定用）
// 毎月末に翌月の月額会員へリマインドを送信
// ============================================================
function sendMonthlyReminder() {
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAME_MEMBERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const type = data[i][2];
    const email = data[i][3];
    const name = data[i][1];
    if (!email || !MEMBER_TYPES[type]?.isMonthly) continue;

    const subject = `【${CONFIG.SPACE_NAME}】来月もよろしくお願いします`;
    const body = `${name} 様

いつも${CONFIG.SPACE_NAME}をご利用いただきありがとうございます。
来月もご利用をお待ちしております。

${CONFIG.SPACE_NAME}`;
    try {
      GmailApp.sendEmail(email, subject, body, { name: CONFIG.SPACE_NAME });
    } catch (e) {
      console.log('リマインドメールエラー:', name, e.message);
    }
  }
}

// ============================================================
// 会員登録フォーム回答から会員マスタを同期
// GASエディタから手動実行 または トリガーで定期実行
// ============================================================
// ============================================================
// 会員登録フォーム回答から会員マスタを同期
// GASエディタから手動実行 または フォーム送信トリガーで自動実行
// ============================================================
function syncMembersFromForm(e) {
  const PLAN_MAP = {
    'ドロップイン（一時利用）': 'dropin',
    '月額会員（一般）':         'monthly_general',
    '月額会員（学生）':         'monthly_student',
    '月額プラン（土日祝）':     'monthly_weekend',
    '月額プラン（平日）':       'monthly_weekday',
    'レンタルオフィス入居者':   'rental_office',
  };

  const COL_EMAIL     = 14;
  const COL_NAME      = 2;
  const COL_PLAN      = 8;
  const COL_MEMBER_ID = 11;

  const SOURCE_SS_ID = '1BIIPZKcEppdvrUoD2TGcGIOXFZVKMqYpsB-fmYJDOzs';

  let row;
  if (e && e.values) {
    // トリガー経由：送信された行のデータが e.values に入っている（一瞬で処理）
    row = e.values;
  } else {
    // 手動実行：A列に値がある最終行を取得（H列の数式による誤検知を防ぐ）
    const srcSS = SpreadsheetApp.openById(SOURCE_SS_ID);
    const srcSheet = srcSS.getSheets()[0];
    const colA = srcSheet.getRange('A:A').getValues();
    let lastRow = 1;
    for (let i = colA.length - 1; i >= 1; i--) {
      if (colA[i][0] !== '' && colA[i][0] !== null) {
        lastRow = i + 1;
        break;
      }
    }
    if (lastRow < 2) return { success: false, error: 'データがありません' };
    row = srcSheet.getRange(lastRow, 1, 1, 16).getValues()[0];
    console.log('手動実行：最終行番号=' + lastRow, 'L列=' + row[COL_MEMBER_ID]);
  }

  const memberId = String(row[COL_MEMBER_ID]).trim();
  const name     = String(row[COL_NAME]).trim();
  const email    = String(row[COL_EMAIL]).trim();
  const planRaw  = String(row[COL_PLAN]).trim();
  const planCode = PLAN_MAP[planRaw] || 'dropin';

  if (!memberId || memberId === '' || memberId === 'undefined') {
    console.log('会員番号なし、スキップ');
    return { success: false, error: '会員番号が空です' };
  }

  const destSheet = getOrCreateSheet(CONFIG.SHEET_NAME_MEMBERS);

  // 既存チェック
  const existingData = destSheet.getDataRange().getValues();
  for (let i = 1; i < existingData.length; i++) {
    if (String(existingData[i][0]).trim() === memberId) {
      console.log('既存会員のためスキップ: ' + memberId);
      return { success: false, error: '既存会員: ' + memberId };
    }
  }

  destSheet.appendRow([memberId, name, planCode, email, planRaw]);
  // H列に経過日数の数式を自動挿入
  const newRow = destSheet.getLastRow();
  destSheet.getRange(newRow, 8).setFormula(`=IF(G${newRow}="","",TODAY()-G${newRow})`);
  console.log('追加完了: ' + memberId + ' ' + name);
  return { success: true, memberId: memberId, name: name };
}

// ============================================================
// フォーム送信トリガーをプログラムで登録する関数
// ★一度だけGASエディタから手動実行してください
// ============================================================
function setupFormTrigger() {
  const SOURCE_SS_ID = '1BIIPZKcEppdvrUoD2TGcGIOXFZVKMqYpsB-fmYJDOzs';

  // 既存の同名トリガーを削除（重複防止）
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncMembersFromForm') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 会員登録スプレッドシートのフォーム送信トリガーを登録
  const srcSS = SpreadsheetApp.openById(SOURCE_SS_ID);
  ScriptApp.newTrigger('syncMembersFromForm')
    .forSpreadsheet(srcSS)
    .onFormSubmit()
    .create();

  console.log('トリガーを登録しました。フォーム送信時に自動同期されます。');
}

function debugStripeForMember() {
  const STRIPE_SECRET_KEY = PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY');
  const email = 'sugicoko315@icloud.com';
  
  // 顧客検索
  const res = UrlFetchApp.fetch(
    `https://api.stripe.com/v1/customers?email=${encodeURIComponent(email)}&limit=1`,
    { headers: { Authorization: 'Basic ' + Utilities.base64Encode(STRIPE_SECRET_KEY + ':') } }
  );
  const customers = JSON.parse(res.getContentText()).data;
  if (!customers.length) { console.log('顧客が見つかりません'); return; }
  
  const customerId = customers[0].id;
  console.log('顧客ID:', customerId);
  
  // サブスクリプション取得
  const subRes = UrlFetchApp.fetch(
    `https://api.stripe.com/v1/subscriptions?customer=${customerId}&status=active&limit=5`,
    { headers: { Authorization: 'Basic ' + Utilities.base64Encode(STRIPE_SECRET_KEY + ':') } }
  );
  const subs = JSON.parse(subRes.getContentText()).data;
  console.log('サブスク数:', subs.length);
  
  subs.forEach(sub => {
    sub.items.data.forEach(item => {
      const productId = item.price.product;
      const productRes = UrlFetchApp.fetch(
        `https://api.stripe.com/v1/products/${productId}`,
        { headers: { Authorization: 'Basic ' + Utilities.base64Encode(STRIPE_SECRET_KEY + ':') } }
      );
      const product = JSON.parse(productRes.getContentText());
      console.log('商品名:', product.name, '/ ステータス:', sub.status);
    });
  });
}