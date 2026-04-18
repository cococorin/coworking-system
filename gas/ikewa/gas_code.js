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
  'monthly_general':  { label: '月額会員（一般）',     monthly: 8000, isMonthly: true },
  'monthly_student':  { label: '月額会員（学生）',     monthly: 4000, isMonthly: true },
  'monthly_weekend':  { label: '月額プラン（土日祝）', monthly: 4000, isMonthly: true },
  'monthly_weekday':  { label: '月額プラン（平日）',   monthly: 6000, isMonthly: true },
  'dropin':           { label: 'ドロップイン',         monthly: null, isMonthly: false },
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
    const hours = Math.ceil(diffMin / 60);
    fee = Math.min(hours * CONFIG.DROPIN_HOURLY, CONFIG.DROPIN_DAILY_MAX);
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
// ヘルパー：曜日制限チェック
// ============================================================
function checkDayRestriction(type) {
  const now = new Date();
  const day = now.getDay(); // 0=日, 1=月..., 6=土
  const isWeekend = (day === 0 || day === 6);

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
    if (data[i][0]) settings[String(data[i][0]).trim()] = String(data[i][1] || '');
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
  const settings = getMailSettings().settings || {};
  const subject = settings['first_visit_subject'] || `【${CONFIG.SPACE_NAME}】はじめてのご利用ありがとうございました`;
  const feeText = fee > 0 ? `ご利用料金：¥${fee.toLocaleString()}` : '月額会員のため追加料金はありません';
  const unsubscribeUrl = `${ScriptApp.getService().getUrl()}?action=unsubscribe&memberId=${member.id}`;

  let body = settings['first_visit_body'] ||
`{name} 様

はじめて${CONFIG.SPACE_NAME}をご利用いただきありがとうございました。

■ ご利用内容
　入館時刻：{checkin}
　退館時刻：{checkout}
　利用時間：{duration}
　{fee}

またのご来館をお待ちしております。

${CONFIG.SPACE_NAME}

---
メールの受信停止はこちら：{unsubscribe_url}`;

  // \n を実際の改行に変換（スプレッドシートから読み込んだ場合の対処）
  body = body.replace(/\\n/g, '\n').replace(/\n/g, '\n');

  body = body
    .replace(/{name}/g, member.name)
    .replace(/{checkin}/g, checkinTime.substring(0, 5))
    .replace(/{checkout}/g, checkoutTime)
    .replace(/{duration}/g, duration)
    .replace(/{fee}/g, feeText)
    .replace(/{unsubscribe_url}/g, unsubscribeUrl);

  // HTMLメール用に変換
  const htmlBody = `
<div style="font-family:sans-serif;max-width:560px;margin:0 auto;color:#333;">
  <div style="background:#2e2826;padding:20px 24px;border-radius:8px 8px 0 0;">
    <p style="color:#fff;font-size:18px;font-weight:bold;margin:0;">cococorin</p>
    <p style="color:#aaa;font-size:12px;margin:4px 0 0;">半田市創造・連携・実践センター</p>
  </div>
  <div style="background:#fff;padding:24px;border:1px solid #e0e0e0;">
    <p style="font-size:15px;">${member.name} 様</p>
    <p style="font-size:14px;line-height:1.8;">はじめてcococorinをご利用いただきありがとうございました。<br>またのご来館をお待ちしております。</p>
    <div style="background:#f8f8f6;border-radius:8px;padding:16px 20px;margin:20px 0;">
      <p style="font-size:12px;color:#888;margin:0 0 10px;font-weight:bold;">ご利用内容</p>
      <table style="font-size:13px;width:100%;border-collapse:collapse;">
        <tr><td style="color:#888;padding:4px 0;width:90px;">入館時刻</td><td>${checkinTime.substring(0, 5)}</td></tr>
        <tr><td style="color:#888;padding:4px 0;">退館時刻</td><td>${checkoutTime}</td></tr>
        <tr><td style="color:#888;padding:4px 0;">利用時間</td><td>${duration}</td></tr>
        <tr><td style="color:#888;padding:4px 0;">料金</td><td style="font-weight:bold;">${feeText}</td></tr>
      </table>
    </div>
    <p style="font-size:14px;line-height:1.8;">ご不明な点はお気軽にお問い合わせください。</p>
    <p style="font-size:14px;"><a href="mailto:info@handanotane.com" style="color:#00a3af;">info@handanotane.com</a></p>
  </div>
  <div style="background:#f4f4f0;padding:12px 24px;border-radius:0 0 8px 8px;border:1px solid #e0e0e0;border-top:none;">
    <p style="font-size:11px;color:#aaa;margin:0;">メールの受信を停止する場合は<a href="${unsubscribeUrl}" style="color:#00a3af;">こちら</a>からお手続きください。</p>
  </div>
</div>`;

  try {
    GmailApp.sendEmail(member.email, subject, body, {
      name: CONFIG.SPACE_NAME,
      from: CONFIG.FROM_EMAIL,
      replyTo: CONFIG.ADMIN_EMAIL,
      htmlBody: htmlBody,
    });
    console.log('初回利用メール送信完了:', member.name);
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
    // 手動実行：ソースシートの最終行を取得
    const srcSS = SpreadsheetApp.openById(SOURCE_SS_ID);
    const srcSheet = srcSS.getSheets()[0];
    const lastRow = srcSheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'データがありません' };
    row = srcSheet.getRange(lastRow, 1, 1, 16).getValues()[0];
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