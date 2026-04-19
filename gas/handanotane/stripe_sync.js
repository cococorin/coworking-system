// ============================================================
// Stripe サブスクリプション同期
// 「はんだのたね」アカウントのGASに追加してください
// ============================================================

// ★スクリプトプロパティに以下を設定してください（コードに直書きしない）
//   キー名: STRIPE_SECRET_KEY
//   値:     sk_live_xxxx...（StripeダッシュボードのAPIキー）
//
// 設定方法：GASエディタ → 「プロジェクトの設定」→「スクリプトプロパティ」

// 書き込み先スプレッドシートID（コワーキングスペース会員登録（回答））
const DEST_SS_ID_STRIPE = '1BIIPZKcEppdvrUoD2TGcGIOXFZVKMqYpsB-fmYJDOzs';

// Stripe商品名 → 会員マスタコード
const PLAN_MAP = {
  'ドロップイン（一時利用）': 'dropin',
  '月額会員（一般）':         'monthly_general',
  '月額会員（学生）':         'monthly_student',
  '月額プラン（土日祝）':     'monthly_weekend',
  '月額プラン（平日）':       'monthly_weekday',
  'レンタルオフィス入居者':   'rental_office',   // ← これを追加
};

// アクティブとみなすStripeサブスクリプションのステータス
const ACTIVE_STATUSES = ['active', 'trialing'];

// ============================================================
// メイン：Stripeサブスク情報を取得して会員マスタを更新
// ★毎日1回、時間トリガーで実行してください
// ============================================================
function syncStripeSubscriptions() {
  const stripeKey = PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY');
  if (!stripeKey) {
    throw new Error('STRIPE_SECRET_KEY がスクリプトプロパティに設定されていません');
  }

  // 回答シート取得
  const ss = SpreadsheetApp.openById(DEST_SS_ID_STRIPE);
  const sheet = ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();

  // 列インデックス（0始まり）
  const COL_EMAIL_FORMAL = 14; // O列：正式メールアドレス
  const COL_EMAIL_APPLY  = 1;  // B列：申込時メールアドレス（Oが空の場合の代替）
  const COL_PLAN         = 8;  // I列：プラン

  // メールアドレス → 行インデックスのマップ（高速検索用）
  const emailToRow = {};
  for (let i = 1; i < data.length; i++) {
    const emailFormal = String(data[i][COL_EMAIL_FORMAL]).trim().toLowerCase();
    const emailApply  = String(data[i][COL_EMAIL_APPLY]).trim().toLowerCase();
    const email = emailFormal || emailApply;
    if (email) emailToRow[email] = i;
  }

  // Stripeから全サブスクリプション取得
  const subscriptions = fetchAllStripeSubscriptions(stripeKey);
  console.log('Stripeサブスク取得件数: ' + subscriptions.length);

  let updated = 0;
  let notFound = 0;
  const processedEmails = new Set();

  for (const sub of subscriptions) {
    const email = String(sub.customerEmail || '').trim().toLowerCase();
    if (!email) continue;

    const rowIdx = emailToRow[email];
    if (rowIdx === undefined) { notFound++; continue; }

    const planName = sub.planName || '';
    const planCode = STRIPE_PLAN_MAP[planName] || null;
    const isActive = ACTIVE_STATUSES.includes(sub.status);

    let newPlanLabel;
    if (isActive && planCode) {
      // コード → 日本語ラベルに変換してI列に書き込む
      const labelMap = {
        'monthly_weekday':  '月額会員（平日）',
        'monthly_weekend':  '月額会員（土日祝）',
        'monthly_student':  '月額会員（学生）',
        'monthly_general':  '月額会員（一般）',
      };
      newPlanLabel = labelMap[planCode] || planName;
    } else {
      // 解約・支払いエラー → ドロップインに戻す
      newPlanLabel = 'ドロップイン（一時利用）';
    }

    const currentPlan = String(data[rowIdx][COL_PLAN]).trim();
    if (currentPlan !== newPlanLabel) {
      sheet.getRange(rowIdx + 1, COL_PLAN + 1).setValue(newPlanLabel);
      console.log(`更新: ${data[rowIdx][2]} (${email}) "${currentPlan}" → "${newPlanLabel}"`);
      updated++;
    }

    processedEmails.add(email);
  }

  // Stripeに存在しないメール（＝サブスク無し）はドロップインに戻す
  for (const [email, rowIdx] of Object.entries(emailToRow)) {
    if (processedEmails.has(email)) continue;
    const currentPlan = String(data[rowIdx][COL_PLAN]).trim();
    if (currentPlan !== 'ドロップイン（一時利用）' && currentPlan.startsWith('月額')) {
      sheet.getRange(rowIdx + 1, COL_PLAN + 1).setValue('ドロップイン（一時利用）');
      console.log(`解約扱いに変更: ${data[rowIdx][2]} (${email}) "${currentPlan}" → "ドロップイン（一時利用）"`);
      updated++;
    }
  }

  console.log(`回答シート同期完了: 更新 ${updated}件 / Stripeに未登録 ${notFound}件`);

  // 回答シートの最新プランを会員マスタにも反映
  syncMasterFromAnswerSheet();

  return { success: true, updated, notFound };
}

// ============================================================
// 回答シートのプランを会員マスタに反映
// ============================================================
function syncMasterFromAnswerSheet() {
  const PLAN_MAP = {
    'ドロップイン（一時利用）': 'dropin',
    '月額会員（一般）':         'monthly_general',
    '月額会員（学生）':         'monthly_student',
    '月額会員（土日祝）':       'monthly_weekend',
    '月額会員（平日）':         'monthly_weekday',
  };

  const COL_MEMBER_ID = 11; // L列：会員番号
  const COL_PLAN      = 8;  // I列：プラン

  // 回答シートを読み込む
  const srcSS = SpreadsheetApp.openById(DEST_SS_ID_STRIPE);
  const srcSheet = srcSS.getSheets()[0];
  const srcData = srcSheet.getDataRange().getValues();

  // 会員マスタを読み込む
  const destSS = SpreadsheetApp.openById('1knYE9NMyYkVAWQqqNb5DsUoUHLAF4RFVQ3k6MOzTgCU');
  const destSheet = destSS.getSheetByName('会員マスタ');
  if (!destSheet) { console.log('会員マスタシートが見つかりません'); return; }
  const destData = destSheet.getDataRange().getValues();

  // 会員マスタの会員番号 → 行インデックスマップ
  const idToRow = {};
  for (let i = 1; i < destData.length; i++) {
    const id = parseInt(String(destData[i][0]).trim(), 10);
    if (!isNaN(id)) idToRow[id] = i;
  }

  let updated = 0;
  for (let i = 1; i < srcData.length; i++) {
    const memberId = parseInt(String(srcData[i][COL_MEMBER_ID]).trim(), 10);
    const planLabel = String(srcData[i][COL_PLAN]).trim();
    const planCode  = PLAN_MAP[planLabel] || 'dropin';

    if (isNaN(memberId)) continue;
    const rowIdx = idToRow[memberId];
    if (rowIdx === undefined) continue;

    const currentCode = String(destData[rowIdx][2]).trim();
    if (currentCode !== planCode) {
      destSheet.getRange(rowIdx + 1, 3).setValue(planCode);
      console.log(`会員マスタ更新: ${destData[rowIdx][1]} (#${memberId}) ${currentCode} → ${planCode}`);
      updated++;
    }
  }

  console.log(`会員マスタ同期完了: 更新 ${updated}件`);
}

// ============================================================
// Stripe API：全サブスクリプションをページネーションで取得
// ============================================================
function fetchAllStripeSubscriptions(stripeKey) {
  const results = [];
  let startingAfter = null;

  // 商品IDと商品名のキャッシュ
  const productCache = {};

  while (true) {
    let url = 'https://api.stripe.com/v1/subscriptions?limit=100&expand[]=data.customer&expand[]=data.items';
    if (startingAfter) url += '&starting_after=' + startingAfter;

    const response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + stripeKey },
      muteHttpExceptions: true,
    });

    const json = JSON.parse(response.getContentText());
    if (json.error) throw new Error('Stripe APIエラー: ' + json.error.message);

    for (const sub of json.data) {
      const customer = sub.customer;
      const email = (typeof customer === 'object') ? customer.email : null;

      // 商品名取得：price.productのIDから商品情報を別途取得
      let planName = '';
      const priceItem = sub.items?.data?.[0];
      if (priceItem) {
        const productId = (typeof priceItem.price?.product === 'string')
          ? priceItem.price.product
          : priceItem.price?.product?.id || '';

        if (productId) {
          // キャッシュに無ければAPIで取得
          if (!productCache[productId]) {
            const prodRes = UrlFetchApp.fetch('https://api.stripe.com/v1/products/' + productId, {
              headers: { Authorization: 'Bearer ' + stripeKey },
              muteHttpExceptions: true,
            });
            const prodJson = JSON.parse(prodRes.getContentText());
            productCache[productId] = prodJson.name || '';
          }
          planName = productCache[productId];
        }

        // フォールバック：nicknameも試みる
        if (!planName) planName = priceItem.plan?.nickname || '';
      }

      console.log(`サブスク取得: email="${email}" status="${sub.status}" planName="${planName}"`);

      results.push({
        id:            sub.id,
        status:        sub.status,
        customerEmail: email,
        planName:      planName,
      });
    }

    if (!json.has_more) break;
    startingAfter = json.data[json.data.length - 1].id;
  }

  return results;
}

// ============================================================
// Stripe商品名を確認する（初回実行時に使用）
// GASエディタから手動実行 → ログで商品名一覧を確認
// ============================================================
function checkStripeProductNames() {
  const stripeKey = PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY');
  const response = UrlFetchApp.fetch('https://api.stripe.com/v1/products?limit=20&active=true', {
    headers: { Authorization: 'Bearer ' + stripeKey },
    muteHttpExceptions: true,
  });
  const json = JSON.parse(response.getContentText());
  if (json.error) throw new Error(json.error.message);
  json.data.forEach(p => console.log(`商品名: "${p.name}" / ID: ${p.id}`));
}

// ============================================================
// 毎日実行トリガーを登録（一度だけ手動実行）
// ============================================================
function setupStripeSyncTrigger() {
  // 既存トリガー削除
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncStripeSubscriptions') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 毎日午前3時に実行
  ScriptApp.newTrigger('syncStripeSubscriptions')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();

  console.log('Stripe同期トリガーを登録しました（毎日午前3時）');
}