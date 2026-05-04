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
const STRIPE_PLAN_MAP = {
  '月額会員（平日）':  'monthly_weekday',
  '月額会員（土日祝）': 'monthly_weekend',
  '月額会員（学生）':  'monthly_student',
  '月額会員（一般）':  'monthly_general',
};

// 会員マスタコード → 回答シートI列に書き込む日本語ラベル
const PLAN_CODE_TO_LABEL = {
  'monthly_weekday':  '月額会員（平日）',
  'monthly_weekend':  '月額会員（土日祝）',
  'monthly_student':  '月額会員（学生）',
  'monthly_general':  '月額会員（一般）',
};

const DROPIN_LABEL = 'ドロップイン（一時利用）';

// アクティブとみなすStripeサブスクリプションのステータス
const ACTIVE_STATUSES = ['active', 'trialing'];

// ============================================================
// メイン：Stripeサブスク情報を取得して会員マスタを更新
// ★毎日1回、時間トリガーで実行してください
// ============================================================
function syncStripeSubscriptions() {
  return _runStripeSync({ dryRun: false });
}

// ============================================================
// テスト用：書き込みは行わず、何が起こるかだけログに出す
// GASエディタから手動実行 → 実行ログで確認
// ============================================================
function dryRunSyncStripeSubscriptions() {
  return _runStripeSync({ dryRun: true });
}

function _runStripeSync(options) {
  const dryRun = !!(options && options.dryRun);
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

  // メールアドレスごとにサブスクをグルーピング（複数サブスク対応）
  const subsByEmail = {};
  for (const sub of subscriptions) {
    const email = String(sub.customerEmail || '').trim().toLowerCase();
    if (!email) continue;
    (subsByEmail[email] = subsByEmail[email] || []).push(sub);
  }

  let updated = 0;
  let notFound = 0;
  let preserved = 0;
  const processedEmails = new Set();

  for (const emailKey of Object.keys(subsByEmail)) {
    const subs = subsByEmail[emailKey];
    const rowIdx = emailToRow[emailKey];
    if (rowIdx === undefined) { notFound++; continue; }

    const decision = decidePlanLabel(subs);
    processedEmails.add(emailKey);

    const currentPlan = String(data[rowIdx][COL_PLAN]).trim();

    // active ありだが商品名が未知 → 危険な上書きを避け、既存値を維持
    if (decision.preserve) {
      console.log(`保留: ${data[rowIdx][2]} (${emailKey}) — 既存値「${currentPlan}」を維持 [${decision.reason}]`);
      preserved++;
      continue;
    }

    if (currentPlan !== decision.label) {
      if (!dryRun) sheet.getRange(rowIdx + 1, COL_PLAN + 1).setValue(decision.label);
      console.log(`${dryRun ? '[DRY-RUN] ' : ''}更新: ${data[rowIdx][2]} (${emailKey}) "${currentPlan}" → "${decision.label}" [${decision.reason}]`);
      updated++;
    }
  }

  // Stripeに登録されていないメール（＝サブスク無し）はドロップインに戻す
  for (const [emailKey, rowIdx] of Object.entries(emailToRow)) {
    if (processedEmails.has(emailKey)) continue;
    const currentPlan = String(data[rowIdx][COL_PLAN]).trim();
    if (currentPlan !== DROPIN_LABEL && currentPlan.startsWith('月額')) {
      if (!dryRun) sheet.getRange(rowIdx + 1, COL_PLAN + 1).setValue(DROPIN_LABEL);
      console.log(`${dryRun ? '[DRY-RUN] ' : ''}解約扱いに変更: ${data[rowIdx][2]} (${emailKey}) "${currentPlan}" → "${DROPIN_LABEL}"`);
      updated++;
    }
  }

  console.log(`${dryRun ? '[DRY-RUN] ' : ''}回答シート同期完了: 更新 ${updated}件 / Stripe未登録 ${notFound}件 / 保留 ${preserved}件`);

  // 本番のみ会員マスタへ反映
  if (!dryRun) syncMasterFromAnswerSheet();

  return { success: true, updated, notFound, preserved, dryRun };
}

// ============================================================
// プラン決定ロジック
//   優先順位:
//     1. active/trialing かつ商品名が STRIPE_PLAN_MAP に登録あり
//     2. active/trialing だが商品名不明 → 既存値を維持（preserve）
//     3. それ以外（incomplete / past_due / paused / 解約等のみ）→ ドロップイン
// ============================================================
function decidePlanLabel(subs) {
  const activeKnown = subs.find(function(s) {
    return ACTIVE_STATUSES.indexOf(s.status) !== -1
      && STRIPE_PLAN_MAP[String(s.planName || '').trim()];
  });
  if (activeKnown) {
    const planCode = STRIPE_PLAN_MAP[String(activeKnown.planName).trim()];
    return {
      label: PLAN_CODE_TO_LABEL[planCode] || activeKnown.planName,
      reason: 'active+known status=' + activeKnown.status + ' product="' + activeKnown.planName + '"',
      preserve: false,
    };
  }

  const activeUnknown = subs.find(function(s) {
    return ACTIVE_STATUSES.indexOf(s.status) !== -1;
  });
  if (activeUnknown) {
    return {
      label: null,
      reason: 'active but unknown product="' + activeUnknown.planName + '" status=' + activeUnknown.status,
      preserve: true,
    };
  }

  const statuses = subs.map(function(s) { return s.status; }).join(',');
  return {
    label: DROPIN_LABEL,
    reason: 'no active subscription (statuses=' + statuses + ')',
    preserve: false,
  };
}

// ============================================================
// 回答シートのプランを会員マスタに反映
// ============================================================
function syncMasterFromAnswerSheet() {
  // 「月額会員」と「月額プラン」両方の表記揺れを許容する
  const PLAN_MAP = {
    'ドロップイン（一時利用）': 'dropin',
    '月額会員（一般）':         'monthly_general',
    '月額会員（学生）':         'monthly_student',
    '月額会員（土日祝）':       'monthly_weekend',
    '月額会員（平日）':         'monthly_weekday',
    '月額プラン（土日祝）':     'monthly_weekend',
    '月額プラン（平日）':       'monthly_weekday',
    'レンタルオフィス入居者':   'rental_office',
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
  let skipped = 0;
  for (let i = 1; i < srcData.length; i++) {
    const memberId = parseInt(String(srcData[i][COL_MEMBER_ID]).trim(), 10);
    const planLabel = String(srcData[i][COL_PLAN]).trim();
    if (!planLabel) continue;

    const planCode = PLAN_MAP[planLabel];
    if (!planCode) {
      // I列が想定外の値の場合は会員マスタを上書きしない（既存値を維持）
      console.log('会員マスタ更新スキップ: I列="' + planLabel + '" は未知の値 (memberId=' + memberId + ')');
      skipped++;
      continue;
    }

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

  console.log(`会員マスタ同期完了: 更新 ${updated}件 / スキップ ${skipped}件`);
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
// 【検証用】decidePlanLabel のユニットテスト（Stripe API不要）
// 修正後ロジックが期待通りに動くかをコンソールだけで確認できる
// ============================================================
function testDecidePlanLabel() {
  const cases = [
    {
      name: 'active 月額（土日祝）のみ',
      input: [{ status: 'active', planName: '月額会員（土日祝）' }],
      expect: { label: '月額会員（土日祝）', preserve: false },
    },
    {
      name: 'active(土日祝) + incomplete(他) ← 元のバグの再現',
      input: [
        { status: 'incomplete', planName: '何か別商品' },
        { status: 'active',     planName: '月額会員（土日祝）' },
      ],
      expect: { label: '月額会員（土日祝）', preserve: false },
    },
    {
      name: 'active(土日祝) + past_due(土日祝) — 順序逆',
      input: [
        { status: 'active',   planName: '月額会員（土日祝）' },
        { status: 'past_due', planName: '月額会員（土日祝）' },
      ],
      expect: { label: '月額会員（土日祝）', preserve: false },
    },
    {
      name: 'active だが商品名が STRIPE_PLAN_MAP に未登録 → 既存値維持',
      input: [{ status: 'active', planName: '月額プラン（土日祝）' }],
      expect: { preserve: true },
    },
    {
      name: 'active なし（incomplete のみ） → ドロップイン',
      input: [{ status: 'incomplete', planName: '月額会員（土日祝）' }],
      expect: { label: DROPIN_LABEL, preserve: false },
    },
    {
      name: 'past_due のみ → ドロップイン',
      input: [{ status: 'past_due', planName: '月額会員（一般）' }],
      expect: { label: DROPIN_LABEL, preserve: false },
    },
    {
      name: 'trialing 月額（一般） → 月額（一般）',
      input: [{ status: 'trialing', planName: '月額会員（一般）' }],
      expect: { label: '月額会員（一般）', preserve: false },
    },
    {
      name: '商品名前後に空白 → trim許容',
      input: [{ status: 'active', planName: ' 月額会員（土日祝） ' }],
      expect: { label: '月額会員（土日祝）', preserve: false },
    },
  ];

  let pass = 0, fail = 0;
  for (const c of cases) {
    const r = decidePlanLabel(c.input);
    const labelOk = c.expect.label === undefined || r.label === c.expect.label;
    const preserveOk = c.expect.preserve === undefined || r.preserve === c.expect.preserve;
    const ok = labelOk && preserveOk;
    if (ok) {
      console.log('  ✓ ' + c.name);
      pass++;
    } else {
      console.log('  ✗ ' + c.name + ' / 期待=' + JSON.stringify(c.expect)
        + ' / 実際 label="' + r.label + '" preserve=' + r.preserve + ' reason=' + r.reason);
      fail++;
    }
  }
  console.log('--- testDecidePlanLabel: ' + pass + ' pass / ' + fail + ' fail ---');
  return { pass, fail };
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

// ============================================================
// 【検証用】特定メールアドレスのStripe同期判定をシミュレーション
// 実シートには書き込まず、ログだけ出力する
//   GASエディタで debugSyncForEmail() の引数を編集 → 実行 → 実行ログ確認
//
// 使い方の例：
//   function runDebug() { debugSyncForEmail('xxx@example.com'); }
// ============================================================
function debugSyncForEmail(email) {
  const stripeKey = PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY');
  if (!stripeKey) throw new Error('STRIPE_SECRET_KEY 未設定');
  const targetEmail = String(email || '').trim().toLowerCase();
  if (!targetEmail) { console.log('emailを引数で指定してください'); return; }

  // [1] Stripe顧客検索（同一メールに複数顧客レコードがある可能性も考慮）
  const customerRes = UrlFetchApp.fetch(
    'https://api.stripe.com/v1/customers?email=' + encodeURIComponent(targetEmail) + '&limit=10',
    { headers: { Authorization: 'Bearer ' + stripeKey }, muteHttpExceptions: true }
  );
  const customers = JSON.parse(customerRes.getContentText()).data || [];
  console.log('[1] Stripe顧客検索 email="' + targetEmail + '": ' + customers.length + '件');
  if (customers.length === 0) { console.log('  該当顧客なし'); }

  // [2] 各顧客の全サブスク（status=allで解約済も含めて確認）を取得し商品名を解決
  const allSubsForLogic = [];   // 同期ロジックに渡す（解約は除外）
  for (const c of customers) {
    const subRes = UrlFetchApp.fetch(
      'https://api.stripe.com/v1/subscriptions?customer=' + c.id + '&status=all&limit=20',
      { headers: { Authorization: 'Bearer ' + stripeKey }, muteHttpExceptions: true }
    );
    const subs = JSON.parse(subRes.getContentText()).data || [];
    console.log('[2] 顧客 ' + c.id + ' (' + c.email + ') のサブスク: ' + subs.length + '件');
    for (const s of subs) {
      let productName = '';
      const item = s.items && s.items.data && s.items.data[0];
      const productId = item && item.price && item.price.product;
      if (productId) {
        const prodRes = UrlFetchApp.fetch('https://api.stripe.com/v1/products/' + productId, {
          headers: { Authorization: 'Bearer ' + stripeKey }, muteHttpExceptions: true,
        });
        productName = JSON.parse(prodRes.getContentText()).name || '';
      }
      const mapped = STRIPE_PLAN_MAP[String(productName).trim()] || '(未登録)';
      console.log('    - subId=' + s.id + ' status=' + s.status + ' product="' + productName + '" → ' + mapped);
      // 本番sync は status=all を使わない（canceled/incomplete_expired は返らない）
      if (s.status !== 'canceled' && s.status !== 'incomplete_expired') {
        allSubsForLogic.push({
          id: s.id, status: s.status, customerEmail: c.email, planName: productName,
        });
      }
    }
  }

  // [3] 修正後の同期ロジックの判定をシミュレーション
  const decision = decidePlanLabel(allSubsForLogic);
  console.log('[3] decidePlanLabel: label="' + decision.label + '" preserve=' + decision.preserve + ' reason=' + decision.reason);

  // [4] 回答シートの現在値を確認
  const ss = SpreadsheetApp.openById(DEST_SS_ID_STRIPE);
  const sheet = ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  let hits = 0;
  for (let i = 1; i < data.length; i++) {
    const ef = String(data[i][14]).trim().toLowerCase();
    const ea = String(data[i][1]).trim().toLowerCase();
    if (ef === targetEmail || ea === targetEmail) {
      hits++;
      console.log('[4] 回答シート行 ' + (i + 1)
        + ': 名前="' + data[i][2] + '" 現在のI列="' + String(data[i][8]).trim()
        + '" 会員番号=' + String(data[i][11]).trim());
    }
  }
  if (hits === 0) console.log('[4] 回答シートにこのメールの行は見つかりませんでした');

  // [5] 会員マスタの現在値も確認
  const masterSS = SpreadsheetApp.openById('1knYE9NMyYkVAWQqqNb5DsUoUHLAF4RFVQ3k6MOzTgCU');
  const masterSheet = masterSS.getSheetByName('会員マスタ');
  if (masterSheet) {
    const mData = masterSheet.getDataRange().getValues();
    for (let i = 1; i < mData.length; i++) {
      const memberEmail = String(mData[i][3]).trim().toLowerCase();
      if (memberEmail === targetEmail) {
        console.log('[5] 会員マスタ行 ' + (i + 1)
          + ': 会員番号=' + mData[i][0] + ' 名前="' + mData[i][1] + '" 現在の種別="' + mData[i][2] + '"');
      }
    }
  }
}

// ============================================================
// 【検証用】Stripe商品名 と STRIPE_PLAN_MAP の対応を一覧表示
// 商品名のスペース差異・全角差異等の検出に使う
// ============================================================
function debugProductNamesMatch() {
  const stripeKey = PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY');
  if (!stripeKey) throw new Error('STRIPE_SECRET_KEY 未設定');
  const res = UrlFetchApp.fetch('https://api.stripe.com/v1/products?limit=100&active=true', {
    headers: { Authorization: 'Bearer ' + stripeKey }, muteHttpExceptions: true,
  });
  const products = JSON.parse(res.getContentText()).data || [];
  console.log('Stripeアクティブ商品: ' + products.length + '件');
  for (const p of products) {
    const name = String(p.name || '');
    const trimmed = name.trim();
    const code = STRIPE_PLAN_MAP[trimmed];
    const trimDiff = name !== trimmed ? ' [前後の空白あり！]' : '';
    const matchInfo = code ? 'OK → ' + code : '❌ STRIPE_PLAN_MAP に未登録';
    console.log('  "' + name + '"' + trimDiff + ' / id=' + p.id + ' : ' + matchInfo);
  }
  console.log('--- STRIPE_PLAN_MAP 登録キー ---');
  for (const key of Object.keys(STRIPE_PLAN_MAP)) {
    console.log('  "' + key + '" → ' + STRIPE_PLAN_MAP[key]);
  }
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
