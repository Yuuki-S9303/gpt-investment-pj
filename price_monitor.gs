/**
 * price_monitor.gs
 * 保有銘柄の現在値を10分毎に監視し、閾値到達時にDiscord通知
 *
 * 依存（他ファイルで定義済み）:
 *  - 10_bootstrap.gs : sh(), CFG, notifyDiscord()
 *  - qce_lite_dto.gs : COL_L_*, SHEET_LEDGER, postDiscordBorderline_()
 *
 * セットアップ（初回のみ）:
 *  1. setupPriceMonitorTrigger() を手動実行 → 10分毎トリガー登録
 *
 * テスト:
 *  - checkPriceAlerts() を手動実行 → ログで価格・変化率を確認
 *  - pm_resetAlertFlags() でフラグリセット（再通知テスト用）
 */

// ===== アラート閾値 =====
const PM_PLUS_PCT  = 0.05;  // +5%  → 半利確タイミング
const PM_MINUS_PCT = 0.03;  // -3%  → HardStop発動

// 東京市場の取引時間（分で管理）
const PM_SESSIONS = [
  { start: 9 * 60,       end: 11 * 60 + 30 },  // 前場: 9:00-11:30
  { start: 12 * 60 + 30, end: 15 * 60 + 30 },  // 後場: 12:30-15:30
];

// =========================================================
// メイン（時間トリガーから実行）
// =========================================================

function checkPriceAlerts() {
  if (!pm_isMarketOpen_()) {
    console.log('[PriceMonitor] 市場時間外 → スキップ');
    return;
  }

  const positions = pm_getOpenPositions_();
  if (positions.length === 0) {
    console.log('[PriceMonitor] 保有ポジションなし');
    return;
  }

  const props = PropertiesService.getScriptProperties();

  for (const pos of positions) {
    const price = pm_fetchCurrentPrice_(pos.code);
    if (price === null) {
      console.log(`[PriceMonitor] ${pos.code} 価格取得失敗`);
      continue;
    }

    const pct = (price - pos.entry) / pos.entry;
    console.log(
      `[PriceMonitor] ${pos.code} ${pos.name}` +
      ` ¥${price.toLocaleString()} entry=¥${pos.entry.toLocaleString()} ${pm_sign_(pct)}`
    );

    pm_checkPlus_(pos, price, pct, props);
    pm_checkMinus_(pos, price, pct, props);
  }
}

// =========================================================
// アラート判定
// =========================================================

function pm_checkPlus_(pos, price, pct, props) {
  const key = `PM_alerted_plus_${pos.code}`;

  if (pct >= PM_PLUS_PCT) {
    if (props.getProperty(key)) return; // 既に通知済み

    const msg = pos.halfDone
      ? `⚠️ **[価格監視] ${pos.code} ${pos.name}**\n` +
        `現在値: ¥${price.toLocaleString()} (${pm_sign_(pct)})\n` +
        `📌 残ポジション追加利確タイミング（+5%到達）`
      : `📈 **[価格監視] ${pos.code} ${pos.name}**\n` +
        `現在値: ¥${price.toLocaleString()} (${pm_sign_(pct)})\n` +
        `📌 半利確タイミング（+5%到達）`;

    postDiscordBorderline_(msg);
    props.setProperty(key, '1');

  } else if (pct < PM_PLUS_PCT * 0.9) {
    // 閾値から10%以上戻ったらリセット（次回到達時に再通知）
    props.deleteProperty(key);
  }
}

function pm_checkMinus_(pos, price, pct, props) {
  const key = `PM_alerted_minus_${pos.code}`;

  if (pct <= -PM_MINUS_PCT) {
    if (props.getProperty(key)) return;

    const msg =
      `🛑 **[価格監視] ${pos.code} ${pos.name}**\n` +
      `現在値: ¥${price.toLocaleString()} (${pm_sign_(pct)})\n` +
      `🚨 HardStop発動（-3%到達）`;

    postDiscordBorderline_(msg);
    props.setProperty(key, '1');

  } else if (pct > -PM_MINUS_PCT * 0.9) {
    props.deleteProperty(key);
  }
}

// =========================================================
// Ledger 読み取り（全保有ポジション）
// =========================================================

function pm_getOpenPositions_() {
  // sh() / COL_L_* / SHEET_LEDGER は既存ファイルで定義済み
  const sheet   = sh(SHEET_LEDGER);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const positions = [];

  for (const row of values) {
    const sellDate = row[COL_L_SELL_DATE - 1];
    const qty      = Number(row[COL_L_QTY   - 1] || 0);

    // 保有中の条件: 売却日が空 かつ 株数 > 0
    if (sellDate || qty <= 0) continue;

    const code  = String(row[COL_L_CODE  - 1] || '').trim();
    const name  = String(row[COL_L_NAME  - 1] || '').trim();
    const entry = Number(row[COL_L_PRICE - 1] || 0);
    const memo  = String(row[COL_L_MEMO  - 1] || '').trim();

    if (!code || entry <= 0) continue;

    positions.push({
      code,
      name,
      entry,
      halfDone: memo.includes('半利確'),
    });
  }

  return positions;
}

// =========================================================
// 価格取得（Yahoo Finance 非公式API）
// =========================================================

function pm_fetchCurrentPrice_(code) {
  // 東証銘柄は "{code}.T" 形式
  const url =
    `https://query1.finance.yahoo.com/v8/finance/chart/${code}.T` +
    `?interval=1d&range=1d`;
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) {
      console.log(`[PriceMonitor] HTTP ${res.getResponseCode()} for ${code}`);
      return null;
    }
    const json = JSON.parse(res.getContentText());
    return json?.chart?.result?.[0]?.meta?.regularMarketPrice ?? null;
  } catch (e) {
    console.error(`[PriceMonitor] 価格取得エラー ${code}: ${e}`);
    return null;
  }
}

// =========================================================
// ユーティリティ
// =========================================================

function pm_sign_(pct) {
  const s = (pct * 100).toFixed(2);
  return pct >= 0 ? `+${s}%` : `${s}%`;
}

function pm_isMarketOpen_() {
  const tz  = (typeof CFG !== 'undefined' && CFG.CLOCK_TZ) || 'Asia/Tokyo';
  const now = new Date();
  const dow = Number(Utilities.formatDate(now, tz, 'u')); // 1=Mon...7=Sun
  if (dow >= 6) return false; // 土日

  const h   = Number(Utilities.formatDate(now, tz, 'H'));
  const m   = Number(Utilities.formatDate(now, tz, 'm'));
  const cur = h * 60 + m;
  return PM_SESSIONS.some(s => cur >= s.start && cur <= s.end);
}

// =========================================================
// セットアップ（初回のみ手動実行）
// =========================================================

/**
 * 10分毎トリガーを登録する。
 * 重複防止のため既存トリガーを削除してから作成。
 */
function setupPriceMonitorTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'checkPriceAlerts')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('checkPriceAlerts')
    .timeBased()
    .everyMinutes(10)
    .inTimezone((typeof CFG !== 'undefined' && CFG.CLOCK_TZ) || 'Asia/Tokyo')
    .create();

  console.log('トリガー登録完了: checkPriceAlerts を10分毎に実行');
}

/**
 * 市場時間外でも強制実行（テスト用）
 * 動作確認が終わったら削除してOK
 */
function checkPriceAlerts_forceRun() {
  const positions = pm_getOpenPositions_();
  if (positions.length === 0) {
    console.log('[PriceMonitor] 保有ポジションなし');
    return;
  }

  const props = PropertiesService.getScriptProperties();

  for (const pos of positions) {
    const price = pm_fetchCurrentPrice_(pos.code);
    if (price === null) {
      console.log(`[PriceMonitor] ${pos.code} 価格取得失敗`);
      continue;
    }
    const pct = (price - pos.entry) / pos.entry;
    console.log(
      `[PriceMonitor] ${pos.code} ${pos.name}` +
      ` ¥${price.toLocaleString()} entry=¥${pos.entry.toLocaleString()} ${pm_sign_(pct)}`
    );
    pm_checkPlus_(pos, price, pct, props);
    pm_checkMinus_(pos, price, pct, props);
  }
}

/** アラート送信済みフラグをリセット（テスト・再通知用） */
function pm_resetAlertFlags() {
  const props = PropertiesService.getScriptProperties();
  let count = 0;
  for (const key of Object.keys(props.getProperties())) {
    if (key.startsWith('PM_alerted_')) {
      props.deleteProperty(key);
      count++;
    }
  }
  console.log(`${count} 件のアラートフラグをリセットしました`);
}
