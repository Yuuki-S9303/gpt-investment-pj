/** メインスプレッドシート（Fills / Ledger があるやつ）のID */
const MAIN_SS_ID = '1EBlWYTlCCQqlWfyNWXsb2VcNxKH6YCxyLL4OXg4_NAA';

/** ====== Fills（新規約定検知用） ====== */

// Fills シート名
const SHEET_FILLS = 'Fills';

// 列番号（1始まり）
// A:Date / B:Side / C:Code / D:Name / E:Price / F:Qty
// G:Account / H:OrderNo / I:ExecType / J:Source / K:InsertedAt / L:ProcessedAt
const COL_F_DATE        = 1;
const COL_F_SIDE        = 2;
const COL_F_CODE        = 3;
const COL_F_NAME        = 4;
const COL_F_PRICE       = 5;
const COL_F_QTY         = 6;
const COL_F_ACCOUNT     = 7;
const COL_F_ORDER_NO    = 8;
const COL_F_EXEC_TYPE   = 9;
const COL_F_SOURCE      = 10;
const COL_F_INSERTED_AT = 11;
const COL_F_PROCESSED   = 12;

// 「Fills をどこまで見たか」（新規BUY検知用）
const PROP_BUY_NOTIFY_LAST_FILLS_ROW = 'BUY_NOTIFY_LAST_FILLS_ROW';


/** ====== Ledger（ボーダー計算用） ====== */

// Ledger シート名（必要ならここは実際の名前に変更OK）
const SHEET_LEDGER = 'Ledger';

// Ledger のヘッダー：
// 購入日 / 銘柄コード / 銘柄名 / 購入単価 / 株数 / 売却日 / 売却単価 / 実現損益 / 購入金額 / 1R(想定リスク) / R / 保有日数 / メモ
//  1        2          3        4        5       6        7         8          9          10             11  12       13
const COL_L_DATE       = 1;  // 購入日
const COL_L_CODE       = 2;  // 銘柄コード
const COL_L_NAME       = 3;  // 銘柄名
const COL_L_PRICE      = 4;  // 購入単価（＝平均取得単価として利用）
const COL_L_QTY        = 5;  // 株数（現在の保有株数）
const COL_L_SELL_DATE  = 6;  // 売却日（未決済ポジションはここが空）
const COL_L_SELL_PRICE = 7;  // 売却単価
const COL_L_MEMO       = 13; // メモ（必要なら後で使えるように定義）

/** Discord Webhook: ボーダーライン用（ボーダー通知チャンネル） */
const PROP_DISCORD_BORDERLINE = 'https://discord.com/api/webhooks/1439176930594127984/Qh31MBud679lE1WZLF7o3p6c_tLP7TG5ozBZeToIGe6dhTMFoyzTVM_AHP8jDWE702a4';

/** Discord Webhook: リマインド用（S株締切リマインドチャンネル） */
const PROP_DISCORD_REMIND = 'https://discord.com/api/webhooks/1439176281597153434/eSjZ2DEFHGK88uB32k6-IywPwVH6VPkY-QHGTIc5VZT8KANLkZTIhtrTzcndxUFa6ebF';

/** ボーダーライン用チャンネルに送る */
function postDiscordBorderline_(content){
   const url = PROP_DISCORD_BORDERLINE;  // ←ここだけ変える
    if (!url) {
    console.warn('postDiscordBorderline_: webhook URL not set');
    return;
  }
  const payload = { content };
  const params = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };
  UrlFetchApp.fetch(url, params);
}

/** リマインド用チャンネルに送る */
function postDiscordRemind_(content){
  const url = PROP_DISCORD_REMIND;      // ←ここだけ変える
    if (!url) {
    console.warn('postDiscordRemind_: webhook URL not set');
    return;
  }
  const payload = { content };
  const params = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };
  UrlFetchApp.fetch(url, params);
}

/** S株オークション受付締切（日本時間） */
const AUCTION_DEADLINES = [
  { key: 'AM_OPEN', label: '前場寄り（09:00 約定）', cutoff: '07:00' },
  { key: 'PM_OPEN', label: '後場寄り（12:30 約定）', cutoff: '10:30' },
  { key: 'EOD',     label: '大引け（15:00 約定）',   cutoff: '14:00' },
];

/** 同じ締切で1日に1回しかリマインドしないようにする用のプレフィックス */
const PROP_AUCTION_REMIND_PREFIX = 'AUCTION_REMIND_';


/** ====== Discord レポーティング用ヘルパー ====== */

// ScriptProperties に DISCORD_WEBHOOK_REPORTING を入れておく前提
const PROP_DISCORD_REPORTING = 'DISCORD_WEBHOOK_REPORTING';

function postDiscordReporting_(content){
  const props = PropertiesService.getScriptProperties();
  const url = props.getProperty(PROP_DISCORD_REPORTING);
  if (!url) {
    console.warn('postDiscordReporting_: webhook URL not set');
    return;
  }
  const payload = { content: content };
  const params = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };
  UrlFetchApp.fetch(url, params);
}

function formatPrice_(value){
  if (value === null || value === '' || !isFinite(Number(value))) return '-';

  const n = Math.round(Number(value));
  // 3桁区切りにする
  const withComma = n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return '¥' + withComma;
}

/**
 * Ledger から対象コードの「現在ポジション」を1件探す。
 *
 * 条件:
 *  - 銘柄コード一致
 *  - 売却日 が空（未決済ポジション）
 *  - 購入単価 > 0 かつ 株数 > 0
 * 探し方:
 *  - 下から上に向かって検索（最新の行を優先）
 *
 * 戻り値:
 *  - 見つかった場合:
 *    {
 *      date:     購入日,
 *      code:     銘柄コード（文字列）,
 *      name:     銘柄名,
 *      avgPrice: 購入単価（＝平均取得単価として使用）,
 *      qty:      株数,
 *      memo:     メモ（必要に応じて）
 *    }
 *  - 見つからなかった場合: null
 */
function findLedgerPositionByCode_(code){
  // ★ トリガー実行時でも確実に開けるように MAIN_SS_ID 経由で開く
  const ss = SpreadsheetApp.openById(MAIN_SS_ID);
  const sh = ss.getSheetByName(SHEET_LEDGER);
  if (!sh) {
    console.warn('findLedgerPositionByCode_: sheet not found:', SHEET_LEDGER);
    return null;
  }

  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return null; // ヘッダーしかない

  const lastCol = sh.getLastColumn();
  // 2行目以降すべてを取得
  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const targetCode = String(code);

  // ▼下から上に向かって検索（最新行を優先）
  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];

    const rowCode   = String(row[COL_L_CODE - 1]);       // 銘柄コード
    if (rowCode !== targetCode) continue;

    const sellDate  = row[COL_L_SELL_DATE - 1];          // 売却日
    const price     = Number(row[COL_L_PRICE     - 1]);  // 購入単価（平均取得単価）
    const qty       = Number(row[COL_L_QTY       - 1]);  // 株数

    // 売却済み（売却日あり）はスキップ
    if (sellDate) continue;

    // 単価 or 株数が0/空ならスキップ
    if (!price || !qty) continue;

    const date      = row[COL_L_DATE - 1];               // 購入日
    const name      = row[COL_L_NAME - 1];               // 銘柄名
    const memo      = row[COL_L_MEMO - 1];               // メモ

    // 見つかったらこの行を現在ポジションとして返す
    return {
      date,
      code: rowCode,
      name,
      avgPrice: price,
      qty,
      memo,
    };
  }

  // 条件に合う行が見つからなかった
  return null;
}


/**
 * ボーダー通知用メッセージ生成
 *
 * info = {
 *   date:     購入日（Date or 文字列でもOK）,
 *   code:     銘柄コード（文字列）,
 *   name:     銘柄名,
 *   avgPrice: 平均取得単価,
 *   qty:      株数,
 *   note:     補足メモ（任意）
 * }
 */
function buildBuyBorderMessage_(info){
  const tz = Session.getScriptTimeZone();

  const dateStr = info.date
    ? Utilities.formatDate(new Date(info.date), tz, 'yyyy/MM/dd')
    : '';

  const avgStr = formatPrice_(info.avgPrice);
  const qtyStr = Utilities.formatString('%s株', info.qty);

  const hardStop = formatPrice_(info.avgPrice * 0.97); // -3%
  const tpHalf   = formatPrice_(info.avgPrice * 1.05); // +5%

  return [
    '🆕【新規エントリー / 買い増し】',
    dateStr ? `約定日: ${dateStr}` : '',
    `銘柄: ${info.code} ${info.name}`,
    `平均取得単価: ${avgStr} × ${qtyStr}`,
    '',
    '--- ボーダーライン（QCE-Lite） ---',
    `🔺 +5% 半利確 : ${tpHalf}`,
    `🔻 HardStop -3% : ${hardStop}`,
    info.note ? `※ ${info.note}` : '',
  ].filter(Boolean).join('\n');
}

/**
 * 新しく追加された Fills の BUY 約定を検知し、
 * Ledger の平均取得単価・株数を使ってボーダーを Discord 通知する。
 *
 * - 検知トリガー：Fills（1約定 = 1イベント）
 * - ボーダー計算：Ledger（平均取得単価ベース）
 *
 * 依存:
 *  - 定数: SHEET_FILLS, COL_F_*, PROP_BUY_NOTIFY_LAST_FILLS_ROW
 *  - 関数: findLedgerPositionByCode_, buildBuyBorderMessage_, postDiscordReporting_
 */
function notifyNewBuyBorders(){
  const ss = SpreadsheetApp.openById(MAIN_SS_ID);
  const shFills = ss.getSheetByName(SHEET_FILLS);
  if (!shFills) {
    console.warn('notifyNewBuyBorders_: sheet not found:', SHEET_FILLS);
    return 0;
  }

  const lastRow = shFills.getLastRow();
  if (lastRow <= 1) return 0; // ヘッダのみ

  const props = PropertiesService.getScriptProperties();
  let lastSeenStr = props.getProperty(PROP_BUY_NOTIFY_LAST_FILLS_ROW);

  // 🔰 初回：過去分は通知せず、ポインタだけ今の最終行に合わせて終了
  if (!lastSeenStr) {
    props.setProperty(PROP_BUY_NOTIFY_LAST_FILLS_ROW, String(lastRow));
    console.log('[notifyNewBuyBorders_] init pointer at Fills row ' + lastRow);
    return 0;
  }

  let lastSeen = Number(lastSeenStr);
  if (lastSeen < 1) lastSeen = 1;

  const startRow = lastSeen + 1;
  if (startRow > lastRow) {
    // 新規行なし
    return 0;
  }

  const numRows = lastRow - startRow + 1;
  const lastCol = shFills.getLastColumn();
  const values  = shFills.getRange(startRow, 1, numRows, lastCol).getValues();

  let sent = 0;
  let latestRowUsed = lastSeen;

  values.forEach((row, idx) => {
    const rowNo = startRow + idx;

    const side = String(row[COL_F_SIDE - 1] || '').toUpperCase();
    if (side !== 'BUY') {
      latestRowUsed = rowNo;
      return;
    }

    const fillDate = row[COL_F_DATE  - 1];                 // Date
    const code     = row[COL_F_CODE  - 1];                 // 銘柄コード
    const nameFill = row[COL_F_NAME  - 1];                 // Fills側銘柄名
    const price    = Number(row[COL_F_PRICE - 1]);         // 約定単価
    const qty      = Number(row[COL_F_QTY   - 1]);         // 約定株数

    if (!code || !nameFill || !price || !qty) {
      latestRowUsed = rowNo;
      return;
    }

    // Ledger から「現在ポジション」（平均取得単価＆保有株数）を取得
    const pos = findLedgerPositionByCode_(code);

    let info;
    if (pos) {
      info = {
        date: pos.date || fillDate,
        code: pos.code,
        name: pos.name || nameFill,
        avgPrice: pos.avgPrice,
        qty: pos.qty,
        note: 'Ledger（平均取得単価）ベース',
      };
    } else {
      // 万一 Ledger に見つからなかった場合は Fills 約定単価で代用
      info = {
        date: fillDate,
        code: String(code),
        name: String(nameFill),
        avgPrice: price,
        qty: qty,
        note: 'Ledger未検出のため Fills 約定単価ベース',
      };
      console.warn('notifyNewBuyBorders_: Ledger position not found for code=' + code + ', fallback to Fills.');
    }

    const msg = buildBuyBorderMessage_(info);
    postDiscordBorderline_(msg);

    sent++;
    latestRowUsed = rowNo;
  });

  // ここまで見たよ、というポインタを更新
  props.setProperty(PROP_BUY_NOTIFY_LAST_FILLS_ROW, String(latestRowUsed));
  console.log('[notifyNewBuyBorders_] sent=' + sent + ', lastFillsRow=' + latestRowUsed);

  return sent;
}

function debugOpenMainSheet(){
  const ss = SpreadsheetApp.openById(MAIN_SS_ID);
  console.log('MAIN_SS_ID = ' + MAIN_SS_ID);
  console.log('Sheet name = ' + ss.getName());
}

/**
 * S株オークションの受付締切「10分前」に Discord へリマインドを飛ばす。
 *
 * - 5分おきの時間トリガーで実行する想定。
 * - 07:00, 10:30, 14:00 の「10分以内 & まだ締切前」のタイミングで、
 *   その日その締切については1回だけ通知する。
 *
 * 依存:
 *  - AUCTION_DEADLINES
 *  - PROP_AUCTION_REMIND_PREFIX
 *  - postDiscordRemind_
 */
function jobAuctionDeadlineReminder(){
  const tz  = (typeof CFG !== 'undefined' && CFG.CLOCK_TZ) || 'Asia/Tokyo';
  const now = new Date();

  // ★ 土日はスキップ
  const dow = Number(Utilities.formatDate(now, tz, 'u')); // 1=Mon ... 7=Sun
  if (dow === 6 || dow === 7) {
    console.log('[jobAuctionDeadlineReminder] weekend, skip');
    return;
  }

  const today      = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const h          = Number(Utilities.formatDate(now, tz, 'H'));   // 0-23
  const m          = Number(Utilities.formatDate(now, tz, 'm'));   // 0-59
  const minutesNow = h * 60 + m;


  const props = PropertiesService.getScriptProperties();

  AUCTION_DEADLINES.forEach(d => {
    const [ch, cm] = d.cutoff.split(':').map(Number);
    const cutoffMinutes = ch * 60 + cm;
    const diff = cutoffMinutes - minutesNow; // 単位: 分

    // diff > 0 かつ 10分以内 → 「締切まで10分以内＆まだ締切前」
    if (diff > 0 && diff <= 10) {
      const key = `${PROP_AUCTION_REMIND_PREFIX}${today}_${d.key}`;
      if (props.getProperty(key)) {
        // その日のその締切はすでにリマインド済み
        console.log(`[jobAuctionDeadlineReminder] already sent for ${d.key} on ${today}`);
        return;
      }

       const msg = [
        '⏰【S株オークション締切リマインド】',
        `対象: ${d.label}`,
        `SBI受付締切: ${d.cutoff}`,
        '',
        'あと10分で締切です。',
        '注文の出し忘れがないか、ざっとチェックしてください ✅',
      ].join('\n');

      // リマインド専用チャンネルへ送信
      postDiscordRemind_(msg);

      // 同じ日・同じ締切 key で2度通知しないようにマーク
      props.setProperty(key, 'sent');
      console.log(`[jobAuctionDeadlineReminder] sent reminder for ${d.key} (${d.cutoff})`);
    }
  });
}
