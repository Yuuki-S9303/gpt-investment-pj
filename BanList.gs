/**
 * Risk Guard (v1.2) - 連敗0.5R + 総量上限（base×mult） + 出禁(BanList)
 *
 * ✅変更点（v1.2）
 * - cap_1r_yen を固定円（RISK_CAP_YEN）ではなく、
 *   「BASE_CAP_1R_YEN × risk_multiplier（実効cap）」に変更
 *   → 連敗で 0.5R のときは cap も 0.5倍（自動でリスク総量を絞る）
 *
 * 前提:
 *  - Ledgerシートに以下の見出しがある（日本語そのまま）
 *    「購入日」「銘柄コード」「銘柄名」「売却日」「1R(想定リスク)」「R」
 *
 * 追加シート:
 *  - BanList: code, name, last_R, ban_until, reason, updated_at
 *  - RiskState: risk_multiplier, open_total_1r_yen, ref_1r_yen, cap_1r_yen, remaining_1r_yen, updated_at
 */

const RG = {
  // ✅あなたのブックID（固定）
  SS_ID: "1EBlWYTlCCQqlWfyNWXsb2VcNxKH6YCxyLL4OXg4_NAA",

  LEDGER_SHEET: "Ledger",
  BAN_SHEET: "BanList",
  STATE_SHEET: "RiskState",

  // 出禁ルール
  BAN_THRESHOLD_R: -1.5,
  BAN_DAYS: 14,            // “2週間”を暦日で扱う（営業日厳密化は後で可）

  // 連敗ルール
  LOSS_STREAK_N: 2,        // 2連敗で
  LOSS_MULTIPLIER: 0.5,    // 次を0.5R

  // ✅総量上限（ベース円）→ 実効capは base × risk_multiplier
  // 元金100万ならまず 50,000（=5%）を基準に運用、など
  BASE_CAP_1R_YEN: 50000,

  // 基準1R（参考値として出すだけ。上限判定には使わない）
  REF_TRADES: 30,
};

function rgGetSS_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  return SpreadsheetApp.openById(RG.SS_ID);
}

function rgEnsureSheets_() {
  const ss = rgGetSS_();
  const ban = ss.getSheetByName(RG.BAN_SHEET) || ss.insertSheet(RG.BAN_SHEET);
  const st  = ss.getSheetByName(RG.STATE_SHEET) || ss.insertSheet(RG.STATE_SHEET);

  if (ban.getLastRow() === 0) {
    ban.appendRow(["code","name","last_R","ban_until","reason","updated_at"]);
  }
  if (st.getLastRow() === 0) {
    st.appendRow(["risk_multiplier","open_total_1r_yen","ref_1r_yen","cap_1r_yen","remaining_1r_yen","updated_at"]);
  }
}

function rgGetHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    if (!h) return;
    map[String(h).trim()] = i;
  });
  return map;
}

function rgParseNumber_(v) {
  if (v === null || v === "") return null;
  if (typeof v === "number") return v;
  const s = String(v).replace(/,/g, "").trim();
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

function rgMedian_(arr) {
  const a = arr.slice().sort((x, y) => x - y);
  const n = a.length;
  if (n === 0) return null;
  const mid = Math.floor(n / 2);
  return (n % 2) ? a[mid] : (a[mid - 1] + a[mid]) / 2;
}

function rgAddDays_(dateObj, days) {
  const d = new Date(dateObj);
  d.setDate(d.getDate() + days);
  d.setHours(0, 0, 0, 0);
  return d;
}

/**
 * ✅実効cap算出（base × multiplier）
 */
function rgCalcEffectiveCap1R_(baseCap, multiplier) {
  const eff = Math.round(Number(baseCap) * Number(multiplier));
  return Math.max(0, eff);
}

/**
 * 1) 出禁リスト更新:
 *    Ledgerの売却済み行のうち、R<=-1.5 の銘柄を BanList に登録/更新
 */
function rgUpdateBanListFromLedger() {
  rgEnsureSheets_();
  const ss = rgGetSS_();
  const ledger = ss.getSheetByName(RG.LEDGER_SHEET);
  const ban = ss.getSheetByName(RG.BAN_SHEET);

  const hm = rgGetHeaderMap_(ledger);
  const idxCode = hm["銘柄コード"];
  const idxName = hm["銘柄名"];
  const idxSell = hm["売却日"];
  const idxR    = hm["R"];

  if ([idxCode, idxName, idxSell, idxR].some(v => v === undefined)) {
    throw new Error("Ledgerヘッダー不足: 銘柄コード/銘柄名/売却日/R が必要");
  }

  const lastRow = ledger.getLastRow();
  if (lastRow < 2) return;

  const data = ledger.getRange(2, 1, lastRow - 1, ledger.getLastColumn()).getValues();

  // BanListをMap化（code->rowIndex）
  const banLast = ban.getLastRow();
  const banMap = new Map();
  if (banLast >= 2) {
    const banData = ban.getRange(2, 1, banLast - 1, 6).getValues();
    banData.forEach((r, i) => {
      const code = String(r[0] || "").trim();
      if (code) banMap.set(code, i + 2);
    });
  }

  const now = new Date();
  const updates = [];
  const appends = [];

  data.forEach(row => {
    const sellDate = row[idxSell];
    if (!sellDate) return; // 未売却は対象外

    const code = String(row[idxCode] || "").trim();
    if (!code) return;

    const name = String(row[idxName] || "").trim();
    const rVal = rgParseNumber_(row[idxR]);
    if (rVal === null) return;

    if (rVal <= RG.BAN_THRESHOLD_R) {
      const banUntil = rgAddDays_(now, RG.BAN_DAYS);
      const reason = `R<=${RG.BAN_THRESHOLD_R}`;

      if (banMap.has(code)) {
        updates.push({
          row: banMap.get(code),
          values: [code, name, rVal, banUntil, reason, now],
        });
      } else {
        appends.push([code, name, rVal, banUntil, reason, now]);
      }
    }
  });

  updates.forEach(u => ban.getRange(u.row, 1, 1, 6).setValues([u.values]));
  if (appends.length) {
    ban.getRange(ban.getLastRow() + 1, 1, appends.length, 6).setValues(appends);
  }
}

/**
 * 2) リスク状態更新:
 *   - risk_multiplier: 直近2回が負けなら0.5、それ以外1.0
 *   - ref_1r_yen: 直近売却済みの 1R(想定リスク) の中央値（参考値）
 *   - open_total_1r_yen: 保有中（売却日空欄）の1R合計
 *   - cap_1r_yen: ✅ base_cap × risk_multiplier（実効cap）
 */
function rgRefreshRiskState() {
  rgEnsureSheets_();
  const ss = rgGetSS_();
  const ledger = ss.getSheetByName(RG.LEDGER_SHEET);
  const state = ss.getSheetByName(RG.STATE_SHEET);

  const hm = rgGetHeaderMap_(ledger);
  const idxSell = hm["売却日"];
  const idx1R   = hm["1R(想定リスク)"];
  const idxR    = hm["R"];

  if ([idxSell, idx1R, idxR].some(v => v === undefined)) {
    throw new Error("Ledgerヘッダー不足: 売却日/1R(想定リスク)/R が必要");
  }

  const lastRow = ledger.getLastRow();
  if (lastRow < 2) return;

  const data = ledger.getRange(2, 1, lastRow - 1, ledger.getLastColumn()).getValues();

  // 直近の売却済みトレードを抽出（後ろから）
  const closed = [];
  for (let i = data.length - 1; i >= 0; i--) {
    const row = data[i];
    if (!row[idxSell]) continue;
    const rVal = rgParseNumber_(row[idxR]);
    const oneR = rgParseNumber_(row[idx1R]);
    if (rVal === null || oneR === null) continue;
    closed.push({ r: rVal, oneR: oneR });
    if (closed.length >= RG.REF_TRADES) break;
  }

  // 連敗判定（直近N回がマイナスなら0.5）
  let riskMultiplier = 1.0;
  if (closed.length >= RG.LOSS_STREAK_N) {
    const lastN = closed.slice(0, RG.LOSS_STREAK_N);
    const allLoss = lastN.every(t => t.r < 0);
    if (allLoss) riskMultiplier = RG.LOSS_MULTIPLIER;
  }

  // 基準1R（参考値）
  const ref1R = rgMedian_(closed.map(t => t.oneR)) || 0;

  // ✅上限は base×multiplier（実効cap）
  const cap1R = rgCalcEffectiveCap1R_(RG.BASE_CAP_1R_YEN, riskMultiplier);

  // 保有中の1R合計（売却日空欄）
  let openTotal = 0;
  data.forEach(row => {
    if (row[idxSell]) return;
    const oneR = rgParseNumber_(row[idx1R]);
    if (oneR !== null) openTotal += oneR;
  });

  const remaining = cap1R - openTotal;
  const now = new Date();

  // RiskStateは2行目に1行で上書き
  if (state.getLastRow() < 2) state.appendRow(["","","","","",""]);
  state.getRange(2, 1, 1, 6).setValues([[
    riskMultiplier,
    openTotal,
    ref1R,
    cap1R,
    remaining,
    now
  ]]);
}

/**
 * 3) その銘柄が出禁か判定（TRUEなら新規禁止）
 */
function rgIsBanned(code) {
  rgEnsureSheets_();
  const ss = rgGetSS_();
  const ban = ss.getSheetByName(RG.BAN_SHEET);
  const last = ban.getLastRow();
  if (last < 2) return false;

  const data = ban.getRange(2, 1, last - 1, 6).getValues();
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const target = String(code).trim();

  for (const r of data) {
    const c = String(r[0] || "").trim();
    if (c !== target) continue;
    const until = r[3];
    if (until && until >= today) return true;
    return false;
  }
  return false;
}

/**
 * 4) 発注前チェック（例）
 *  - 出禁ならNG
 *  - 総量上限を超えるならNG
 *  - OKなら multiplier（1 or 0.5）を返す
 *
 * newTrade1RYenOpt を渡せるなら厳密。渡せないなら0扱いで「総量判定は保有分のみ」にする。
 *
 * ※ここは v1.2 でも同じ。cap_1r_yen が実効capに変わるので、
 *   "openTotal + new1R > cap1R" がそのまま機能する。
 */
function rgPreflight(code, newTrade1RYenOpt) {
  rgRefreshRiskState(); // 最新化

  if (rgIsBanned(code)) {
    return { ok: false, reason: "BAN", multiplier: null };
  }

  const ss = rgGetSS_();
  const state = ss.getSheetByName(RG.STATE_SHEET);
  const row = state.getRange(2, 1, 1, 6).getValues()[0];

  const mult = Number(row[0]) || 1.0;
  const openTotal = Number(row[1]) || 0;
  const cap1R = Number(row[3]) || 0;

  const new1R = (newTrade1RYenOpt != null ? Number(newTrade1RYenOpt) : 0) * mult;

  if (openTotal + new1R > cap1R) {
    return { ok: false, reason: "RISK_CAP", multiplier: mult };
  }
  return { ok: true, reason: "OK", multiplier: mult };
}

/**
 * 手動実行用：全部更新
 */
function rgRunAll() {
  rgUpdateBanListFromLedger();
  rgRefreshRiskState();
}

/**
 * デバッグ：どのブックに書いているか可視化
 */
function rgDebugWhere() {
  const ss = rgGetSS_();
  Logger.log("SS Name: " + ss.getName());
  Logger.log("SS ID: " + ss.getId());
  Logger.log("SS URL: " + ss.getUrl());
  Logger.log("Sheets: " + ss.getSheets().map(s => s.getName()).join(", "));

  const stamp = new Date();
  const sh = ss.getSheetByName(RG.STATE_SHEET) || ss.insertSheet(RG.STATE_SHEET);
  sh.getRange("H1").setValue("rgDebugWhere stamp");
  sh.getRange("H2").setValue(stamp);
  sh.getRange("H3").setValue(ss.getUrl());
}

/**
 * メニュー追加（スプレッドシートから開いた時のみ表示）
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("RiskGuard")
    .addItem("RunAll（出禁+リスク更新）", "rgRunAll")
    .addItem("Update BanList", "rgUpdateBanListFromLedger")
    .addItem("Refresh RiskState", "rgRefreshRiskState")
    .addItem("Debug Where", "rgDebugWhere")
    .addToUi();
}
