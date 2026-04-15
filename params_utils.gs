/** ==========================================================
 * Paramsユーティリティ（クリーン版）
 * - Paramsシートを {Key:Value} マップとして取得
 * - 型変換ヘルパー付き（数値／真偽値）
 * 依存: sh() or SpreadsheetApp.getActiveSpreadsheet()
 * ========================================================== */

/** Paramsシートを {Key:Value} マップで取得（空でも落とさない） */
function getParamMap(){
  let ss;
  try {
    ss = (typeof SS === 'function') ? SS() : SpreadsheetApp.getActiveSpreadsheet();
  } catch(e){
    throw new Error('スプレッドシートを開けません（未バインド or 権限不足）');
  }

  const shParams = ss.getSheetByName('Params');
  if (!shParams) return {}; // 無ければ空

  const last = shParams.getLastRow();
  if (last < 2) return {}; // ヘッダーのみ

  const vals = shParams.getRange(2, 1, last - 1, 2).getValues(); // [Key, Value]
  const map = {};
  vals.forEach(([k, v])=>{
    const key = String(k || '').trim();
    if (key) map[key] = v;
  });
  return map;
}

/** Paramを数値で取得（未設定なら def を返す） */
function getParamNumber(map, key, def = 0){
  const v = map[key];
  return (v === undefined || v === '') ? def : Number(v);
}

/** Paramを真偽値で取得（未設定なら def を返す） */
function getParamBool(map, key, def = false){
  const v = map[key];
  if (v === true || String(v).toUpperCase() === 'TRUE') return true;
  if (v === false || String(v).toUpperCase() === 'FALSE') return false;
  return def;
}
