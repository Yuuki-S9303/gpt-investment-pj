/**** Gmail約定取り込み（SBI証券）— 当日手動強制版 ****************************
 * 役割:
 *  - 今日(Asia/Tokyo)の日付のメールだけを、時間帯/Dedupに関係なく取り込み
 *  - 既存の Fills 形式（小文字ヘッダー）に追記
 *  - 取込済みでも再度取り込む（重複許容）※SourceにFORCE_TODAYを付与
 ***************************************************************************/

// ==== ローカルヘルパー（不足対策） =========================================
const _TZ = (typeof CFG!=='undefined' && CFG.CLOCK_TZ) || 'Asia/Tokyo';

function fmtYMD_(d, tz){ return Utilities.formatDate(new Date(d), tz||_TZ, 'yyyy/MM/dd'); }
function addDays_(d, n){ const x = new Date(d); x.setDate(x.getDate()+Number(n||0)); return x; }
function nowJST(){ return Utilities.formatDate(new Date(), _TZ, 'yyyy-MM-dd HH:mm:ss'); }
// extractOrderNo_ が未定義でも動くようフォールバック
function extractOrderNo_(subj){
  try{
    if (typeof this.extractOrderNo_ === 'function') return this.extractOrderNo_(subj);
  }catch(e){}
  const m = String(subj||'').match(/注文番号[:：]\s*(\d+)/);
  return m ? m[1] : '';
}

// ==== 本体 ================================================================

/** 手動：今日のSBI約定メールを強制取り込み（Dedup/時間帯ガード無視） */
function jobMailImportSBI_todayManual(){
  const today = new Date();                       // 実行時刻をJST日付に丸めず使用
  const startStr = fmtYMD_(today, _TZ);           // 例: 2025/10/23
  const endStr   = fmtYMD_(addDays_(today, 1), _TZ); // 翌日（before は排他的）

  const q = `from:sbisec.co.jp subject:"国内株式の約定通知" after:${startStr} before:${endStr}`;
  console.log('[TODAY] manual import started');
  console.log('[TODAY] query=', q);

  const threads = GmailApp.search(q, 0, 200);
  console.log('[TODAY] threads=', threads.length);

  const rows = [];
  let appended = 0, skip = 0;

  threads.forEach(th=>{
    th.getMessages().forEach(msg=>{
      const subj  = msg.getSubject() || '';
      const orderNo = extractOrderNo_(subj);

      const html  = msg.getBody();
      const plain = msg.getPlainBody();

      // ★ 修正点：正しいパーサ関数名を使用
      const rec = parseSbiExecMail_Today_(html) || parseSbiExecMail_Today_(plain);
      if (!rec){ skip++; return; }

      // 念のためJST日付で当日か再確認
      if (rec.date !== fmtYMD_(today, _TZ)){ skip++; return; }

      rows.push({
        Date:       rec.date,
        Side:       rec.side,
        Code:       rec.code,
        Name:       rec.name || '',
        Price:      rec.price,
        Qty:        rec.qty,
        Account:    'SBI',
        OrderNo:    orderNo || '',
        ExecType:   '',
        Source:     'FORCE_TODAY SBI_MAIL',
        InsertedAt: nowJST()
      });

      console.log('[TODAY][OK]', { orderNo, code:rec.code, side:rec.side, price:rec.price, qty:rec.qty });
      appended++;
    });
  });

  if (rows.length){
    appendFillsRows_(rows);
    console.log('[TODAY][DONE]', { appended, skip });
    if (typeof notifyDiscordSafe_==='function'){
      notifyDiscordSafe_(`【Gmail取込(手動/当日)】SBI約定メールから ${rows.length} 件をFillsへ反映（skip=${skip}）`);
    }
  }else{
    console.log('[TODAY][DONE]', { appended:0, skip });
    if (typeof notifyDiscordSafe_==='function'){
      notifyDiscordSafe_('【Gmail取込(手動/当日)】該当なし（本文パース失敗or当日外）');
    }
  }
}

/** ===== パーサ（SBI約定メールHTML/プレーン両対応｜当日手動版）===== */
function parseSbiExecMail_Today_(raw){
  if (!raw) return null;

  const norm = String(raw)
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(h1|p|div|tr|td|th|li|br|table)>/gi, '\n')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<[^>]*>/g, '')
    .replace(/\r/g, '')
    .replace(/\u00A0/g, ' ')
    .replace(/[ \t]+/g, ' ')
    .replace(/[：:]\s*/g, '：')          // ラベルコロンを全角統一
    .replace(/(\d)：(\d)/g, '$1:$2')     // 時刻コロンは半角へ
    .replace(/\n{2,}/g, '\n')
    .trim();

  const flat = norm.replace(/\s+/g, ' ');
  const find = (re, src=flat)=>{ const m = src.match(re); return m ? m[1] : null; };
  const toNum = s => (s==null) ? null : Number(String(s).replace(/,/g,'').trim());

  // 取引種別 → BUY/SELL
  const sideMap = { '現物買':'BUY', '買付':'BUY', '買':'BUY', '現物売':'SELL', '売却':'SELL', '売':'SELL' };
  const sideKey = find(/取引種別\s*：\s*(現物買|現物売|買付|買|売却|売)/);
  const side = sideKey ? sideMap[sideKey] : null;

  // 約定日時 → yyyy/MM/dd（JST）
  const dtStr = find(/約定日時\s*：\s*(\d{4}\/\d{1,2}\/\d{1,2}\s+\d{1,2}:\d{2})/);
  const date = dtStr
    ? Utilities.formatDate(new Date(dtStr), _TZ, 'yyyy/MM/dd')
    : null;

  // 銘柄コード（4桁 + 任意英字2まで：例 285A）
  let code = find(/銘柄コード\s*：\s*([0-9]{3,4}[A-Z]{0,2})/);
  if (!code){
    const m = flat.match(/[（(]([0-9]{3,4}[A-Z]{0,2})[)）]/);
    if (m) code = m[1];
  }

  // 銘柄名：改行まで、なければ次ラベル手前まで
  let name = null;
  const nameLine = norm.match(/銘柄名\s*：\s*([^\n\r]+)/);
  if (nameLine) {
    name = nameLine[1].trim();
  } else {
    const nameSpan = flat.match(/銘柄名\s*：\s*([^：\n\r]+?)(?=\s*(?:取引種別|株数|市場|約定価格|約定単価|約定金額)\s*：|$)/);
    if (nameSpan) name = nameSpan[1].trim();
  }

  // 数量
  let qty = toNum(find(/株数\s*：\s*([0-9,]+)/));
  if (qty == null) qty = toNum(find(/約定(?:数量|株数)\s*：\s*([0-9,]+)/));

  // 単価（小数も許容） or 金額/数量で逆算
  let price = toNum(find(/約定(?:価格|単価)\s*：\s*([0-9,]+(?:\.[0-9]+)?)/));
  if (price == null && qty){
    const amount = toNum(find(/約定金額\s*：\s*([0-9,]+)\s*円?/));
    if (amount != null) price = +(amount / qty).toFixed(2);
  }

  if (!side || !date || !code || !qty || !price) return null;
  return { side, date, code, name: name||'', qty, price };
}
