/**** Gmail約定取り込み（SBI証券）— 本番ミニマル版 *************************
 * 役割:
 *  - 平日 9:00〜18:00 だけ Gmail を取り込み（ガード付き）※必要なら後段で追加
 *  - 件名の「注文番号」で重複排除（ScriptPropertiesに保存）
 *  - Fills へ {Date,Side,Code,Name,Price,Qty,Account,OrderNo,ExecType,Source,InsertedAt,ProcessedAt}
 *
 * トリガー:
 *  - 30分おきの CLOCK トリガーを `jobMailImportSBI` に1本だけ
 ***************************************************************************/

/** ====== 運用パラメータ（再宣言ガード付き） ====== */
if (typeof MAIL_QUERY === 'undefined') {
  var MAIL_QUERY = 'from:sbisec.co.jp subject:"国内株式の約定通知" newer_than:30d';
}
if (typeof DEDUP_PROP_KEY === 'undefined') {
  var DEDUP_PROP_KEY = 'MAIL_DEDUP_SET';
}
if (typeof DEDUP_EXPIRE_DAYS === 'undefined') {
  var DEDUP_EXPIRE_DAYS = 120;
}

/** ====== トリガー入口（ガード＆ロック） ====== */
function jobMailImportSBI(){
  const DEDUP_PROP = DEDUP_PROP_KEY;

  const threads = GmailApp.search(MAIL_QUERY, 0, 50);
  console.log('[INGEST] threads=', threads.length);
  if (!threads.length){ log_('INFO','gmail no threads'); return; }

  const dedup = loadDedup_(DEDUP_PROP);
  const rowsObj = [];
  let ok=0, skipDup=0, skipParse=0;

  threads.forEach(th=>{
    th.getMessages().forEach(msg=>{
      const subj = msg.getSubject() || '';
      const orderNo = extractOrderNo_(subj);
      const key = orderNo ? `SBI-${orderNo}` : null;
      if (key && dedup[key]){ skipDup++; return; }

      const html  = msg.getBody();
      const plain = msg.getPlainBody();
      const rec   = parseSbiExecMail_(html) || parseSbiExecMail_(plain);
      if (!rec){ skipParse++; return; }

      rowsObj.push({
        Date:       rec.date,
        Side:       rec.side,
        Code:       rec.code,
        Name:       rec.name || '',
        Price:      rec.price,
        Qty:        rec.qty,
        Account:    'SBI',
        OrderNo:    orderNo || '',
        ExecType:   '',
        Source:     'SBI_MAIL',
        InsertedAt: nowJST()
      });

      if (key) dedup[key] = new Date().toISOString();
      if (++ok <= 3) console.log('[OK]', { orderNo, code:rec.code, side:rec.side, price:rec.price, qty:rec.qty });
    });
  });

  if (rowsObj.length){
    appendFillsRows_(rowsObj);
    // ★ 有効期限でDedupキーを掃除して保存（重複呼び出しを1回に整理）
    pruneDedup_(dedup, DEDUP_EXPIRE_DAYS);
    saveDedup_(DEDUP_PROP, dedup);

    console.log('[INGEST DONE]', {appended: rowsObj.length, skipDup, skipParse});
    // ▼ ログチャンネルへ通知（postDiscordTo_ が無ければフォールバック）
    notifyDiscordSafe_(`【Gmail取込】SBI約定メールから ${rowsObj.length} 件をFillsへ反映`);

    // 取込があった時だけ下流を実行（ロック＆デバウンスは先方関数内）
    try {
      jobPipelineAfterIngest?.();
    } catch(e) {
      console.log('[INGEST→PIPE] err', e && e.message ? e.message : e);
    }
  }else{
    console.log('[INGEST DONE] appended=0', {skipDup, skipParse});
  }
}

/** ===== HTML優先でパース（失敗時テキスト） ===== */
function parseSbiExecMail_(raw){
  if (!raw) return null;
  const isHtml = /<\s*(html|body|table|tr|td|div|span|br)\b/i.test(raw);
  if (isHtml) {
    const r = parseSbiExecMailHtml_(raw);
    if (r) return r;
    return parseSbiExecMailText_(raw.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]*>/g,''));
  }
  return parseSbiExecMailText_(raw);
}

/** ===== パーサ（HTML：<th>…</th><td>…</td>） ===== */
function parseSbiExecMailHtml_(html){
  if (!html) return null;
  const pick = (label) => {
    const re = new RegExp(label + '\\s*<\\/th>[\\s\\S]*?<td[^>]*>([\\s\\S]*?)<\\/td>', 'i');
    const m = html.match(re);
    return m ? html2text_(m[1]) : '';
  };

  const dtStr = pick('約定日時');
  const code  = pick('銘柄コード');
  const name  = pick('銘柄名');
  const sideJ = pick('取引種別');
  const qty   = Number(String(pick('株数')).replace(/,/g,'')) || 0;
  const px    = Number(String(pick('約定価格')).replace(/,/g,'')) || 0;
  if (!dtStr || !code || !name || !sideJ || !qty || !px) return null;

  const side     = /売/.test(sideJ) ? 'SELL' : 'BUY';
  const dateOnly = dtStr.split(/\s+/)[0].replace(/-/g,'/');
  if (!side || !dateOnly || !code || !qty || !px) return null;
  return { side, date: dateOnly, code, name, qty, price: px, market: (pick('市場')||'---') };
}

/** ===== パーサ（テキストfallback） ===== */
function parseSbiExecMailText_(raw){
  if (!raw) return null;
  const t = String(raw).replace(/\r/g,'').replace(/\u3000/g,' ').trim();

  const m = {
    dt   : t.match(/約定日時\s*[:：]\s*([0-9/]{4,}\s+[0-9:]{4,})/),
    code : t.match(/銘柄コード\s*[:：]\s*([0-9A-Za-z]{3,})/),
    name : t.match(/銘柄名\s*[:：]\s*([^\n\r]+)/),
    side : t.match(/取引種別\s*[:：]\s*([^\n\r]+)/),
    qty  : t.match(/株数\s*[:：]\s*([0-9,]+)/),
    px   : t.match(/約定価格\s*[:：]\s*([0-9,]+(?:\.[0-9]+)?)/),
    mkt  : t.match(/市場\s*[:：]\s*([^\s\n\r]+)/),
  };
  if (!m.dt || !m.code || !m.name || !m.side || !m.qty || !m.px) return null;

  // 銘柄名の暴走カット
  const stopWords = ['取引種別','株数','市場','約定価格','注文内容','その他注文履歴','本メール','ご案内メール'];
  let name = m.name[1].trim();
  for (const sw of stopWords){
    const i = name.indexOf(sw);
    if (i > 0) { name = name.slice(0, i).trim(); break; }
  }

  const side     = /売/.test(m.side[1]) ? 'SELL' : 'BUY';
  const qty      = Number(m.qty[1].replace(/,/g,'')) || 0;
  const price    = Number(m.px[1].replace(/,/g,'')) || 0;
  const dateOnly = m.dt[1].split(/\s+/)[0].replace(/-/g,'/');
  if (!side || !dateOnly || !m.code[1] || !qty || !price) return null;

  return { side, date: dateOnly, code: m.code[1], name, qty, price, market: (m.mkt ? m.mkt[1] : '---') };
}

/** ===== HTML→テキスト整形 ===== */
function html2text_(s){
  return String(s||'')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<[^>]*>/g, '')
    .replace(/&nbsp;/g,' ')
    .replace(/&amp;/g,'&')
    .replace(/\s+/g,' ')
    .trim();
}

/** ===== ユーティリティ ===== */
function extractOrderNo_(subject){
  const m = subject && subject.match(/注文番号[:：]\s*([0-9]+)/);
  return m ? m[1] : null;
}
function loadDedup_(key){
  const js = PropertiesService.getScriptProperties().getProperty(key) || '{}';
  try { return JSON.parse(js); } catch(e){ return {}; }
}
function saveDedup_(key, map){
  PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(map));
}
function ensureHeadersRow_(sheet, headers){
  const lastCol  = Math.max(sheet.getLastColumn(), headers.length) || headers.length;
  const hasRow1  = sheet.getLastRow() >= 1;
  const firstRow = hasRow1 ? sheet.getRange(1,1,1,lastCol).getValues()[0] : [];
  const nonEmpty = firstRow.filter(v=>String(v||'').trim()!=='').length;

  if (!hasRow1 || nonEmpty===0){
    sheet.clearContents();
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    return;
  }
  if (firstRow.length < headers.length){
    sheet.getRange(1, firstRow.length+1, 1, headers.length-firstRow.length)
         .setValues([headers.slice(firstRow.length)]);
  }
}
// 期限で重複キーを掃除
function pruneDedup_(map, days){
  const cutoff = Date.now() - days*24*60*60*1000;
  for (const k in map){
    const t = Date.parse(map[k] || '');
    if (!t || t < cutoff) delete map[k];
  }
}

/** ---- Discord送信セーフラッパ（logsへルーティング） ----
 * postDiscordTo_('logs', ...) が存在すればそれを使用。
 * 無ければ postDiscord_ / discord_ → 最後は console にフォールバック。
 */
function notifyDiscordSafe_(text){
  if (!text) return;
  try{
    if (typeof postDiscordTo_ === 'function') { postDiscordTo_('logs', String(text)); return; }
    if (typeof postDiscord_   === 'function') { postDiscord_(String(text)); return; }
    if (typeof discord_       === 'function') { discord_(String(text)); return; }
  }catch(e){
    console.log('[DISCORD ERR]', e && e.message ? e.message : e);
  }
  console.log('[DISCORD]', text);
}

/** ===== Fills: 既存スキーマに合わせた見出し & 行書き込み ===== */
function ensureFillsHeaders_() {
  const sF = sh('Fills');
  const headers = ['Date','Side','Code','Name','Price','Qty','Account','OrderNo','ExecType','Source','InsertedAt','ProcessedAt'];
  const lastCol = Math.max(sF.getLastColumn(), headers.length) || headers.length;
  const firstRow = (sF.getLastRow()===0) ? [] : sF.getRange(1,1,1,lastCol).getValues()[0];
  const nonEmpty = firstRow.filter(v=>String(v||'').trim()!=='').length;

  if (sF.getLastRow()===0 || nonEmpty===0){
    sF.clearContents();
    sF.getRange(1,1,1,headers.length).setValues([headers]);
    return headers;
  }
  // 足りない見出しは右側に追加
  if (firstRow.length < headers.length){
    sF.getRange(1, firstRow.length+1, 1, headers.length-firstRow.length)
      .setValues([headers.slice(firstRow.length)]);
  }
  // 既存ヘッダーを返す
  const row = sF.getRange(1,1,1,Math.max(sF.getLastColumn(), headers.length)).getValues()[0];
  return row;
}
function getHeaderMapByName_(sheet){
  const row = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const map = {};
  row.forEach((v,i)=> map[String(v).trim()] = i);
  return map;
}
// rowsObjs: {Date,Side,Code,Name,Price,Qty,Account,OrderNo,ExecType,Source,InsertedAt}[]
function appendFillsRows_(rowsObjs){
  if (!rowsObjs || !rowsObjs.length) return;
  const sF = sh('Fills');
  ensureFillsHeaders_();
  const h = getHeaderMapByName_(sF);

  const headers = ['Date','Side','Code','Name','Price','Qty','Account','OrderNo','ExecType','Source','InsertedAt','ProcessedAt'];
  const out = rowsObjs.map(o=>{
    const row = new Array(headers.length).fill('');
    row[h['Date']]       = o.Date || '';
    row[h['Side']]       = o.Side || '';
    row[h['Code']]       = o.Code || '';
    row[h['Name']]       = o.Name || '';
    row[h['Price']]      = o.Price || '';
    row[h['Qty']]        = o.Qty || '';
    row[h['Account']]    = o.Account || '';
    row[h['OrderNo']]    = o.OrderNo || '';
    row[h['ExecType']]   = o.ExecType || '';
    row[h['Source']]     = o.Source || '';
    row[h['InsertedAt']] = o.InsertedAt || '';
    // ProcessedAt は空で入れておく
    return row;
  });

  const start = sF.getLastRow() + 1;
  sF.getRange(start, 1, out.length, Math.max(sF.getLastColumn(), headers.length)).setValues(out);
}

/** ===== 互換ログ（log_→sh の橋渡し） ===== */
function log_(level, msg){
  try{
    const m = String(level||'').toUpperCase();
    const kind = (m==='ERROR') ? 'エラー' : (m==='INFO' ? '情報' : 'お知らせ');
    sh(kind, String(msg||''));  // 環境の sh() に委譲（無ければ Logger に落ちる）
  }catch(e){
    Logger.log(`[${level}] ${msg}`);
  }
}

/** ===== nowJST（環境に合わせて適宜差し替え可） ===== */
function nowJST(){
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}
