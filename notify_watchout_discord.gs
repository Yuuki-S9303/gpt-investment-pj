/** ==========================================================
 * Discord通知：Watch_OUT — クリーン本番版
 * 依存: sh() / notifyDiscord()（bootstrap）
 * 機能: Watch_OUT の上位N件をDiscordへ送信（2000字対策で分割）
 * ========================================================== */

function notifyWatchOutToDiscord(topN){
  const shW = sh('Watch_OUT');
  if (!shW) throw new Error('Watch_OUT シートが見つかりません。');

  const last = shW.getLastRow();
  if (last < 2){
    notifyDiscord('📣 Watch_OUT に候補がありません（本日の該当なし）');
    return;
  }

  // 1行目ヘッダー＋データ
  const vals = shW.getRange(1,1,last,8).getValues();
  const data = vals.slice(1); // Rank,Code,Name,Reason,Bucket,Side,SuggestQty,Budget

  // Rank昇順（念のため）
  data.sort((a,b)=> (Number(a[0]||99999) - Number(b[0]||99999)));

  // 上位N件（未指定なら全部）
  const rows = (typeof topN === 'number' && topN > 0) ? data.slice(0, Math.min(topN, data.length)) : data;

  // Discordは表の見栄えが不安定なのでコードブロックで整形
  const lines = [];
  lines.push('**本日のスクリーニング結果（Watch_OUT）**', '');
  lines.push('```');
  lines.push(pad('Rank',4)+'  '+pad('Code',6)+'  '+pad('Name',14)+'  '+pad('Bucket',6)+'  '+pad('Side',4)+'  Qty   Reason');
  lines.push('----  ------  --------------  ------  ----  ----  --------------------------------');

  for (const r of rows){
    const rank   = String(r[0]||'');
    const code   = String(r[1]||'');
    const name   = String(r[2]||'');
    const reason = String(r[3]||'');
    const bucket = String(r[4]||'');
    const side   = String(r[5]||'');
    const qty    = (r[6]===null||r[6]==='') ? '-' : String(r[6]);

    lines.push(
      pad(rank,4)+'  '+
      pad(code,6)+'  '+
      pad(truncate(name,14),14)+'  '+
      pad(bucket,6)+'  '+
      pad(side,4)+'  '+
      pad(qty,4)+'  '+
      truncate(reason,60)
    );
  }
  lines.push('```');

  const header = `📊 ${Utilities.formatDate(new Date(), CFG.CLOCK_TZ || Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm')} Watch_OUT（上位${rows.length}件）\n`;
  const chunks = chunkByLength(header + lines.join('\n'), 1800);

  if (!chunks.length){
    notifyDiscord('📣 Watch_OUT は空でした');
    return;
  }
  // 分割送信
  chunks.forEach(text => notifyDiscord(text));
}

/** ====== 内部ユーティリティ（アンダースコア無し） ====== */

// 左寄せ固定幅
function pad(s, width){
  s = String(s||'');
  if (s.length >= width) return s.slice(0, width);
  return s + ' '.repeat(width - s.length);
}

// 末尾省略で幅に収める
function truncate(s, width){
  s = String(s||'');
  if (s.length <= width) return s;
  if (width <= 1) return s.slice(0,width);
  return s.slice(0, width-1) + '…';
}

// Discord 2000字制限対策で分割
function chunkByLength(s, maxLen){
  const out = [];
  let i = 0;
  while (i < s.length){
    out.push(s.slice(i, i+maxLen));
    i += maxLen;
  }
  return out;
}
