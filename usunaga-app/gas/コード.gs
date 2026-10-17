/**
 * 神戸うすなが牧場 月次報告アプリ — データの保存先（Googleスプレッドシート）
 *
 * 使い方
 * 1. 新しいGoogleスプレッドシートを作る（名前は「うすなが月次報告データ」など）
 * 2. メニュー「拡張機能」→「Apps Script」を開き、このファイルの中身を全部貼り付けて保存
 * 3. 下の PASS の '' の中に、スタッフで共有する合言葉を入れて保存
 * 4. 右上「デプロイ」→「新しいデプロイ」→ 歯車から種類「ウェブアプリ」
 *    ・次のユーザーとして実行：自分
 *    ・アクセスできるユーザー：全員
 *    →「デプロイ」（初回は Google の許可画面が出るので許可する）
 *    表示された「ウェブアプリのURL」（…/exec で終わる）をコピー
 * 5. アプリを開くと出る「共有データにつなぐ」画面に、そのURLと合言葉を入れる
 *
 * データは「data」シートに1行1件で入ります。シートを直接編集しないでください。
 * 消した記録は削除済みの印（_del）を付けて残します（他の人の画面からも消えるようにするため）。
 */

const PASS = ''; // ← ここに合言葉を入れる（例：'usunaga2026'）

const SHEET = 'data';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.getRange(1, 1, 1, 4).setValues([['col', 'id', 'json', 'saved_at']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return json_({ ok: true, app: 'usunaga-geppo' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad request' }); }
  if (!PASS || body.pass !== PASS) return json_({ ok: false, error: 'wrong pass' });

  const sh = sheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const last = sh.getLastRow();
    const rows = last > 1 ? sh.getRange(2, 1, last - 1, 4).getValues() : [];

    if (body.op === 'get') {
      const since = Number(body.since || 0);
      const docs = rows
        .filter(r => Number(r[3] || 0) > since)
        .map(r => ({ col: r[0], id: String(r[1]), doc: JSON.parse(r[2]) }));
      return json_({ ok: true, docs: docs, now: Date.now() });
    }

    if (body.op === 'put') {
      const index = {};
      rows.forEach((r, i) => { index[r[0] + '/' + r[1]] = i; });
      const append = [];
      const changed = {};
      const now = Date.now();
      const uOf = r => { try { return Number(JSON.parse(r[2])._u || 0); } catch (err) { return 0; } };
      (body.docs || []).forEach(q => {
        if (!q || !q.col || !q.id || !q.doc) return;
        const key = q.col + '/' + q.id;
        const u = Number(q.doc._u || 0);
        const row = [q.col, q.id, JSON.stringify(q.doc), now];
        if (key in index) {
          const i = index[key];
          if (i < rows.length) {
            if (u >= uOf(rows[i])) { rows[i] = row; changed[i] = true; }
          } else {
            const j = i - rows.length;
            if (u >= uOf(append[j])) append[j] = row;
          }
        } else {
          index[key] = rows.length + append.length;
          append.push(row);
        }
      });
      Object.keys(changed).forEach(k => { const i = Number(k); sh.getRange(i + 2, 1, 1, 4).setValues([rows[i]]); });
      if (append.length) sh.getRange(rows.length + 2, 1, append.length, 4).setValues(append);
      return json_({ ok: true });
    }

    return json_({ ok: false, error: 'unknown op' });
  } finally {
    lock.releaseLock();
  }
}
