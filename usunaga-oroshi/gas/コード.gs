/**
 * うすなが牧場 卸台帳 — データの保存先（Googleスプレッドシート）
 *
 * 使い方
 * 1. 新しいGoogleスプレッドシートを作る（名前は「うすなが卸台帳データ」など）
 * 2. メニュー「拡張機能」→「Apps Script」を開き、このファイルの中身を全部貼り付けて保存
 * 3. 下の PASS の '' の中に、スタッフで共有する合言葉を入れて保存
 * 4. 右上「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *    ・次のユーザーとして実行：自分
 *    ・アクセスできるユーザー：全員
 *    →「デプロイ」。表示された「ウェブアプリのURL」（…/exec で終わる）をコピー
 * 5. 卸台帳アプリの「設定」→「データの保存先」に、そのURLと合言葉を入れて「保存先を登録してつなぐ」
 *
 * データは「data」シートに1行1件で入ります。シートを直接編集しないでください。
 *
 * 在庫表の読み込み：SRC_ID の在庫表（梅本さんの在庫管理表）を読み取り専用で読み込みます。
 * このスクリプトの持ち主のGoogleアカウントに、その在庫表の閲覧権限が必要です。在庫表には書き込みません。
 */

const PASS = ''; // ← ここに合言葉を入れる（例：'usunaga2026'）

const SRC_ID = '1djgEYAVyspbO-f9GqeiU-N1fh3fg7unA1nONlyiz_zY'; // 在庫管理表2
const SRC_SHEETS = ['但馬・神戸', '冷凍・神戸', '冷凍・和牛'];

const DATA_ID = ''; // 保存先スプレッドシートのID（空なら、このスクリプトがひも付いたスプレッドシート）
const SHEET = 'data';

function sheet_() {
  const ss = DATA_ID ? SpreadsheetApp.openById(DATA_ID) : SpreadsheetApp.getActiveSpreadsheet();
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
  return json_({ ok: true, app: 'usunaga-oroshi' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad request' }); }
  if (!PASS || body.pass !== PASS) return json_({ ok: false, error: 'wrong pass' });

  if (body.op === 'srcsheet') {
    try {
      const ss = SpreadsheetApp.openById(SRC_ID);
      const out = {};
      SRC_SHEETS.forEach(n => { const sh = ss.getSheetByName(n); if (sh) out[n] = sh.getDataRange().getDisplayValues(); });
      return json_({ ok: true, name: ss.getName(), sheets: out, at: Date.now() });
    } catch (err) {
      return json_({ ok: false, error: 'src: ' + err });
    }
  }

  const sh = sheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const last = sh.getLastRow();
    const rows = last > 1 ? sh.getRange(2, 1, last - 1, 4).getValues() : [];

    if (body.op === 'get') {
      const docs = rows.map(r => ({ col: r[0], id: String(r[1]), doc: JSON.parse(r[2]) }));
      return json_({ ok: true, docs: docs });
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
