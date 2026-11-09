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
 * このスクリプトの持ち主のGoogleアカウントに、その在庫表の閲覧権限が必要です。在庫表には書き込みません（セルの色も読むだけ）。
 * 証明書PDFは、このスクリプトの持ち主のドライブ「うすなが卸台帳 証明書」フォルダに保存します。
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
      const bgs = {};
      // 在庫表のシートを全部見て、「部位」「左重量」の見出しがあるシートだけ読む（シート名が変わっても大丈夫なように）
      const names = ss.getSheets().map(sh => sh.getName());
      ss.getSheets().forEach(sh => {
        const n = sh.getName();
        if (sh.isSheetHidden() && SRC_SHEETS.indexOf(n) < 0) return;
        const rg = sh.getDataRange();
        const vals = rg.getDisplayValues();
        const head = vals.slice(0, 8).map(r => r.join('|')).join('|');
        if (SRC_SHEETS.indexOf(n) < 0 && !(head.indexOf('部位') >= 0 && head.indexOf('左') >= 0)) return;
        out[n] = vals;
        // セルの色（白以外だけ）："行,列" → 色。色で 在庫／売約／一部売約／納品済 を見分ける
        const m = {};
        rg.getBackgrounds().forEach((row, i) => row.forEach((c, j) => { if (c && c !== '#ffffff') m[i + ',' + j] = c; }));
        bgs[n] = m;
      });
      return json_({ ok: true, name: ss.getName(), names: names, sheets: out, bgs: bgs, at: Date.now() });
    } catch (err) {
      return json_({ ok: false, error: 'src: ' + err });
    }
  }

  if (body.op && body.op.indexOf('cert') === 0) return cert_(body);

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

/* ---------- 証明書（PDF）の保管：このスクリプトの持ち主のGoogleドライブ「うすなが卸台帳 証明書」フォルダ ---------- */
const CERT_FOLDER = 'うすなが卸台帳 証明書';
function certFolder_() {
  const it = DriveApp.getFoldersByName(CERT_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CERT_FOLDER);
}
function certInfo_(f) {
  let meta = {}; try { meta = JSON.parse(f.getDescription() || '{}'); } catch (err) {}
  return { id: f.getId(), name: f.getName(), edaban: meta.edaban || '', kind: meta.kind || '', size: f.getSize(), at: f.getDateCreated().getTime() };
}
function cert_(body) {
  try {
    const folder = certFolder_();
    if (body.op === 'certlist') {
      const list = []; const it = folder.getFiles();
      while (it.hasNext()) { const f = it.next(); if (!f.isTrashed()) list.push(certInfo_(f)); }
      return json_({ ok: true, files: list });
    }
    if (body.op === 'certput') {
      if (!body.data || !body.edaban) return json_({ ok: false, error: 'no data' });
      const blob = Utilities.newBlob(Utilities.base64Decode(body.data), body.mime || 'application/pdf', String(body.name || 'cert.pdf'));
      const f = folder.createFile(blob);
      f.setDescription(JSON.stringify({ edaban: String(body.edaban), kind: String(body.kind || '') }));
      return json_({ ok: true, file: certInfo_(f) });
    }
    const f = DriveApp.getFileById(String(body.id || ''));
    let inFolder = false; const ps = f.getParents(); while (ps.hasNext()) { if (ps.next().getId() === folder.getId()) inFolder = true; }
    if (!inFolder) return json_({ ok: false, error: 'not found' });
    if (body.op === 'certget') return json_({ ok: true, file: certInfo_(f), mime: f.getMimeType(), data: Utilities.base64Encode(f.getBlob().getBytes()) });
    if (body.op === 'certset') { f.setDescription(JSON.stringify({ edaban: String(body.edaban || ''), kind: String(body.kind || '') })); return json_({ ok: true, file: certInfo_(f) }); }
    if (body.op === 'certdel') { f.setTrashed(true); return json_({ ok: true }); }
    return json_({ ok: false, error: 'unknown op' });
  } catch (err) {
    return json_({ ok: false, error: 'cert: ' + err });
  }
}
