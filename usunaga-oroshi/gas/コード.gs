/**
 * うすなが牧場 卸台帳 — データの保存先（Googleスプレッドシート）
 *
 * 使い方
 * 1. 新しいGoogleスプレッドシートを作る（名前は「卸台帳データ」など）
 * 2. メニュー「拡張機能」→「Apps Script」を開き、このファイルの中身を全部貼り付けて保存
 * 3. 下の PASS の '' の中に、スタッフで共有する合言葉を入れて保存
 * 4. 右上「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *    ・次のユーザーとして実行：自分
 *    ・アクセスできるユーザー：全員
 *    →「デプロイ」。表示された「ウェブアプリのURL」（…/exec で終わる）をコピー
 * 5. 卸台帳アプリの「設定」→「データの保存先」に、そのURLと合言葉を入れて「保存先を登録してつなぐ」
 *
 * データは「data」シートに1行1件で入ります。シートを直接編集しないでください。
 */

const PASS = ''; // ← ここに合言葉を入れる（例：'usunaga2026'）

const SHEET = 'data';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.getRange(1, 1, 1, 4).setValues([['col', 'id', 'json', 'updated']]);
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

  const sh = sheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const last = sh.getLastRow();
    const rows = last > 1 ? sh.getRange(2, 1, last - 1, 4).getValues() : [];

    // ナチュール神戸のアプリから、注文を受注として取り込む（部位名→部位ID、単価は取引先の単価表から）
    if (body.op === 'importOrder') {
      const o = body.order || {};
      const docs = rows.map(r => ({ col: r[0], id: String(r[1]), doc: JSON.parse(r[2]) }));
      const cust = docs.find(d => d.col === 'customers' && d.doc.name === 'ナチュール神戸');
      if (!cust) return json_({ ok: false, error: '取引先「ナチュール神戸」が未登録です' });
      const parts = docs.filter(d => d.col === 'parts');
      const lines = [];
      for (const l of (o.lines || [])) {
        const p = parts.find(d => d.doc.name === l.partName);
        if (!p) return json_({ ok: false, error: '部位が見つかりません：' + l.partName });
        lines.push({ partId: p.id, qty: +l.qty || 0, price: +((cust.doc.prices || {})[p.id]) || 0, stockIds: [], kg: +l.kg || 0, note: '' });
      }
      const id = 'nk-' + String(o.id);
      const doc = { id: id, no: 'NK' + String(o.id), customerId: cust.id, date: o.date || '', deliver: o.shipDate || '', shipDate: o.shipDate || '',
        memo: 'ナチュール神戸（自動取込）', lines: lines, status: o.status === 'shipped' ? 'shipped' : 'alloc', disc: 0,
        createdAt: new Date().toISOString(), source: 'NK', _u: Date.now() };
      const row = ['orders', id, JSON.stringify(doc), doc._u];
      let found = -1;
      rows.forEach((r, i) => { if (r[0] === 'orders' && String(r[1]) === id) found = i; });
      if (found >= 0) sh.getRange(found + 2, 1, 1, 4).setValues([row]);
      else sh.appendRow(row);
      return json_({ ok: true, id: id });
    }

    if (body.op === 'get') {
      const docs = rows.map(r => ({ col: r[0], id: String(r[1]), doc: JSON.parse(r[2]) }));
      return json_({ ok: true, docs: docs });
    }

    if (body.op === 'put') {
      const index = {};
      rows.forEach((r, i) => { index[r[0] + '/' + r[1]] = i; });
      const append = [];
      (body.docs || []).forEach(q => {
        if (!q || !q.col || !q.id || !q.doc) return;
        const key = q.col + '/' + q.id;
        const u = Number(q.doc._u || 0);
        const row = [q.col, q.id, JSON.stringify(q.doc), u];
        if (key in index) {
          const i = index[key];
          if (u >= Number(rows[i][3] || 0)) rows[i] = row;
        } else {
          index[key] = rows.length + append.length;
          append.push(row);
        }
      });
      if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);
      if (append.length) sh.getRange(rows.length + 2, 1, append.length, 4).setValues(append);
      return json_({ ok: true });
    }

    return json_({ ok: false, error: 'unknown op' });
  } finally {
    lock.releaseLock();
  }
}
