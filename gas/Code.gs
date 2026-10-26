/**
 * NATURE KOBE 肉の注文管理 — GAS（スプレッドシート「NATURE KOBE 注文管理」にバインド）
 *
 * ■初回だけ
 *   1. エディタ上部の関数で「setup」を選んで ▶実行（権限を許可）
 *      → シート作成／商品・設定の初期値／9/27奈良販売会の13件を登録
 *      → 管理用・精肉店用パスワードを発行（実行ログに表示）
 *   2. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *      実行ユーザー：自分　／　アクセスできるユーザー：全員
 *   3. 表示されたURLを index.html の API_URL に貼る
 * ■パスワード変更：プロジェクトの設定 → スクリプトプロパティ ADMIN_PASS / SHOP_PASS
 */
const SH = { orders: '注文', events: 'イベント', products: '商品', settings: '設定' };
const ORDER_HEAD = ['注文番号', 'イベント', '受付日時', '入力元', '氏名', 'メール', '電話', '税抜', '合計(税込)', '支払', 'ステータス', '入金', '発送日', '伝票番号', '更新日時', 'JSON'];
const JSON_COL = ORDER_HEAD.length;
const STATUSES = ['受付', '発注済', '発送済', 'キャンセル'];
const TZ = 'Asia/Tokyo';
/* ネット販売（イベントID固定）：振込のみ・送料なし・ワイン特典なし・上乗せなし（仕入れ×1.2） */
const NET_EVENT_ID = 'NET-ONLINE';
const NET_PAY_DAYS = 7;
function isNet_(id) { return String(id || '') === NET_EVENT_ID; }
const SHEET_ID = '1dtIiKiQKuVY1fN65YXXiun4rJ2L5pxmUh70SkmaaLwg'; // スプレッドシート「NATURE KOBE 注文管理」

/* ---------- 共通 ---------- */
function book_() {
  return SpreadsheetApp.openById(SHEET_ID);
}
function sheet_(n) { return book_().getSheetByName(n); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function need_(b) { if (!b) throw new Error('パスワードが違います'); }

function doGet() { return out_({ ok: true, app: 'nature-kobe-order' }); }

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad request' }); }
  if (req && Array.isArray(req.events)) return lineWebhook_(req);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return out_(route_(req)); }
  catch (err) { return out_({ ok: false, error: String(err.message || err) }); }
  finally { lock.releaseLock(); }
}

function route_(r) {
  const P = PropertiesService.getScriptProperties();
  const isAdmin = !!r.pass && r.pass === P.getProperty('ADMIN_PASS');
  const isShop = !!r.pass && r.pass === P.getProperty('SHOP_PASS');
  switch (r.action) {
    case 'public': return publicInfo_(r.e);
    case 'submit': return submit_(r.order);
    case 'admin': need_(isAdmin); return { ok: true, events: events_(), products: products_(), settings: settings_(), orders: orders_() };
    case 'saveOrder': need_(isAdmin); return { ok: true, order: saveOrder_(r.order) };
    case 'bulkStatus': need_(isAdmin); return { ok: true, orders: bulkStatus_(r.ids, r.status) };
    case 'saveEvent': need_(isAdmin); return { ok: true, events: saveEvent_(r.event) };
    case 'saveProducts': need_(isAdmin); return { ok: true, products: saveProducts_(r.products) };
    case 'saveSettings': need_(isAdmin); return { ok: true, settings: saveSettings_(r.settings) };
    case 'deleteEvent': need_(isAdmin); deleteEvent_(r.id); return { ok: true, events: events_(), orders: orders_() };
    case 'deleteOrder': need_(isAdmin); deleteRowsById_(SH.orders, [String(r.id)]); return { ok: true };
    case 'shop': need_(isShop); return { ok: true, events: events_(), settings: { shopName: settings_().shopName || '' }, orders: shopOrders_() };
    case 'shopUpdate': need_(isShop); return { ok: true, order: shopUpdate_(r) };
    case 'mailPreview': need_(isAdmin || isShop); return mailPreview_(r, isAdmin);
    case 'sendMail': need_(isAdmin || isShop); return sendMail_(r, isAdmin);
    case 'setupNet': need_(isAdmin); ensureNetEvent_(); return { ok: true, events: events_() };
    case 'installReminder': need_(isAdmin); installReminderTrigger_(); return { ok: true };
    case 'setLine': need_(isAdmin); { const P = PropertiesService.getScriptProperties(); if (String(r.token || '').trim()) P.setProperty('LINE_TOKEN', String(r.token).trim()); return { ok: true, groupSet: !!P.getProperty('LINE_GROUP_ID') }; }
    case 'testLine': need_(isAdmin); return testLine_();
    case 'receiptPdf': need_(isAdmin); { const o = orders_().find(x => x.id === r.id); return { ok: true, pdf: Utilities.base64Encode(receiptPdf_(o).getBytes()) }; }
    case 'setRelay': need_(isAdmin); PropertiesService.getScriptProperties().setProperties({ MAIL_RELAY_URL: String(r.url || ''), MAIL_RELAY_KEY: String(r.key || '') }); return { ok: true };
    case 'testRelay': need_(isAdmin); relaySend_(String(r.to), '【テスト】NATURE KOBE 注文アプリからの送信テスト', 'このメールは送信テストです。', null); return { ok: true };
    default: throw new Error('unknown action');
  }
}

/* ---------- 読み込み ---------- */
function settings_() {
  const o = {};
  sheet_(SH.settings).getDataRange().getValues().slice(1).forEach(r => { if (r[0]) o[r[0]] = String(r[1]); });
  return o;
}
function pubSettings_() {
  const S = settings_();
  return { taxRate: S.taxRate, shipTaxRate: S.shipTaxRate || '10', shippingFee: S.shippingFee, sellerName: S.sellerName, processorName: S.processorName || '', bankInfo: S.bankInfo, deliveryNote: S.deliveryNote };
}
function products_() {
  return sheet_(SH.products).getDataRange().getValues().slice(1).filter(r => r[0]).map(r => ({
    id: String(r[0]), cat: String(r[1]), name: String(r[2]), price: +r[3] || 0,
    variable: String(r[4]).toUpperCase() === 'TRUE', sort: +r[5] || 0, active: String(r[6]).toUpperCase() !== 'FALSE'
  })).sort((a, b) => a.sort - b.sort);
}
function events_() {
  return sheet_(SH.events).getDataRange().getValues().slice(1).filter(r => r[0]).map(r => ({
    id: String(r[0]), name: String(r[1]), date: (r[2] instanceof Date ? Utilities.formatDate(r[2], TZ, 'yyyy-MM-dd') : String(r[2])), venue: String(r[3]), open: String(r[4]).toUpperCase() === 'TRUE',
    partner: String(r[5] == null ? '' : r[5]), markup: +r[6] || 0, perks: String(r[7] == null ? '' : r[7]), extras: parseExtras_(r[8]), invoice: String(r[9] == null ? '' : r[9])
  })).sort((a, b) => (a.date < b.date ? 1 : -1));
}
// 販売会ごとの追加商品（JSON）。price は税抜・円/kg、cost は原価（卸価格・税抜・円/kg）
function parseExtras_(v) {
  try { const a = JSON.parse(String(v || '[]')); return Array.isArray(a) ? a.filter(x => x && x.name && +x.price > 0).map(x => ({ id: 'x:' + String(x.name), cat: '追加', name: String(x.name), price: +x.price, cost: +x.cost || 0, variable: false, sort: 999, active: true })) : []; }
  catch (e) { return []; }
}
function orders_() {
  const sh = sheet_(SH.orders);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, JSON_COL, sh.getLastRow() - 1, 1).getValues()
    .map(r => { try { return JSON.parse(r[0]); } catch (e) { return null; } }).filter(Boolean);
}

/* ---------- 書き込み ---------- */
function orderRow_(o) {
  return [o.id, o.eventId, o.createdAt, o.source, o.name, o.email || '', o.phone || '', o.subtotal, o.total, o.payment || '', o.status,
    o.paid ? '済' : '未', o.shipDate || '', o.tracking || '', now_(), JSON.stringify(o)];
}
function writeOrder_(o) {
  const sh = sheet_(SH.orders);
  const last = sh.getLastRow();
  if (last >= 2) {
    const ids = sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]));
    const i = ids.indexOf(o.id);
    if (i >= 0) { sh.getRange(i + 2, 1, 1, ORDER_HEAD.length).setValues([orderRow_(o)]); return o; }
  }
  sh.appendRow(orderRow_(o));
  return o;
}
function newId_(all, ymd) {
  ymd = ymd || Utilities.formatDate(new Date(), TZ, 'yyMMdd');
  const pre = 'NK' + ymd + '-';
  let max = 0;
  all.forEach(o => { if (String(o.id).indexOf(pre) === 0) max = Math.max(max, +String(o.id).slice(pre.length) || 0); });
  return pre + ('00' + (max + 1)).slice(-3);
}

/* 金額計算（g × kg単価、1行ごとに四捨五入、消費税は切り捨て）。trust=false のときは商品マスタの単価を使う */
/* 販売店バック用の上乗せ価格（100円単位に切り上げ） */
function markupPrice_(base, m) { m = +m || 0; return m ? Math.ceil(base * (1 + m / 100) / 100) * 100 : base; }
function compute_(o, S, prods, trust) {
  const pm = {};
  prods.forEach(p => pm[p.id] = p);
  const m = +o.markup || 0;
  let sub = 0, back = 0;
  o.items = (o.items || []).filter(it => +it.g > 0 && +it.count > 0 && (trust || pm[it.pid])).map(it => {
    const p = pm[it.pid];
    const base = trust && it.basePrice !== '' && it.basePrice != null ? +it.basePrice
      : trust && it.unitPrice !== '' && it.unitPrice != null ? +it.unitPrice : (p ? p.price : 0);
    const unit = markupPrice_(base, m);
    const amount = Math.round(unit * (+it.g) * (+it.count) / 1000);
    sub += amount;
    back += amount - Math.round(base * (+it.g) * (+it.count) / 1000);
    return { pid: String(it.pid || ''), name: p ? (p.cat + ' ' + p.name) : String(it.name || ''), basePrice: base, unitPrice: unit, variable: p ? p.variable : !!it.variable,
      g: +it.g, count: +it.count, amount: amount, note: String(it.note || ''),
      cost: p && p.cost != null && p.cost !== '' ? +p.cost : (it.cost != null && it.cost !== '' ? +it.cost : '') };
  });
  o.extras = (trust ? (o.extras || []) : []).filter(x => x && String(x.name || '').trim())
    .map(x => ({ name: String(x.name), amount: +x.amount || 0, note: String(x.note || '') }));
  o.extras.forEach(x => sub += x.amount);
  o.subtotal = sub;
  o.markup = m;
  /* 個人販売など、販売店バックを付けない注文は back を 0 にする（売上・原価はそのまま） */
  o.noBack = !!(trust && o.noBack);
  o.back = o.noBack ? 0 : back;
  o.shipping = trust && o.shipping !== '' && o.shipping != null ? +o.shipping : (+S.shippingFee || 0);
  o.taxRate = trust && o.taxRate !== '' && o.taxRate != null ? +o.taxRate : (+S.taxRate || 10);
  /* 軽減税率：お肉は taxRate（8%）、送料は shipTaxRate（10%）。旧データ（shipTaxRate なし）は従来どおり同じ税率 */
  const legacy = trust && (o.shipTaxRate === '' || o.shipTaxRate == null) && o.createdAt && String(o.createdAt) < '2026-10-18 09';
  o.shipTaxRate = trust && o.shipTaxRate !== '' && o.shipTaxRate != null ? +o.shipTaxRate : legacy ? o.taxRate : (+S.shipTaxRate || 10);
  o.taxFood = Math.floor(sub * o.taxRate / 100);
  o.taxShip = Math.floor(o.shipping * o.shipTaxRate / 100);
  o.tax = o.taxFood + o.taxShip;
  o.total = sub + o.shipping + o.tax;
  return o;
}

/* ---------- お客さん ---------- */
function publicInfo_(eid) {
  const ev = events_().find(x => x.id === eid);
  if (!ev) throw new Error('注文ページが見つかりません');
  const prods = products_().filter(p => p.active).concat(ev.extras || []).map(p => Object.assign({}, p, { price: markupPrice_(p.price, ev.markup) }));
  return { ok: true, event: { id: ev.id, name: ev.name, date: ev.date, venue: ev.venue, open: ev.open, perks: ev.perks, net: isNet_(ev.id) }, products: prods, settings: pubSettings_() };
}
function submit_(o) {
  if (!o) throw new Error('bad request');
  const ev = events_().find(x => x.id === o.eventId);
  if (!ev || !ev.open) throw new Error('このイベントのご注文受付は終了しました');
  ['name', 'email', 'phone', 'zip', 'address'].forEach(k => { if (!String(o[k] || '').trim()) throw new Error('未入力の項目があります'); });
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(o.email).trim())) throw new Error('メールアドレスの形式が正しくありません');
  if (['冷凍', '冷蔵'].indexOf(o.storage) < 0) throw new Error('保存状態を選んでください');
  const netOrder = isNet_(ev.id);
  if ((netOrder ? ['振込'] : ['現金', '振込']).indexOf(o.payment) < 0) throw new Error('お支払い方法を選んでください');
  const all = orders_();
  if (o.clientKey) {
    const dup = all.find(x => x.clientKey === o.clientKey);
    if (dup) return { ok: true, order: dup };
  }
  const st = o.shipTo && o.shipTo.on
    ? { on: true, name: String(o.shipTo.name || ''), phone: String(o.shipTo.phone || ''), zip: String(o.shipTo.zip || ''), address: String(o.shipTo.address || '') }
    : { on: false };
  if (st.on && (!st.name.trim() || !st.address.trim())) throw new Error('お届け先のお名前・住所を入力してください');
  const c = {
    id: newId_(all), clientKey: String(o.clientKey || ''), eventId: ev.id, createdAt: now_(), source: 'お客さん',
    name: String(o.name).trim(), email: String(o.email).trim(), phone: String(o.phone).trim(), zip: String(o.zip).trim(), address: String(o.address).trim(),
    shipTo: st, storage: o.storage, payment: o.payment, receipt: !!o.receipt, receiptTo: o.receipt ? String(o.receiptTo || o.name) : '',
    note: String(o.note || '').slice(0, 1000), items: o.items, extras: [], status: '受付', paid: false, paidAt: '',
    mail: { confirm: false, ship: false }, checks: '', memo: '', markup: ev.markup
  };
  compute_(c, settings_(), products_().filter(p => p.active).concat(ev.extras || []), false);
  if (!c.items.length) throw new Error('商品が選ばれていません');
  if (c.items.some(it => it.g < 300 || it.g % 300)) throw new Error('お肉は300g単位で選んでください');
  // 現金の注文は、注文が入った時点で「発注済」にして、うすなが牧場（碓永さん）へ通知する
  if (c.payment === '現金') { c.status = '発注済'; c.orderedAt = now_(); }
  writeOrder_(c);
  if (c.payment === '現金') { try { notifySupplier_(c); } catch (err) { Logger.log('発注通知失敗: ' + err.message); } try { sendUsunagaOrder_(c); } catch (err) { Logger.log('卸台帳送信失敗: ' + err.message); } }
  try { notifyLine_(c, ev); } catch (err) { Logger.log('LINE通知失敗: ' + err.message); }
  if (c.payment === '振込') { try { sendBankMail_(c, ev); } catch (err) { Logger.log('振込案内メール失敗: ' + err.message); } }
  try { sendOrderMails_(c, ev); } catch (err) { Logger.log('注文メール失敗: ' + err.message); }
  return { ok: true, order: c };
}

/* ---------- 管理 ---------- */
function saveOrder_(o) {
  if (!o) throw new Error('bad request');
  const all = orders_();
  const prev = o.id ? all.find(x => x.id === o.id) : null;
  if (!prev) {
    o.id = newId_(all);
    o.createdAt = now_();
    o.source = o.source || 'スタッフ';
    o.mail = o.mail || { confirm: false, ship: false };
  }
  if (STATUSES.indexOf(o.status) < 0) o.status = '受付';
  if (o.markup == null || o.markup === '' || !prev || prev.eventId !== o.eventId) {
    const ev = events_().find(x => x.id === o.eventId);
    o.markup = ev ? ev.markup : 0;
  }
  compute_(o, settings_(), products_(), true);
  const prevStatus = prev ? prev.status : '';
  // 振込の注文が「入金済」になった時、自動で「発注済」にして、うすなが牧場（碓永さん）へ通知する
  if (prev && o.paid && !prev.paid && o.status === '受付') {
    o.status = '発注済';
    o.orderedAt = now_();
    try { notifySupplier_(o); } catch (err) { Logger.log('発注通知失敗: ' + err.message); }
  }
  if (o.status !== prevStatus && (o.status === '発注済' || o.status === '発送済')) { try { sendUsunagaOrder_(o); } catch (err) { Logger.log('卸台帳送信失敗: ' + err.message); } }
  return writeOrder_(o);
}

// うすなが牧場（碓永さん）へ、発注内容をLINEグループに通知する
function notifySupplier_(o) {
  const P = PropertiesService.getScriptProperties();
  const token = P.getProperty('LINE_TOKEN'), to = P.getProperty('LINE_GROUP_ID');
  if (!token || !to) return;
  const items = (o.items || []).map(i => '・' + i.name + ' ' + gl_(i.g) + '×' + i.count).concat((o.extras || []).map(x => '・' + x.name));
  const text = '【発注】ナチュール神戸\n'
    + '注文番号：' + o.id + '\n'
    + '保管：' + (o.storage === '冷蔵' ? '冷蔵' : '冷凍') + '\n'
    + items.join('\n') + '\n'
    + '※発注の内容をご確認ください';
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ to: to, messages: [{ type: 'text', text: text }] })
  });
}
// 卸台帳の部位名は、ナチュール神戸の商品名（例：スライス 赤身）と同じ名前で登録してもらう。そのため、変換は不要

// うすなが牧場の卸台帳へ、注文を受注として自動で送る（発注済・発送済になった時）
function sendUsunagaOrder_(o) {
  const P = PropertiesService.getScriptProperties();
  const url = P.getProperty('USUNAGA_URL'), pass = P.getProperty('USUNAGA_PASS');
  if (!url || !pass) { Logger.log('卸台帳の送信先が未設定のため、送りません: ' + o.id); return; }
  const lines = (o.items || []).map(i => ({
    partName: i.name,
    kg: (+i.g || 0) * (+i.count || 0) / 1000,
    qty: +i.count || 0
  }));
  const payload = { op: 'importOrder', pass: pass, order: {
    id: o.id, date: String(o.createdAt || '').slice(0, 10), shipDate: o.shipDate || '',
    status: o.status === '発送済' ? 'shipped' : 'alloc', lines: lines
  } };
  const res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'text/plain;charset=utf-8', payload: JSON.stringify(payload), muteHttpExceptions: true, followRedirects: true });
  let j = null;
  try { j = JSON.parse(res.getContentText()); } catch (e) {}
  if (!j || !j.ok) throw new Error('卸台帳への送信に失敗：' + (j && j.error ? j.error : '応答が不正です'));
}

function bulkStatus_(ids, status) {
  if (STATUSES.indexOf(status) < 0) throw new Error('bad status');
  const all = orders_();
  const res = [];
  (ids || []).forEach(id => {
    const o = all.find(x => x.id === id);
    if (o) {
      const prevStatus = o.status;
      o.status = status; if (status === '発注済') o.orderedAt = now_();
      writeOrder_(o); res.push(o);
      if (status !== prevStatus && (status === '発注済' || status === '発送済')) { try { sendUsunagaOrder_(o); } catch (err) { Logger.log('卸台帳送信失敗: ' + err.message); } }
    }
  });
  return res;
}
function saveEvent_(ev) {
  const sh = sheet_(SH.events);
  if (!ev.id) ev.id = 'e' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmmss');
  if (sh.getRange(1, 8).getValue() === '') sh.getRange(1, 6, 1, 3).setValues([['販売店', '上乗せ率(%)', '購入特典']]);
  if (sh.getRange(1, 9).getValue() === '') sh.getRange(1, 9).setValue('追加商品');
  if (sh.getRange(1, 10).getValue() === '') sh.getRange(1, 10).setValue('請求書リンク');
  const extras = (Array.isArray(ev.extras) ? ev.extras : []).filter(x => x && String(x.name || '').trim() && +x.price > 0)
    .map(x => ({ name: String(x.name).trim(), price: +x.price, cost: +x.cost || 0 }));
  const row = [ev.id, ev.name || '', ev.date || '', ev.venue || '', ev.open ? 'TRUE' : 'FALSE', ev.partner || '', +ev.markup || 0, ev.perks || '', JSON.stringify(extras), String(ev.invoice || '').trim()];
  const v = sh.getDataRange().getValues();
  const i = v.findIndex(r => String(r[0]) === ev.id);
  if (i >= 1) sh.getRange(i + 1, 1, 1, 10).setValues([row]); else sh.appendRow(row);
  return events_();
}
function saveProducts_(list) {
  const sh = sheet_(SH.products);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 7).clearContent();
  const rows = (list || []).filter(p => p.name).map((p, i) => [p.id || ('p' + Date.now() + i), p.cat || '', p.name, +p.price || 0,
    p.variable ? 'TRUE' : 'FALSE', i + 1, p.active === false ? 'FALSE' : 'TRUE']);
  if (rows.length) sh.getRange(2, 1, rows.length, 7).setValues(rows);
  return products_();
}
function saveSettings_(s) {
  const sh = sheet_(SH.settings);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  const rows = Object.keys(s || {}).map(k => [k, String(s[k])]);
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setValues(rows);
  return settings_();
}

/* 削除（イベントを消すと、そのイベントの注文も消える） */
function deleteRowsById_(name, ids) {
  const sh = sheet_(name);
  if (sh.getLastRow() < 2) return;
  const v = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = v.length - 1; i >= 1; i--) if (ids.indexOf(String(v[i][0])) >= 0) sh.deleteRow(i + 1);
}
function deleteEvent_(id) {
  id = String(id);
  deleteRowsById_(SH.orders, orders_().filter(o => o.eventId === id).map(o => o.id));
  deleteRowsById_(SH.events, [id]);
}

/* ---------- 精肉店（金額・メールは見せない） ---------- */
function shopView_(o) {
  return {
    id: o.id, eventId: o.eventId, createdAt: o.createdAt, orderedAt: o.orderedAt || '', name: o.name, phone: o.phone, zip: o.zip, address: o.address,
    shipTo: o.shipTo, storage: o.storage, note: o.note,
    items: (o.items || []).map(i => ({ name: i.name, g: i.g, count: i.count, note: i.note })),
    extras: (o.extras || []).map(x => ({ name: x.name, note: x.note })),
    status: o.status, shipDate: o.shipDate || '', carrier: o.carrier || '', tracking: o.tracking || '',
    hasEmail: !!o.email, shipMailSent: !!(o.mail && o.mail.ship)
  };
}
function shopOrders_() {
  return orders_().filter(o => o.status === '発注済' || o.status === '発送済').map(shopView_);
}
function shopUpdate_(r) {
  const o = orders_().find(x => x.id === r.id);
  if (!o || (o.status !== '発注済' && o.status !== '発送済')) throw new Error('この注文は更新できません');
  o.shipDate = String(r.shipDate || '');
  o.carrier = String(r.carrier || '');
  o.tracking = String(r.tracking || '');
  if (o.shipDate && o.tracking) { o.status = '発送済'; o.shippedAt = o.shippedAt || now_(); }
  writeOrder_(o);
  return shopView_(o);
}

/* ---------- 自動メール（info@naturekobe.com から送信。送信用の中継スクリプト経由） ---------- */
function yen_(n) { return '¥' + Math.round(+n || 0).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function gl_(g) { g = +g || 0; return g >= 1000 ? (Math.round(g / 10) / 100) + 'kg' : g + 'g'; }
function jdate_(s) { const m = String(s || '').match(/(\d{4})-(\d{1,2})-(\d{1,2})/); return m ? (+m[1]) + '年' + (+m[2]) + '月' + (+m[3]) + '日' : String(s || ''); }
function mdw_(s) { const m = String(s || '').match(/(\d{4})-(\d{1,2})-(\d{1,2})/); if (!m) return String(s || ''); const d = new Date(+m[1], +m[2] - 1, +m[3]); return (+m[2]) + '月' + (+m[3]) + '日（' + '日月火水木金土'.charAt(d.getDay()) + '）'; }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function payLabel_(p) { return p === 'クレジットカード' ? 'クレジットカード' : p === '現金' || p === '当日現金' ? '現金' : p === '振込' ? '銀行振込' : (p || ''); }
function trackUrl_(carrier, no) {
  const n = String(no || '').replace(/[^0-9A-Za-z]/g, '');
  if (!n) return '';
  if (carrier === 'ヤマト運輸') return 'https://jizen.kuronekoyamato.co.jp/jizen/servlet/crjz.b.NQ0010?id=' + n;
  if (carrier === '佐川急便') return 'https://k2k.sagawa-exp.co.jp/p/web/okurijosearch.do?okurijoNo=' + n;
  if (carrier === '日本郵便') return 'https://trackings.post.japanpost.jp/services/srv/search/direct?reqCodeNo1=' + n + '&locale=ja';
  return '';
}
function sig_(S) {
  return String(S.mailSignature || '').trim() || [S.sellerName || 'NATURE KOBE', S.sellerAddress, S.sellerTel ? 'TEL ' + S.sellerTel : '', S.sellerEmail].filter(String).join('\n');
}
function buildMail_(o, kind) {
  const S = settings_(), ev = events_().find(x => x.id === o.eventId) || {};
  const seller = (S.sellerName || 'NATURE KOBE') + (S.processorName ? '（加工元：' + S.processorName + '）' : '');
  const items = (o.items || []).map(i => '・' + i.name + ' ' + gl_(i.g) + '×' + i.count + (i.note ? '（' + i.note + '）' : ''))
    .concat((o.extras || []).map(x => '・' + x.name));
  const dest = o.shipTo && o.shipTo.on ? '〒' + (o.shipTo.zip || '') + ' ' + (o.shipTo.address || '') + '\n' + (o.shipTo.name || '') + ' 様'
    : '〒' + (o.zip || '') + ' ' + (o.address || '') + '\n' + o.name + ' 様';
  if (kind === 'receipt') {
    return {
      subject: '【' + (S.sellerName || 'NATURE KOBE') + '】お支払いありがとうございました／領収書のお送り（注文番号 ' + o.id + '）',
      body: o.name + ' 様\n\nこのたびは' + (ev.name ? '「' + ev.name + '」にて、' : '') + '但馬牛・神戸ビーフをご注文いただき、誠にありがとうございます。\n'
        + 'お支払いを確認いたしましたので、領収書をPDFで添付いたします。\n\n'
        + '■注文番号：' + o.id + '\n■お支払い金額：' + yen_(o.total) + '（税込）\n■お支払い日：' + jdate_(o.paidAt) + '\n■お支払い方法：' + payLabel_(o.payment) + '\n'
        + '■領収書の宛名：' + (o.receiptTo || o.name) + ' 様\n\n■ご注文内容\n' + items.join('\n') + '\n\n■お届け先\n' + dest + '\n\n'
        + (S.deliveryNote || '') + '\n発送の際は、配送伝票番号をあらためてメールでお知らせいたします。\n\n■販売者：' + seller + '\n\n'
        + 'ご不明な点がございましたら、本メールにご返信ください。\n今後ともよろしくお願いいたします。\n\n' + sig_(S)
    };
  }
  if (kind === 'ship') {
    const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    const future = o.shipDate && o.shipDate > today;
    const url = trackUrl_(o.carrier, o.tracking);
    return {
      subject: '【' + (S.sellerName || 'NATURE KOBE') + '】商品' + (future ? '発送日' : '発送') + 'のお知らせ（注文番号 ' + o.id + '）',
      body: o.name + ' 様\n\nこのたびはご注文いただき、誠にありがとうございます。\nご注文の商品を' + (o.shipDate ? mdw_(o.shipDate) : '') + 'に' + (future ? '発送予定です。' : '発送いたしました。') + '\n\n'
        + '■配送業者：' + (o.carrier || '') + '（' + (o.storage === '冷蔵' ? 'クール便・冷蔵' : 'クール便・冷凍') + '）\n■お問い合わせ伝票番号：' + (o.tracking || '') + '\n'
        + (url ? '■配送状況の確認：' + url + '\n' : '') + '\n■お届け先\n' + dest + '\n\n■ご注文内容\n' + items.join('\n') + '\n\n■販売者：' + seller + '\n\n'
        + 'お届け日時のご変更は、上記の伝票番号で配送業者へお問い合わせください。\nお受け取り後は、お早めに' + (o.storage === '冷蔵' ? '冷蔵庫' : '冷凍庫') + 'で保存してください。\n\n'
        + 'お肉をお楽しみいただけましたら幸いです。ご感想もぜひ本メールにご返信ください。\n今後ともよろしくお願いいたします。\n\n' + sig_(S)
    };
  }
  throw new Error('unknown mail');
}
function receiptHtml_(o) {
  const S = settings_();
  const card = o.payment === 'クレジットカード';
  return '<html><head><meta charset="utf-8"><style>body{font-family:"Noto Serif JP","Hiragino Mincho ProN",serif;color:#111;margin:40px 50px}'
    + 'h1{text-align:center;font-size:30px;letter-spacing:12px;margin:0 0 24px}.row{display:flex;justify-content:space-between;font-size:12px}'
    + '.to{font-size:22px;border-bottom:1px solid #111;padding-bottom:4px;margin:26px 0;width:70%}.amt{font-size:28px;text-align:center;border:1px solid #111;background:#f4f1ea;padding:14px;margin:0 0 12px}'
    + 'table{border-collapse:collapse;font-size:12px}td{border:1px solid #111;padding:5px 10px}.iss{text-align:right;font-size:13px;line-height:1.8;margin-top:30px}</style></head><body>'
    + '<h1>領収書</h1><table style="width:100%;border:0"><tr><td style="border:0">No. ' + esc_(o.id) + '</td><td style="border:0;text-align:right">発行日　' + esc_(jdate_(o.paidAt)) + '</td></tr></table>'
    + '<div class="to">' + esc_(o.receiptTo || o.name) + '　様</div>'
    + '<div class="amt">金額　' + yen_(o.total) + ' −（税込）</div>'
    + '<p style="font-size:14px;margin:6px 0">但し　但馬牛・神戸ビーフ代として' + (card ? '（クレジットカード決済）' : '') + '</p><p style="font-size:14px;margin:6px 0">上記正に領収いたしました。</p>'
    + '<table style="margin-top:22px"><tr><td colspan="2">内訳</td></tr><tr><td>税抜金額</td><td style="text-align:right">' + yen_(o.subtotal + (o.shipping || 0)) + '</td></tr>'
    + (o.shipping && o.shipTaxRate === o.taxRate ? '<tr><td>　うち送料</td><td style="text-align:right">' + yen_(o.shipping) + '</td></tr>' : '')
    + (o.shipping && o.shipTaxRate !== o.taxRate
      ? '<tr><td>' + o.taxRate + '%対象（税抜）</td><td style="text-align:right">' + yen_(o.subtotal) + '</td></tr><tr><td>　消費税額（' + o.taxRate + '%）</td><td style="text-align:right">' + yen_(o.taxFood) + '</td></tr>'
        + '<tr><td>' + o.shipTaxRate + '%対象（送料・税抜）</td><td style="text-align:right">' + yen_(o.shipping) + '</td></tr><tr><td>　消費税額（' + o.shipTaxRate + '%）</td><td style="text-align:right">' + yen_(o.taxShip) + '</td></tr>'
      : '<tr><td>消費税額（' + o.taxRate + '%）</td><td style="text-align:right">' + yen_(o.tax) + '</td></tr>')
    + '<tr><td>消費税額 合計</td><td style="text-align:right">' + yen_(o.tax) + '</td></tr></table>'
    + (o.taxRate === 8 ? '<p style="font-size:11px;margin:6px 0">※お肉は軽減税率（8%）対象です。</p>' : '')
    + '<div class="iss"><b style="font-size:17px">' + esc_(S.sellerName || 'NATURE KOBE') + '</b><br>' + esc_(S.sellerAddress || '') + (S.sellerTel ? '<br>TEL ' + esc_(S.sellerTel) : '') + (S.sellerEmail ? '<br>' + esc_(S.sellerEmail) : '') + '</div>'
    + '</body></html>';
}
function receiptPdf_(o) {
  return Utilities.newBlob(receiptHtml_(o), 'text/html', 'receipt.html').getAs('application/pdf').setName('領収書_' + o.id + '.pdf');
}
function relaySend_(to, subject, body, pdf) {
  const P = PropertiesService.getScriptProperties();
  const url = P.getProperty('MAIL_RELAY_URL'), key = P.getProperty('MAIL_RELAY_KEY');
  if (!url || !key) throw new Error('送信元（info@naturekobe.com）の設定がまだ終わっていません');
  const payload = { key: key, to: to, subject: subject, body: body };
  if (pdf) { payload.pdf = Utilities.base64Encode(pdf.getBytes()); payload.pdfName = pdf.getName(); }
  const res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'text/plain;charset=utf-8', payload: JSON.stringify(payload), muteHttpExceptions: true, followRedirects: true });
  let j = null;
  try { j = JSON.parse(res.getContentText()); } catch (e) { throw new Error('送信に失敗しました（中継の応答が不正です）'); }
  if (!j.ok) throw new Error('送信に失敗しました：' + (j.error || ''));
}
function mailTarget_(r, isAdmin) {
  const o = orders_().find(x => x.id === r.id);
  if (!o) throw new Error('注文が見つかりません');
  if (r.kind !== 'ship' && !isAdmin) throw new Error('この操作はできません');
  if (r.kind === 'ship' && !isAdmin && o.status !== '発送済') throw new Error('先に発送日と伝票番号を保存してください');
  if (r.kind === 'receipt' && !o.paid) throw new Error('まだ入金済みになっていません');
  if (r.kind === 'ship' && (!o.shipDate || !o.tracking)) throw new Error('発送日と伝票番号が入っていません');
  return o;
}
function mailPreview_(r, isAdmin) {
  const o = mailTarget_(r, isAdmin);
  const m = buildMail_(o, r.kind);
  return { ok: true, to: o.email || '', hasEmail: !!o.email, subject: m.subject, body: m.body, attach: r.kind === 'receipt' ? '領収書_' + o.id + '.pdf' : '' };
}
function sendMail_(r, isAdmin) {
  const o = mailTarget_(r, isAdmin);
  if (!o.email) throw new Error('メールアドレスが登録されていません');
  const m = buildMail_(o, r.kind);
  relaySend_(o.email, m.subject, m.body, r.kind === 'receipt' ? receiptPdf_(o) : null);
  o.mail = o.mail || {};
  o.mail[r.kind] = true;
  o.mailLog = (o.mailLog || []).concat([{ kind: r.kind, at: now_(), to: o.email }]);
  writeOrder_(o);
  return { ok: true, order: isAdmin ? o : shopView_(o) };
}
/* 最初に一度だけエディタで実行して「外部への接続」を許可する */
function authorizeMail() { UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true }); Logger.log('OK'); }

/* ---------- 初回セットアップ ---------- */
function setup() {
  const ss = book_();
  const P = PropertiesService.getScriptProperties();
  const mk = (name, head) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(head);
    sh.getRange(1, 1, sh.getMaxRows(), head.length).setNumberFormat('@');
    sh.setFrozenRows(1);
    return sh;
  };
  const so = mk(SH.orders, ORDER_HEAD);
  const se = mk(SH.events, ['ID', '名前', '日付', '会場', '受付中']);
  const sp = mk(SH.products, ['ID', '区分', '商品名', '単価(税抜/kg)', '価格変動(〜)', '並び順', '販売中']);
  const sset = mk(SH.settings, ['項目', '値']);
  ['シート1', 'Sheet1'].forEach(n => { const s = ss.getSheetByName(n); if (s && ss.getSheets().length > 1 && s.getLastRow() === 0) ss.deleteSheet(s); });

  if (sp.getLastRow() < 2) {
    const rows = SEED.products.map(p => [p.id, p.cat, p.name, p.price, p.variable ? 'TRUE' : 'FALSE', p.sort, 'TRUE']);
    sp.getRange(2, 1, rows.length, 7).setValues(rows);
  }
  if (sset.getLastRow() < 2) {
    const rows = Object.keys(SEED.settings).map(k => [k, SEED.settings[k]]);
    sset.getRange(2, 1, rows.length, 2).setValues(rows);
  }
  if (se.getLastRow() < 2) {
    const e = SEED.event;
    se.appendRow([e.id, e.name, e.date, e.venue, e.open ? 'TRUE' : 'FALSE']);
  }
  if (so.getLastRow() < 2) {
    const prods = products_();
    const S = settings_();
    const all = [];
    SEED.orders.forEach(s => {
      const o = Object.assign({
        eventId: SEED.event.id, createdAt: '2026-09-27 00:00', source: '紙の注文票', email: '', zip: '', shipTo: { on: false },
        receipt: false, receiptTo: '', payment: '', note: '', extras: [], status: '受付', paid: false, paidAt: '',
        mail: { confirm: false, ship: false }, memo: '', checks: ''
      }, JSON.parse(JSON.stringify(s)));
      o.id = newId_(all, '260927');
      compute_(o, S, prods, false);
      o.extras = s.extras || [];
      compute_(o, S, prods, true);
      all.push(o);
      so.appendRow(orderRow_(o));
    });
  }
  const gen = () => Math.random().toString(36).slice(2, 8);
  if (!P.getProperty('ADMIN_PASS')) P.setProperty('ADMIN_PASS', 'nk-' + gen());
  if (!P.getProperty('SHOP_PASS')) P.setProperty('SHOP_PASS', 'shop-' + gen());
  Logger.log('管理用パスワード（栗原さん・牛尾さん）: ' + P.getProperty('ADMIN_PASS'));
  Logger.log('精肉店用パスワード: ' + P.getProperty('SHOP_PASS'));
}

/* ---------- 9/27 奈良販売会の紙の注文票（13件）と商品の初期値 ---------- */
const SEED = /*SEED*/null;


/* ---------- ネット販売：振込期限の督促（期限の翌日に1回だけ） ---------- */
function ensureNetEvent_() {
  const sh = sheet_(SH.events);
  if (sh.getRange(1, 8).getValue() === '') sh.getRange(1, 6, 1, 3).setValues([['販売店', '上乗せ率(%)', '購入特典']]);
  if (events_().some(e => e.id === NET_EVENT_ID)) return;
  sh.appendRow([NET_EVENT_ID, 'ネット販売', '', '', 'TRUE', '', 0, '']);
}
function addDays_(ymd, n) {
  const t = new Date(ymd + 'T00:00:00+09:00').getTime() + n * 864e5;
  return Utilities.formatDate(new Date(t), TZ, 'yyyy-MM-dd');
}
function installReminderTrigger_() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'remindUnpaidNet_') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('remindUnpaidNet_').timeBased().everyDays(1).atHour(9).create();
}
function remindUnpaidNet_() {
  const S = settings_();
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  orders_().forEach(o => {
    if (!isNet_(o.eventId) || o.paid || o.status === 'キャンセル' || !o.email) return;
    if (o.mail && o.mail.remind) return;
    const ymd = String(o.createdAt || '').slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(ymd)) return;
    const due = addDays_(ymd, NET_PAY_DAYS);
    if (today <= due) return; // 期限当日までは送らない。期限の翌日から送る
    const subject = '【' + (S.sellerName || 'NATURE KOBE') + '】お振込みのご確認（注文番号 ' + o.id + '）';
    const body = o.name + ' 様\n\nこのたびはご注文いただき、ありがとうございます。\n'
      + 'お振込みの期限（' + jdate_(due) + '）を過ぎてもご入金を確認できておりません。\n'
      + 'お手数ですが、下記の口座へお振込みをお願いいたします。\n\n'
      + '■ご請求額：' + yen_(o.total) + '（税込）\n■お振込先：\n' + (S.bankInfo || '（別途ご案内いたします）') + '\n\n'
      + 'すでにお振込みいただいている場合は、行き違いとなりますのでこのメールはお見過ごしください。\n\n' + sig_(S);
    try {
      relaySend_(o.email, subject, body, null);
      o.mail = o.mail || {};
      o.mail.remind = true;
      o.mailLog = (o.mailLog || []).concat([{ kind: 'remind', at: now_(), to: o.email }]);
      writeOrder_(o);
    } catch (e) { Logger.log('督促メール失敗 ' + o.id + ': ' + e.message); }
  });
}


/* ---------- LINE通知（Messaging API・グループへ送信） ---------- */
/* LINE公式アカウントのWebhook URLにこのアプリのURLを設定し、グループに招待すると、グループIDを自動で保存します */
function lineWebhook_(req) {
  req.events.forEach(ev => {
    if (ev.type === 'join' && ev.source && ev.source.type === 'group') {
      PropertiesService.getScriptProperties().setProperty('LINE_GROUP_ID', ev.source.groupId);
    }
  });
  return out_({ ok: true });
}
function notifyLine_(o, ev) {
  const P = PropertiesService.getScriptProperties();
  const token = P.getProperty('LINE_TOKEN'), to = P.getProperty('LINE_GROUP_ID');
  if (!token || !to) return; // 未設定なら何もしない（注文は通常どおり受付）
  const items = (o.items || []).map(i => '・' + i.name + ' ' + gl_(i.g) + '×' + i.count).join('\n');
  const text = '【新規注文】' + (ev && ev.name ? ev.name : '') + '\n'
    + '注文番号：' + o.id + '\n'
    + 'お名前：' + o.name + ' 様\n'
    + items + '\n'
    + '合計（税込）：' + yen_(o.total) + '\n'
    + 'お支払い：' + (o.payment || '') + '\n'
    + '※入金確認後、ステータスを更新してください';
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ to: to, messages: [{ type: 'text', text: text }] })
  });
}

function testLine_() {
  const P = PropertiesService.getScriptProperties();
  const token = P.getProperty('LINE_TOKEN'), to = P.getProperty('LINE_GROUP_ID');
  if (!token || !to) return { ok: true, sent: false, reason: 'トークンまたはグループ未登録です' };
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ to: to, messages: [{ type: 'text', text: 'テストです' }] })
  });
  const code = res.getResponseCode();
  return { ok: true, sent: code === 200, reason: code === 200 ? '' : 'LINE応答コード ' + code };
}

// 注文が入った時に、お客さんへ受付メール（振込の方は振込先つきの案内が別で届くので送らない）と、お店への通知メールを自動で送る
function sendOrderMails_(o, ev) {
  const S = settings_();
  const seller = S.sellerName || 'NATURE KOBE';
  const items = (o.items || []).map(i => '・' + i.name + ' ' + gl_(i.g) + '×' + i.count).concat((o.extras || []).map(x => '・' + x.name));
  if (o.email && o.payment !== '振込') {
    try {
      const subject = '【' + seller + '】ご注文を受け付けました（注文番号 ' + o.id + '）';
      const body = o.name + ' 様\n\nこのたびはご注文いただき、誠にありがとうございます。\n以下の内容で受け付けました。\n\n'
        + '■注文番号：' + o.id + '\n■ご請求額：' + yen_(o.total) + '（税込）\n■お支払い方法：' + payLabel_(o.payment) + '\n\n'
        + '■ご注文内容\n' + items.join('\n') + '\n\n'
        + '■お届け先\n' + (o.shipTo && o.shipTo.on ? '〒' + (o.shipTo.zip || '') + ' ' + (o.shipTo.address || '') + '\n' + (o.shipTo.name || '') + ' 様' : '〒' + (o.zip || '') + ' ' + (o.address || '') + '\n' + o.name + ' 様') + '\n\n'
        + (S.deliveryNote ? S.deliveryNote + '\n\n' : '')
        + 'ご不明な点がございましたら、本メールにご返信ください。\n今後ともよろしくお願いいたします。\n\n' + sig_(S);
      relaySend_(o.email, subject, body, null);
      o.mail = o.mail || {};
      o.mail.confirm = true;
      o.mailLog = (o.mailLog || []).concat([{ kind: 'order', at: now_(), to: o.email }]);
    } catch (err) { Logger.log('受付メール失敗: ' + err.message); }
  }
  try {
    const notify = S.sellerEmail || 'info@naturekobe.com';
    const subject = '【新規注文】' + (ev && ev.name ? ev.name + ' ' : '') + '注文番号 ' + o.id;
    const body = '新しい注文が入りました。\n\n'
      + '■注文番号：' + o.id + '\n■お名前：' + o.name + ' 様\n■メール：' + (o.email || '') + '\n■電話：' + (o.phone || '') + '\n'
      + '■お届け先：〒' + (o.zip || '') + ' ' + (o.address || '') + '\n■ご請求額：' + yen_(o.total) + '（税込）\n■お支払い方法：' + payLabel_(o.payment) + '\n\n'
      + '■ご注文内容\n' + items.join('\n') + '\n';
    relaySend_(notify, subject, body, null);
  } catch (err) { Logger.log('通知メール失敗: ' + err.message); }
  writeOrder_(o);
}

// 銀行振込の注文に、振込先つきの確認メールを自動で送る（注文の受付時）
function sendBankMail_(o, ev) {
  const S = settings_();
  if (!o.email || !S.bankInfo) { Logger.log('振込案内メールを送れません（メールアドレスまたは振込先が未設定）: ' + o.id); return; }
  const ymd = String(o.createdAt || '').slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ymd)) return;
  const net = isNet_(ev.id);
  const deadlineText = net
    ? '■お振込期限：ご注文から7日以内（' + jdate_(addDays_(ymd, 7)) + 'まで）\n'
    : '■お振込期限：ご注文から3日以内（' + jdate_(addDays_(ymd, 3)) + 'まで）\n'
      + '　※3日を過ぎても、7日以内（' + jdate_(addDays_(ymd, 7)) + 'まで）にご入金いただければ受付いたします。\n'
      + '　※7日を過ぎた場合は、キャンセル扱いとさせていただきます。\n';
  const subject = '【' + (S.sellerName || 'NATURE KOBE') + '】ご注文の受付とお振込みのご案内（注文番号 ' + o.id + '）';
  const body = o.name + ' 様\n\nこのたびはご注文いただき、ありがとうございます。\n'
    + '下記の口座へお振込みをお願いいたします。\n\n'
    + '■ご注文番号：' + o.id + '\n'
    + '■ご請求額：' + yen_(o.total) + '（税込）\n'
    + '■お振込先：\n' + S.bankInfo + '\n'
    + deadlineText + '\n'
    + 'ご入金を確認次第、発送の準備をいたします。\n\n' + sig_(S);
  relaySend_(o.email, subject, body, null);
  o.mail = o.mail || {};
  o.mail.confirm = true;
  o.mailLog = (o.mailLog || []).concat([{ kind: 'bank', at: now_(), to: o.email }]);
  writeOrder_(o);
}
