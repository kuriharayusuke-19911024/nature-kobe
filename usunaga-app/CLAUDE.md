# うすなが牧場 月次報告アプリ — 改善リクエストの自動対応のきまり

このフォルダは、神戸うすなが牧場（碓永さん）の月次報告アプリです。保守は栗原優介さん。
アプリの中の「💬 改善リクエスト」に書かれた要望を、Claudeの定期実行が1日3回読んで対応します。

## ファイル
- `index.html`：アプリ本体（これ1枚）。GitHub Pages で https://kuriharayusuke-19911024.github.io/nature-kobe/usunaga-app/ に公開される
- `gas/コード.gs`：データの保存先（Google Apps Script）。**自動対応では触らない**
- `../usunaga-demo/index.html`：デモ版。画面の変更は、デモ版にも同じように入れてよい（入れられない時は無理にしない）

## 自動で対応してよいこと（画面だけで済むもの）
- 文言・ラベル・説明文の変更、並び順、表の列の追加・入れ替え（すでにあるデータを表示するだけのもの）
- 入力欄の追加（新しい項目を保存して表示するだけで、計算に使わないもの）
- 見た目（文字の大きさ・色・余白）、スマホでの見やすさ
- 絞り込み・検索・並べ替えの追加
- 表示の不具合の修正（原因がはっきりしていて、直し方が1つに決まるもの）

## 自動で対応しないこと（「確認待ち」にして栗原さんに回す）
- 計算の仕方の変更（コスト・按分・粗利・在庫・目標の判定など、数字が変わるもの）
- お金・請求・会計に関わるもの
- データの削除・まとめての書き換え、データの形（保存する項目名）の変更
- 保存先（gas/）・合言葉・AI読み取りの仕組みの変更
- 外部への送信（メール・LINEなど）を増やすこと
- 要望の意味が2通り以上に取れるもの、大きな機能の追加
- 迷ったら「確認待ち」

## 改善リクエストの読み書き（PCのブラウザで、ページの中で実行）
データの保存先には、クラウドのコンテナからは届かない。栗原さんのPCの内蔵ブラウザでアプリのページを開き、javascript_tool で実行する。`PASS` は定期実行の指示に書かれた合言葉（ここには書かない）。

一覧を取る：
```js
await (async()=>{const U='https://script.google.com/macros/s/AKfycbwFChTuoFMYvY45n7mffKezzBJP34WWOaf4v9Vfaonbm34Z083W38xE6WFnUoCtvwCV/exec';
const j=await fetch(U,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({op:'get',since:0,pass:PASS})}).then(r=>r.json());
const m={};(j.docs||[]).filter(d=>d.col==='requests').forEach(d=>{if(!m[d.id]||Number(d.doc._u||0)>Number(m[d.id]._u||0))m[d.id]=d.doc});
return Object.entries(m).map(([id,d])=>({id,...d}));})()
```

状態と返事を書く（`ID`・`STATUS`・`REPLY` を置き換える。ほかの項目は変えない）：
```js
await (async()=>{const U='https://script.google.com/macros/s/AKfycbwFChTuoFMYvY45n7mffKezzBJP34WWOaf4v9Vfaonbm34Z083W38xE6WFnUoCtvwCV/exec';
const j=await fetch(U,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({op:'get',since:0,pass:PASS})}).then(r=>r.json());
let cur=null;(j.docs||[]).filter(d=>d.col==='requests'&&d.id===ID).forEach(d=>{if(!cur||Number(d.doc._u||0)>Number(cur._u||0))cur=d.doc});
if(!cur)return 'not found';
const now=new Date().toISOString();const doc={...cur,status:STATUS,reply:REPLY,repliedAt:now,updatedAt:now,_u:Math.max(Date.now(),Number(cur._u||0)+1)};
return await fetch(U,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({op:'put',docs:[{col:'requests',id:ID,doc}],pass:PASS})}).then(r=>r.json());})()
```
- 状態は「対応済み」「確認待ち」「対応しない」のどれか
- 返事は、牧場の人向けに専門用語を使わず、何をどう変えたか・どの画面で見られるか・「ページを開き直すと変わります」を書く

## 直したあとの確認（必須）
1. `node --check` で、index.html の一番最後の `<script>` の中身に文法エラーがないこと
2. Playwright（`require('playwright')`、chromium）で index.html を `file://` で開き、ページのエラー（pageerror）が出ないこと。データにはつながらないので「データを読み込んでいます」や合言葉の画面が出るのは正常
3. 変更した画面の部分を、実際の公開ページ（PCのブラウザ）で開いて目で確認する。データを書き換える操作はしない

## コミット
- 日本語1行（例：「うすなが月次報告：請求書の一覧に請求元の電話番号を表示（改善リクエスト対応）」）、末尾に `Co-Authored-By: Claude <noreply@anthropic.com>`
- `git pull --rebase` してから push。このリポジトリには他のアプリ（nature-kobe など）も入っているので、`usunaga-app/`（と `usunaga-demo/`）以外のファイルは触らない
- 合言葉やAPIキーは絶対にコミットしない（このリポジトリは公開）
