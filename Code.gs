// ============================================================
// 出品登録（撮影班 → Googleドライブ → 出品一覧）
// スプレッドシートの「拡張機能 → Apps Script」に貼り付けて使う
// ============================================================

// ---- 設定（ここだけ好きに変えてOK） ----
const PIN = '1105';                 // 撮影画面の合言葉
const PHOTOGRAPHERS = ['福井', '仲田', '岩崎', '佐塚']; // 撮影する人の名前に書き換える
const STAFF = ['①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩']; // 営業担当の番号
const CONDITIONS = ['新品', '未使用', '未組立', '中古', '現状品', '動作確認のみ', '動作未確認', 'ジャンク'];
const SHIPPING = ['通常', '家財便', '直接引取', '家財便・直接引取']; // 配送方法
const APPEARANCES = ['綺麗', '普通', '悪い'];              // 外観（見た目）
const NO_APPEARANCE = ['新品', '未使用', '未組立'];         // この状態の時は外観を選ばない
const MAX_PHOTOS = 10;
const SHEET_NAME = '出品一覧';
const ROOT_FOLDER_NAME = 'ヤフオク出品写真';

// ---- ここから下は触らなくて大丈夫 ----
const TZ = 'Asia/Tokyo';
const HEAD = ['発番日時', '月', 'Y番号', '進捗', 'GPT用', 'GPT出力', 'タイトル', '説明HTML', 'カテゴリ', 'CSV出力', '写真フォルダ', '撮影者', '仕入れ担当', '商品の状態', '外観', '配送方法', 'サイズ', '重さ', 'メモ', '写真枚数', '受付日時'];
const C = { at: 1, month: 2, y: 3, status: 4, gpt: 5, gptOut: 6, title: 7, html: 8, category: 9, exported: 10, folder: 11, photographer: 12, staff: 13, condition: 14, appearance: 15, shipping: 16, size: 17, weight: 18, memo: 19, count: 20, received: 21 };
const CATEGORY_SHEET = 'カテゴリ一覧';
const DICT_SHEET = 'カテゴリ辞書';
const AUTO_CAT_BG = '#fff2cc';  // 自動で入れたカテゴリ候補の色
const PARTS_SHEET = 'テンプレート部品';
const TITLE_MAX = 65;
// 前の版の列の並び。setup で今の形に自動で直す
const HEAD_V1 = ['発番日時', '月', 'Y番号', '進捗', '撮影者', '仕入れ担当', '商品の状態', 'サイズ', '重さ', 'メモ', '写真枚数', '写真フォルダ', '受付日時'];
const HEAD_V2 = ['発番日時', '月', '写真フォルダ', 'Y番号', '進捗', '撮影者', '仕入れ担当', '商品の状態', 'サイズ', '重さ', 'メモ', '写真枚数', '受付日時'];
const HEAD_V3 = ['発番日時', '月', '写真フォルダ', 'Y番号', '進捗', '撮影者', '仕入れ担当', '商品の状態', '配送方法', 'サイズ', '重さ', 'メモ', '写真枚数', '受付日時'];
const HEAD_V4 = ['発番日時', '月', '写真フォルダ', 'Y番号', '進捗', '撮影者', '仕入れ担当', '商品の状態', '外観', '配送方法', 'サイズ', '重さ', 'メモ', '写真枚数', '受付日時'];
const HEAD_V5 = ['発番日時', '月', '写真フォルダ', 'GPT用', 'Y番号', '進捗', '撮影者', '仕入れ担当', '商品の状態', '外観', '配送方法', 'サイズ', '重さ', 'メモ', '写真枚数', '受付日時'];
const HEAD_V7 = ['発番日時', '月', '写真フォルダ', 'GPT用', 'GPT出力', 'タイトル', '説明HTML', 'カテゴリ', 'CSV出力', 'Y番号', '進捗', '撮影者', '仕入れ担当', '商品の状態', '外観', '配送方法', 'サイズ', '重さ', 'メモ', '写真枚数', '受付日時'];
const HEAD_V6 = ['発番日時', '月', '写真フォルダ', 'GPT用', 'GPT出力', 'タイトル', '説明HTML', 'Y番号', '進捗', '撮影者', '仕入れ担当', '商品の状態', '外観', '配送方法', 'サイズ', '重さ', 'メモ', '写真枚数', '受付日時'];
const STATUS = ['撮影中', '出品待ち', '作成中', '出品済', '取消'];

// 最初に1回だけ実行：シートと写真フォルダを用意する
function setup() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  const headNow = () => sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].join() : '';
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEAD);
  }
  if (headNow() === HEAD_V1.join()) {
    // 写真フォルダ列（L列）を月の右（C列）へ移動。入っているデータもそのまま移る
    sh.moveColumns(sh.getRange('L:L'), 3);
  }
  if (headNow() === HEAD_V2.join()) {
    // 商品の状態の右に「配送方法」列を足す
    sh.insertColumnAfter(8);
    sh.getRange(1, 9).setValue('配送方法');
    sh.getRange(2, 9, sh.getMaxRows() - 1, 1).clearDataValidations();
  }
  if (headNow() === HEAD_V3.join()) {
    // 商品の状態の右に「外観」列を足す
    sh.insertColumnAfter(8);
    sh.getRange(1, 9).setValue('外観');
    sh.getRange(2, 9, sh.getMaxRows() - 1, 1).clearDataValidations();
  }
  if (headNow() === HEAD_V4.join()) {
    // 写真フォルダの右に「GPT用」列を足す
    sh.insertColumnAfter(3);
    sh.getRange(1, 4).setValue('GPT用');
    sh.getRange(2, 4, sh.getMaxRows() - 1, 1).clearDataValidations().clearFormat();
  }
  if (headNow() === HEAD_V5.join()) {
    // GPT用の右に「GPT出力」「タイトル」「説明HTML」列を足す
    sh.insertColumnsAfter(4, 3);
    sh.getRange(1, 5, 1, 3).setValues([['GPT出力', 'タイトル', '説明HTML']]);
    sh.getRange(2, 5, sh.getMaxRows() - 1, 3).clearDataValidations().clearFormat();
  }
  if (headNow() === HEAD_V6.join()) {
    // 説明HTMLの右に「カテゴリ」「CSV出力」列を足す
    sh.insertColumnsAfter(7, 2);
    sh.getRange(1, 8, 1, 2).setValues([['カテゴリ', 'CSV出力']]);
    sh.getRange(2, 8, sh.getMaxRows() - 1, 2).clearDataValidations().clearFormat();
  }
  if (headNow() === HEAD_V7.join()) {
    // 並べ替え：Y番号・進捗を月の右へ、写真フォルダをCSV出力の右へ（中のデータも一緒に動く）
    sh.moveColumns(sh.getRange('J:K'), 3);
    sh.moveColumns(sh.getRange('E:E'), 12);
  }
  if (headNow() !== HEAD.join()) {
    throw new Error(`「${SHEET_NAME}」シートの列が古い形です。シートを削除してから、もう一度 setup を実行してください`);
  }
  sh.setFrozenRows(1);
  // 見出しはオレンジ地に白の太字
  sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold').setBackground('#ff6d01').setFontColor('#ffffff').setHorizontalAlignment('center');
  sh.setRowHeight(1, 32);
  const n = sh.getMaxRows() - 1;
  // プルダウン。すでに設定済みの列は触らない（手で付けたチップの色を消さないため）
  const list = (col, vals, strict) => {
    if (sh.getRange(2, col).getDataValidation()) return;
    sh.getRange(2, col, n, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(vals, true).setAllowInvalid(!strict).build());
  };
  list(C.status, STATUS, true);
  list(C.photographer, PHOTOGRAPHERS, false);
  list(C.staff, STAFF, false);
  list(C.condition, CONDITIONS, false);
  list(C.appearance, APPEARANCES, false);
  list(C.shipping, SHIPPING, false);
  setupCategories_();
  setupDict_();
  if (!sh.getRange(2, C.category).getDataValidation()) {
    const cat = SpreadsheetApp.getActive().getSheetByName(CATEGORY_SHEET).getRange('A2:A');
    sh.getRange(2, C.category, n, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInRange(cat, true).setAllowInvalid(true).build());
  }

  // 色分け：進捗のセルを 撮影中＝黄／出品待ち＝緑／作成中＝青 に。
  // 出品済は行全体を濃いグレーに白文字、取消は薄い文字に取り消し線
  const all = sh.getRange(2, 1, n, HEAD.length);
  const cell = sh.getRange(2, C.status, n, 1);
  const col = String.fromCharCode(64 + C.status); // 進捗の列の文字（E）
  const when = (text, range) => SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=$${col}2="${text}"`).setRanges([range]);
  sh.setConditionalFormatRules([
    when('出品済', all).setBackground('#5f6368').setFontColor('#ffffff').build(),
    when('取消', all).setFontColor('#bdbdbd').setStrikethrough(true).build(),
    when('撮影中', cell).setBackground('#ffe599').setFontColor('#7f6000').setBold(true).build(),
    when('出品待ち', cell).setBackground('#b6d7a8').setFontColor('#274e13').setBold(true).build(),
    when('作成中', cell).setBackground('#9fc5e8').setFontColor('#073763').setBold(true).build(),
  ]);
  sh.setColumnWidth(C.memo, 300);
  sh.setColumnWidth(C.folder, 70);
  sh.setColumnWidth(C.gpt, 90);
  sh.getRange(2, C.gpt, n, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setFontColor('#9b9a97');
  sh.setColumnWidth(C.gptOut, 110);
  sh.setColumnWidth(C.title, 220);
  sh.setColumnWidth(C.html, 90);
  sh.setColumnWidth(C.category, 160);
  sh.setColumnWidth(C.exported, 110);
  sh.getRange(2, C.gptOut, n, 3).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sh.getRange(2, C.html, n, 1).setFontColor('#9b9a97');
  setupParts_();
  // すでに受付済みの行にもGPT用の文を入れる
  rows_().forEach(r => { if (r.status && r.status !== '撮影中') sh.getRange(r.row, C.gpt).setFormula(gptFormula_(r.row)); });
  rootFolder_();
}

// ============================================================
// 説明HTMLの組み立て
//   GPT出力の列に「キーワード｜外観の追記｜動作の追記」の1行を貼ると、
//   タイトルと説明HTMLが自動でできる
// ============================================================

// シートに何か入力された時に自動で動く
function onEdit(e) {
  const r = e.range, sh = r.getSheet();
  if (sh.getName() !== SHEET_NAME || r.getLastRow() < 2) return;
  // カテゴリを人が選び直したら「自動候補」の印を消す
  if (r.getColumn() <= C.category && r.getLastColumn() >= C.category) {
    sh.getRange(Math.max(2, r.getRow()), C.category, r.getLastRow() - Math.max(2, r.getRow()) + 1, 1).setBackground(null).clearNote();
  }
  const watch = [C.gptOut, C.staff, C.condition, C.appearance, C.shipping, C.size, C.weight];
  let hit = false;
  for (let c = r.getColumn(); c <= r.getLastColumn(); c++) if (watch.indexOf(c) >= 0) hit = true;
  if (!hit) return;
  for (let row = Math.max(2, r.getRow()); row <= r.getLastRow(); row++) buildListing_(sh, row);
}

// メニュー「出品ツール」
function onOpen() {
  SpreadsheetApp.getUi().createMenu('出品ツール')
    .addItem('選んだ行のタイトル・HTMLを作り直す', 'rebuildSelected')
    .addSeparator()
    .addItem('オークタウン用zipを作る', 'makeAuctownZip')
    .addToUi();
}
function rebuildSelected() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== SHEET_NAME) { SpreadsheetApp.getUi().alert(`「${SHEET_NAME}」シートで行を選んでください`); return; }
  const r = sh.getActiveRange();
  for (let row = Math.max(2, r.getRow()); row <= r.getLastRow(); row++) buildListing_(sh, row);
}

function buildListing_(sh, row) {
  const v = sh.getRange(row, 1, 1, HEAD.length).getValues()[0];
  const get = c => String(v[c - 1] == null ? '' : v[c - 1]).trim();
  const out = get(C.gptOut);
  const titleCell = sh.getRange(row, C.title);
  if (!out) { sh.getRange(row, C.title, 1, 2).clearContent(); titleCell.setNote('').setFontColor(null); return; }

  const f = out.replace(/^[`\s]+|[`\s]+$/g, '').split(/[｜|]/).map(s => s.trim());
  const none = s => !s || /^(-|ー|－|なし|無し)$/.test(s);
  const keywords = (f[0] || '').replace(/^1円[～〜~].*?】\s*/, '');
  const gaiAdd = none(f[1]) ? '' : f[1];
  const noAction = /^動作なし$/.test(f[2] || '');
  const douAdd = none(f[2]) || noAction ? '' : f[2];

  // タイトル：先頭の「1円～④Y3【中古】」はシートの情報から作る
  const condition = get(C.condition);
  const title = `1円～${get(C.staff)}${get(C.y)}【${condition}】${keywords}`;
  titleCell.setValue(title);
  if (title.length > TITLE_MAX) titleCell.setNote(`${title.length}文字です。${TITLE_MAX}文字以内に縮めてください`).setFontColor('#e03e3e');
  else titleCell.setNote('').setFontColor(null);

  // 説明HTML
  const P = parts_();
  const app = NO_APPEARANCE.indexOf(condition) >= 0 ? condition : (get(C.appearance) || '普通');
  const size = get(C.size), weight = get(C.weight);
  let h = P['枠_前'] + '<br>【外観】<br>' + (P['外観_' + app] || '') + (gaiAdd ? '<br>' + gaiAdd : '') + '<br><br>';
  if (!noAction) h += '【動作系】<br>' + (P['動作_' + condition] || '') + (douAdd ? '<br>' + douAdd : '') + '<br><br>';
  if (size || weight) h += '【サイズ・重さ】<br>' + (size ? 'サイズ：' + size + '<br>' : '') + (weight ? '重さ：' + weight + '<br>' : '') + '<br>';
  h += P['付属品'] + (app === '綺麗' ? '<br>' + P['美品'] + '<br>' : '') + '<br>' + P['商品注意事項'] + P['枠_中'];
  const ship = SHIPPING.indexOf(get(C.shipping)) >= 0 ? get(C.shipping) : SHIPPING[0];
  h += String(P['発送_' + ship] || '').replace(/\{\{3辺合計\}\}/g, sum3_(size)) + P['枠_後'];
  sh.getRange(row, C.html).setValue(h.replace(/[\r\n]+/g, ''));

  // カテゴリが空なら、過去の出品（カテゴリ辞書）から似た商品を探して候補を入れる
  if (!get(C.category)) {
    const hit = suggestCategory_(keywords);
    if (hit) {
      sh.getRange(row, C.category).setValue(hit.label).setBackground(AUTO_CAT_BG)
        .setNote(`自動候補：似ている過去の出品\n${hit.example}\n合っていればそのままでOK。違えば選び直してください`);
    }
  }
}

// ---- カテゴリの自動候補 ----
const CAT_STOP = ['まとめ', 'まとめ売り', '大量', 'セット', '中古', '未使用', '新品', 'ジャンク', '現状品', '当時物', '動作未確認', '動作確認済', '通電確認済', '美品', 'レトロ', '昭和レトロ', '本体のみ', '付属品', '箱付', '箱なし', '希少', 'レア'];
function catTokens_(title) {
  const t = String(title || '').replace(/^.*?】/, '')
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xfee0)).toLowerCase()
    .replace(/[【】（）()［］\[\]「」『』・,，、。!！?？＋+/／]/g, ' ');
  const seen = {};
  return t.split(/[\s\u3000]+/).filter(w => w.length >= 2 && CAT_STOP.indexOf(w) < 0 && !/^\d+$/.test(w) && !seen[w] && (seen[w] = true));
}
function suggestCategory_(keywords) {
  const sh = SpreadsheetApp.getActive().getSheetByName(DICT_SHEET);
  if (!sh || sh.getLastRow() < 2) return null;
  const rows = sh.getRange(2, 3, sh.getLastRow() - 1, 2).getValues().filter(r => r[0] && r[1]);
  const mine = catTokens_(keywords);
  if (!mine.length) return null;
  // 珍しい言葉ほど重く数える（「バンダイ」より「ボウケンジャー」の一致を重視）
  const docs = rows.map(r => catTokens_(r[0]));
  const df = {};
  docs.forEach(ws => ws.forEach(w => { df[w] = (df[w] || 0) + 1; }));
  const N = docs.length;
  let best = null;
  docs.forEach((ws, i) => {
    const shared = mine.filter(w => ws.indexOf(w) >= 0);
    if (!shared.length) return;
    const score = shared.reduce((a, w) => a + Math.log(1 + N / df[w]), 0);
    if (!best || score > best.score) best = { score, shared, i };
  });
  // 1語だけの一致なら、その言葉が辞書の中で珍しい時だけ採用する
  const need = Math.min(3, 0.9 * Math.log(1 + N));
  if (!best || (best.shared.length < 2 && best.score < need)) return null;
  const id = String(rows[best.i][1]).replace(/\D/g, '');
  const names = categoryMap_();
  const label = Object.keys(names).find(k => names[k] === id) || id;
  return { label, example: String(rows[best.i][0]) };
}

// zipにした商品を、タイトルとカテゴリIDの組として辞書に覚えさせる
function learnCategories_(items) {
  const sh = setupDict_();
  const when = now_();
  const add = items.map(t => [when, t.y, t.get(C.title), "'" + t.catId]);
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 4).setValues(add);
}
function setupDict_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DICT_SHEET);
  if (sh) return sh;
  sh = ss.insertSheet(DICT_SHEET);
  const rows = [['覚えた日', 'Y番号', 'タイトル', 'カテゴリID']].concat(A_DICT_SEED.map(r => [r[0], r[1], r[2], "'" + r[3]]));
  sh.getRange(1, 1, rows.length, 4).setValues(rows);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#ff6d01').setFontColor('#ffffff');
  sh.setColumnWidth(1, 120); sh.setColumnWidth(2, 70); sh.setColumnWidth(3, 520); sh.setColumnWidth(4, 110);
  return sh;
}

// 「30×20×15cm」→ 65（3つ数字がなければ ◯◯ のまま）
function sum3_(size) {
  const n = String(size || '').replace(/[０-９．]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xfee0)).match(/\d+(\.\d+)?/g);
  if (!n || n.length < 3) return '◯◯';
  return String(Math.round(n.slice(0, 3).reduce((a, b) => a + Number(b), 0)));
}

// テンプレート部品シートを読む（なければ初期値）
function parts_() {
  const P = Object.assign({}, DEFAULT_PARTS);
  const sh = SpreadsheetApp.getActive().getSheetByName(PARTS_SHEET);
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(([k, val]) => { if (k) P[String(k).trim()] = String(val); });
  }
  return P;
}

// テンプレート部品シートを作る（すでにあれば触らない）
function setupParts_() {
  const ss = SpreadsheetApp.getActive();
  const old = ss.getSheetByName(PARTS_SHEET);
  if (old) {
    // 新しく増えた部品だけ下に足す（書き換えた中身は触らない）
    const have = old.getLastRow() > 1 ? old.getRange(2, 1, old.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim()) : [];
    const add = Object.keys(DEFAULT_PARTS).filter(k => have.indexOf(k) < 0).map(k => [k, DEFAULT_PARTS[k]]);
    if (add.length) old.getRange(old.getLastRow() + 1, 1, add.length, 2).setValues(add);
    return;
  }
  const sh = ss.insertSheet(PARTS_SHEET);
  const rows = [['部品の名前（変えない）', '中身（ここを書き換えると、次に作るHTMLから反映される）']]
    .concat(Object.keys(DEFAULT_PARTS).map(k => [k, DEFAULT_PARTS[k]]));
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#ff6d01').setFontColor('#ffffff');
  sh.setColumnWidth(1, 180); sh.setColumnWidth(2, 700);
  sh.getRange(2, 2, rows.length - 1, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
}

// ============================================================
// オークタウン一括出品用のzip（CSV＋写真）を作る
//   対象：進捗が「作成中」で、タイトル・説明HTML・カテゴリがあり、CSV出力が空の行
// ============================================================
const AUCTOWN_FOLDER_NAME = 'オークタウン一括出品';
const ZIP_MAX_BYTES = 35 * 1024 * 1024;  // 1つのzipの上限（Apps Scriptの制限より少し小さく）

function makeAuctownZip() {
  const ui = SpreadsheetApp.getUi();
  const sh = sheet_();
  const P = parts_();

  // 終了はその週の日曜22時：開催期間＝次の日曜までの日数（日曜なら翌週の日曜で7日）
  const dow = Number(Utilities.formatDate(new Date(), TZ, 'u')) % 7; // 0=日曜
  const days = dow === 0 ? 7 : 7 - dow;
  if (days < 2) { ui.alert('土曜日は、日曜終了にすると開催期間が1日になってしまうため作れません（最短2日）。日曜か月曜に作ってください。'); return; }
  const endHour = String(P['オークタウン_終了時間'] || '22').replace(/\D/g, '') || '22';

  const n = sh.getLastRow() - 1;
  if (n <= 0) { ui.alert('行がありません'); return; }
  const data = sh.getRange(2, 1, n, HEAD.length).getValues();
  const cats = categoryMap_();
  const targets = [], skipped = [];
  data.forEach((v, i) => {
    const get = c => String(v[c - 1] == null ? '' : v[c - 1]).trim();
    if (get(C.status) !== '作成中' || get(C.exported)) return;
    const y = get(C.y);
    if (!get(C.title) || !get(C.html)) { skipped.push(`${y}：タイトルか説明HTMLがない`); return; }
    const catRaw = get(C.category);
    const catId = /^\d+$/.test(catRaw) ? catRaw : (cats[catRaw] || '');
    if (!catId) { skipped.push(`${y}：カテゴリが未選択か、カテゴリ一覧にない`); return; }
    targets.push({ row: i + 2, get, y, catId });
  });
  if (!targets.length) {
    ui.alert('zipにできる行がありませんでした。\n進捗を「作成中」にして、タイトル・説明HTML・カテゴリを入れてください。' + (skipped.length ? '\n\n' + skipped.join('\n') : ''));
    return;
  }
  const ok = ui.alert(`${targets.length}件をオークタウン用zipにします。\n開催期間：${days}日（日曜${endHour}時終了）\n\n作ったあと、この${targets.length}件は「出品済」になります。よろしいですか？`, ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;

  const out = child_(rootFolder_(), AUCTOWN_FOLDER_NAME, true);
  const stamp = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmm');
  const made = [];
  let batch = { rows: [], blobs: [], bytes: 0, items: [] };
  const flush = () => {
    if (!batch.items.length) return;
    const no = made.length + 1;
    const csv = [A_HEAD].concat(batch.rows).map(r => r.map(csvCell_).join(',')).join('\r\n') + '\r\n';
    const csvBlob = Utilities.newBlob('', 'text/csv', `auctown_${stamp}_${no}.csv`);
    try { csvBlob.setDataFromString(csv, 'windows-31j'); } catch (e) { csvBlob.setDataFromString(csv, 'Shift_JIS'); }
    const zip = Utilities.zip([csvBlob].concat(batch.blobs), `auctown_${stamp}_${no}.zip`);
    const file = out.createFile(zip);
    made.push({ name: file.getName(), url: file.getUrl(), items: batch.items.slice() });
    batch = { rows: [], blobs: [], bytes: 0, items: [] };
  };

  for (const t of targets) {
    const folder = findFolder_(t.get(C.month), t.y);
    const files = [];
    if (folder) { const it = folder.getFiles(); while (it.hasNext()) files.push(it.next()); }
    files.sort((a, b) => a.getName() < b.getName() ? -1 : 1);
    const prefix = t.get(C.month).replace(/\D/g, '').slice(2) + t.y; // 例：2609Y3
    const imgs = files.slice(0, 10).map((f, k) => f.getBlob().setName(`${prefix}_${String(k + 1).padStart(2, '0')}.jpg`));
    const bytes = imgs.reduce((a, b) => a + b.getBytes().length, 0);
    if (batch.items.length && batch.bytes + bytes > ZIP_MAX_BYTES) flush();
    batch.rows.push(auctownRow_(t.get, t.catId, imgs.map(b => b.getName()), days, endHour, P));
    batch.blobs.push(...imgs);
    batch.bytes += bytes;
    batch.items.push(t);
  }
  flush();

  // カテゴリ辞書に覚えさせる
  learnCategories_(made.reduce((a, m) => a.concat(m.items), []));

  // 出品済にしてCSV出力日を入れる
  const when = now_();
  made.forEach(m => m.items.forEach(t => {
    sh.getRange(t.row, C.exported).setValue(when);
    sh.getRange(t.row, C.status).setValue('出品済');
  }));

  const html = '<div style="font-family:sans-serif;font-size:14px;line-height:1.7">'
    + `<p>${targets.length}件を${made.length}個のzipにしました（開催期間${days}日・日曜${endHour}時終了）。<br>ダウンロードして、オークタウンの一括出品からアップロードしてください。</p>`
    + made.map(m => `<p><a href="${m.url}" target="_blank">${m.name}</a>（${m.items.map(t => t.y).join('・')}）</p>`).join('')
    + (skipped.length ? '<p style="color:#b00">入れなかった行：<br>' + skipped.join('<br>') + '</p>' : '')
    + '</div>';
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(460).setHeight(320), 'オークタウン用zipができました');
}

// 1商品分の行（114列）。共通の設定はもらったCSVの値をそのまま使う
function auctownRow_(get, catId, imageNames, days, endHour, P) {
  const r = Object.assign({}, A_DEFAULT);
  r['カテゴリ'] = catId;
  r['タイトル'] = get(C.title).replace(/\s*,\s*/g, ' ');
  // CSVでトラブルになりやすい「"」「,」を消す（属性の値はクォートなしでもHTMLとして有効）
  r['説明'] = get(C.html).replace(/(\w+)="([^"\s,]*)"/g, '$1=$2').replace(/"/g, '').replace(/,/g, '、');
  r['開始価格'] = '1';
  r['個数'] = '1';
  r['開催期間'] = String(days);
  r['終了時間'] = endHour;
  imageNames.forEach((name, k) => { r[`画像${k + 1}`] = name; });
  const condition = get(C.condition);
  const key = NO_APPEARANCE.indexOf(condition) >= 0 ? condition : (get(C.appearance) || '普通');
  r['商品の状態'] = P['ヤフオク状態_' + key] || 'やや傷や汚れあり';
  return A_HEAD.map(h => r[h] == null ? '' : String(r[h]));
}

function csvCell_(x) {
  const s = String(x == null ? '' : x).replace(/[\r\n]+/g, '');
  return /[",]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

// カテゴリ一覧：名前 → カテゴリID
function categoryMap_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CATEGORY_SHEET);
  const m = {};
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(([k, id]) => {
    if (k && id) m[String(k).trim()] = String(id).replace(/\D/g, '');
  });
  return m;
}
function setupCategories_() {
  const ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName(CATEGORY_SHEET)) return;
  const sh = ss.insertSheet(CATEGORY_SHEET);
  const rows = [['カテゴリの名前（出品一覧のプルダウンに出る）', 'カテゴリID']].concat(A_CATEGORIES);
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(2, 2, rows.length - 1, 1).setNumberFormat('@');
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#ff6d01').setFontColor('#ffffff');
  sh.setColumnWidth(1, 320); sh.setColumnWidth(2, 140);
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('出品登録')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // アイコン付きの入口ページの中で表示できるように
}

// ---- 画面から呼ばれる関数 ----

function getInit(pin) {
  check_(pin);
  const m = month_();
  const open = rows_().filter(r => r.month === m && r.status === '撮影中').map(r => r.y);
  return {
    photographers: PHOTOGRAPHERS, staff: STAFF, conditions: CONDITIONS, shipping: SHIPPING, maxPhotos: MAX_PHOTOS,
    appearances: APPEARANCES, noAppearance: NO_APPEARANCE,
    month: m, open: open,
  };
}

// 次のY番号を配る。2ブース同時に押しても重複しないよう順番待ちさせる
function reserve(pin) {
  check_(pin);
  return withLock_(() => {
    const sh = sheet_();
    const m = month_();
    const mine = rows_().filter(r => r.month === m);
    const reuse = mine.filter(r => r.status === '取消').sort((a, b) => a.num - b.num)[0];
    if (reuse) {
      sh.getRange(reuse.row, C.at).setValue(now_());
      sh.getRange(reuse.row, C.status).setValue('撮影中');
      return { month: m, y: reuse.y };
    }
    const max = mine.reduce((mx, r) => Math.max(mx, r.num), 0);
    const y = 'Y' + (max + 1);
    const row = new Array(HEAD.length).fill('');
    row[C.at - 1] = now_(); row[C.month - 1] = "'" + m; row[C.y - 1] = y; row[C.status - 1] = '撮影中';
    sh.appendRow(row);
    SpreadsheetApp.flush();
    return { month: m, y: y };
  });
}

function cancel(pin, month, y) {
  check_(pin);
  return withLock_(() => {
    const r = find_(month, y);
    if (!r || r.status !== '撮影中') throw new Error(`${y} は取り消せる状態ではありません`);
    sheet_().getRange(r.row, C.status).setValue('取消');
    const f = findFolder_(month, y);
    if (f) f.setTrashed(true);
    return true;
  });
}

function uploadPhoto(pin, month, y, index, base64) {
  check_(pin);
  const r = find_(month, y);
  if (!r || r.status !== '撮影中') throw new Error(`${y} は受付できる状態ではありません`);
  const folder = folder_(month, y);
  const name = `${y}_${String(index).padStart(2, '0')}.jpg`;
  const old = folder.getFilesByName(name);
  while (old.hasNext()) old.next().setTrashed(true);
  folder.createFile(Utilities.newBlob(Utilities.base64Decode(base64), 'image/jpeg', name));
  return true;
}

// 送信完了：一覧を「出品待ち」にする
// info = { photographer, staff, condition, appearance, shipping, size, weight, memo }
function finish(pin, month, y, info) {
  check_(pin);
  info = info || {};
  if (PHOTOGRAPHERS.indexOf(info.photographer) < 0) throw new Error('撮影者を選んでください');
  if (CONDITIONS.indexOf(info.condition) < 0) throw new Error('商品の状態を選んでください');
  if (STAFF.indexOf(info.staff) < 0) throw new Error('営業担当を選んでください');
  const shipping = SHIPPING.indexOf(info.shipping) >= 0 ? info.shipping : SHIPPING[0];
  const needApp = NO_APPEARANCE.indexOf(info.condition) < 0;
  if (needApp && APPEARANCES.indexOf(info.appearance) < 0) throw new Error('外観を選んでください');
  const appearance = needApp ? info.appearance : '';
  return withLock_(() => {
    const r = find_(month, y);
    if (!r || r.status !== '撮影中') throw new Error(`${y} は受付できる状態ではありません`);
    const folder = findFolder_(month, y);
    if (!folder) throw new Error(`${y} の写真が届いていません`);
    let count = 0;
    const it = folder.getFiles();
    while (it.hasNext()) { it.next(); count++; }
    const t = s => String(s || '').trim();
    const sh = sheet_();
    sh.getRange(r.row, C.photographer, 1, C.received - C.photographer + 1).setValues([[
      info.photographer, STAFF.indexOf(info.staff) >= 0 ? info.staff : '', info.condition, appearance, shipping,
      t(info.size), t(info.weight), t(info.memo), count, now_(),
    ]]);
    sh.getRange(r.row, C.status).setValue('出品待ち');
    sh.getRange(r.row, C.gpt).setFormula(gptFormula_(r.row));
    // 写真フォルダは「開く」という短いリンクにして列の幅を取らないようにする
    sh.getRange(r.row, C.folder).setRichTextValue(
      SpreadsheetApp.newRichTextValue().setText('開く').setLinkUrl(folder.getUrl()).build());
    return { y: y, count: count };
  });
}

// ---- 内部の道具 ----

// GPTに貼る1行の文章を作る式。シートで値を直すと自動で変わる
function gptFormula_(row) {
  const col = c => '$' + String.fromCharCode(64 + c) + row;
  const part = (label, c) => `"${label}："&${col(c)}`;
  const memo = `"メモ："&SUBSTITUTE(${col(C.memo)},CHAR(10)," ")`;
  return `=IF(${col(C.y)}="","",` + [
    part('Y番号', C.y), part('仕入れ担当', C.staff), part('商品の状態', C.condition), part('外観', C.appearance),
    part('配送方法', C.shipping), part('サイズ', C.size), part('重さ', C.weight), memo,
  ].join('&"／"&') + ')';
}

function check_(pin) {
  const norm = s => String(s == null ? '' : s).replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xfee0)).replace(/\s+/g, '');
  if (norm(pin) !== norm(PIN)) throw new Error('合言葉が違います');
}
function sheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('シートがありません。Apps Scriptで setup を実行してください');
  return sh;
}
function month_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM'); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm'); }
function monthText_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM') : String(v).trim(); }

function rows_() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  return sh.getRange(2, 1, n, HEAD.length).getValues().map((v, i) => {
    const y = String(v[C.y - 1]).trim();
    return { row: i + 2, month: monthText_(v[C.month - 1]), y: y, num: parseInt(y.replace(/\D/g, ''), 10) || 0, status: String(v[C.status - 1]).trim() };
  });
}
function find_(month, y) {
  return rows_().filter(r => r.month === month && r.y === y).pop() || null;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function rootFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('ROOT_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* 消されていたら作り直す */ } }
  const f = DriveApp.createFolder(ROOT_FOLDER_NAME);
  props.setProperty('ROOT_ID', f.getId());
  return f;
}
function child_(parent, name, create) {
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return create ? parent.createFolder(name) : null;
}
function findFolder_(month, y) {
  const m = child_(rootFolder_(), month, false);
  return m ? child_(m, y, false) : null;
}
function folder_(month, y) {
  const found = findFolder_(month, y);
  if (found) return found;
  return withLock_(() => child_(child_(rootFolder_(), month, true), y, true));
}

// ---- テンプレート部品の初期値（シート「テンプレート部品」で上書きできる） ----
const DEFAULT_PARTS = {
  "オークタウン_終了時間": "22",
  "ヤフオク状態_新品": "新品",
  "ヤフオク状態_未使用": "未使用に近い",
  "ヤフオク状態_未組立": "未使用に近い",
  "ヤフオク状態_綺麗": "目立った傷や汚れなし",
  "ヤフオク状態_普通": "やや傷や汚れあり",
  "ヤフオク状態_悪い": "全体的に状態が悪い",
  "枠_前": "<br><center><table cellspacing=\"0\" border=\"0\" cellpadding=\"0\" width=\"520\"><tbody><tr><td width=\"1\"></td><td width=\"1\"></td><td width=\"1\"></td><td width=\"2\"></td><td width=\"2\"></td><td width=\"506\"></td><td width=\"2\"></td><td width=\"2\"></td><td width=\"1\"></td><td width=\"1\"></td><td width=\"1\"></td></tr><tr><td height=\"1\" colspan=\"4\"></td><td height=\"1\" bgcolor=\"#339966\"></td><td rowspan=\"5\" bgcolor=\"#339966\" align=\"center\"><font color=\"#CCFFCC\" size=\"4\"><b></b></font></td><td height=\"1\" bgcolor=\"#339966\"></td><td height=\"1\" colspan=\"4\"></td></tr><tr><td height=\"1\" colspan=\"3\"></td><td height=\"1\" colspan=\"2\" bgcolor=\"#339966\"></td><td height=\"1\" colspan=\"2\" bgcolor=\"#339966\"></td><td height=\"1\" colspan=\"3\"></td></tr><tr><td height=\"1\" colspan=\"2\"></td><td height=\"1\" colspan=\"3\" bgcolor=\"#339966\"></td><td height=\"1\" colspan=\"3\" bgcolor=\"#339966\"></td><td height=\"1\" colspan=\"2\"></td></tr><tr><td height=\"2\"></td><td height=\"1\" colspan=\"4\" bgcolor=\"#339966\"></td><td height=\"1\" colspan=\"4\" bgcolor=\"#339966\"></td><td height=\"2\"></td></tr><tr><td colspan=\"5\" bgcolor=\"#339966\"><br><br></td><td colspan=\"5\" bgcolor=\"#339966\"><br><br></td></tr><tr><td colspan=\"5\" bgcolor=\"#339966\"></td><td><table width=\"100%\" cellspacing=\"1\" border=\"0\" cellpadding=\"10\" bgcolor=\"#66CC99\"><tbody><tr><td align=\"center\"><font color=\"#003333\" size=\"3\"><b>商品詳細</b></font></td></tr><tr><td align=\"left\" bgcolor=\"#FFFFCC\"><font color=\"#660000\" size=\"2\">",
  "付属品": "【付属品】<br>写真に写っているものが全てです。<br>（写真で判別しにくいものがあれば質問欄からお問い合わせください）<br>",
  "美品": "美品は個人的感想のため画像をご確認お願いします。",
  "商品注意事項": "【商品についての注意事項】<br>・買取品のため、購入時期・使用期間・修理歴等の詳細は不明です。<br>・当方は専門店ではないため、記載内容以外の細かな確認は行っておりません。<br>・商品情報はカメラアプリ等による検索結果も参考にしているため、記載内容に誤りが含まれる場合がございます。ご不明な点はご購入前にお問い合わせください。<br>・経年保管に伴う小傷、細かな汚れがある場合がございます。<br>・状態の感じ方には個人差がありますので、画像をよくご確認のうえご判断ください。<br>・画像に写っているものが全てとなります。<br>・簡易的に緩衝材を使用して発送致します。<br>・中古品、現状品にご理解のある方のみご入札ください。<br>・神経質な方、完璧な状態をお求めの方は入札をご遠慮ください。<br>・落札後はノークレーム・ノーリターンでお願いいたします。<br>",
  "枠_中": "</font></td></tr><tr><td align=\"center\"><font color=\"#003333\" size=\"3\"><b>支払詳細</b></font></td></tr><tr><td align=\"left\" bgcolor=\"#FFFFCC\"><font color=\"#660000\" size=\"2\">・Yahoo!かんたん決済<br>・落札後、2日以内にお支払い手続きをお願いいたします。<br>・やむを得ない事情により2日以内のお支払いが難しい場合は、必ず事前に取引メッセージよりご連絡ください。<br>・ご連絡がないまま期限を過ぎた場合は、落札者都合にてキャンセルさせていただく場合があります。<br></font></td></tr><tr><td align=\"center\"><font color=\"#003333\" size=\"3\"><b>発送詳細</b></font></td></tr><tr><td align=\"left\" bgcolor=\"#FFFFCC\">",
  "発送_通常": "<font color=\"#660000\" size=\"2\">・発送はクロネコヤマトの着払いで静岡から発送致します。<br>※商品の大きさ・長さによっては、ヤマト運輸以外（西濃運輸・佐川急便など）で発送する場合がございます。その場合は落札後に個別にご連絡いたします。<br><br><b>【送料について】</b><br>・送料は落札者様のご負担となり、お届け先の地域・サイズによって異なります。<br>・ヤフオクの出品システム上、商品ページに送料が表示される場合がありますが、表示されている金額が実際の送料とは限りません。<br>・送料の金額についての個別のご連絡は行っておりません。発送後に送り状番号を登録しますので、料金は直接クロネコヤマトにお問い合わせください。<br>・送料を理由とした落札後のキャンセルはお受けできません。<br>上記内容をご理解・ご了承いただいた上でご入札をお願いいたします。</font>",
  "発送_家財便": "<font color=\"#660000\" size=\"2\"><b>【発送について】</b><br>アートセッティングデリバリー（家財便）にて発送予定です。<br><br>3辺合計：約{{3辺合計}}cm<br><br>※実際の計測や配送会社の判断により、サイズ・送料が変更となる場合があります。<br>※配送不可地域等がある場合がありますのでご了承ください。<br><br><b>【送料について】</b><br>・送料は落札者様のご負担となり、サイズ・お届け先の地域によって異なります。<br>・送料を理由とした落札後のキャンセルはお受けできません。<br>上記内容をご理解・ご了承いただいた上でご入札をお願いいたします。</font>",
  "発送_直接引取": "<font color=\"#660000\" size=\"2\"><b>【お引き取りについて】</b><br>こちらの商品は、静岡県静岡市での直接引き取り限定となります。<br>発送には対応しておりませんのでご了承ください。<br>お引き取りの日時は、落札後に取引メッセージにてご相談ください。</font>",
  "発送_家財便・直接引取": "<font color=\"#660000\" size=\"2\"><b>【発送について】</b><br>アートセッティングデリバリー（家財便）にて発送予定です。<br><br>3辺合計：約{{3辺合計}}cm<br><br>※実際の計測や配送会社の判断により、サイズ・送料が変更となる場合があります。<br>※配送不可地域等がある場合がありますのでご了承ください。<br><br><b>【送料について】</b><br>・送料は落札者様のご負担となり、サイズ・お届け先の地域によって異なります。<br>・送料を理由とした落札後のキャンセルはお受けできません。<br>上記内容をご理解・ご了承いただいた上でご入札をお願いいたします。<br><br>静岡県静岡市での直接引き取りも可能です。<br>直接引き取りをご希望の場合は、落札後に取引メッセージにてお知らせください。</font>",
  "枠_後": "</td></tr><tr><td align=\"center\"><font color=\"#003333\" size=\"3\"><b>注意事項</b></font></td></tr><tr><td align=\"left\" bgcolor=\"#FFFFCC\"><font color=\"#660000\" size=\"2\">〇領収書、納品書などにつきまして<br>・ご希望の方は、落札後に取引メッセージにて個別にご連絡をお願いいたします。<br><br>〇入札・キャンセルにつきまして<br><b>・いかなる理由でも入札後のキャンセル・取り消しはお受けできません。<br>入札＝購入意思ありと判断いたしますので、十分ご確認のうえご入札ください。</b><br>・トラブル防止のため、悪い評価が多数あるお客様の入札は削除させていただく場合がございます。<br>・評価が0やマイナスの方はご入札をご遠慮ください。入札を削除させていただきます。<br></font></td></tr><tr><td align=\"center\"><font color=\"#003333\" size=\"3\"><b>コメント</b></font></td></tr><tr><td align=\"left\" bgcolor=\"#FFFFCC\"><font color=\"#660000\" size=\"2\">何か、不明な点がございましたらお答えできる範囲でご回答させていただきますのでお気軽にお問い合わせ下さいませ。宜しくお願い申し上げます。</font></td></tr></tbody></table></td><td colspan=\"5\" bgcolor=\"#339966\"></td></tr><tr><td height=\"7\" colspan=\"11\" bgcolor=\"#339966\"></td></tr><tr><td height=\"2\" colspan=\"11\"></td></tr><tr><td height=\"2\" colspan=\"11\" bgcolor=\"#339966\"></td></tr></tbody></table><br><font color=\"#FFFFFF\" size=\"1\"><br>No.212.001.005</font><br></center>",
  "外観_新品": "新品・未開封品です。",
  "外観_未使用": "未使用品です。",
  "外観_未組立": "未組立品です。",
  "外観_綺麗": "写真でお分かりのように使用感もなく<br>非常に綺麗です。（要写真確認）",
  "外観_普通": "傷汚れあります。（要写真確認）",
  "外観_悪い": "全体的に状態悪いです。（要写真確認）",
  "動作_新品": "未開封のため動作確認は行なっておりません。",
  "動作_未使用": "未使用のため動作確認は行なっておりません。<br>現状引き渡し商品として出品します。",
  "動作_未組立": "パーツの欠品確認は行なっておりません。<br>現状引き渡し商品として出品します。",
  "動作_中古": "動作確認済みです。",
  "動作_現状品": "メンテナンス等は行っておりません。<br>現状品として出品します。",
  "動作_動作確認のみ": "一部の動作のみ確認しております。<br>それ以外の動作は未確認のため、現状品として出品します。",
  "動作_動作未確認": "動作確認は行なっておりません。<br>動作未確認の現状品として出品します。",
  "動作_ジャンク": "動作未確認のジャンク品として出品します。"
};

// ---- オークタウンCSVの列（もらったCSVと同じ114列）と共通の初期値 ----
const A_HEAD = ["カテゴリ", "タイトル", "説明", "開始価格", "即決価格", "個数", "開催期間", "終了時間", "JANコード", "画像1", "画像1コメント", "画像2", "画像2コメント", "画像3", "画像3コメント", "画像4", "画像4コメント", "画像5", "画像5コメント", "画像6", "画像6コメント", "画像7", "画像7コメント", "画像8", "画像8コメント", "画像9", "画像9コメント", "画像10", "画像10コメント", "商品発送元の都道府県", "商品発送元の市区町村", "送料負担", "代金支払い", "Yahoo!かんたん決済", "かんたん取引", "商品代引", "商品の状態", "商品の状態備考", "返品の可否", "返品の可否備考", "入札者評価制限", "悪い評価の割合での制限", "入札者認証制限", "自動延長", "早期終了", "値下げ交渉", "自動再出品", "自動値下げ", "自動値下げ価格変更率", "注目のオークション", "おすすめコレクション", "送料固定", "荷物の大きさ", "荷物の重量", "ネコポス", "ネコ宅急便コンパクト", "ネコ宅急便", "ゆうパケット", "ゆうパック", "ゆうパケットポストmini", "ゆうパケットプラス", "発送までの日数", "配送方法1", "配送方法1全国一律価格", "北海道料金1", "沖縄料金1", "離島料金1", "配送方法2", "配送方法2全国一律価格", "北海道料金2", "沖縄料金2", "離島料金2", "配送方法3", "配送方法3全国一律価格", "北海道料金3", "沖縄料金3", "離島料金3", "配送方法4", "配送方法4全国一律価格", "北海道料金4", "沖縄料金4", "離島料金4", "配送方法5", "配送方法5全国一律価格", "北海道料金5", "沖縄料金5", "離島料金5", "配送方法6", "配送方法6全国一律価格", "北海道料金6", "沖縄料金6", "離島料金6", "配送方法7", "配送方法7全国一律価格", "北海道料金7", "沖縄料金7", "離島料金7", "配送方法8", "配送方法8全国一律価格", "北海道料金8", "沖縄料金8", "離島料金8", "配送方法9", "配送方法9全国一律価格", "北海道料金9", "沖縄料金9", "離島料金9", "配送方法10", "配送方法10全国一律価格", "北海道料金10", "沖縄料金10", "離島料金10", "受け取り後決済サービス", "海外発送"];
const A_DEFAULT = {
  "開始価格": "1",
  "即決価格": "",
  "個数": "1",
  "JANコード": "",
  "商品発送元の都道府県": "静岡県",
  "商品発送元の市区町村": "",
  "送料負担": "落札者",
  "代金支払い": "先払い",
  "Yahoo!かんたん決済": "はい",
  "かんたん取引": "はい",
  "商品代引": "いいえ",
  "返品の可否": "返品不可",
  "返品の可否備考": "",
  "入札者評価制限": "いいえ",
  "悪い評価の割合での制限": "いいえ",
  "入札者認証制限": "いいえ",
  "自動延長": "はい",
  "早期終了": "はい",
  "値下げ交渉": "いいえ",
  "自動再出品": "0",
  "自動値下げ": "いいえ",
  "自動値下げ価格変更率": "",
  "注目のオークション": "",
  "おすすめコレクション": "",
  "送料固定": "着払い",
  "荷物の大きさ": "",
  "荷物の重量": "",
  "ネコポス": "いいえ",
  "ネコ宅急便コンパクト": "いいえ",
  "ネコ宅急便": "いいえ",
  "ゆうパケット": "いいえ",
  "ゆうパック": "いいえ",
  "ゆうパケットポストmini": "いいえ",
  "ゆうパケットプラス": "いいえ",
  "発送までの日数": "3日～7日",
  "配送方法1": "宅急便（ヤマト運輸）",
  "配送方法1全国一律価格": "",
  "北海道料金1": "",
  "沖縄料金1": "",
  "離島料金1": "",
  "配送方法2": "",
  "配送方法2全国一律価格": "",
  "北海道料金2": "",
  "沖縄料金2": "",
  "離島料金2": "",
  "配送方法3": "",
  "配送方法3全国一律価格": "",
  "北海道料金3": "",
  "沖縄料金3": "",
  "離島料金3": "",
  "配送方法4": "",
  "配送方法4全国一律価格": "",
  "北海道料金4": "",
  "沖縄料金4": "",
  "離島料金4": "",
  "配送方法5": "",
  "配送方法5全国一律価格": "",
  "北海道料金5": "",
  "沖縄料金5": "",
  "離島料金5": "",
  "配送方法6": "",
  "配送方法6全国一律価格": "",
  "北海道料金6": "",
  "沖縄料金6": "",
  "離島料金6": "",
  "配送方法7": "",
  "配送方法7全国一律価格": "",
  "北海道料金7": "",
  "沖縄料金7": "",
  "離島料金7": "",
  "配送方法8": "",
  "配送方法8全国一律価格": "",
  "北海道料金8": "",
  "沖縄料金8": "",
  "離島料金8": "",
  "配送方法9": "",
  "配送方法9全国一律価格": "",
  "北海道料金9": "",
  "沖縄料金9": "",
  "離島料金9": "",
  "配送方法10": "",
  "配送方法10全国一律価格": "",
  "北海道料金10": "",
  "沖縄料金10": "",
  "離島料金10": "",
  "受け取り後決済サービス": "いいえ",
  "海外発送": "いいえ"
};
const A_CATEGORIES = [
  [
    "SUZUKI SX650RII 発電機 と同じカテゴリ",
    "2084304064"
  ],
  [
    "KARCHER ケルヒャー JTK Si と同じカテゴリ",
    "2084207591"
  ],
  [
    "HP ProDesk 600 G3 S と同じカテゴリ",
    "2084306954"
  ],
  [
    "XEBEX MIGHTY PT-1200 と同じカテゴリ",
    "2084261981"
  ],
  [
    "R.KLAAS SOLINGEN 鶴型ハ と同じカテゴリ",
    "2084306778"
  ],
  [
    "National Panasonic C と同じカテゴリ",
    "2084024142"
  ]
];
const A_DICT_SEED = [
  [
    "",
    "Y247",
    "1円～④Y247【ジャンク】SUZUKI SX650RII 発電機 100V 650W 60Hz エンジン不動 直接引取限定",
    "2084304064"
  ],
  [
    "",
    "Y246",
    "1円～④Y246【未使用】KARCHER ケルヒャー JTK Silent S 高圧洗浄機 静音モデル 箱付 付属品完備 通電確認済",
    "2084207591"
  ],
  [
    "",
    "Y245",
    "1円～④Y245【ジャンク】 HP ProDesk 600 G3 SFF Core i5 デスクトップPC 本体のみ　2017年頃",
    "2084306954"
  ],
  [
    "",
    "Y249",
    "1円～③Y249【現状品】XEBEX MIGHTY PT-1200A 業務用ストロボ 電源部 ジェネレーター 通電確認済",
    "2084261981"
  ],
  [
    "",
    "Y205",
    "1円～③Y205【現状品】R.KLAAS SOLINGEN 鶴型ハサミ 大小2本 指ぬき セット ドイツ ゾーリンゲン 裁縫 手芸 刺繍 工芸品",
    "2084306778"
  ],
  [
    "",
    "Y206",
    "1円～③Y206【ジャンク】National Panasonic COUGAR 2200 RF-2200 1970年代 BCL 8バンドラジオ 昭和レトロ",
    "2084024142"
  ]
];
