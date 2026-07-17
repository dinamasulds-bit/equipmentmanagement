/**
 * 発注管理 — Google スプレッドシート バックエンド (Apps Script)
 *
 * セットアップ手順:
 *  1. Google スプレッドシートを新規作成
 *  2. 拡張機能 → Apps Script を開く
 *  3. 既定のコードを全消しして、このファイルの中身を貼り付けて保存
 *  4. 右上「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *       - 説明: 任意
 *       - 次のユーザーとして実行: 自分
 *       - アクセスできるユーザー: 全員  ← ここ重要
 *  5. デプロイ → 表示される「ウェブアプリ URL」(.../exec) をコピー
 *  6. 発注管理アプリ(index.html)の右上「⚙ 共有設定」に貼り付け
 *
 * シートは Orders / Master の2タブが自動で作られます。
 * 履歴はスプレッドシートを直接開いても閲覧できます。
 */

var ORDERS = "Orders";
var MASTER = "Master";
var ORDER_HEADERS = ["id","保存日時","発注日","ディレクター","現場","着日","依頼者","受取者","ホテル名","郵便番号","住所","電話","宛先","あいさつ","品目","合計金額","data"];

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name, headers) {
  var s = ss_().getSheetByName(name);
  if (!s) s = ss_().insertSheet(name);
  if (s.getLastRow() === 0) s.appendRow(headers);
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function n_(v) {
  var x = Number(String(v).replace(/[^0-9.\-]/g, ""));
  return isNaN(x) ? 0 : x;
}

// ---------- 読み取り ----------
function doGet(e) {
  return json_(readAll_());
}

function readAll_() {
  var os = sheet_(ORDERS, ORDER_HEADERS);
  var ov = os.getDataRange().getValues();
  var orders = [];
  for (var i = 1; i < ov.length; i++) {
    if (!ov[i][0]) continue;
    var data = {};
    try { data = JSON.parse(ov[i][ORDER_HEADERS.length - 1]); } catch (err) {}
    orders.push({ id: String(ov[i][0]), data: data });
  }
  // 新しい順
  orders.reverse();

  var ms = sheet_(MASTER, ["name", "price", "unit"]);
  var mv = ms.getDataRange().getValues();
  var master = [];
  for (var j = 1; j < mv.length; j++) {
    if (mv[j][0]) master.push({ name: String(mv[j][0]), price: n_(mv[j][1]), unit: String(mv[j][2] || "") });
  }
  return { ok: true, orders: orders, master: master };
}

// ---------- 書き込み ----------
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var req = JSON.parse(e.postData.contents);
    if (req.type === "order")  return json_(saveOrder_(req.record));
    if (req.type === "delete") return json_(deleteOrder_(req.id));
    if (req.type === "master") return json_(saveMaster_(req.master));
    return json_({ ok: false, error: "unknown type" });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function rowFromRecord_(rec) {
  var d = rec.data || {};
  var items = d.items || [];
  var itemsText = items.map(function (it) {
    return "・" + (it.name || "") + (it.qty ? "×" + it.qty + (it.unit || "") : "");
  }).join("\n");
  var total = items.reduce(function (s, it) { return s + n_(it.qty) * n_(it.price); }, 0);
  return [
    rec.id, new Date(), d.orderDate || "", d.director || "", d.site || "", d.arrive || "",
    d.requester || "", d.receiver || "", d.hotelName || "", d.hotelZip || "", d.hotelAddr || "",
    d.hotelTel || "", d.toName || "", d.greeting || "", itemsText, total, JSON.stringify(d)
  ];
}

function findRow_(sheet, id) {
  var ids = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 1;
  return -1;
}

function saveOrder_(rec) {
  var s = sheet_(ORDERS, ORDER_HEADERS);
  var row = rowFromRecord_(rec);
  var r = findRow_(s, rec.id);
  if (r > 1) s.getRange(r, 1, 1, row.length).setValues([row]);
  else s.appendRow(row);
  return { ok: true, id: rec.id };
}

function deleteOrder_(id) {
  var s = sheet_(ORDERS, ORDER_HEADERS);
  var r = findRow_(s, id);
  if (r > 1) s.deleteRow(r);
  return { ok: true };
}

function saveMaster_(master) {
  var s = sheet_(MASTER, ["name", "price", "unit"]);
  s.clear();
  s.appendRow(["name", "price", "unit"]);
  (master || []).forEach(function (m) {
    s.appendRow([m.name || "", n_(m.price), m.unit || ""]);
  });
  return { ok: true };
}
