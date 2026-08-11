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
var INVENTORY = "Inventory";
var MOVEMENTS = "Movements";
var ORDER_HEADERS = ["id","保存日時","発注日","ディレクター","現場","着日","依頼者","受取者","ホテル名","郵便番号","住所","電話","宛先","あいさつ","品目","合計金額","data"];
var INV_HEADERS = ["エリア","品名","在庫数","しきい値","更新日時"];
var MOV_HEADERS = ["日時","エリア","品名","種別","数量","後在庫","実行者","発注ID","メモ"];

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

  // 在庫
  var is = sheet_(INVENTORY, INV_HEADERS);
  var iv = is.getDataRange().getValues();
  var inventory = [];
  for (var k = 1; k < iv.length; k++) {
    if (iv[k][0] === "" && iv[k][1] === "") continue;
    inventory.push({ area: String(iv[k][0] || ""), name: String(iv[k][1] || ""), qty: n_(iv[k][2]), threshold: n_(iv[k][3]) });
  }

  // 変動ログ（新しい順・直近300件）
  var mvS = sheet_(MOVEMENTS, MOV_HEADERS);
  var mvv = mvS.getDataRange().getValues();
  var movements = [];
  for (var p = 1; p < mvv.length; p++) {
    if (mvv[p][0] === "" && mvv[p][1] === "") continue;
    movements.push({ ts: String(mvv[p][0]), area: String(mvv[p][1]), name: String(mvv[p][2]), kind: String(mvv[p][3]), delta: n_(mvv[p][4]), after: n_(mvv[p][5]), by: String(mvv[p][6] || ""), orderId: String(mvv[p][7] || ""), memo: String(mvv[p][8] || "") });
  }
  movements.reverse();
  if (movements.length > 300) movements = movements.slice(0, 300);

  return { ok: true, orders: orders, master: master, inventory: inventory, movements: movements };
}

// ---------- 在庫ヘルパー ----------
function invFind_(sheet, area, name) {
  var v = sheet.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(area) && String(v[i][1]) === String(name)) return i + 1;
  }
  return -1;
}
function logMove_(area, name, kind, delta, after, by, orderId, memo) {
  var m = sheet_(MOVEMENTS, MOV_HEADERS);
  m.appendRow([new Date(), area, name, kind, Number(delta), after, by || "", orderId || "", memo || ""]);
}
// 在庫を delta 分だけ増減し、変動ログに記録。新在庫数を返す
function invApplyOne_(area, name, delta, kind, by, orderId, memo) {
  var s = sheet_(INVENTORY, INV_HEADERS);
  var r = invFind_(s, area, name);
  var cur = 0;
  if (r > 1) cur = n_(s.getRange(r, 3).getValue());
  var next = cur + Number(delta);
  if (r > 1) { s.getRange(r, 3).setValue(next); s.getRange(r, 5).setValue(new Date()); }
  else { s.appendRow([area, name, next, 0, new Date()]); }
  logMove_(area, name, kind, delta, next, by, orderId, memo);
  return next;
}

// 手動の入荷/持出/補正（delta指定）
function invMove_(req) {
  var q = invApplyOne_(req.area, req.name, n_(req.delta), req.kind || "調整", req.by, req.orderId || "", req.memo || "");
  return { ok: true, qty: q };
}
// 棚卸し（絶対数を設定）＋しきい値更新
function invSet_(req) {
  var s = sheet_(INVENTORY, INV_HEADERS);
  var r = invFind_(s, req.area, req.name);
  var cur = 0;
  if (r > 1) cur = n_(s.getRange(r, 3).getValue());
  var target = n_(req.qty);
  var delta = target - cur;
  if (r > 1) {
    s.getRange(r, 3).setValue(target);
    if (req.threshold !== undefined && req.threshold !== null) s.getRange(r, 4).setValue(n_(req.threshold));
    s.getRange(r, 5).setValue(new Date());
  } else {
    s.appendRow([req.area, req.name, target, n_(req.threshold || 0), new Date()]);
  }
  if (delta !== 0) logMove_(req.area, req.name, "棚卸し", delta, target, req.by, "", req.memo || "棚卸し補正");
  return { ok: true, qty: target };
}
// 発注連動：発注分＝入荷(+)、押入れ出し＝持出(-)。reverse=trueで取消
function invApplyOrder_(req) {
  var sign = req.reverse ? -1 : 1;
  (req.ins || []).forEach(function (it) {
    var q = n_(it.qty); if (q > 0) invApplyOne_(req.area, it.name, sign * q, req.reverse ? "入荷取消" : "入荷", req.by, req.orderId, "発注連動");
  });
  (req.outs || []).forEach(function (it) {
    var q = n_(it.qty); if (q > 0) invApplyOne_(req.area, it.name, sign * (-q), req.reverse ? "持出取消" : "持出", req.by, req.orderId, "発注連動");
  });
  return { ok: true };
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
    if (req.type === "inv_move")        return json_(invMove_(req));
    if (req.type === "inv_set")         return json_(invSet_(req));
    if (req.type === "inv_apply_order") return json_(invApplyOrder_(req));
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
