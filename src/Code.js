/**
 * Time to Roll — 抽菸紀錄 (GAS backend)
 * Google Sheet = database. Frontend (Index.html) talks to these functions via google.script.run.
 *
 * Phases:
 *   P1 (this file, live)  記錄：原因自訂 + 菸品清單(名稱/類別/常用) + 記一根 + 今日/近期列表
 *   P2 (schema ready)     庫存：每盒支數/售價/剩餘支數 連動；捲菸「菸草包」每包捲數統計
 *   P3                    統計：頻率/原因/時段/花費
 *
 * The 菸品 / 菸草包 / 紀錄 sheets already carry all P2/P3 columns; P1 just leaves them blank.
 * Never address a fixed column by HEADERS.length — always by the *_COL constants below.
 */

var TZ  = 'Asia/Taipei';
var CUR = 'NT$';

var SHEET_SETTINGS = '設定';
var SHEET_REASONS  = '原因';
var SHEET_PRODUCTS = '菸品';
var SHEET_POUCHES  = '菸草包';   // P2

var CATEGORIES = ['加熱菸', '盒菸', '捲菸'];

// 菸品 columns (1-based, fixed)
// 剩餘支數(P_LEFT) 給加熱菸/盒菸用；未開包數(P_POUCHES) 給捲菸用（未開封的菸草包數）
var P_ID = 1, P_CAT = 2, P_NAME = 3, P_PERBOX = 4, P_PRICE = 5,
    P_PRICEUNIT = 6, P_LEFT = 7, P_DEFAULT = 8, P_STATUS = 9, P_CREATED = 10, P_POUCHES = 11;
var PRODUCT_HEADERS = ['id', '類別', '名稱', '每盒支數', '售價', '售價單位', '剩餘支數', '常用預設', '狀態', '建立時間', '未開包數'];

var STICK_CATS = ['加熱菸', '盒菸'];   // 支為單位、可自動連動扣除
function isStick(cat) { return STICK_CATS.indexOf(cat) >= 0; }

// 紀錄 columns (1-based, fixed) — one sheet per month, named yyyy-MM
var R_ID = 1, R_TIME = 2, R_REASON = 3, R_PID = 4, R_PNAME = 5, R_CAT = 6, R_POUCH = 7, R_COST = 8, R_NOTE = 9;
var RECORD_HEADERS = ['id', '時間', '原因', '菸品id', '菸品名稱', '類別', '菸草包id', '成本', '備註'];

// 原因 columns (1-based, fixed)
var RS_NAME = 1, RS_ORDER = 2, RS_ACTIVE = 3;
var REASON_HEADERS = ['名稱', '排序', '啟用'];

// 菸草包 columns (P2, 1-based, fixed)
var G_ID = 1, G_PID = 2, G_NAME = 3, G_OPEN = 4, G_DONE = 5, G_ROLLED = 6, G_PRICE = 7, G_STATUS = 8;
var POUCH_HEADERS = ['id', '菸品id', '口味', '開封日', '用完日', '已捲支數', '售價', '狀態'];
var POUCH_USING = '使用中', POUCH_DONE = '已用完';

/* ------------------------------------------------------------------ web entry */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Time to Roll')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover');
}

/* ------------------------------------------------------------------ helpers */

function ss() { return SpreadsheetApp.getActive(); }
function sh(name) { return ss().getSheetByName(name); }
function monthTab(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM'); }

function styleHeader(sheet, n) {
  sheet.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#e3efed');
  sheet.setFrozenRows(1);
}

function getOrCreateMonthSheet(name) {
  var s = sh(name);
  if (!s) {
    s = ss().insertSheet(name);
    s.getRange(1, 1, 1, RECORD_HEADERS.length).setValues([RECORD_HEADERS]);
    styleHeader(s, RECORD_HEADERS.length);
  }
  return s;
}

/* ------------------------------------------------------------------ one-time setup (run once in the editor) */

function setup() {
  var book = ss();

  if (!sh(SHEET_SETTINGS)) {
    var s = book.insertSheet(SHEET_SETTINGS, 0);
    s.getRange(1, 1, 1, 2).setValues([['設定', '值']]);
    s.getRange(2, 1, 3, 2).setValues([['時區', TZ], ['幣別', CUR], ['預設每盒支數', 20]]);
    styleHeader(s, 2);
  }

  if (!sh(SHEET_REASONS)) {
    var r = book.insertSheet(SHEET_REASONS);
    r.getRange(1, 1, 1, REASON_HEADERS.length).setValues([REASON_HEADERS]);
    var seed = [['壓力', 1, true], ['飯後', 2, true], ['無聊', 3, true], ['社交', 4, true], ['習慣', 5, true]];
    r.getRange(2, 1, seed.length, 3).setValues(seed);
    styleHeader(r, REASON_HEADERS.length);
  }

  if (!sh(SHEET_PRODUCTS)) {
    var p = book.insertSheet(SHEET_PRODUCTS);
    p.getRange(1, 1, 1, PRODUCT_HEADERS.length).setValues([PRODUCT_HEADERS]);
    styleHeader(p, PRODUCT_HEADERS.length);
  }

  if (!sh(SHEET_POUCHES)) {
    var g = book.insertSheet(SHEET_POUCHES);
    g.getRange(1, 1, 1, POUCH_HEADERS.length).setValues([POUCH_HEADERS]);
    styleHeader(g, POUCH_HEADERS.length);
  }

  getOrCreateMonthSheet(monthTab(new Date()));
  return 'setup 完成';
}

/* ------------------------------------------------------------------ boot / reads */

function getBootData() {
  ensureReady_();
  var now = new Date();
  var tab = monthTab(now);
  return {
    settings: readSettings(),
    reasons: getReasons(),
    products: getProducts(),
    pouches: getPouches(),
    records: getMonthRecords(tab),
    thisMonth: tab
  };
}

function ensureReady_() {
  if (!sh(SHEET_SETTINGS) || !sh(SHEET_REASONS) || !sh(SHEET_PRODUCTS) || !sh(SHEET_POUCHES)) setup();
  migrate_();
}

// 幫既有分頁補上後來新增的欄位表頭（P1 → P2）
function migrate_() {
  var p = sh(SHEET_PRODUCTS);
  if (p && String(p.getRange(1, P_POUCHES).getValue()).trim() !== '未開包數') {
    p.getRange(1, P_POUCHES).setValue('未開包數').setFontWeight('bold').setBackground('#e3efed');
  }
}

function readSettings() {
  var o = { tz: TZ, cur: CUR, perBox: 20 };
  var s = sh(SHEET_SETTINGS);
  if (!s || s.getLastRow() < 2) return o;
  var v = s.getRange(2, 1, s.getLastRow() - 1, 2).getValues();
  v.forEach(function (row) {
    var k = String(row[0]).trim();
    if (k === '時區') o.tz = String(row[1]);
    else if (k === '幣別') o.cur = String(row[1]);
    else if (k === '預設每盒支數') o.perBox = Number(row[1]) || 20;
  });
  return o;
}

function getReasons() {
  var s = sh(SHEET_REASONS);
  if (!s || s.getLastRow() < 2) return [];
  var v = s.getRange(2, 1, s.getLastRow() - 1, 3).getValues();
  var out = [];
  v.forEach(function (row, i) {
    if (String(row[RS_NAME - 1]).trim() === '') return;
    var active = row[RS_ACTIVE - 1] !== false && String(row[RS_ACTIVE - 1]).toUpperCase() !== 'FALSE';
    if (!active) return;
    out.push({ name: String(row[RS_NAME - 1]).trim(), order: Number(row[RS_ORDER - 1]) || i, row: i + 2 });
  });
  out.sort(function (a, b) { return a.order - b.order; });
  return out;
}

function getProducts() {
  var s = sh(SHEET_PRODUCTS);
  if (!s || s.getLastRow() < 2) return [];
  var v = s.getRange(2, 1, s.getLastRow() - 1, PRODUCT_HEADERS.length).getValues();
  var out = [];
  v.forEach(function (row, i) {
    if (String(row[P_ID - 1]).trim() === '') return;
    if (String(row[P_STATUS - 1] || '在用') === '停用') return;
    out.push({
      id: String(row[P_ID - 1]),
      cat: String(row[P_CAT - 1]),
      name: String(row[P_NAME - 1]),
      perBox: Number(row[P_PERBOX - 1]) || 0,
      price: Number(row[P_PRICE - 1]) || 0,
      priceUnit: String(row[P_PRICEUNIT - 1] || ''),
      left: Number(row[P_LEFT - 1]) || 0,
      pouchesLeft: Number(row[P_POUCHES - 1]) || 0,
      isStick: isStick(String(row[P_CAT - 1])),
      isDefault: row[P_DEFAULT - 1] === true || String(row[P_DEFAULT - 1]).toUpperCase() === 'TRUE',
      row: i + 2
    });
  });
  return out;
}

function getMonthRecords(tab) {
  var s = sh(tab);
  if (!s || s.getLastRow() < 2) return [];
  var v = s.getRange(2, 1, s.getLastRow() - 1, RECORD_HEADERS.length).getValues();
  var out = [];
  v.forEach(function (row) {
    if (String(row[R_ID - 1]).trim() === '') return;
    var t = row[R_TIME - 1];
    var d = (t instanceof Date) ? t : new Date(t);
    out.push({
      id: String(row[R_ID - 1]),
      time: d.getTime(),
      timeStr: Utilities.formatDate(d, TZ, 'HH:mm'),
      dateStr: Utilities.formatDate(d, TZ, 'yyyy-MM-dd'),
      reason: String(row[R_REASON - 1] || ''),
      productId: String(row[R_PID - 1] || ''),
      productName: String(row[R_PNAME - 1] || ''),
      cat: String(row[R_CAT - 1] || ''),
      pouchId: String(row[R_POUCH - 1] || ''),
      month: tab
    });
  });
  out.sort(function (a, b) { return b.time - a.time; });
  return out;
}

/* ------------------------------------------------------------------ 記錄 CRUD */

// 記一根之後回傳整包最新狀態，前端一次刷新（記錄 + 庫存 + 菸草包）
function state_(tab) {
  return { records: getMonthRecords(tab), products: getProducts(), pouches: getPouches(), month: tab };
}

function addSmoke(payload) {
  ensureReady_();
  var prod = findProductById_(payload.productId);
  var cat = prod ? prod.cat : '';
  var pouchId = '';
  if (cat === '捲菸') {
    pouchId = String(payload.pouchId || '');
    if (!pouchId) throw new Error('捲菸請先選一包使用中的菸草包');
  }
  var now = new Date();
  var tab = monthTab(now);
  var s = getOrCreateMonthSheet(tab);
  var id = 'r' + now.getTime();
  var row = new Array(RECORD_HEADERS.length).fill('');
  row[R_ID - 1] = id;
  row[R_TIME - 1] = now;
  row[R_REASON - 1] = payload.reason || '';
  row[R_PID - 1] = prod ? prod.id : (payload.productId || '');
  row[R_PNAME - 1] = prod ? prod.name : '';
  row[R_CAT - 1] = cat;
  row[R_POUCH - 1] = pouchId;
  s.appendRow(row);
  consumeInventory_(cat, row[R_PID - 1], pouchId);
  return state_(tab);
}

function updateSmoke(payload) {
  var loc = findRecordRow_(payload.month, payload.id);
  if (!loc) throw new Error('找不到這筆紀錄');
  var old = readRecordRow_(loc.sheet, loc.row);

  if (payload.reason != null) loc.sheet.getRange(loc.row, R_REASON).setValue(payload.reason);

  if (payload.productId != null) {
    var prod = findProductById_(payload.productId);
    var newCat = prod ? prod.cat : '';
    var newPouch = '';
    if (newCat === '捲菸') {
      newPouch = String(payload.pouchId || (old.cat === '捲菸' ? old.pouchId : '') || '');
      if (!newPouch) throw new Error('捲菸請選一包菸草包');
    }
    var changed = (String(prod ? prod.id : payload.productId) !== old.pid) || (newPouch !== old.pouchId);
    if (changed) {
      restoreInventory_(old.cat, old.pid, old.pouchId);
      consumeInventory_(newCat, prod ? prod.id : payload.productId, newPouch);
    }
    loc.sheet.getRange(loc.row, R_PID).setValue(prod ? prod.id : payload.productId);
    loc.sheet.getRange(loc.row, R_PNAME).setValue(prod ? prod.name : '');
    loc.sheet.getRange(loc.row, R_CAT).setValue(newCat);
    loc.sheet.getRange(loc.row, R_POUCH).setValue(newPouch);
  }

  if (payload.timeMillis) loc.sheet.getRange(loc.row, R_TIME).setValue(new Date(payload.timeMillis));
  return state_(payload.month);
}

function deleteSmoke(payload) {
  var loc = findRecordRow_(payload.month, payload.id);
  if (!loc) throw new Error('找不到這筆紀錄');
  var old = readRecordRow_(loc.sheet, loc.row);
  loc.sheet.deleteRow(loc.row);
  restoreInventory_(old.cat, old.pid, old.pouchId);
  return state_(payload.month);
}

function findRecordRow_(tab, id) {
  var s = sh(tab);
  if (!s || s.getLastRow() < 2) return null;
  var ids = s.getRange(2, R_ID, s.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return { sheet: s, row: i + 2 };
  return null;
}

function readRecordRow_(sheet, row) {
  var v = sheet.getRange(row, R_PID, 1, 4).getValues()[0]; // R_PID..R_POUCH
  return { pid: String(v[0] || ''), cat: String(v[2] || ''), pouchId: String(v[3] || '') };
}

/* ------------------------------------------------------------------ 庫存連動 */

function consumeInventory_(cat, pid, pouchId) {
  if (isStick(cat)) adjustProductLeft_(pid, -1);
  else if (cat === '捲菸' && pouchId) adjustPouchRolled_(pouchId, +1);
}
function restoreInventory_(cat, pid, pouchId) {
  if (isStick(cat)) adjustProductLeft_(pid, +1);
  else if (cat === '捲菸' && pouchId) adjustPouchRolled_(pouchId, -1);
}
function adjustProductLeft_(pid, delta) {
  var loc = findProductRow_(pid);
  if (!loc) return;
  var nv = (Number(loc.sheet.getRange(loc.row, P_LEFT).getValue()) || 0) + delta;
  if (nv < 0) nv = 0;
  loc.sheet.getRange(loc.row, P_LEFT).setValue(nv);
}
function adjustPouchRolled_(pouchId, delta) {
  var loc = findPouchRow_(pouchId);
  if (!loc) return;
  var nv = (Number(loc.sheet.getRange(loc.row, G_ROLLED).getValue()) || 0) + delta;
  if (nv < 0) nv = 0;
  loc.sheet.getRange(loc.row, G_ROLLED).setValue(nv);
}

/* ------------------------------------------------------------------ 原因 CRUD */

function addReason(name) {
  ensureReady_();
  name = String(name || '').trim();
  if (!name) throw new Error('原因不能空白');
  if (getReasons().some(function (r) { return r.name === name; })) throw new Error('這個原因已經有了');
  var s = sh(SHEET_REASONS);
  s.appendRow([name, s.getLastRow(), true]);
  return getReasons();
}

function deleteReason(name) {
  var s = sh(SHEET_REASONS);
  if (!s || s.getLastRow() < 2) return getReasons();
  var v = s.getRange(2, 1, s.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < v.length; i++)
    if (String(v[i][0]).trim() === String(name).trim()) s.getRange(i + 2, RS_ACTIVE).setValue(false);
  return getReasons();
}

function renameReason(oldName, newName) {
  newName = String(newName || '').trim();
  if (!newName) throw new Error('名稱不能空白');
  var s = sh(SHEET_REASONS);
  var v = s.getRange(2, 1, s.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < v.length; i++)
    if (String(v[i][0]).trim() === String(oldName).trim()) s.getRange(i + 2, RS_NAME).setValue(newName);
  return getReasons();
}

/* ------------------------------------------------------------------ 菸品 CRUD */

function addProduct(p) {
  ensureReady_();
  if (CATEGORIES.indexOf(p.cat) < 0) throw new Error('菸種不對');
  var name = String(p.name || '').trim();
  if (!name) throw new Error('名稱不能空白');
  var s = sh(SHEET_PRODUCTS);
  var id = 'p' + Date.now();
  var isRoll = (p.cat === '捲菸');
  var row = new Array(PRODUCT_HEADERS.length).fill('');
  row[P_ID - 1] = id;
  row[P_CAT - 1] = p.cat;
  row[P_NAME - 1] = name;
  row[P_PERBOX - 1] = isRoll ? '' : (Number(p.perBox) || readSettings().perBox);
  row[P_PRICE - 1] = Number(p.price) || 0;
  row[P_PRICEUNIT - 1] = isRoll ? '包' : '盒';
  row[P_LEFT - 1] = 0;
  row[P_DEFAULT - 1] = !!p.isDefault;
  row[P_STATUS - 1] = '在用';
  row[P_CREATED - 1] = new Date();
  s.appendRow(row);
  if (p.isDefault) setDefaultProduct(id);
  return getProducts();
}

function updateProduct(p) {
  var loc = findProductRow_(p.id);
  if (!loc) throw new Error('找不到菸品');
  if (p.name != null) loc.sheet.getRange(loc.row, P_NAME).setValue(String(p.name).trim());
  if (p.perBox != null) loc.sheet.getRange(loc.row, P_PERBOX).setValue(p.cat === '捲菸' ? '' : (Number(p.perBox) || ''));
  if (p.price != null) loc.sheet.getRange(loc.row, P_PRICE).setValue(Number(p.price) || 0);
  return getProducts();
}

function deleteProduct(id) {
  var loc = findProductRow_(id);
  if (loc) loc.sheet.getRange(loc.row, P_STATUS).setValue('停用');
  return getProducts();
}

function setDefaultProduct(id) {
  var s = sh(SHEET_PRODUCTS);
  if (!s || s.getLastRow() < 2) return getProducts();
  var ids = s.getRange(2, P_ID, s.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++)
    s.getRange(i + 2, P_DEFAULT).setValue(String(ids[i][0]) === String(id));
  return getProducts();
}

function findProductRow_(id) {
  var s = sh(SHEET_PRODUCTS);
  if (!s || s.getLastRow() < 2) return null;
  var ids = s.getRange(2, P_ID, s.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return { sheet: s, row: i + 2 };
  return null;
}

function findProductById_(id) {
  var all = getProducts();
  for (var i = 0; i < all.length; i++) if (all[i].id === String(id)) return all[i];
  return null;
}

/* ------------------------------------------------------------------ 買入 / 菸草包 (P2) */

// qty = 盒數(加熱菸/盒菸) 或 包數(捲菸)
function buyStock(payload) {
  ensureReady_();
  var prod = findProductById_(payload.id);
  if (!prod) throw new Error('找不到菸品');
  var qty = Math.max(0, Math.floor(Number(payload.qty) || 0));
  if (!qty) throw new Error('數量要大於 0');
  var loc = findProductRow_(prod.id);
  if (prod.isStick) {
    adjustProductLeft_(prod.id, qty * (prod.perBox || readSettings().perBox));
  } else {   // 捲菸：加未開包數
    var cur = Number(loc.sheet.getRange(loc.row, P_POUCHES).getValue()) || 0;
    loc.sheet.getRange(loc.row, P_POUCHES).setValue(cur + qty);
  }
  return { products: getProducts(), pouches: getPouches() };
}

// 捲菸：開一包 → 建立「使用中」的菸草包，未開包數 -1
function openPouch(payload) {
  ensureReady_();
  var prod = findProductById_(payload.productId);
  if (!prod || prod.cat !== '捲菸') throw new Error('只有捲菸能開包');
  var loc = findProductRow_(prod.id);
  var cur = Number(loc.sheet.getRange(loc.row, P_POUCHES).getValue()) || 0;
  if (cur > 0) loc.sheet.getRange(loc.row, P_POUCHES).setValue(cur - 1);
  var g = sh(SHEET_POUCHES);
  var row = new Array(POUCH_HEADERS.length).fill('');
  row[G_ID - 1] = 'g' + Date.now();
  row[G_PID - 1] = prod.id;
  row[G_NAME - 1] = prod.name;
  row[G_OPEN - 1] = new Date();
  row[G_ROLLED - 1] = 0;
  row[G_PRICE - 1] = prod.price || 0;
  row[G_STATUS - 1] = POUCH_USING;
  g.appendRow(row);
  return { products: getProducts(), pouches: getPouches() };
}

// 這包用完 → 歸檔，記用完日（共捲支數＝已捲支數，每支成本＝售價/已捲支數，由前端算）
function finishPouch(payload) {
  var loc = findPouchRow_(payload.pouchId);
  if (!loc) throw new Error('找不到菸草包');
  loc.sheet.getRange(loc.row, G_STATUS).setValue(POUCH_DONE);
  loc.sheet.getRange(loc.row, G_DONE).setValue(new Date());
  return { products: getProducts(), pouches: getPouches() };
}

function getPouches() {
  var s = sh(SHEET_POUCHES);
  if (!s || s.getLastRow() < 2) return { using: [], done: [] };
  var v = s.getRange(2, 1, s.getLastRow() - 1, POUCH_HEADERS.length).getValues();
  var using = [], done = [];
  v.forEach(function (row) {
    if (String(row[G_ID - 1]).trim() === '') return;
    var openD = row[G_OPEN - 1] instanceof Date ? row[G_OPEN - 1] : (row[G_OPEN - 1] ? new Date(row[G_OPEN - 1]) : null);
    var doneD = row[G_DONE - 1] instanceof Date ? row[G_DONE - 1] : (row[G_DONE - 1] ? new Date(row[G_DONE - 1]) : null);
    var rolled = Number(row[G_ROLLED - 1]) || 0;
    var price = Number(row[G_PRICE - 1]) || 0;
    var o = {
      id: String(row[G_ID - 1]),
      productId: String(row[G_PID - 1] || ''),
      name: String(row[G_NAME - 1] || ''),
      openStr: openD ? Utilities.formatDate(openD, TZ, 'M/d') : '',
      doneStr: doneD ? Utilities.formatDate(doneD, TZ, 'M/d') : '',
      rolled: rolled,
      price: price,
      perStick: (rolled > 0) ? Math.round(price / rolled * 10) / 10 : 0,
      status: String(row[G_STATUS - 1] || '')
    };
    if (o.status === POUCH_DONE) done.push(o); else using.push(o);
  });
  done.reverse();
  return { using: using, done: done.slice(0, 20) };
}

function findPouchRow_(id) {
  var s = sh(SHEET_POUCHES);
  if (!s || s.getLastRow() < 2) return null;
  var ids = s.getRange(2, G_ID, s.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return { sheet: s, row: i + 2 };
  return null;
}
