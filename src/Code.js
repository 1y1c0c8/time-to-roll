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
var P_ID = 1, P_CAT = 2, P_NAME = 3, P_PERBOX = 4, P_PRICE = 5,
    P_PRICEUNIT = 6, P_LEFT = 7, P_DEFAULT = 8, P_STATUS = 9, P_CREATED = 10;
var PRODUCT_HEADERS = ['id', '類別', '名稱', '每盒支數', '售價', '售價單位', '剩餘支數', '常用預設', '狀態', '建立時間'];

// 紀錄 columns (1-based, fixed) — one sheet per month, named yyyy-MM
var R_ID = 1, R_TIME = 2, R_REASON = 3, R_PID = 4, R_PNAME = 5, R_CAT = 6, R_POUCH = 7, R_COST = 8, R_NOTE = 9;
var RECORD_HEADERS = ['id', '時間', '原因', '菸品id', '菸品名稱', '類別', '菸草包id', '成本', '備註'];

// 原因 columns (1-based, fixed)
var RS_NAME = 1, RS_ORDER = 2, RS_ACTIVE = 3;
var REASON_HEADERS = ['名稱', '排序', '啟用'];

// 菸草包 columns (P2)
var POUCH_HEADERS = ['id', '菸品id', '口味', '開封日', '用完日', '已捲支數', '售價', '狀態'];

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
    records: getMonthRecords(tab),
    thisMonth: tab
  };
}

function ensureReady_() {
  if (!sh(SHEET_SETTINGS) || !sh(SHEET_REASONS) || !sh(SHEET_PRODUCTS) || !sh(SHEET_POUCHES)) setup();
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
      month: tab
    });
  });
  out.sort(function (a, b) { return b.time - a.time; });
  return out;
}

/* ------------------------------------------------------------------ 記錄 CRUD */

function addSmoke(payload) {
  ensureReady_();
  var prod = findProductById_(payload.productId);
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
  row[R_CAT - 1] = prod ? prod.cat : '';
  s.appendRow(row);
  return { id: id, month: tab };
}

function updateSmoke(payload) {
  var loc = findRecordRow_(payload.month, payload.id);
  if (!loc) throw new Error('找不到這筆紀錄');
  if (payload.reason != null) loc.sheet.getRange(loc.row, R_REASON).setValue(payload.reason);
  if (payload.productId != null) {
    var prod = findProductById_(payload.productId);
    loc.sheet.getRange(loc.row, R_PID).setValue(prod ? prod.id : payload.productId);
    loc.sheet.getRange(loc.row, R_PNAME).setValue(prod ? prod.name : '');
    loc.sheet.getRange(loc.row, R_CAT).setValue(prod ? prod.cat : '');
  }
  if (payload.timeMillis) loc.sheet.getRange(loc.row, R_TIME).setValue(new Date(payload.timeMillis));
  return true;
}

function deleteSmoke(payload) {
  var loc = findRecordRow_(payload.month, payload.id);
  if (!loc) throw new Error('找不到這筆紀錄');
  loc.sheet.deleteRow(loc.row);
  return true;
}

function findRecordRow_(tab, id) {
  var s = sh(tab);
  if (!s || s.getLastRow() < 2) return null;
  var ids = s.getRange(2, R_ID, s.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return { sheet: s, row: i + 2 };
  return null;
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
