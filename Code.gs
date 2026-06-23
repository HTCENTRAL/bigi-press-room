// BIGI PRESS ROOM アパレル貸出管理システム

var SHEET_LOAN       = '貸出管理';
var SHEET_PUBLISH    = '掲載リストマスタ';
var SHEET_SETTINGS   = '設定';
var SHEET_CREDIT     = 'クレジットマスタ';
var SHEET_ITEM       = '品番マスタ';
var SHEET_ITEM_PASTE = '品番マスタコピペ先';
var SHEET_CONTACT    = '貸出先マスタ';
var START_SLIP_NO = 8877;

// 貸出管理シート 列番号（1-based）
var COL_LOAN = {
  SLIP_NO:     1,   // A: 伝票番号
  DATE:        2,   // B: 起票日
  NAME:        3,   // C: お名前
  PHONE:       4,   // D: 電話番号
  TYPE:        5,   // E: 種別（パイプ区切り）
  MEDIA:       6,   // F: 媒体名（パイプ区切り）
  ISSUE:       7,   // G: 月号/放送局（パイプ区切り）
  THEME:       8,   // H: テーマ/着用者（パイプ区切り）
  PUB_DATE:    9,   // I: 公開日（パイプ区切り）
  SHOOT_DATE:  10,  // J: 撮影日（パイプ区切り）
  PICKUP_DATE: 11,  // K: ピックアップ日
  RETURN_PLAN: 12,  // L: 返却予定日
  PARTIAL_RTN: 13,  // M: 一部返却日
  MAIN_BRAND:  14,  // N: メインブランド
  BRAND:       15,  // O: ブランド
  ITEM_CODE:   16,  // P: 品番
  COLOR:       17,  // Q: カラー
  COLOR_NAME:  18,  // R: カラー名
  ITEM_NAME:   19,  // S: アイテム名
  PRICE:       20,  // T: 単価(税込)
  STATUS:      21,  // U: 返却ステータス
  RETURN_DATE: 22,  // V: 返却日
  PUBLISHED:   23,  // W: 掲載済
  NOTE:        24,  // X: 備考
  STAFF:       25,  // Y: 起票者名
  TOTAL:       25
};

// 掲載リストマスタ 列番号（1-based）
var COL_PUB = {
  SLIP_NO:    1,   // A: 伝票番号
  TYPE:       2,   // B: 種別
  MEDIA:      3,   // C: 媒体名
  ISSUE:      4,   // D: 月号/放送局
  PUB_DATE:   5,   // E: 公開日
  THEME:      6,   // F: テーマ/着用者
  STYLIST:    7,   // G: スタイリスト
  MAIN_BRAND: 8,   // H: メインブランド
  BRAND:      9,   // I: ブランド名
  ITEM_CODE:  10,  // J: 品番
  COLOR:      11,  // K: カラー
  COLOR_NAME: 12,  // L: カラー名
  PRICE:      13,  // M: 単価(税込)
  ITEM_NAME:  14,  // N: アイテム
  PAGE:       15,  // O: 頁
  IMAGE:      16,  // P: 画像/リンク
  TOTAL:      16
};

var TYPE_CONFIG = {
  '雑誌(紙面)': { title: '雑誌掲載リスト',  col1: '雑誌名' },
  '雑誌(WEB)':  { title: 'WEB掲載リスト',   col1: '雑誌名' },
  'TV':         { title: 'テレビ着用リスト', col1: '番組名' },
  'その他':     { title: 'その他リスト',     col1: '媒体名' }
};
var TYPE_ORDER = ['雑誌(紙面)', '雑誌(WEB)', 'TV', 'その他'];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('貸出管理')
    .addItem('新規貸出登録', 'openLoanSidebar')
    .addItem('伝票印刷', 'openPrintSlipDialog')
    .addSeparator()
    .addItem('返却・掲載登録処理', 'openReturnDialog')
    .addSeparator()
    .addItem('掲載リスト作成', 'openFormattedSheetDialog')
    .addSeparator()
    .addItem('品番マスタ更新', 'importItemMaster')
    .addSeparator()
    .addItem('伝票編集', 'openMediaSetDialog')
    .addItem('掲載リストマスタ編集', 'openPublishListEditDialog')
    .addToUi();
}

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // === 貸出管理シート（25列） ===
  var loanSheet = ss.getSheetByName(SHEET_LOAN);
  if (!loanSheet) loanSheet = ss.insertSheet(SHEET_LOAN);

  var loanHeaders = [
    '伝票番号', '起票日', 'お名前', '電話番号',
    '種別', '媒体名', '月号/放送局', 'テーマ/着用者', '公開日', '撮影日',
    'ピックアップ日', '返却予定日', '一部返却日',
    'メインブランド', 'ブランド', '品番', 'カラー', 'カラー名', 'アイテム名', '単価(税込)',
    '返却ステータス', '返却日', '掲載済', '備考', '起票者名'
  ];
  loanSheet.getRange(1, 1, 1, loanHeaders.length).setValues([loanHeaders]);
  loanSheet.getRange(1, 1, 1, loanHeaders.length)
    .setBackground('#4a4a4a').setFontColor('#ffffff').setFontWeight('bold');
  loanSheet.setFrozenRows(1);

  var loanColWidths = [80,85,120,110,100,120,90,180,100,85,95,95,90,110,100,100,90,90,180,80,90,85,65,150,120];
  for (var ci = 0; ci < loanColWidths.length; ci++) {
    loanSheet.setColumnWidth(ci + 1, loanColWidths[ci]);
  }

  // 条件付き書式: U列（返却ステータス）=「未返却」→ 赤
  var cfRange = loanSheet.getRange(2, 1, 1000, COL_LOAN.TOTAL);
  loanSheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$U2="未返却"')
      .setBackground('#ffcccc')
      .setRanges([cfRange])
      .build()
  ]);

  // === 掲載リストマスタシート（16列） ===
  var pubSheet = ss.getSheetByName(SHEET_PUBLISH);
  if (!pubSheet) pubSheet = ss.insertSheet(SHEET_PUBLISH);
  var pubHeaders = [
    '伝票番号', '種別', '媒体名', '月号/放送局', '公開日', 'テーマ/着用者', 'スタイリスト',
    'メインブランド', 'ブランド名', '品番', 'カラー', 'カラー名', '単価(税込)', 'アイテム', '頁', '画像/リンク'
  ];
  pubSheet.getRange(1, 1, 1, pubHeaders.length).setValues([pubHeaders]);
  pubSheet.getRange(1, 1, 1, pubHeaders.length)
    .setBackground('#4a4a4a').setFontColor('#ffffff').setFontWeight('bold');
  pubSheet.setFrozenRows(1);

  // === 設定シート ===
  var settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!settingsSheet) settingsSheet = ss.insertSheet(SHEET_SETTINGS);
  if (!settingsSheet.getRange('A3').getValue()) {
    settingsSheet.getRange('A3').setValue('ブランド名').setFontWeight('bold');
  }
  if (!settingsSheet.getRange('B3').getValue()) {
    settingsSheet.getRange('B3').setValue("L'EQUIPE");
  }
  if (!settingsSheet.getRange('A5').getValue()) {
    var prRows = [
      ['PRESS ROOM 1 名称', 'BIGI CO.,LTD. PRESS ROOM（石橋ビル）'],
      ['PRESS ROOM 1 住所', '〒153-8610 東京都目黒区青葉台2-1-4, 2F'],
      ['PRESS ROOM 1 TEL',  '03-6861-7702'],
      ['PRESS ROOM 1 FAX',  '03-6861-7703'],
      ['', ''],
      ['PRESS ROOM 2 名称', 'BIGI CO.,LTD. PRESS ROOM（ダミービル）'],
      ['PRESS ROOM 2 住所', '（ダミー住所）'],
      ['PRESS ROOM 2 TEL',  '03-xxxx-xxxx'],
      ['PRESS ROOM 2 FAX',  '03-xxxx-xxxx']
    ];
    settingsSheet.getRange(5, 1, prRows.length, 2).setValues(prRows);
    settingsSheet.getRange(5, 1, prRows.length, 1).setFontWeight('bold');
  }

  // === クレジットマスタシート ===
  var creditSheet = ss.getSheetByName(SHEET_CREDIT);
  if (!creditSheet) {
    creditSheet = ss.insertSheet(SHEET_CREDIT);
    var creditHeaders = ['表示名', 'カナ', 'TEL', 'WEB SITE', 'ONLINE STORE', 'Instagram'];
    creditSheet.getRange(1, 1, 1, creditHeaders.length).setValues([creditHeaders]);
    creditSheet.getRange(1, 1, 1, creditHeaders.length)
      .setBackground('#4a4a4a').setFontColor('#ffffff').setFontWeight('bold');
    creditSheet.setFrozenRows(1);
  }

  // === 品番マスタシート（新規） ===
  var itemSheet = ss.getSheetByName(SHEET_ITEM);
  if (!itemSheet) {
    itemSheet = ss.insertSheet(SHEET_ITEM);
    var itemHeaders = ['メインブランド', 'ブランド', '品番', 'カラー', 'カラー名', '単価(税込)', 'アイテム名'];
    itemSheet.getRange(1, 1, 1, itemHeaders.length).setValues([itemHeaders]);
    itemSheet.getRange(1, 1, 1, itemHeaders.length)
      .setBackground('#4a4a4a').setFontColor('#ffffff').setFontWeight('bold');
    itemSheet.setFrozenRows(1);
    var itemColWidths = [120, 120, 100, 90, 120, 90, 200];
    for (var ii = 0; ii < itemColWidths.length; ii++) {
      itemSheet.setColumnWidth(ii + 1, itemColWidths[ii]);
    }
  }

  // === 品番マスタコピペ先シート（素データ貼り付け用） ===
  var itemPasteSheet = ss.getSheetByName(SHEET_ITEM_PASTE);
  if (!itemPasteSheet) {
    itemPasteSheet = ss.insertSheet(SHEET_ITEM_PASTE);
  }

  // === 貸出先マスタシート（新規） ===
  var contactSheet = ss.getSheetByName(SHEET_CONTACT);
  if (!contactSheet) {
    contactSheet = ss.insertSheet(SHEET_CONTACT);
    var contactHeaders = ['お名前', '電話番号'];
    contactSheet.getRange(1, 1, 1, contactHeaders.length).setValues([contactHeaders]);
    contactSheet.getRange(1, 1, 1, contactHeaders.length)
      .setBackground('#4a4a4a').setFontColor('#ffffff').setFontWeight('bold');
    contactSheet.setFrozenRows(1);
  }

  Logger.log('setup done');
}

// === ユーティリティ ===

function getLastDataRow(sheet) {
  var lastMatch = sheet.getRange('A:A')
    .createTextFinder('.').useRegularExpression(true)
    .findPrevious();
  return lastMatch ? lastMatch.getRow() : 1;
}

function cellToStr(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  }
  if (v === null || v === undefined) return '';
  return String(v);
}

function getNextSlipNumber() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOAN);
  var lastDataRow = getLastDataRow(sheet);
  var maxNo = START_SLIP_NO;
  if (lastDataRow > 1) {
    var colA = sheet.getRange(2, 1, lastDataRow - 1, 1).getValues();
    for (var i = 0; i < colA.length; i++) {
      var val = parseInt(String(colA[i][0]).replace(/^0+/, ''), 10);
      if (!isNaN(val) && val > maxNo) maxNo = val;
    }
  }
  var s = String(maxNo + 1);
  while (s.length < 5) s = '0' + s;
  return s;
}

function todayStr() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

// === 品番マスタ検索 ===

function lookupItemMaster(itemCode, colorCode) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ITEM);
    if (!sheet || sheet.getLastRow() <= 1) return null;
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
    var code  = String(itemCode).trim();
    var color = String(colorCode).trim();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][2]).trim() === code && String(data[i][3]).trim() === color) {
        return {
          mainBrand: cellToStr(data[i][0]),
          brand:     cellToStr(data[i][1]),
          colorName: cellToStr(data[i][4]),
          price:     data[i][5] !== '' ? Number(data[i][5]) : '',
          itemName:  cellToStr(data[i][6])
        };
      }
    }
    return null;
  } catch(e) {
    return null;
  }
}

function getMainBrandList() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ITEM);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    var brands = [];
    data.forEach(function(row) {
      var b = String(row[0]).trim();
      if (b && brands.indexOf(b) === -1) brands.push(b);
    });
    return brands;
  } catch(e) {
    return [];
  }
}

// === 品番マスタ インポート ===

function importItemMaster() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var srcSheet = ss.getSheetByName(SHEET_ITEM_PASTE);
  if (!srcSheet) {
    SpreadsheetApp.getUi().alert('「' + SHEET_ITEM_PASTE + '」シートが見つかりません。');
    return;
  }
  var srcLastRow = srcSheet.getLastRow();
  if (srcLastRow < 1) {
    SpreadsheetApp.getUi().alert('コピペ先にデータがありません。');
    return;
  }

  // コピペ先から抽出・ユニーク化（品番+カラーをキー）
  var srcData = srcSheet.getRange(1, 1, srcLastRow, 17).getValues();
  var seen = {};
  var srcItems = [];
  for (var i = 0; i < srcData.length; i++) {
    var r = srcData[i];
    var itemCode  = String(r[5]).trim();   // F列
    var color     = String(r[10]).trim();  // K列
    if (!itemCode) continue;
    var key = itemCode + '\t' + color;
    if (seen[key]) continue;
    seen[key] = true;
    var priceRaw = r[16]; // Q列
    var price = '';
    if (priceRaw !== '' && priceRaw !== null && priceRaw !== undefined) {
      var p = parseFloat(String(priceRaw).replace(/,/g, ''));
      if (!isNaN(p)) price = Math.floor(p * 1.1);
    }
    srcItems.push({
      key:       key,
      mainBrand: String(r[2]).trim(),   // C列
      itemCode:  itemCode,
      color:     color,
      colorName: String(r[11]).trim(),  // L列
      price:     price,
      itemName:  String(r[13]).trim()   // N列
    });
  }
  if (srcItems.length === 0) {
    SpreadsheetApp.getUi().alert('有効なデータが見つかりませんでした（F列が空の行はスキップされます）。');
    return;
  }

  // 既存マスタを読み込んでキーマップ化（ブランドを保持するため）
  var destSheet = ss.getSheetByName(SHEET_ITEM);
  var masterMap = {};
  var masterLastRow = getLastDataRow(destSheet);
  if (masterLastRow > 1) {
    var masterData = destSheet.getRange(2, 1, masterLastRow - 1, 7).getValues();
    for (var j = 0; j < masterData.length; j++) {
      if (!masterData[j][2]) continue;
      var mKey = String(masterData[j][2]).trim() + '\t' + String(masterData[j][3]).trim();
      masterMap[mKey] = true;
    }
  }

  // 品番マスタに存在しない行だけ抽出
  var appendRows = [];
  for (var k = 0; k < srcItems.length; k++) {
    if (!masterMap[srcItems[k].key]) {
      var it = srcItems[k];
      appendRows.push([it.mainBrand, '', it.itemCode, it.color, it.colorName, it.price, it.itemName]);
    }
  }

  var ui = SpreadsheetApp.getUi();
  if (appendRows.length === 0) {
    ui.alert('追加対象なし', 'コピペ先の全レコードはすでに品番マスタに存在します。', ui.ButtonSet.OK);
    return;
  }

  var skipCount = srcItems.length - appendRows.length;
  var msg = '品番マスタに追記します。\n\n' +
    '  新規追加: ' + appendRows.length + ' 件\n' +
    '  スキップ（既存）: ' + skipCount + ' 件\n\n' +
    '既存レコードは変更されません。続けますか？';
  if (ui.alert('品番マスタ更新確認', msg, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  var nextRow = masterLastRow + 1;
  destSheet.getRange(nextRow, 1, appendRows.length, 7).setValues(appendRows);
  destSheet.getRange(nextRow, 3, appendRows.length, 1).setNumberFormat('@');
  destSheet.getRange(nextRow, 4, appendRows.length, 1).setNumberFormat('@');

  ui.alert('完了', appendRows.length + ' 件を追加しました。', ui.ButtonSet.OK);
}

// === 貸出先マスタ ===

function getContactList() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONTACT);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    var contacts = [];
    data.forEach(function(row) {
      if (!row[0]) return;
      contacts.push({ name: cellToStr(row[0]), phone: cellToStr(row[1]) });
    });
    return contacts;
  } catch(e) {
    return [];
  }
}

function upsertContactMaster(name, phone) {
  try {
    if (!name) return;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONTACT);
    if (!sheet) return;
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      for (var i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim() === String(name).trim()) {
          if (phone && cellToStr(data[i][1]) !== String(phone)) {
            sheet.getRange(i + 2, 2).setValue(phone);
          }
          return;
        }
      }
    }
    var newRow = lastRow < 1 ? 2 : lastRow + 1;
    sheet.getRange(newRow, 1).setValue(name);
    sheet.getRange(newRow, 2).setValue(phone || '');
  } catch(e) {
    Logger.log('upsertContactMaster error: ' + e.message);
  }
}

// === 新規貸出登録 ===

function openLoanSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('新規貸出登録').setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function registerLoan(data) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOAN);
    var slipNo = getNextSlipNumber();
    var today  = todayStr();

    var typeStr    = data.mediaSets.map(function(m) { return m.type || ''; }).join(' | ');
    var mediaStr   = data.mediaSets.map(function(m) { return m.media; }).join(' | ');
    var issueStr   = data.mediaSets.map(function(m) { return m.issue; }).join(' | ');
    var themeStr   = data.mediaSets.map(function(m) { return m.theme; }).join(' | ');
    var relDateStr = data.mediaSets.map(function(m) { return m.releaseDate; }).join(' | ');
    var shootArr   = data.mediaSets.map(function(m) { return m.shootDate || ''; });
    var shootStr   = shootArr.some(function(d) { return !!d; }) ? shootArr.join(' | ') : '';

    var pickupDate = data.pickupDate || today;
    var partialRtn = data.partialReturnDate || '';

    var rows = [];
    for (var ii = 0; ii < data.items.length; ii++) {
      var item = data.items[ii];
      var pr = (item.price !== '' && item.price !== null && item.price !== undefined)
        ? Number(String(item.price).replace(/,/g, '')) : '';
      rows.push([
        slipNo,                      // A: 伝票番号
        today,                       // B: 起票日
        data.name,                   // C: お名前
        data.phone || '',            // D: 電話番号
        typeStr,                     // E: 種別
        mediaStr,                    // F: 媒体名
        issueStr,                    // G: 月号/放送局
        themeStr,                    // H: テーマ/着用者
        relDateStr,                  // I: 公開日
        shootStr,                    // J: 撮影日
        pickupDate,                  // K: ピックアップ日
        data.plannedReturnDate || '',// L: 返却予定日
        partialRtn,                  // M: 一部返却日
        item.mainBrand || '',        // N: メインブランド
        item.brand || '',            // O: ブランド
        item.itemCode || '',         // P: 品番
        item.colorCode || '',        // Q: カラー
        item.colorName || '',        // R: カラー名
        item.itemName || '',         // S: アイテム名
        pr,                          // T: 単価(税込)
        '未返却',                    // U: 返却ステータス
        '',                          // V: 返却日
        false,                       // W: 掲載済
        '',                          // X: 備考
        data.staff || ''             // Y: 起票者名
      ]);
    }
    if (rows.length === 0) return { success: false, message: 'アイテムが入力されていません。' };

    var startRow = getLastDataRow(sheet) + 1;
    sheet.getRange(startRow, 1, rows.length, COL_LOAN.TOTAL).setValues(rows);
    sheet.getRange(startRow, COL_LOAN.SLIP_NO,   rows.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, COL_LOAN.PHONE,     rows.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, COL_LOAN.ITEM_CODE, rows.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, COL_LOAN.COLOR,     rows.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, COL_LOAN.PUBLISHED, rows.length, 1).insertCheckboxes();

    upsertContactMaster(data.name, data.phone || '');

    return { success: true, slipNo: slipNo, message: buildLoanMessage(slipNo, data) };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

function buildLoanMessage(slipNo, data) {
  var uniqueMedia = [];
  var uniqueTheme = [];
  data.mediaSets.forEach(function(m) {
    if (m.media && uniqueMedia.indexOf(m.media) === -1) uniqueMedia.push(m.media);
    if (m.theme && uniqueTheme.indexOf(m.theme) === -1) uniqueTheme.push(m.theme);
  });

  var lines = [];
  lines.push('【BIGI PRESS ROOM 貸出伝票 No.' + slipNo + '】');
  lines.push('お名前：' + data.name + '様');
  lines.push('');
  lines.push('想定媒体：' + uniqueMedia.join(' / '));
  if (uniqueTheme.length > 0) lines.push('想定テーマ/着用者：' + uniqueTheme.join(' / '));

  var CIRCLED = ['①','②','③','④','⑤','⑥','⑦','⑧','⑨','⑩'];
  var shootDates = data.mediaSets.map(function(m) { return m.shootDate || ''; }).filter(Boolean);
  var dateStr = '';
  if (shootDates.length === 1) {
    dateStr += '撮影日：' + shootDates[0];
  } else if (shootDates.length > 1) {
    dateStr += '撮影日：' + shootDates.map(function(d, i) {
      return (CIRCLED[i] || (i + 1) + '.') + d;
    }).join(' / ');
  }
  if (data.pickupDate) dateStr += (dateStr ? '　' : '') + 'ピックアップ日：' + data.pickupDate;
  if (data.plannedReturnDate) dateStr += (dateStr ? '　' : '') + '返却予定日：' + data.plannedReturnDate;
  if (dateStr) lines.push(dateStr);
  lines.push('---');
  for (var i = 0; i < data.items.length; i++) {
    var item = data.items[i];
    var priceStr = (item.price !== '' && item.price !== undefined)
      ? ('¥' + Number(String(item.price).replace(/,/g, '')).toLocaleString()) : '';
    var brandPart = item.brand ? item.brand + ' ' : '';
    lines.push((i + 1) + '. ' + brandPart + (item.itemCode || '') + ' ' + (item.colorCode || '') + ' ' + (item.itemName || '') + ' ' + priceStr);
  }
  lines.push('---');
  lines.push('ご返却は10:00〜18:00（12:00〜13:00を除く）でお願いいたします。');
  lines.push('環境への配慮からハンガーカバーを再利用させていただく場合がございます。ご了承くださいませ。');
  return lines.join('\n');
}

// === 伝票編集 ===

function openMediaSetDialog() {
  var html = HtmlService.createHtmlOutputFromFile('MediaSetDialog')
    .setWidth(560).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, '伝票編集');
}

// === 伝票印刷 ===

function openPrintSlipDialog() {
  var html = HtmlService.createHtmlOutputFromFile('PrintSlipDialog')
    .setWidth(520).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, '貸出伝票 印刷');
}

function openPrintSlipDialogWithNo(slipNo) {
  PropertiesService.getScriptProperties().setProperty('PRINT_SLIP_NO', slipNo);
  var html = HtmlService.createHtmlOutputFromFile('PrintSlipDialog')
    .setWidth(520).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, '貸出伝票 印刷');
}

function getInitialSlipNo() {
  var no = PropertiesService.getScriptProperties().getProperty('PRINT_SLIP_NO') || '';
  PropertiesService.getScriptProperties().deleteProperty('PRINT_SLIP_NO');
  return no;
}

function getSlipDataAndCredits(slipNo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_LOAN);
    var s = String(slipNo);
    while (s.length < 5) s = '0' + s;

    var lastDataRow = getLastDataRow(sheet);
    if (lastDataRow <= 1) return { found: false };
    var data = sheet.getRange(2, 1, lastDataRow - 1, COL_LOAN.TOTAL).getValues();

    var rowCount = 0;
    var loanDate = '', name = '', phone = '';
    var pickupDate = '', plannedReturnDate = '', partialReturnDate = '';
    var mediaSets = [];
    var items = [];

    for (var i = 0; i < data.length; i++) {
      if (!data[i][0]) continue;
      var rowSlip = String(data[i][0]);
      while (rowSlip.length < 5) rowSlip = '0' + rowSlip;
      if (rowSlip !== s) continue;
      rowCount++;
      if (rowCount === 1) {
        loanDate          = cellToStr(data[i][COL_LOAN.DATE - 1]);
        name              = cellToStr(data[i][COL_LOAN.NAME - 1]);
        phone             = cellToStr(data[i][COL_LOAN.PHONE - 1]);
        pickupDate        = cellToStr(data[i][COL_LOAN.PICKUP_DATE - 1]);
        plannedReturnDate = cellToStr(data[i][COL_LOAN.RETURN_PLAN - 1]);
        partialReturnDate = cellToStr(data[i][COL_LOAN.PARTIAL_RTN - 1]);
        var types      = cellToStr(data[i][COL_LOAN.TYPE - 1]).split(' | ');
        var medias     = cellToStr(data[i][COL_LOAN.MEDIA - 1]).split(' | ');
        var issues     = cellToStr(data[i][COL_LOAN.ISSUE - 1]).split(' | ');
        var themes     = cellToStr(data[i][COL_LOAN.THEME - 1]).split(' | ');
        var relDates   = cellToStr(data[i][COL_LOAN.PUB_DATE - 1]).split(' | ');
        var shootDates = cellToStr(data[i][COL_LOAN.SHOOT_DATE - 1]).split(' | ');
        for (var j = 0; j < medias.length; j++) {
          if (!medias[j]) continue;
          mediaSets.push({
            type:        types[j]      || '',
            media:       medias[j],
            issue:       issues[j]     || '',
            theme:       themes[j]     || '',
            releaseDate: relDates[j]   || '',
            shootDate:   shootDates[j] || ''
          });
        }
      }
      items.push({
        mainBrand: cellToStr(data[i][COL_LOAN.MAIN_BRAND - 1]),
        brand:     cellToStr(data[i][COL_LOAN.BRAND - 1]),
        itemCode:  cellToStr(data[i][COL_LOAN.ITEM_CODE - 1]),
        colorCode: cellToStr(data[i][COL_LOAN.COLOR - 1]),
        colorName: cellToStr(data[i][COL_LOAN.COLOR_NAME - 1]),
        itemName:  cellToStr(data[i][COL_LOAN.ITEM_NAME - 1]),
        price:     data[i][COL_LOAN.PRICE - 1] !== '' ? Number(data[i][COL_LOAN.PRICE - 1]) : ''
      });
    }
    if (rowCount === 0) return { found: false };

    // クレジットマスタ
    var credits = [];
    var creditSheet = ss.getSheetByName(SHEET_CREDIT);
    if (creditSheet && creditSheet.getLastRow() > 1) {
      var creditData = creditSheet.getRange(2, 1, creditSheet.getLastRow() - 1, 6).getValues();
      creditData.forEach(function(cr) {
        if (!cr[0]) return;
        credits.push({
          name: cellToStr(cr[0]), kana: cellToStr(cr[1]), tel: cellToStr(cr[2]),
          web: cellToStr(cr[3]), onlineStore: cellToStr(cr[4]), instagram: cellToStr(cr[5])
        });
      });
    }

    // PRESS ROOM情報
    var pressRooms = [];
    var settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
    if (settingsSheet) {
      var prRange = settingsSheet.getRange(5, 2, 9, 1).getValues();
      if (prRange[0][0]) pressRooms.push({
        label: cellToStr(prRange[0][0]), address: cellToStr(prRange[1][0]),
        tel: cellToStr(prRange[2][0]), fax: cellToStr(prRange[3][0])
      });
      if (prRange[5][0]) pressRooms.push({
        label: cellToStr(prRange[5][0]), address: cellToStr(prRange[6][0]),
        tel: cellToStr(prRange[7][0]), fax: cellToStr(prRange[8][0])
      });
    }

    return {
      found: true, slipNo: s,
      loanDate: loanDate, name: name, phone: phone,
      pickupDate: pickupDate, plannedReturnDate: plannedReturnDate, partialReturnDate: partialReturnDate,
      mediaSets: mediaSets, items: items, credits: credits, pressRooms: pressRooms
    };
  } catch (e) {
    return { found: false, error: e.message };
  }
}

function getMediaSetsForSlip(slipNo) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOAN);
    var s = String(slipNo);
    while (s.length < 5) s = '0' + s;

    var lastDataRow = getLastDataRow(sheet);
    if (lastDataRow <= 1) return { found: false };
    var data = sheet.getRange(2, 1, lastDataRow - 1, COL_LOAN.SHOOT_DATE).getValues();

    var rowCount = 0;
    var mediaSets = [];
    for (var i = 0; i < data.length; i++) {
      if (!data[i][0]) continue;
      var rowSlip = String(data[i][0]);
      while (rowSlip.length < 5) rowSlip = '0' + rowSlip;
      if (rowSlip !== s) continue;
      rowCount++;
      if (rowCount === 1) {
        var types      = cellToStr(data[i][COL_LOAN.TYPE - 1]).split(' | ');
        var medias     = cellToStr(data[i][COL_LOAN.MEDIA - 1]).split(' | ');
        var issues     = cellToStr(data[i][COL_LOAN.ISSUE - 1]).split(' | ');
        var themes     = cellToStr(data[i][COL_LOAN.THEME - 1]).split(' | ');
        var relDates   = cellToStr(data[i][COL_LOAN.PUB_DATE - 1]).split(' | ');
        var shootDates = cellToStr(data[i][COL_LOAN.SHOOT_DATE - 1]).split(' | ');
        for (var j = 0; j < medias.length; j++) {
          if (!medias[j]) continue;
          mediaSets.push({
            type:        types[j]      || '',
            media:       medias[j],
            issue:       issues[j]     || '',
            theme:       themes[j]     || '',
            releaseDate: relDates[j]   || '',
            shootDate:   shootDates[j] || ''
          });
        }
      }
    }
    if (rowCount === 0) return { found: false };
    return { found: true, rowCount: rowCount, mediaSets: mediaSets };
  } catch (e) {
    return { found: false, error: e.message };
  }
}

// === クレジットマスタ ダミーデータ ===

function insertCreditDummyData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CREDIT);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('クレジットマスタシートが存在しません。先に setupSheets() を実行してください。');
    return;
  }
  if (sheet.getLastRow() > 1) {
    var ui = SpreadsheetApp.getUi();
    var res = ui.alert('既存データの確認', 'クレジットマスタにすでにデータがあります。上書きしますか？', ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) return;
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).clearContent();
  }
  var dummy = [
    ['UNTITLED',          'アンタイトル',         '03-6861-7600', 'https://untitled.co.jp',           'https://store.bigi.co.jp/untitled',    '@untitled_official_jp'],
    ["L'EQUIPE",          'レキップ',             '03-6861-7610', 'https://lequipe.bigi.co.jp',       'https://store.bigi.co.jp/lequipe',     '@lequipe_official'],
    ['MOGA',              'モガ',                 '03-6861-7620', 'https://moga.co.jp',               'https://store.bigi.co.jp/moga',        '@moga_official_jp'],
    ['DÉPAREILLÉE',       'デパリエ',             '03-6861-7630', 'https://depareillee.bigi.co.jp',   'https://store.bigi.co.jp/depareillee', '@depareillee_official'],
    ['ADIEU TRISTESSE',   'アデュー トリステス', '03-6861-7658', 'https://adieu-tristesse.jp',       'https://store.bigi.co.jp/adieu',       '@adieu_tristesse_official'],
    ['LOISIR',            'ロワジール',           '03-6861-7640', 'https://loisir.bigi.co.jp',        'https://store.bigi.co.jp/loisir',      '@loisir_official'],
    ['ef-de',             'エフデ',               '03-6861-7650', 'https://ef-de.jp',                 'https://store.bigi.co.jp/efde',        '@efde_official_jp'],
    ['NATURAL BEAUTY',    'ナチュラルビューティ', '03-6861-7660', 'https://naturalbeauty.bigi.co.jp', 'https://store.bigi.co.jp/nb',          '@naturalbeauty_official'],
    ['ICEBERG',           'アイスバーグ',         '03-6861-7670', 'https://iceberg.bigi.co.jp',       '',                                     '@iceberg_jp_official'],
    ['RENOMA',            'レノマ',               '03-6861-7680', 'https://renoma.bigi.co.jp',        '',                                     '']
  ];
  sheet.getRange(2, 1, dummy.length, 6).setValues(dummy);
  SpreadsheetApp.getUi().alert('完了', 'クレジットマスタに ' + dummy.length + ' 件のダミーデータを投入しました。', SpreadsheetApp.getUi().ButtonSet.OK);
}

function addMediaSetToSlip(params) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOAN);
    var s = String(params.slipNo);
    while (s.length < 5) s = '0' + s;

    var lastDataRow = getLastDataRow(sheet);
    if (lastDataRow <= 1) return { success: false, message: '該当データがありません。' };
    var data = sheet.getRange(2, 1, lastDataRow - 1, COL_LOAN.SHOOT_DATE).getValues();

    var updated = 0;
    for (var i = 0; i < data.length; i++) {
      if (!data[i][0]) continue;
      var rowSlip = String(data[i][0]);
      while (rowSlip.length < 5) rowSlip = '0' + rowSlip;
      if (rowSlip !== s) continue;
      var r = i + 2;
      var sep = ' | ';
      var append = function(col, val) {
        var current = cellToStr(sheet.getRange(r, col).getValue());
        sheet.getRange(r, col).setValue(current ? current + sep + val : val);
      };
      append(COL_LOAN.TYPE,       params.mediaSet.type || '');
      append(COL_LOAN.MEDIA,      params.mediaSet.media);
      append(COL_LOAN.ISSUE,      params.mediaSet.issue);
      append(COL_LOAN.THEME,      params.mediaSet.theme);
      append(COL_LOAN.PUB_DATE,   params.mediaSet.releaseDate);
      if (params.mediaSet.shootDate) append(COL_LOAN.SHOOT_DATE, params.mediaSet.shootDate);
      updated++;
    }
    if (updated === 0) return { success: false, message: '該当する伝票番号が見つかりませんでした。' };
    return { success: true, updated: updated };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

// === 返却処理 ===

function openReturnDialog() {
  var html = HtmlService.createHtmlOutputFromFile('ReturnDialog')
    .setWidth(650).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '返却処理');
}

function searchUnreturnedItems(slipNo) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOAN);
    var s = String(slipNo);
    while (s.length < 5) s = '0' + s;

    var lastDataRow = getLastDataRow(sheet);
    if (lastDataRow <= 1) return [];
    var data = sheet.getRange(2, 1, lastDataRow - 1, COL_LOAN.TOTAL).getValues();

    var results = [];
    for (var i = 0; i < data.length; i++) {
      if (!data[i][0]) continue;
      var rowSlip = String(data[i][0]);
      while (rowSlip.length < 5) rowSlip = '0' + rowSlip;
      if (rowSlip === s && data[i][COL_LOAN.STATUS - 1] === '未返却') {
        results.push({
          rowIndex:   i + 2,
          mainBrand:  cellToStr(data[i][COL_LOAN.MAIN_BRAND - 1]),
          brand:      cellToStr(data[i][COL_LOAN.BRAND - 1]),
          itemCode:   cellToStr(data[i][COL_LOAN.ITEM_CODE - 1]),
          colorCode:  cellToStr(data[i][COL_LOAN.COLOR - 1]),
          colorName:  cellToStr(data[i][COL_LOAN.COLOR_NAME - 1]),
          itemName:   cellToStr(data[i][COL_LOAN.ITEM_NAME - 1]),
          price:      data[i][COL_LOAN.PRICE - 1] !== '' ? Number(data[i][COL_LOAN.PRICE - 1]) : 0,
          stylist:    cellToStr(data[i][COL_LOAN.NAME - 1]),
          mediaStr:   cellToStr(data[i][COL_LOAN.MEDIA - 1]),
          issueStr:   cellToStr(data[i][COL_LOAN.ISSUE - 1]),
          themeStr:   cellToStr(data[i][COL_LOAN.THEME - 1]),
          relDateStr: cellToStr(data[i][COL_LOAN.PUB_DATE - 1]),
          typeStr:    cellToStr(data[i][COL_LOAN.TYPE - 1]),
          staff:      cellToStr(data[i][COL_LOAN.STAFF - 1])
        });
      }
    }
    return results;
  } catch (e) {
    Logger.log('searchError: ' + e.message);
    return [];
  }
}

function processReturn(params) {
  try {
    var loanSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOAN);
    for (var i = 0; i < params.rows.length; i++) {
      var r = params.rows[i].rowIndex;
      loanSheet.getRange(r, COL_LOAN.STATUS).setValue('返却済');
      loanSheet.getRange(r, COL_LOAN.RETURN_DATE).setValue(params.returnDate);
    }
    return { success: true };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

function processPublish(params) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var loanSheet = ss.getSheetByName(SHEET_LOAN);
    var pubSheet  = ss.getSheetByName(SHEET_PUBLISH);
    var pubLastRow = pubSheet.getLastRow();
    for (var i = 0; i < params.combinations.length; i++) {
      var c = params.combinations[i];
      loanSheet.getRange(c.rowIndex, COL_LOAN.PUBLISHED).setValue(true);
      var slipNo    = cellToStr(loanSheet.getRange(c.rowIndex, COL_LOAN.SLIP_NO).getValue());
      var mainBrand = cellToStr(loanSheet.getRange(c.rowIndex, COL_LOAN.MAIN_BRAND).getValue());
      var brand     = cellToStr(loanSheet.getRange(c.rowIndex, COL_LOAN.BRAND).getValue());
      var colorName = cellToStr(loanSheet.getRange(c.rowIndex, COL_LOAN.COLOR_NAME).getValue());
      pubLastRow++;
      pubSheet.getRange(pubLastRow, 1, 1, COL_PUB.TOTAL).setValues([[
        slipNo,                       // A: 伝票番号
        c.type,                       // B: 種別
        c.media,                      // C: 媒体名
        c.issue,                      // D: 月号/放送局
        c.releaseDate,                // E: 公開日
        c.theme,                      // F: テーマ/着用者
        c.stylist,                    // G: スタイリスト
        mainBrand,                    // H: メインブランド
        brand || c.brand || '',       // I: ブランド名
        c.itemCode,                   // J: 品番
        c.colorCode,                  // K: カラー
        colorName,                    // L: カラー名
        c.price,                      // M: 単価(税込)
        c.itemName,                   // N: アイテム
        '',                           // O: 頁
        ''                            // P: 画像/リンク
      ]]);
    }
    return { success: true };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

function processReturnAndPublish(params) {
  try {
    var retResult = processReturn(params.returnData);
    if (!retResult.success) return retResult;
    if (params.publishData && params.publishData.combinations && params.publishData.combinations.length > 0) {
      var pubResult = processPublish(params.publishData);
      if (!pubResult.success) return pubResult;
    }
    return { success: true };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

function getPublishDataForMonth(year, month) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PUBLISH);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, COL_PUB.TOTAL).getValues();
  var results = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[COL_PUB.MEDIA - 1] && !row[COL_PUB.ISSUE - 1]) continue;
    var pubDate = row[COL_PUB.PUB_DATE - 1];
    var matched = false;
    if (pubDate) {
      var d = new Date(pubDate);
      if (d.getFullYear() === year && (d.getMonth() + 1) === month) matched = true;
    } else {
      matched = true;
    }
    if (matched) {
      var rdStr = pubDate ? Utilities.formatDate(new Date(pubDate), Session.getScriptTimeZone(), 'yyyy/MM/dd') : '';
      results.push({
        mainBrand:   cellToStr(row[COL_PUB.MAIN_BRAND - 1]),
        brand:       cellToStr(row[COL_PUB.BRAND - 1]),
        media:       cellToStr(row[COL_PUB.MEDIA - 1]),
        type:        cellToStr(row[COL_PUB.TYPE - 1]),
        issue:       cellToStr(row[COL_PUB.ISSUE - 1]),
        releaseDate: rdStr,
        theme:       cellToStr(row[COL_PUB.THEME - 1]),
        stylist:     cellToStr(row[COL_PUB.STYLIST - 1]),
        itemCode:    cellToStr(row[COL_PUB.ITEM_CODE - 1]),
        colorCode:   cellToStr(row[COL_PUB.COLOR - 1]),
        colorName:   cellToStr(row[COL_PUB.COLOR_NAME - 1]),
        price:       row[COL_PUB.PRICE - 1] !== '' ? Number(row[COL_PUB.PRICE - 1]) : 0,
        itemName:    cellToStr(row[COL_PUB.ITEM_NAME - 1]),
        page:        cellToStr(row[COL_PUB.PAGE - 1]),
        imageLink:   cellToStr(row[COL_PUB.IMAGE - 1])
      });
    }
  }
  return results;
}

function getSlipDataForEdit(slipNo) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOAN);
    var s = String(slipNo);
    while (s.length < 5) s = '0' + s;

    var lastDataRow = getLastDataRow(sheet);
    if (lastDataRow <= 1) return { found: false };
    var data = sheet.getRange(2, 1, lastDataRow - 1, COL_LOAN.TOTAL).getValues();

    var rowCount = 0, returnedCount = 0, publishedCount = 0;
    var name = '', phone = '', pickupDate = '', plannedReturnDate = '', partialReturnDate = '', staff = '';
    var mediaSets = [];

    for (var i = 0; i < data.length; i++) {
      if (!data[i][0]) continue;
      var rowSlip = String(data[i][0]);
      while (rowSlip.length < 5) rowSlip = '0' + rowSlip;
      if (rowSlip !== s) continue;
      rowCount++;
      if (data[i][COL_LOAN.STATUS - 1] === '返却済') returnedCount++;
      if (data[i][COL_LOAN.PUBLISHED - 1] === true) publishedCount++;
      if (rowCount === 1) {
        name              = cellToStr(data[i][COL_LOAN.NAME - 1]);
        phone             = cellToStr(data[i][COL_LOAN.PHONE - 1]);
        pickupDate        = cellToStr(data[i][COL_LOAN.PICKUP_DATE - 1]);
        plannedReturnDate = cellToStr(data[i][COL_LOAN.RETURN_PLAN - 1]);
        partialReturnDate = cellToStr(data[i][COL_LOAN.PARTIAL_RTN - 1]);
        staff             = cellToStr(data[i][COL_LOAN.STAFF - 1]);
        var types      = cellToStr(data[i][COL_LOAN.TYPE - 1]).split(' | ');
        var medias     = cellToStr(data[i][COL_LOAN.MEDIA - 1]).split(' | ');
        var issues     = cellToStr(data[i][COL_LOAN.ISSUE - 1]).split(' | ');
        var themes     = cellToStr(data[i][COL_LOAN.THEME - 1]).split(' | ');
        var relDates   = cellToStr(data[i][COL_LOAN.PUB_DATE - 1]).split(' | ');
        var shootDates = cellToStr(data[i][COL_LOAN.SHOOT_DATE - 1]).split(' | ');
        for (var j = 0; j < medias.length; j++) {
          if (!medias[j]) continue;
          mediaSets.push({
            type:        types[j]      || '',
            media:       medias[j],
            issue:       issues[j]     || '',
            theme:       themes[j]     || '',
            releaseDate: relDates[j]   || '',
            shootDate:   shootDates[j] || ''
          });
        }
      }
    }
    if (rowCount === 0) return { found: false };
    return {
      found: true, rowCount: rowCount, returnedCount: returnedCount, publishedCount: publishedCount,
      name: name, phone: phone, pickupDate: pickupDate,
      plannedReturnDate: plannedReturnDate, partialReturnDate: partialReturnDate, staff: staff,
      mediaSets: mediaSets
    };
  } catch (e) {
    return { found: false, error: e.message };
  }
}

function updateSlipData(slipNo, params) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOAN);
    var s = String(slipNo);
    while (s.length < 5) s = '0' + s;

    var lastDataRow = getLastDataRow(sheet);
    if (lastDataRow <= 1) return { success: false, message: '該当データがありません。' };
    var data = sheet.getRange(2, 1, lastDataRow - 1, 1).getValues();

    var typeStr      = params.mediaSets.map(function(m) { return m.type || ''; }).join(' | ');
    var mediaStr     = params.mediaSets.map(function(m) { return m.media; }).join(' | ');
    var issueStr     = params.mediaSets.map(function(m) { return m.issue; }).join(' | ');
    var themeStr     = params.mediaSets.map(function(m) { return m.theme; }).join(' | ');
    var relDateStr   = params.mediaSets.map(function(m) { return m.releaseDate; }).join(' | ');
    var shootArr     = params.mediaSets.map(function(m) { return m.shootDate || ''; });
    var shootDateStr = shootArr.some(function(d) { return !!d; }) ? shootArr.join(' | ') : '';

    var updated = 0;
    for (var i = 0; i < data.length; i++) {
      if (!data[i][0]) continue;
      var rowSlip = String(data[i][0]);
      while (rowSlip.length < 5) rowSlip = '0' + rowSlip;
      if (rowSlip !== s) continue;
      var r = i + 2;
      sheet.getRange(r, COL_LOAN.NAME).setValue(params.name);
      sheet.getRange(r, COL_LOAN.PHONE).setValue(params.phone || '');
      sheet.getRange(r, COL_LOAN.TYPE).setValue(typeStr);
      sheet.getRange(r, COL_LOAN.MEDIA).setValue(mediaStr);
      sheet.getRange(r, COL_LOAN.ISSUE).setValue(issueStr);
      sheet.getRange(r, COL_LOAN.THEME).setValue(themeStr);
      sheet.getRange(r, COL_LOAN.PUB_DATE).setValue(relDateStr);
      sheet.getRange(r, COL_LOAN.SHOOT_DATE).setValue(shootDateStr);
      sheet.getRange(r, COL_LOAN.PICKUP_DATE).setValue(params.pickupDate || '');
      sheet.getRange(r, COL_LOAN.RETURN_PLAN).setValue(params.plannedReturnDate || '');
      sheet.getRange(r, COL_LOAN.PARTIAL_RTN).setValue(params.partialReturnDate || '');
      sheet.getRange(r, COL_LOAN.STAFF).setValue(params.staff || '');
      updated++;
    }
    if (updated === 0) return { success: false, message: '該当する伝票番号が見つかりませんでした。' };
    return { success: true, updated: updated };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

function getPublishRowsForSlip(slipNo) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PUBLISH);
    var s = String(slipNo);
    while (s.length < 5) s = '0' + s;
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    var data = sheet.getRange(2, 1, lastRow - 1, COL_PUB.TOTAL).getValues();
    var results = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      var rowSlip = String(row[0]);
      while (rowSlip.length < 5) rowSlip = '0' + rowSlip;
      if (rowSlip !== s) continue;
      var pubDate = row[COL_PUB.PUB_DATE - 1];
      var rdStr = pubDate ? Utilities.formatDate(new Date(pubDate), Session.getScriptTimeZone(), 'yyyy/MM/dd') : '';
      results.push({
        rowIndex:    i + 2,
        slipNo:      cellToStr(row[COL_PUB.SLIP_NO - 1]),
        type:        cellToStr(row[COL_PUB.TYPE - 1]),
        media:       cellToStr(row[COL_PUB.MEDIA - 1]),
        issue:       cellToStr(row[COL_PUB.ISSUE - 1]),
        releaseDate: rdStr,
        theme:       cellToStr(row[COL_PUB.THEME - 1]),
        stylist:     cellToStr(row[COL_PUB.STYLIST - 1]),
        mainBrand:   cellToStr(row[COL_PUB.MAIN_BRAND - 1]),
        brand:       cellToStr(row[COL_PUB.BRAND - 1]),
        itemCode:    cellToStr(row[COL_PUB.ITEM_CODE - 1]),
        colorCode:   cellToStr(row[COL_PUB.COLOR - 1]),
        colorName:   cellToStr(row[COL_PUB.COLOR_NAME - 1]),
        price:       row[COL_PUB.PRICE - 1] !== '' ? Number(row[COL_PUB.PRICE - 1]) : 0,
        itemName:    cellToStr(row[COL_PUB.ITEM_NAME - 1]),
        page:        cellToStr(row[COL_PUB.PAGE - 1]),
        imageLink:   cellToStr(row[COL_PUB.IMAGE - 1])
      });
    }
    return results;
  } catch (e) {
    return [];
  }
}

function updatePublishRow(rowIndex, data) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PUBLISH);
    sheet.getRange(rowIndex, 1, 1, COL_PUB.TOTAL).setValues([[
      data.slipNo, data.type, data.media, data.issue, data.releaseDate,
      data.theme, data.stylist, data.mainBrand || '', data.brand || '',
      data.itemCode, data.colorCode, data.colorName || '', data.price,
      data.itemName, data.page, data.imageLink
    ]]);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deletePublishRow(rowIndex) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PUBLISH);
    sheet.deleteRow(rowIndex);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function openPublishListEditDialog() {
  var html = HtmlService.createHtmlOutputFromFile('PublishListEditDialog')
    .setWidth(900).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, '掲載リストマスタ編集');
}

function revertToUnreturned(slipNo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var loanSheet = ss.getSheetByName(SHEET_LOAN);
    var pubSheet  = ss.getSheetByName(SHEET_PUBLISH);
    var s = String(slipNo);
    while (s.length < 5) s = '0' + s;

    var lastDataRow = getLastDataRow(loanSheet);
    var revertedCount = 0;
    if (lastDataRow > 1) {
      var loanData = loanSheet.getRange(2, 1, lastDataRow - 1, 1).getValues();
      for (var i = 0; i < loanData.length; i++) {
        if (!loanData[i][0]) continue;
        var rowSlip = String(loanData[i][0]);
        while (rowSlip.length < 5) rowSlip = '0' + rowSlip;
        if (rowSlip !== s) continue;
        var r = i + 2;
        loanSheet.getRange(r, COL_LOAN.STATUS).setValue('未返却');
        loanSheet.getRange(r, COL_LOAN.RETURN_DATE).setValue('');
        loanSheet.getRange(r, COL_LOAN.PUBLISHED).setValue(false);
        revertedCount++;
      }
    }

    var deletedPublishCount = 0;
    var pubLastRow = pubSheet.getLastRow();
    if (pubLastRow > 1) {
      var pubData = pubSheet.getRange(2, 1, pubLastRow - 1, 1).getValues();
      var rowsToDelete = [];
      for (var j = 0; j < pubData.length; j++) {
        if (!pubData[j][0]) continue;
        var rowSlipPub = String(pubData[j][0]);
        while (rowSlipPub.length < 5) rowSlipPub = '0' + rowSlipPub;
        if (rowSlipPub === s) rowsToDelete.push(j + 2);
      }
      for (var k = rowsToDelete.length - 1; k >= 0; k--) {
        pubSheet.deleteRow(rowsToDelete[k]);
        deletedPublishCount++;
      }
    }
    return { success: true, revertedCount: revertedCount, deletedPublishCount: deletedPublishCount };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

function testGetLastDataRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOAN);
  var row = getLastDataRow(sheet);
  Logger.log('最終データ行: ' + row);
}

// === 掲載リスト作成 ===

function getBrandName() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  if (!sheet) return 'BIGI PRESS ROOM';
  var v = sheet.getRange('B3').getValue();
  return v ? String(v) : 'BIGI PRESS ROOM';
}

function openFormattedSheetDialog() {
  var html = HtmlService.createHtmlOutputFromFile('FormattedSheetDialog')
    .setWidth(380).setHeight(280);
  SpreadsheetApp.getUi().showModalDialog(html, '掲載リスト作成');
}

function generateFormattedSheet(params) {
  try {
    var year      = Number(params.year);
    var mainBrand = String(params.mainBrand).trim();

    var pubSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PUBLISH);
    var lastRow  = pubSheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: '掲載リストマスタにデータがありません。' };
    var rawData = pubSheet.getRange(2, 1, lastRow - 1, COL_PUB.TOTAL).getValues();

    // メインブランド＋年でフィルタリング
    var items = [];
    for (var i = 0; i < rawData.length; i++) {
      var row = rawData[i];
      if (!row[COL_PUB.MEDIA - 1] && !row[COL_PUB.ISSUE - 1]) continue;
      if (String(row[COL_PUB.MAIN_BRAND - 1]).trim() !== mainBrand) continue;
      var pubDate = row[COL_PUB.PUB_DATE - 1];
      if (!pubDate) continue;
      var d = new Date(pubDate);
      if (d.getFullYear() !== year) continue;
      var rdStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      items.push({
        month:      d.getMonth() + 1,
        type:       cellToStr(row[COL_PUB.TYPE - 1]),
        media:      cellToStr(row[COL_PUB.MEDIA - 1]),
        issue:      cellToStr(row[COL_PUB.ISSUE - 1]),
        releaseDate: rdStr,
        theme:      cellToStr(row[COL_PUB.THEME - 1]),
        stylist:    cellToStr(row[COL_PUB.STYLIST - 1]),
        itemCode:   cellToStr(row[COL_PUB.ITEM_CODE - 1]),
        colorName:  cellToStr(row[COL_PUB.COLOR_NAME - 1]),
        price:      row[COL_PUB.PRICE - 1] !== '' ? Number(row[COL_PUB.PRICE - 1]) : '',
        itemName:   cellToStr(row[COL_PUB.ITEM_NAME - 1]),
        page:       cellToStr(row[COL_PUB.PAGE - 1]),
        imageLink:  cellToStr(row[COL_PUB.IMAGE - 1])
      });
    }
    if (items.length === 0) return { success: false, message: '対象データが0件でした。' };

    var sheetName = mainBrand + '_' + year + '年_掲載リスト';
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var outSheet = ss.getSheetByName(sheetName);
    if (outSheet) {
      outSheet.clearContents();
      outSheet.clearFormats();
    } else {
      outSheet = ss.insertSheet(sheetName, ss.getSheets().length);
    }

    var colWidths = [120, 90, 55, 200, 100, 100, 100, 70, 200, 50, 150];
    for (var ci = 0; ci < colWidths.length; ci++) {
      outSheet.setColumnWidth(ci + 1, colWidths[ci]);
    }

    // 必須表示列（0-based index in row array）: 品番(5)・カラー名(6)・単価(7)・アイテム(8)・頁(9)・画像/リンク(10)
    var ALWAYS_SHOW = [5, 6, 7, 8, 9, 10];

    var currentRow = 1;

    for (var m = 1; m <= 12; m++) {
      var monthItems = items.filter(function(it) { return it.month === m; });
      if (monthItems.length === 0) continue;
      var monthLabel = m + '月';

      for (var ti = 0; ti < TYPE_ORDER.length; ti++) {
        var type = TYPE_ORDER[ti];
        var typeItems = monthItems.filter(function(it) { return it.type === type; });
        if (typeItems.length === 0) continue;

        typeItems.sort(function(a, b) {
          return new Date(a.releaseDate) - new Date(b.releaseDate);
        });

        var config = TYPE_CONFIG[type];

        // タイトル行（オレンジ）
        outSheet.getRange(currentRow, 1, 1, 11).setBackground('#FF9900');
        outSheet.getRange(currentRow, 1).setValue(mainBrand);
        outSheet.getRange(currentRow, 4).setValue(monthLabel);
        outSheet.getRange(currentRow, 8).setValue(config.title);
        currentRow++;

        // ヘッダー行（薄オレンジ）
        var headers = [config.col1, '掲載号', '発売日', 'テーマ', 'スタイリスト', '品番', 'カラー名', '単価(税込)', 'アイテム', '頁', '画像/リンク'];
        outSheet.getRange(currentRow, 1, 1, 11).setValues([headers]).setBackground('#FFD9B3').setFontWeight('bold');
        currentRow++;

        var prevMedia  = null;
        var prevDedup  = [null, null, null, null, null];

        for (var di = 0; di < typeItems.length; di++) {
          var item = typeItems[di];
          if (prevMedia !== null && prevMedia !== item.media) {
            currentRow++;
            prevDedup = [null, null, null, null, null];
          }
          var relDateMD = '';
          if (item.releaseDate) {
            var dd = new Date(item.releaseDate);
            relDateMD = (dd.getMonth() + 1) + '/' + dd.getDate();
          }
          var row = [
            item.media, item.issue, relDateMD, item.theme, item.stylist,
            item.itemCode, item.colorName, item.price !== '' ? item.price : '',
            item.itemName, item.page, item.imageLink
          ];
          var outputRow = row.map(function(v, idx) {
            if (ALWAYS_SHOW.indexOf(idx) !== -1) return v;
            if (v !== '' && v !== null && v !== undefined && String(v) === String(prevDedup[idx])) return '';
            return v;
          });
          prevDedup  = [row[0], row[1], row[2], row[3], row[4]];
          prevMedia  = item.media;
          outSheet.getRange(currentRow, 1, 1, 11).setValues([outputRow]);
          if (row[7] !== '') outSheet.getRange(currentRow, 8).setNumberFormat('¥#,##0');
          currentRow++;
        }
      }
    }

    return { success: true, sheetName: sheetName, rowCount: currentRow - 1 };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}
