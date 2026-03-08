// BIGI PRESS ROOM アパレル貸出管理システム

var SHEET_LOAN = '貸出管理';
var SHEET_PUBLISH = '掲載リスト';
var SHEET_SETTINGS = '設定';
var START_SLIP_NO = 8877;

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
    .addItem('返却処理', 'openReturnDialog')
    .addItem('媒体セット追加', 'openMediaSetDialog')
    .addSeparator()
    .addItem('整形シート生成', 'openFormattedSheetDialog')
    .addToUi();
}

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var loanSheet = ss.getSheetByName(SHEET_LOAN);
  if (!loanSheet) loanSheet = ss.insertSheet(SHEET_LOAN);

  var loanHeaders = [
    '伝票番号', '貸出日', 'スタイリスト名', '電話番号',
    '媒体名', '月号', 'テーマ/着用者', '公開日', '種別',
    '撮影日', '返却予定日',
    'ブランド', '品番', '色番', 'アイテム名', '単価',
    '返却ステータス', '返却日', '掲載済', '備考'
  ];
  loanSheet.getRange(1, 1, 1, loanHeaders.length).setValues([loanHeaders]);
  loanSheet.getRange(1, 1, 1, loanHeaders.length)
    .setBackground('#4a4a4a').setFontColor('#ffffff').setFontWeight('bold');
  loanSheet.setFrozenRows(1);

  var colWidths = [80, 85, 120, 110, 120, 70, 180, 100, 100, 85, 95, 100, 100, 70, 180, 80, 90, 85, 65, 150];
  for (var ci = 0; ci < colWidths.length; ci++) {
    loanSheet.setColumnWidth(ci + 1, colWidths[ci]);
  }

  // 条件付き書式: 未返却=赤のみ（1ルール）
  var cfRange = loanSheet.getRange(2, 1, 1000, 20);
  loanSheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$Q2="未返却"')
      .setBackground('#ffcccc')
      .setRanges([cfRange])
      .build()
  ]);

  var pubSheet = ss.getSheetByName(SHEET_PUBLISH);
  if (!pubSheet) pubSheet = ss.insertSheet(SHEET_PUBLISH);
  var pubHeaders = ['雑誌名', '種別', '掲載号', '公開日', 'テーマ/着用者', 'スタイリスト', 'ブランド名', '品番', '色番', '上代', 'アイテム', '頁', '画像/リンク'];
  pubSheet.getRange(1, 1, 1, pubHeaders.length).setValues([pubHeaders]);
  pubSheet.getRange(1, 1, 1, pubHeaders.length)
    .setBackground('#4a4a4a').setFontColor('#ffffff').setFontWeight('bold');
  pubSheet.setFrozenRows(1);

  var settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SHEET_SETTINGS);
  }
  // ブランド名セル（既存設定シートにも追記）
  if (!settingsSheet.getRange('A3').getValue()) {
    settingsSheet.getRange('A3').setValue('ブランド名').setFontWeight('bold');
  }
  if (!settingsSheet.getRange('B3').getValue()) {
    settingsSheet.getRange('B3').setValue("L'EQUIPE");
  }

  Logger.log('setup done');
}

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

function openLoanSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('新規貸出登録').setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function registerLoan(data) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOAN);
    var slipNo = getNextSlipNumber();
    var today = todayStr();

    var mediaStr   = data.mediaSets.map(function(m) { return m.media; }).join(' | ');
    var issueStr   = data.mediaSets.map(function(m) { return m.issue; }).join(' | ');
    var themeStr   = data.mediaSets.map(function(m) { return m.theme; }).join(' | ');
    var relDateStr = data.mediaSets.map(function(m) { return m.releaseDate; }).join(' | ');
    var typeStr    = data.mediaSets.map(function(m) { return m.type || ''; }).join(' | ');

    var rows = [];
    for (var ii = 0; ii < data.items.length; ii++) {
      var item = data.items[ii];
      var pr = item.price ? Number(String(item.price).replace(/,/g, '')) : '';
      rows.push([
        slipNo, today, data.stylist, data.phone,
        mediaStr, issueStr, themeStr, relDateStr, typeStr,  // E-I
        data.shootDate, data.plannedReturnDate,              // J-K
        item.brand, item.itemCode, item.colorCode, item.itemName, pr,  // L-P
        '未返却', '', false, ''                              // Q-T
      ]);
    }
    if (rows.length === 0) {
      return { success: false, message: 'アイテムが入力されていません。' };
    }

    var startRow = getLastDataRow(sheet) + 1;
    sheet.getRange(startRow, 1, rows.length, 20).setValues(rows);
    // 0落ち防止: 伝票番号(A)・電話番号(D)・品番(M)・色番(N) をテキスト形式に
    sheet.getRange(startRow, 1,  rows.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, 4,  rows.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, 13, rows.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, 14, rows.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, 19, rows.length, 1).insertCheckboxes();

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
  lines.push('【BIGI PRESS ROOM 貸出内容 No.' + slipNo + '】');
  lines.push('スタイリスト：' + data.stylist + '様');
  lines.push('');
  lines.push('想定媒体：' + uniqueMedia.join(' / '));
  if (uniqueTheme.length > 0) lines.push('想定テーマ/着用者：' + uniqueTheme.join(' / '));
  var dateStr = '';
  if (data.shootDate) dateStr += '撮影日：' + data.shootDate;
  if (data.plannedReturnDate) dateStr += (dateStr ? '\u3000' : '') + '返却予定日：' + data.plannedReturnDate;
  if (dateStr) lines.push(dateStr);
  lines.push('---');
  for (var i = 0; i < data.items.length; i++) {
    var item = data.items[i];
    var priceStr = item.price ? ('¥' + Number(String(item.price).replace(/,/g, '')).toLocaleString()) : '';
    lines.push((i + 1) + '. ' + (item.brand ? item.brand + ' ' : '') + item.itemCode + ' ' + item.colorCode + ' ' + item.itemName + ' ' + priceStr);
  }
  lines.push('---');
  lines.push('ご返却の際はご連絡ください。');
  return lines.join('\n');
}

function openMediaSetDialog() {
  var html = HtmlService.createHtmlOutputFromFile('MediaSetDialog')
    .setWidth(500).setHeight(480);
  SpreadsheetApp.getUi().showModalDialog(html, '媒体セット追加');
}

function getMediaSetsForSlip(slipNo) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOAN);
    var s = String(slipNo);
    while (s.length < 5) s = '0' + s;

    var lastDataRow = getLastDataRow(sheet);
    if (lastDataRow <= 1) return { found: false };
    var data = sheet.getRange(2, 1, lastDataRow - 1, 9).getValues();

    var rowCount = 0;
    var mediaSets = [];
    for (var i = 0; i < data.length; i++) {
      if (!data[i][0]) continue;
      var rowSlip = String(data[i][0]);
      while (rowSlip.length < 5) rowSlip = '0' + rowSlip;
      if (rowSlip !== s) continue;
      rowCount++;
      // 先頭行の媒体セットのみ解析（全行同じ）
      if (rowCount === 1) {
        var medias    = cellToStr(data[i][4]).split(' | ');
        var issues    = cellToStr(data[i][5]).split(' | ');
        var themes    = cellToStr(data[i][6]).split(' | ');
        var relDates  = cellToStr(data[i][7]).split(' | ');
        var types     = cellToStr(data[i][8]).split(' | ');
        for (var j = 0; j < medias.length; j++) {
          if (!medias[j]) continue;
          mediaSets.push({
            media:       medias[j],
            issue:       issues[j]    || '',
            theme:       themes[j]    || '',
            releaseDate: relDates[j]  || '',
            type:        types[j]     || ''
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

function addMediaSetToSlip(params) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOAN);
    var s = String(params.slipNo);
    while (s.length < 5) s = '0' + s;

    var lastDataRow = getLastDataRow(sheet);
    if (lastDataRow <= 1) return { success: false, message: '該当データがありません。' };
    var data = sheet.getRange(2, 1, lastDataRow - 1, 9).getValues();

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
      append(5, params.mediaSet.media);
      append(6, params.mediaSet.issue);
      append(7, params.mediaSet.theme);
      append(8, params.mediaSet.releaseDate);
      append(9, params.mediaSet.type || '');
      updated++;
    }
    if (updated === 0) return { success: false, message: '該当する伝票番号が見つかりませんでした。' };
    return { success: true, updated: updated };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}

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
    // A〜Q列（17列）を読み込む
    var data = sheet.getRange(2, 1, lastDataRow - 1, 17).getValues();

    var results = [];
    for (var i = 0; i < data.length; i++) {
      if (!data[i][0]) continue;
      var rowSlip = String(data[i][0]);
      while (rowSlip.length < 5) rowSlip = '0' + rowSlip;
      // Q列(index16)が「未返却」の行のみ
      if (rowSlip === s && data[i][16] === '未返却') {
        results.push({
          rowIndex:   i + 2,
          brand:      cellToStr(data[i][11]),
          itemCode:   cellToStr(data[i][12]),
          colorCode:  cellToStr(data[i][13]),
          itemName:   cellToStr(data[i][14]),
          price:      data[i][15] ? Number(data[i][15]) : 0,
          stylist:    cellToStr(data[i][2]),
          mediaStr:   cellToStr(data[i][4]),
          issueStr:   cellToStr(data[i][5]),
          themeStr:   cellToStr(data[i][6]),
          relDateStr: cellToStr(data[i][7]),
          typeStr:    cellToStr(data[i][8])
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
      loanSheet.getRange(r, 17).setValue('返却済');       // Q列
      loanSheet.getRange(r, 18).setValue(params.returnDate); // R列
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
    var pubSheet = ss.getSheetByName(SHEET_PUBLISH);
    for (var i = 0; i < params.combinations.length; i++) {
      var c = params.combinations[i];
      loanSheet.getRange(c.rowIndex, 19).setValue(true); // S列（掲載済）
      var pubLastRow = pubSheet.getLastRow() + 1;
      pubSheet.getRange(pubLastRow, 1, 1, 13).setValues([[
        c.media, c.type, c.issue, c.releaseDate, c.theme, c.stylist,
        c.brand || '', c.itemCode, c.colorCode, c.price, c.itemName, '', ''
      ]]);
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
  var data = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  var results = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[0] && !row[2]) continue; // A=雑誌名, C=掲載号 で空行判定
    var pubDate = row[3]; // D列: 公開日
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
        media:       cellToStr(row[0]),
        type:        cellToStr(row[1]),
        issue:       cellToStr(row[2]),
        releaseDate: rdStr,
        theme:       cellToStr(row[4]),
        stylist:     cellToStr(row[5]),
        brand:       cellToStr(row[6]),
        itemCode:    cellToStr(row[7]),
        colorCode:   cellToStr(row[8]),
        price:       row[9] ? Number(row[9]) : 0,
        itemName:    cellToStr(row[10]),
        page:        cellToStr(row[11]),
        imageLink:   cellToStr(row[12])
      });
    }
  }
  return results;
}

function testGetLastDataRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOAN);
  var row = getLastDataRow(sheet);
  Logger.log('最終データ行: ' + row);
}

// ===== 整形用シート生成 =====

function getBrandName() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  if (!sheet) return 'BIGI PRESS ROOM';
  var v = sheet.getRange('B3').getValue();
  return v ? String(v) : 'BIGI PRESS ROOM';
}

function openFormattedSheetDialog() {
  var html = HtmlService.createHtmlOutputFromFile('FormattedSheetDialog')
    .setWidth(400).setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, '整形シート生成');
}

function generateFormattedSheet(params) {
  try {
    var year  = Number(params.year);
    var month = Number(params.month);
    var splitByBrand = params.splitByBrand === true;
    var items = getPublishDataForMonth(year, month);
    var brandName = getBrandName();
    var monthLabel = month + '月';
    var sheetName = '整形用_' + year + '年' + month + '月';

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var outSheet = ss.getSheetByName(sheetName);
    if (outSheet) {
      outSheet.clearContents();
      outSheet.clearFormats();
    } else {
      outSheet = ss.insertSheet(sheetName);
    }

    var currentRow = 1;

    if (!splitByBrand) {
      // ===== 分割なし（既存処理・10列）=====
      var colWidths = [120, 70, 55, 200, 100, 90, 90, 70, 200, 50];

      TYPE_ORDER.forEach(function(type) {
        var typeItems = items.filter(function(it) { return it.type === type; });
        if (typeItems.length === 0) return;

        typeItems.sort(function(a, b) {
          var da = a.releaseDate ? new Date(a.releaseDate) : new Date(0);
          var db = b.releaseDate ? new Date(b.releaseDate) : new Date(0);
          return da - db;
        });

        var config = TYPE_CONFIG[type];

        var h1Range = outSheet.getRange(currentRow, 1, 1, 10);
        h1Range.setBackground('#FF9900');
        outSheet.getRange(currentRow, 1).setValue(brandName);
        outSheet.getRange(currentRow, 4).setValue(monthLabel);
        outSheet.getRange(currentRow, 8).setValue(config.title);
        currentRow++;

        var headers = [config.col1, '掲載号', '発売日', 'テーマ', 'スタイリスト', '品番', '色番', '上代', 'アイテム', '頁'];
        outSheet.getRange(currentRow, 1, 1, 10).setValues([headers]).setBackground('#FFD9B3').setFontWeight('bold');
        currentRow++;

        var prevMedia = null;
        var prevValues = [null, null, null, null, null, null, null, null, null, null];

        typeItems.forEach(function(item) {
          if (prevMedia !== null && prevMedia !== item.media) {
            currentRow++;
            prevValues = [null, null, null, null, null, null, null, null, null, null];
          }

          var relDateMD = '';
          if (item.releaseDate) {
            var d = new Date(item.releaseDate);
            relDateMD = (d.getMonth() + 1) + '/' + d.getDate();
          }

          var row = [item.media, item.issue, relDateMD, item.theme, item.stylist,
                     item.itemCode, item.colorCode, item.price || '', item.itemName, item.page];

          var outputRow = row.map(function(v, i) {
            if (v !== '' && v !== null && v !== undefined && String(v) === String(prevValues[i])) return '';
            return v;
          });

          outSheet.getRange(currentRow, 1, 1, 10).setValues([outputRow]);
          if (row[7] !== '') outSheet.getRange(currentRow, 8).setNumberFormat('¥#,##0');

          prevValues = row;
          prevMedia = item.media;
          currentRow++;
        });
      });

      for (var ci = 0; ci < colWidths.length; ci++) {
        outSheet.setColumnWidth(ci + 1, colWidths[ci]);
      }

    } else {
      // ===== ブランド別分割（11列）=====
      var colWidths11 = [90, 120, 70, 55, 200, 100, 90, 90, 70, 200, 50];

      // ブランド名を出現順にユニーク化
      var brands = [];
      items.forEach(function(it) {
        if (brands.indexOf(it.brand) === -1) brands.push(it.brand);
      });

      brands.forEach(function(brand) {
        TYPE_ORDER.forEach(function(type) {
          var brandTypeItems = items.filter(function(it) { return it.brand === brand && it.type === type; });
          if (brandTypeItems.length === 0) return;

          brandTypeItems.sort(function(a, b) {
            var da = a.releaseDate ? new Date(a.releaseDate) : new Date(0);
            var db = b.releaseDate ? new Date(b.releaseDate) : new Date(0);
            return da - db;
          });

          var config = TYPE_CONFIG[type];

          // タイトル行（オレンジ・11列）
          outSheet.getRange(currentRow, 1, 1, 11).setBackground('#FF9900');
          outSheet.getRange(currentRow, 1).setValue(brand);
          outSheet.getRange(currentRow, 5).setValue(monthLabel);
          outSheet.getRange(currentRow, 9).setValue(config.title);
          currentRow++;

          // ヘッダー行（薄オレンジ・11列）
          var headers11 = ['ブランド名', config.col1, '掲載号', '発売日', 'テーマ', 'スタイリスト', '品番', '色番', '上代', 'アイテム', '頁'];
          outSheet.getRange(currentRow, 1, 1, 11).setValues([headers11]).setBackground('#FFD9B3').setFontWeight('bold');
          currentRow++;

          var prevMedia = null;
          var prevValues = [null, null, null, null, null, null, null, null, null, null, null];

          brandTypeItems.forEach(function(item) {
            if (prevMedia !== null && prevMedia !== item.media) {
              currentRow++;
              prevValues = [null, null, null, null, null, null, null, null, null, null, null];
            }

            var relDateMD = '';
            if (item.releaseDate) {
              var d = new Date(item.releaseDate);
              relDateMD = (d.getMonth() + 1) + '/' + d.getDate();
            }

            var row = [item.brand, item.media, item.issue, relDateMD, item.theme, item.stylist,
                       item.itemCode, item.colorCode, item.price || '', item.itemName, item.page];

            var outputRow = row.map(function(v, i) {
              if (v !== '' && v !== null && v !== undefined && String(v) === String(prevValues[i])) return '';
              return v;
            });

            outSheet.getRange(currentRow, 1, 1, 11).setValues([outputRow]);
            if (row[8] !== '') outSheet.getRange(currentRow, 9).setNumberFormat('¥#,##0');

            prevValues = row;
            prevMedia = item.media;
            currentRow++;
          });
        });
      });

      for (var ci = 0; ci < colWidths11.length; ci++) {
        outSheet.setColumnWidth(ci + 1, colWidths11[ci]);
      }
    }

    return { success: true, sheetName: sheetName, rowCount: currentRow - 1 };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.message };
  }
}
