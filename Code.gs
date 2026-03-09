// BIGI PRESS ROOM アパレル貸出管理システム

var SHEET_LOAN = '貸出管理';
var SHEET_PUBLISH = '掲載リストマスタ';
var SHEET_SETTINGS = '設定';
var SHEET_CREDIT = 'クレジットマスタ';
var SHEET_STAFF = '担当者マスタ';
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
    .addItem('伝票印刷', 'openPrintSlipDialog')
    .addSeparator()
    .addItem('返却・掲載登録処理', 'openReturnDialog')
    .addSeparator()
    .addItem('掲載リスト作成', 'openFormattedSheetDialog')
    .addSeparator()
    .addItem('伝票編集', 'openMediaSetDialog')
    .addItem('掲載リストマスタ編集', 'openPublishListEditDialog')
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
    '返却ステータス', '返却日', '掲載済', '備考', '担当者名'
  ];
  loanSheet.getRange(1, 1, 1, loanHeaders.length).setValues([loanHeaders]);
  loanSheet.getRange(1, 1, 1, loanHeaders.length)
    .setBackground('#4a4a4a').setFontColor('#ffffff').setFontWeight('bold');
  loanSheet.setFrozenRows(1);

  var colWidths = [80, 85, 120, 110, 120, 70, 180, 100, 100, 85, 95, 100, 100, 70, 180, 80, 90, 85, 65, 150, 120];
  for (var ci = 0; ci < colWidths.length; ci++) {
    loanSheet.setColumnWidth(ci + 1, colWidths[ci]);
  }

  // 条件付き書式: 未返却=赤のみ（1ルール）
  var cfRange = loanSheet.getRange(2, 1, 1000, 21);
  loanSheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$Q2="未返却"')
      .setBackground('#ffcccc')
      .setRanges([cfRange])
      .build()
  ]);

  var pubSheet = ss.getSheetByName(SHEET_PUBLISH);
  if (!pubSheet) pubSheet = ss.insertSheet(SHEET_PUBLISH);
  var pubHeaders = ['伝票番号', '雑誌名', '種別', '掲載号', '公開日', 'テーマ/着用者', 'スタイリスト', 'ブランド名', '品番', '色番', '上代', 'アイテム', '頁', '画像/リンク', '担当者名'];
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

  // PRESS ROOM情報（A5〜B13）
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

  // クレジットマスタシート
  var creditSheet = ss.getSheetByName(SHEET_CREDIT);
  if (!creditSheet) {
    creditSheet = ss.insertSheet(SHEET_CREDIT);
    var creditHeaders = ['表示名', 'カナ', 'TEL', 'WEB SITE', 'ONLINE STORE', 'Instagram'];
    creditSheet.getRange(1, 1, 1, creditHeaders.length).setValues([creditHeaders]);
    creditSheet.getRange(1, 1, 1, creditHeaders.length)
      .setBackground('#4a4a4a').setFontColor('#ffffff').setFontWeight('bold');
    creditSheet.setFrozenRows(1);
  }

  // 担当者マスタシート
  var staffSheet = ss.getSheetByName(SHEET_STAFF);
  if (!staffSheet) {
    staffSheet = ss.insertSheet(SHEET_STAFF);
    staffSheet.getRange(1, 1).setValue('担当者名').setFontWeight('bold')
      .setBackground('#4a4a4a').setFontColor('#ffffff');
    staffSheet.setFrozenRows(1);
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
    var shootDatesArr = data.mediaSets.map(function(m) { return m.shootDate || ''; });
    var anyShootDate  = shootDatesArr.some(function(d) { return !!d; });
    var shootDateStr  = anyShootDate ? shootDatesArr.join(' | ') : '';

    var rows = [];
    for (var ii = 0; ii < data.items.length; ii++) {
      var item = data.items[ii];
      var pr = item.price ? Number(String(item.price).replace(/,/g, '')) : '';
      rows.push([
        slipNo, today, data.stylist, data.phone,
        mediaStr, issueStr, themeStr, relDateStr, typeStr,  // E-I
        shootDateStr, data.plannedReturnDate,                // J-K
        item.brand, item.itemCode, item.colorCode, item.itemName, pr,  // L-P
        '未返却', '', false, '', data.staff || ''            // Q-U
      ]);
    }
    if (rows.length === 0) {
      return { success: false, message: 'アイテムが入力されていません。' };
    }

    var startRow = getLastDataRow(sheet) + 1;
    sheet.getRange(startRow, 1, rows.length, 21).setValues(rows);
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
  lines.push('【BIGI PRESS ROOM 貸出伝票 No.' + slipNo + '】');
  lines.push('スタイリスト：' + data.stylist + '様');
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
  if (data.plannedReturnDate) dateStr += (dateStr ? '\u3000' : '') + '返却予定日：' + data.plannedReturnDate;
  if (dateStr) lines.push(dateStr);
  lines.push('---');
  for (var i = 0; i < data.items.length; i++) {
    var item = data.items[i];
    var priceStr = item.price ? ('¥' + Number(String(item.price).replace(/,/g, '')).toLocaleString()) : '';
    lines.push((i + 1) + '. ' + (item.brand ? item.brand + ' ' : '') + item.itemCode + ' ' + item.colorCode + ' ' + item.itemName + ' ' + priceStr);
  }
  lines.push('---');
  lines.push('ご返却は10:00〜18:00（12:00〜13:00を除く）でお願いいたします。');
  return lines.join('\n');
}

function openMediaSetDialog() {
  var html = HtmlService.createHtmlOutputFromFile('MediaSetDialog')
    .setWidth(560).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '伝票編集');
}

// ===== 伝票印刷 =====

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
    var data = sheet.getRange(2, 1, lastDataRow - 1, 17).getValues();

    var rowCount = 0;
    var loanDate = '', stylist = '', phone = '', plannedReturnDate = '';
    var mediaSets = [];
    var items = [];

    for (var i = 0; i < data.length; i++) {
      if (!data[i][0]) continue;
      var rowSlip = String(data[i][0]);
      while (rowSlip.length < 5) rowSlip = '0' + rowSlip;
      if (rowSlip !== s) continue;
      rowCount++;
      if (rowCount === 1) {
        loanDate          = cellToStr(data[i][1]);
        stylist           = cellToStr(data[i][2]);
        phone             = cellToStr(data[i][3]);
        plannedReturnDate = cellToStr(data[i][10]);
        var medias     = cellToStr(data[i][4]).split(' | ');
        var issues     = cellToStr(data[i][5]).split(' | ');
        var themes     = cellToStr(data[i][6]).split(' | ');
        var relDates   = cellToStr(data[i][7]).split(' | ');
        var types      = cellToStr(data[i][8]).split(' | ');
        var shootDates = cellToStr(data[i][9]).split(' | ');
        for (var j = 0; j < medias.length; j++) {
          if (!medias[j]) continue;
          mediaSets.push({
            media:       medias[j],
            issue:       issues[j]     || '',
            theme:       themes[j]     || '',
            releaseDate: relDates[j]   || '',
            type:        types[j]      || '',
            shootDate:   shootDates[j] || ''
          });
        }
      }
      items.push({
        brand:     cellToStr(data[i][11]),
        itemCode:  cellToStr(data[i][12]),
        colorCode: cellToStr(data[i][13]),
        itemName:  cellToStr(data[i][14]),
        price:     data[i][15] ? Number(data[i][15]) : ''
      });
    }
    if (rowCount === 0) return { found: false };

    // クレジットマスタ
    var credits = [];
    var creditSheet = ss.getSheetByName(SHEET_CREDIT);
    if (creditSheet) {
      var creditLastRow = creditSheet.getLastRow();
      if (creditLastRow > 1) {
        var creditData = creditSheet.getRange(2, 1, creditLastRow - 1, 6).getValues();
        for (var ci = 0; ci < creditData.length; ci++) {
          var cr = creditData[ci];
          if (!cr[0]) continue;
          credits.push({
            name:        cellToStr(cr[0]),
            kana:        cellToStr(cr[1]),
            tel:         cellToStr(cr[2]),
            web:         cellToStr(cr[3]),
            onlineStore: cellToStr(cr[4]),
            instagram:   cellToStr(cr[5])
          });
        }
      }
    }

    // 設定シートからPRESS ROOM情報（B5〜B13 = 9行一括読み込み）
    var pressRooms = [];
    var settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
    if (settingsSheet) {
      var prRange = settingsSheet.getRange(5, 2, 9, 1).getValues();
      // B5=名称1, B6=住所1, B7=TEL1, B8=FAX1, B9=空行, B10=名称2, B11=住所2, B12=TEL2, B13=FAX2
      if (prRange[0][0]) pressRooms.push({
        label: cellToStr(prRange[0][0]),
        address: cellToStr(prRange[1][0]),
        tel: cellToStr(prRange[2][0]),
        fax: cellToStr(prRange[3][0])
      });
      if (prRange[5][0]) pressRooms.push({
        label: cellToStr(prRange[5][0]),
        address: cellToStr(prRange[6][0]),
        tel: cellToStr(prRange[7][0]),
        fax: cellToStr(prRange[8][0])
      });
    }

    return {
      found: true,
      slipNo: s,
      loanDate: loanDate,
      stylist: stylist,
      phone: phone,
      plannedReturnDate: plannedReturnDate,
      mediaSets: mediaSets,
      items: items,
      credits: credits,
      pressRooms: pressRooms
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
    var data = sheet.getRange(2, 1, lastDataRow - 1, 10).getValues();

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
        var medias     = cellToStr(data[i][4]).split(' | ');
        var issues     = cellToStr(data[i][5]).split(' | ');
        var themes     = cellToStr(data[i][6]).split(' | ');
        var relDates   = cellToStr(data[i][7]).split(' | ');
        var types      = cellToStr(data[i][8]).split(' | ');
        var shootDates = cellToStr(data[i][9]).split(' | ');
        for (var j = 0; j < medias.length; j++) {
          if (!medias[j]) continue;
          mediaSets.push({
            media:       medias[j],
            issue:       issues[j]     || '',
            theme:       themes[j]     || '',
            releaseDate: relDates[j]   || '',
            type:        types[j]      || '',
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

// ===== クレジットマスタ ダミーデータ投入 =====
// GASエディタから手動実行する。既存データがある場合は上書きしない。

function insertCreditDummyData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CREDIT);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('クレジットマスタシートが存在しません。先に setupSheets() を実行してください。');
    return;
  }

  // データ行が既にある場合はスキップ
  if (sheet.getLastRow() > 1) {
    var ui = SpreadsheetApp.getUi();
    var res = ui.alert('既存データの確認', 'クレジットマスタにすでにデータがあります。上書きしますか？', ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) return;
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).clearContent();
  }

  // [表示名, カナ, TEL, WEB SITE, ONLINE STORE, Instagram]
  var dummy = [
    ['UNTITLED',          'アンタイトル',         '03-6861-7600', 'https://untitled.co.jp',               'https://store.bigi.co.jp/untitled',    '@untitled_official_jp'],
    ["L'EQUIPE",          'レキップ',             '03-6861-7610', 'https://lequipe.bigi.co.jp',           'https://store.bigi.co.jp/lequipe',     '@lequipe_official'],
    ['MOGA',              'モガ',                 '03-6861-7620', 'https://moga.co.jp',                   'https://store.bigi.co.jp/moga',        '@moga_official_jp'],
    ['DÉPAREILLÉE',       'デパリエ',             '03-6861-7630', 'https://depareillee.bigi.co.jp',       'https://store.bigi.co.jp/depareillee', '@depareillee_official'],
    ['ADIEU TRISTESSE',   'アデュー トリステス', '03-6861-7658', 'https://adieu-tristesse.jp',           'https://store.bigi.co.jp/adieu',       '@adieu_tristesse_official'],
    ['LOISIR',            'ロワジール',           '03-6861-7640', 'https://loisir.bigi.co.jp',            'https://store.bigi.co.jp/loisir',      '@loisir_official'],
    ['ef-de',             'エフデ',               '03-6861-7650', 'https://ef-de.jp',                     'https://store.bigi.co.jp/efde',        '@efde_official_jp'],
    ['NATURAL BEAUTY',    'ナチュラルビューティ', '03-6861-7660', 'https://naturalbeauty.bigi.co.jp',     'https://store.bigi.co.jp/nb',          '@naturalbeauty_official'],
    ['ICEBERG',           'アイスバーグ',         '03-6861-7670', 'https://iceberg.bigi.co.jp',           '',                                     '@iceberg_jp_official'],
    ['RENOMA',            'レノマ',               '03-6861-7680', 'https://renoma.bigi.co.jp',            '',                                     ''],
  ];

  sheet.getRange(2, 1, dummy.length, 6).setValues(dummy);
  Logger.log('クレジットマスタにダミーデータを ' + dummy.length + ' 件投入しました。');
  SpreadsheetApp.getUi().alert('完了', 'クレジットマスタに ' + dummy.length + ' 件のダミーデータを投入しました。', SpreadsheetApp.getUi().ButtonSet.OK);
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
      if (params.mediaSet.shootDate) { append(10, params.mediaSet.shootDate); }
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
    // A〜U列（21列）を読み込む
    var data = sheet.getRange(2, 1, lastDataRow - 1, 21).getValues();

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
          typeStr:    cellToStr(data[i][8]),
          staff:      cellToStr(data[i][20])
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
    var pubLastRow = pubSheet.getLastRow();
    for (var i = 0; i < params.combinations.length; i++) {
      var c = params.combinations[i];
      loanSheet.getRange(c.rowIndex, 19).setValue(true); // S列（掲載済）
      var slipNo = cellToStr(loanSheet.getRange(c.rowIndex, 1).getValue());
      var staff = cellToStr(loanSheet.getRange(c.rowIndex, 21).getValue());
      pubLastRow++;
      pubSheet.getRange(pubLastRow, 1, 1, 15).setValues([[
        slipNo, c.media, c.type, c.issue, c.releaseDate, c.theme, c.stylist,
        c.brand || '', c.itemCode, c.colorCode, c.price, c.itemName, '', '', staff
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
  var data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
  var results = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[1] && !row[3]) continue; // B=雑誌名, D=掲載号 で空行判定
    var pubDate = row[4]; // E列: 公開日
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
        media:       cellToStr(row[1]),
        type:        cellToStr(row[2]),
        issue:       cellToStr(row[3]),
        releaseDate: rdStr,
        theme:       cellToStr(row[5]),
        stylist:     cellToStr(row[6]),
        brand:       cellToStr(row[7]),
        itemCode:    cellToStr(row[8]),
        colorCode:   cellToStr(row[9]),
        price:       row[10] ? Number(row[10]) : 0,
        itemName:    cellToStr(row[11]),
        page:        cellToStr(row[12]),
        imageLink:   cellToStr(row[13]),
        staff:       cellToStr(row[14])
      });
    }
  }
  return results;
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

function getSlipDataForEdit(slipNo) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOAN);
    var s = String(slipNo);
    while (s.length < 5) s = '0' + s;

    var lastDataRow = getLastDataRow(sheet);
    if (lastDataRow <= 1) return { found: false };
    var data = sheet.getRange(2, 1, lastDataRow - 1, 21).getValues();

    var rowCount = 0;
    var returnedCount = 0;
    var publishedCount = 0;
    var stylist = '', phone = '', plannedReturnDate = '', staff = '';
    var mediaSets = [];

    for (var i = 0; i < data.length; i++) {
      if (!data[i][0]) continue;
      var rowSlip = String(data[i][0]);
      while (rowSlip.length < 5) rowSlip = '0' + rowSlip;
      if (rowSlip !== s) continue;
      rowCount++;
      if (data[i][16] === '返却済') returnedCount++;
      if (data[i][18] === true) publishedCount++;
      if (rowCount === 1) {
        stylist           = cellToStr(data[i][2]);
        phone             = cellToStr(data[i][3]);
        plannedReturnDate = cellToStr(data[i][10]);
        staff             = cellToStr(data[i][20]);
        var medias     = cellToStr(data[i][4]).split(' | ');
        var issues     = cellToStr(data[i][5]).split(' | ');
        var themes     = cellToStr(data[i][6]).split(' | ');
        var relDates   = cellToStr(data[i][7]).split(' | ');
        var types      = cellToStr(data[i][8]).split(' | ');
        var shootDates = cellToStr(data[i][9]).split(' | ');
        for (var j = 0; j < medias.length; j++) {
          if (!medias[j]) continue;
          mediaSets.push({
            media:       medias[j],
            issue:       issues[j]     || '',
            theme:       themes[j]     || '',
            releaseDate: relDates[j]   || '',
            type:        types[j]      || '',
            shootDate:   shootDates[j] || ''
          });
        }
      }
    }
    if (rowCount === 0) return { found: false };
    return {
      found: true,
      rowCount: rowCount,
      returnedCount: returnedCount,
      publishedCount: publishedCount,
      stylist: stylist,
      phone: phone,
      plannedReturnDate: plannedReturnDate,
      staff: staff,
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

    var mediaStr    = params.mediaSets.map(function(m) { return m.media; }).join(' | ');
    var issueStr    = params.mediaSets.map(function(m) { return m.issue; }).join(' | ');
    var themeStr    = params.mediaSets.map(function(m) { return m.theme; }).join(' | ');
    var relDateStr  = params.mediaSets.map(function(m) { return m.releaseDate; }).join(' | ');
    var typeStr     = params.mediaSets.map(function(m) { return m.type || ''; }).join(' | ');
    var shootDateArr = params.mediaSets.map(function(m) { return m.shootDate || ''; });
    var anyShootDate = shootDateArr.some(function(d) { return !!d; });
    var shootDateStr = anyShootDate ? shootDateArr.join(' | ') : '';

    var updated = 0;
    for (var i = 0; i < data.length; i++) {
      if (!data[i][0]) continue;
      var rowSlip = String(data[i][0]);
      while (rowSlip.length < 5) rowSlip = '0' + rowSlip;
      if (rowSlip !== s) continue;
      var r = i + 2;
      sheet.getRange(r, 3).setValue(params.stylist);
      sheet.getRange(r, 4).setValue(params.phone);
      sheet.getRange(r, 5).setValue(mediaStr);
      sheet.getRange(r, 6).setValue(issueStr);
      sheet.getRange(r, 7).setValue(themeStr);
      sheet.getRange(r, 8).setValue(relDateStr);
      sheet.getRange(r, 9).setValue(typeStr);
      sheet.getRange(r, 10).setValue(shootDateStr);
      sheet.getRange(r, 11).setValue(params.plannedReturnDate);
      sheet.getRange(r, 21).setValue(params.staff || '');
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
    var data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
    var results = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      var rowSlip = String(row[0]);
      while (rowSlip.length < 5) rowSlip = '0' + rowSlip;
      if (rowSlip !== s) continue;
      var pubDate = row[4];
      var rdStr = pubDate ? Utilities.formatDate(new Date(pubDate), Session.getScriptTimeZone(), 'yyyy/MM/dd') : '';
      results.push({
        rowIndex:  i + 2,
        slipNo:    cellToStr(row[0]),
        media:     cellToStr(row[1]),
        type:      cellToStr(row[2]),
        issue:     cellToStr(row[3]),
        releaseDate: rdStr,
        theme:     cellToStr(row[5]),
        stylist:   cellToStr(row[6]),
        brand:     cellToStr(row[7]),
        itemCode:  cellToStr(row[8]),
        colorCode: cellToStr(row[9]),
        price:     row[10] ? Number(row[10]) : 0,
        itemName:  cellToStr(row[11]),
        page:      cellToStr(row[12]),
        imageLink: cellToStr(row[13]),
        staff:     cellToStr(row[14])
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
    sheet.getRange(rowIndex, 1, 1, 15).setValues([[
      data.slipNo, data.media, data.type, data.issue, data.releaseDate,
      data.theme, data.stylist, data.brand, data.itemCode, data.colorCode,
      data.price, data.itemName, data.page, data.imageLink, data.staff || ''
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
    .setWidth(860).setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, '掲載リストマスタ編集');
}

function revertToUnreturned(slipNo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var loanSheet = ss.getSheetByName(SHEET_LOAN);
    var pubSheet = ss.getSheetByName(SHEET_PUBLISH);
    var s = String(slipNo);
    while (s.length < 5) s = '0' + s;

    // 貸出管理シート: Q列="未返却", R列="", S列=false
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
        loanSheet.getRange(r, 17).setValue('未返却');  // Q列
        loanSheet.getRange(r, 18).setValue('');         // R列
        loanSheet.getRange(r, 19).setValue(false);      // S列
        revertedCount++;
      }
    }

    // 掲載リストマスタ: 該当伝票番号の行を後ろから物理削除
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

function getStaffList() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STAFF);
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    return data.map(function(row) { return String(row[0]); }).filter(Boolean);
  } catch (e) {
    return [];
  }
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
  SpreadsheetApp.getUi().showModalDialog(html, '月次レポート');
}

function generateFormattedSheet(params) {
  try {
    var year  = Number(params.year);
    var month = Number(params.month);
    var splitByBrand = params.splitByBrand === true;
    var items = getPublishDataForMonth(year, month);
    var brandName = getBrandName();
    var monthLabel = month + '月';
    var sheetName = '掲載リスト_' + month + '月';

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
      // ===== 分割なし（11列）=====
      var colWidths = [120, 70, 55, 200, 100, 90, 90, 70, 200, 50, 150];

      TYPE_ORDER.forEach(function(type) {
        var typeItems = items.filter(function(it) { return it.type === type; });
        if (typeItems.length === 0) return;

        typeItems.sort(function(a, b) {
          var da = a.releaseDate ? new Date(a.releaseDate) : new Date(0);
          var db = b.releaseDate ? new Date(b.releaseDate) : new Date(0);
          return da - db;
        });

        var config = TYPE_CONFIG[type];

        var h1Range = outSheet.getRange(currentRow, 1, 1, 11);
        h1Range.setBackground('#FF9900');
        outSheet.getRange(currentRow, 1).setValue(brandName);
        outSheet.getRange(currentRow, 4).setValue(monthLabel);
        outSheet.getRange(currentRow, 8).setValue(config.title);
        currentRow++;

        var headers = [config.col1, '掲載号', '発売日', 'テーマ', 'スタイリスト', '品番', '色番', '上代', 'アイテム', '頁', '画像/リンク'];
        outSheet.getRange(currentRow, 1, 1, 11).setValues([headers]).setBackground('#FFD9B3').setFontWeight('bold');
        currentRow++;

        var prevMedia = null;
        var prevValues = [null, null, null, null, null, null, null, null, null, null, null];

        typeItems.forEach(function(item) {
          if (prevMedia !== null && prevMedia !== item.media) {
            currentRow++;
            prevValues = [null, null, null, null, null, null, null, null, null, null, null];
          }

          var relDateMD = '';
          if (item.releaseDate) {
            var d = new Date(item.releaseDate);
            relDateMD = (d.getMonth() + 1) + '/' + d.getDate();
          }

          var row = [item.media, item.issue, relDateMD, item.theme, item.stylist,
                     item.itemCode, item.colorCode, item.price || '', item.itemName, item.page, item.imageLink];

          var outputRow = row.map(function(v, i) {
            if (v !== '' && v !== null && v !== undefined && String(v) === String(prevValues[i])) return '';
            return v;
          });

          prevValues = row;
          prevMedia = item.media;

          var allEmpty = outputRow.every(function(v) { return v === '' || v === null || v === undefined; });
          if (!allEmpty) {
            outSheet.getRange(currentRow, 1, 1, 11).setValues([outputRow]);
            if (row[7] !== '') outSheet.getRange(currentRow, 8).setNumberFormat('¥#,##0');
            currentRow++;
          }
        });
      });

      for (var ci = 0; ci < colWidths.length; ci++) {
        outSheet.setColumnWidth(ci + 1, colWidths[ci]);
      }

    } else {
      // ===== ブランド別分割（11列）=====
      var colWidths11 = [120, 70, 55, 200, 100, 90, 90, 70, 200, 50, 150];

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
          outSheet.getRange(currentRow, 4).setValue(monthLabel);
          outSheet.getRange(currentRow, 8).setValue(config.title);
          currentRow++;

          // ヘッダー行（薄オレンジ・11列）
          var headers11 = [config.col1, '掲載号', '発売日', 'テーマ', 'スタイリスト', '品番', '色番', '上代', 'アイテム', '頁', '画像/リンク'];
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

            var row = [item.media, item.issue, relDateMD, item.theme, item.stylist,
                       item.itemCode, item.colorCode, item.price || '', item.itemName, item.page, item.imageLink];

            var outputRow = row.map(function(v, i) {
              if (v !== '' && v !== null && v !== undefined && String(v) === String(prevValues[i])) return '';
              return v;
            });

            prevValues = row;
            prevMedia = item.media;

            var allEmpty = outputRow.every(function(v) { return v === '' || v === null || v === undefined; });
            if (!allEmpty) {
              outSheet.getRange(currentRow, 1, 1, 11).setValues([outputRow]);
              if (row[7] !== '') outSheet.getRange(currentRow, 8).setNumberFormat('¥#,##0');
              currentRow++;
            }
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
