/**
 * setup.gs
 * スプレッドシートのシート生成・カラム設定を行うコード。
 * 初回に setupBot() を一度だけ手動実行してください。
 */

// 利用するスプレッドシートのID（URLの /d/ と /edit の間の文字列）
var SPREADSHEET_ID = '1DIss4Wyo-TiA9tIgdboIoFXq0GlJM1Cx5jmyKNbhXIA';

/**
 * カテゴリ（シリーズ）の初期値。
 * 実際の定義は「カテゴリ」シートで管理し、設定画面から追加・編集できる。
 * ここは setupBot() で初回にシートへ書き込む種データ。
 * - id     : URLで使う識別子（英数字）
 * - name   : 画面に出す名前
 * - desc   : メニューに出す補足説明
 * - sheet  : 規程本文を貼り付けるシート名
 * - prefix : 追加ルールのID接頭辞（例 R → R1, R2 ...）
 * - icon   : メニューに出すアイコン
 */
var DEFAULT_CATEGORIES = [
  { id: 'roumu', name: '労務系', desc: '就業規則・勤怠・休暇・給与など', sheet: '就労規則', prefix: 'R', icon: '👤' },
  { id: 'keiri', name: '経理系', desc: '経費精算・出張旅費・請求など',   sheet: '経理規程', prefix: 'K', icon: '💴' },
  { id: 'kobai', name: '購買系', desc: '発注・稟議・取引先・契約など',   sheet: '購買規程', prefix: 'B', icon: '📦' }
];

// 共通シート名
var CATEGORY_SHEET = 'カテゴリ';
var BOT_LOG_SHEET = '質問ログ';
var SETTINGS_SHEET = '設定';

// 各シートのヘッダー
var CATEGORY_HEADER = ['ID', 'カテゴリ名', '説明', '参照シート', 'ID接頭辞', 'アイコン'];
var BOT_LOG_HEADER = ['日時', 'カテゴリ', '社員', '質問', '回答'];
var SETTINGS_HEADER = ['ID', 'カテゴリ', '追加ルール'];

// 規程シートの1行目に入れる説明文
var RULES_SHEET_NOTE = '▼この下(A2以降)に規程の本文を貼り付けてください';

/**
 * AIチャットボット用のシートをまとめて生成する。
 * 何度実行しても既存データは壊さない。
 */
function setupBot() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // カテゴリ定義シート（初回は初期値を書き込む）
  var cats = setupCategorySheet_(ss);

  // カテゴリごとの規程シート
  for (var i = 0; i < cats.length; i++) {
    setupRulesSheet_(ss, cats[i].sheet);
  }

  // 質問ログシート
  setupSheet_(ss, BOT_LOG_SHEET, BOT_LOG_HEADER);

  // 設定シート（1行 = 1つの追加ルール）
  setupSettingsSheet_(ss);

  SpreadsheetApp.flush();
  Logger.log('setupBot 完了');
}

/**
 * カテゴリ定義シートを用意する。空なら初期値を書き込む。
 * @return {Array<Object>} カテゴリ定義の配列
 */
function setupCategorySheet_(ss) {
  var sheet = setupSheet_(ss, CATEGORY_SHEET, CATEGORY_HEADER);

  if (sheet.getLastRow() < 2) {
    for (var i = 0; i < DEFAULT_CATEGORIES.length; i++) {
      var c = DEFAULT_CATEGORIES[i];
      sheet.appendRow([c.id, c.name, c.desc, c.sheet, c.prefix, c.icon]);
    }
  }
  sheet.setColumnWidth(3, 260);
  sheet.setColumnWidth(4, 140);

  return readCategorySheet_(ss);
}

/**
 * カテゴリ定義シートを読み込む。
 * シートが無い・空の場合は初期値を返す。
 * @return {Array<Object>} [{id, name, desc, sheet, prefix, icon}]
 */
function readCategorySheet_(ss) {
  ss = ss || SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CATEGORY_SHEET);
  if (!sheet || sheet.getLastRow() < 2) {
    return DEFAULT_CATEGORIES;
  }
  var values = sheet.getDataRange().getValues();
  var list = [];
  // 列: 0=ID 1=カテゴリ名 2=説明 3=参照シート 4=ID接頭辞 5=アイコン
  for (var i = 1; i < values.length; i++) {
    var id = (values[i][0] || '').toString().trim();
    var name = (values[i][1] || '').toString().trim();
    if (!id || !name) continue;
    list.push({
      id: id,
      name: name,
      desc: (values[i][2] || '').toString(),
      sheet: (values[i][3] || '').toString().trim() || name,
      prefix: (values[i][4] || '').toString().trim() || 'X',
      icon: (values[i][5] || '').toString() || '💬'
    });
  }
  return list.length ? list : DEFAULT_CATEGORIES;
}

/**
 * 規程本文シートを用意する（1行目は説明、2行目以降が本文）。
 */
function setupRulesSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  var head = sheet.getRange(1, 1);
  head.setValue(RULES_SHEET_NOTE);
  head.setFontWeight('bold').setBackground('#fbbc04');
  sheet.setFrozenRows(1);
  return sheet;
}

/**
 * 設定シートを [ID, カテゴリ, 追加ルール] 形式で用意する。
 * 旧形式のデータがあれば、1行1ルールに分解して引き継ぐ。
 */
function setupSettingsSheet_(ss) {
  var sheet = ss.getSheetByName(SETTINGS_SHEET);
  var migrated = [];   // [カテゴリ名, ルール本文]

  if (sheet) {
    var values = sheet.getDataRange().getValues();
    var header0 = values.length ? String(values[0][0]) : '';

    if (header0 === 'カテゴリ') {
      // 旧形式A: [カテゴリ, 追加指示, 追加ルール]
      for (var i = 1; i < values.length; i++) {
        var catName = String(values[i][0] || '');
        if (!catName) continue;
        migrated = migrated.concat(splitToRules_(catName, values[i][1]));
        migrated = migrated.concat(splitToRules_(catName, values[i][2]));
      }
      sheet.clear();

    } else if (header0 === '項目') {
      // 旧形式B: [項目, 内容] … 内容はすべて最初のカテゴリへ引き継ぐ
      var firstCat = readCategorySheet_(ss)[0].name;
      for (var j = 1; j < values.length; j++) {
        migrated = migrated.concat(splitToRules_(firstCat, values[j][1]));
      }
      sheet.clear();
    }
  }

  sheet = setupSheet_(ss, SETTINGS_SHEET, SETTINGS_HEADER);

  // 引き継いだルールにIDを振って書き込む
  if (migrated.length) {
    var counters = {};
    for (var k = 0; k < migrated.length; k++) {
      var name = migrated[k][0];
      var prefix = prefixOfCategoryName_(name);
      counters[prefix] = (counters[prefix] || 0) + 1;
      sheet.appendRow([prefix + counters[prefix], name, migrated[k][1]]);
    }
    Logger.log('旧形式の設定を ' + migrated.length + ' 件引き継ぎました');
  }

  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 600);
  return sheet;
}

/**
 * 複数行のテキストを、1行1ルールの配列に分解する。
 * @return {Array<Array>} [[カテゴリ名, ルール本文], ...]
 */
function splitToRules_(categoryName, text) {
  var out = [];
  var s = (text == null ? '' : String(text)).trim();
  if (!s) {
    return out;
  }
  var lines = s.split('\n');
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (line) {
      out.push([categoryName, line]);
    }
  }
  return out;
}

/**
 * カテゴリ名からIDの接頭辞を返す。見つからなければ 'X'。
 */
function prefixOfCategoryName_(name) {
  var cats = readCategorySheet_();
  for (var i = 0; i < cats.length; i++) {
    if (cats[i].name === name) {
      return cats[i].prefix;
    }
  }
  return 'X';
}

/**
 * 指定名のシートを取得（無ければ作成）し、ヘッダーを設定する。
 * @return {Sheet} 対象シート
 */
function setupSheet_(ss, name, header) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  // ヘッダー行を設定して見やすく整形
  var range = sheet.getRange(1, 1, 1, header.length);
  range.setValues([header]);
  range.setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  return sheet;
}
