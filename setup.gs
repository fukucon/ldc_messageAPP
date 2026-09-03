/**
 * setup.gs
 * スプレッドシートのシート生成・カラム設定を行うコード。
 * 初回に setupBot() を一度だけ手動実行してください。
 */

// 利用するスプレッドシートのID（URLの /d/ と /edit の間の文字列）
var SPREADSHEET_ID = '1DIss4Wyo-TiA9tIgdboIoFXq0GlJM1Cx5jmyKNbhXIA';

/**
 * カテゴリ（シリーズ）の定義。
 * ここに1件足すだけで、メニュー・チャット画面・設定画面すべてに反映される。
 * - id    : URLで使う識別子（英数字）
 * - name  : 画面に出す名前
 * - desc  : メニューに出す補足説明
 * - sheet : 規程本文を貼り付けるシート名
 * - icon  : メニューに出すアイコン
 */
var CATEGORIES = [
  { id: 'roumu', name: '労務系', desc: '就業規則・勤怠・休暇・給与など', sheet: '就労規則', icon: '👤' },
  { id: 'keiri', name: '経理系', desc: '経費精算・出張旅費・請求など',   sheet: '経理規程', icon: '💴' },
  { id: 'kobai', name: '購買系', desc: '発注・稟議・取引先・契約など',   sheet: '購買規程', icon: '📦' }
];

// 共通シート名
var BOT_LOG_SHEET = '質問ログ';
var SETTINGS_SHEET = '設定';

// 各シートのヘッダー
var BOT_LOG_HEADER = ['日時', 'カテゴリ', '社員', '質問', '回答'];
var SETTINGS_HEADER = ['カテゴリ', '追加指示', '追加ルール'];

// 規程シートの1行目に入れる説明文
var RULES_SHEET_NOTE = '▼この下(A2以降)に規程の本文を貼り付けてください';

/**
 * AIチャットボット用のシートをまとめて生成する。
 * 何度実行しても既存データは壊さない。
 */
function setupBot() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // カテゴリごとの規程シート
  for (var i = 0; i < CATEGORIES.length; i++) {
    setupRulesSheet_(ss, CATEGORIES[i].sheet);
  }

  // 質問ログシート
  setupSheet_(ss, BOT_LOG_SHEET, BOT_LOG_HEADER);

  // 設定シート（カテゴリごとに1行）
  setupSettingsSheet_(ss);

  SpreadsheetApp.flush();
  Logger.log('setupBot 完了');
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
 * 設定シートを用意し、カテゴリごとの行を揃える。
 * 旧形式（項目/内容の2列）だった場合は、その内容を労務系へ引き継ぐ。
 */
function setupSettingsSheet_(ss) {
  var sheet = ss.getSheetByName(SETTINGS_SHEET);

  // 旧形式からの引き継ぎ用に、先に中身を読んでおく
  var oldInstruction = '';
  var oldRules = '';
  if (sheet) {
    var values = sheet.getDataRange().getValues();
    var isOldFormat = values.length && values[0][0] === '項目';
    if (isOldFormat) {
      for (var i = 1; i < values.length; i++) {
        if (values[i][0] === '追加指示') oldInstruction = (values[i][1] || '').toString();
        if (values[i][0] === '追加ルール') oldRules = (values[i][1] || '').toString();
      }
      sheet.clear(); // 新形式で作り直す
    }
  }

  sheet = setupSheet_(ss, SETTINGS_SHEET, SETTINGS_HEADER);

  // カテゴリごとの行が無ければ追加する
  var rows = sheet.getDataRange().getValues();
  for (var c = 0; c < CATEGORIES.length; c++) {
    var name = CATEGORIES[c].name;
    var found = false;
    for (var r = 1; r < rows.length; r++) {
      if (rows[r][0] === name) { found = true; break; }
    }
    if (!found) {
      // 旧形式の内容は最初のカテゴリ（労務系）に引き継ぐ
      var inherit = (c === 0);
      sheet.appendRow([name, inherit ? oldInstruction : '', inherit ? oldRules : '']);
    }
  }

  sheet.setColumnWidth(2, 400);
  sheet.setColumnWidth(3, 400);
  return sheet;
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
