/**
 * setup.gs
 * スプレッドシートのシート生成・カラム設定を行うコード。
 * 初回に setup() を一度だけ手動実行してください。
 */

// 利用するスプレッドシートのID（URLの /d/ と /edit の間の文字列）
var SPREADSHEET_ID = '1DIss4Wyo-TiA9tIgdboIoFXq0GlJM1Cx5jmyKNbhXIA';

// シート名
var ROOMS_SHEET = 'Rooms';
var MESSAGES_SHEET = 'Messages';

// 各シートのヘッダー（カラム）定義
var ROOMS_HEADER = ['roomId', 'name', 'createdAt'];
var MESSAGES_HEADER = ['messageId', 'roomId', 'user', 'text', 'createdAt'];

/**
 * シートを生成し、ヘッダー行を設定するメイン処理。
 */
function setup() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  var rooms = setupSheet_(ss, ROOMS_SHEET, ROOMS_HEADER);
  setupSheet_(ss, MESSAGES_SHEET, MESSAGES_HEADER);

  // デフォルトルームが無ければ作成
  if (rooms.getLastRow() < 2) {
    rooms.appendRow([Utilities.getUuid(), '雑談', new Date()]);
  }

  // 初期状態の「Sheet1」などが残っていれば削除（任意）
  var first = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (first && ss.getSheets().length > 1) {
    ss.deleteSheet(first);
  }

  SpreadsheetApp.flush();
  Logger.log('setup 完了');
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

// 就労規則チャットボット用のシート名
var RULES_SHEET = '就労規則';
var BOT_LOG_SHEET = '質問ログ';
var SETTINGS_SHEET = '設定';

// 設定シートの項目名（この名前でキーを探す）
var SETTING_INSTRUCTION = '追加指示';
var SETTING_EXTRA_RULES = '追加ルール';

/**
 * 就労規則チャットボット用のシートを生成する。
 * 初回に一度だけ手動実行してください。
 * - 「就労規則」シート: A2以降に就労規則の本文を貼り付ける
 * - 「質問ログ」シート: 社員の質問と回答を自動記録する
 */
function setupBot() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // 就労規則シート（1行目は説明、2行目以降に本文を貼る）
  var rules = ss.getSheetByName(RULES_SHEET);
  if (!rules) {
    rules = ss.insertSheet(RULES_SHEET);
  }
  var head = rules.getRange(1, 1);
  head.setValue('▼この下(A2以降)に就労規則の本文を貼り付けてください');
  head.setFontWeight('bold').setBackground('#fbbc04');
  rules.setFrozenRows(1);

  // 質問ログシート
  setupSheet_(ss, BOT_LOG_SHEET, ['日時', '社員', '質問', '回答']);

  // 設定シート（設定画面から編集する内容の保存先）
  var settings = ss.getSheetByName(SETTINGS_SHEET);
  if (!settings) {
    settings = setupSheet_(ss, SETTINGS_SHEET, ['項目', '内容']);
    settings.appendRow([SETTING_INSTRUCTION, '']);
    settings.appendRow([SETTING_EXTRA_RULES, '']);
    settings.setColumnWidth(2, 600);
  }

  SpreadsheetApp.flush();
  Logger.log('setupBot 完了');
}
