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
