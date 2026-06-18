/**
 * code.gs
 * Webアプリのメイン動作。Googleチャット風のメッセージ送受信を担当する。
 * 定数（SPREADSHEET_ID やシート名・ヘッダー）は setup.gs を参照する。
 */

/**
 * Webアプリのエントリーポイント. index.html を返す。
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('LDC Chat')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 現在ログイン中のユーザーのメールアドレスを返す。
 * 取得できない（匿名アクセス）場合は 'ゲスト' を返す。
 */
function getCurrentUser() {
  var email = Session.getActiveUser().getEmail();
  return email || 'ゲスト';
}

/**
 * ルーム一覧を取得する。
 * @return {Array<Object>} [{roomId, name}]
 */
function getRooms() {
  var sheet = getSheet_(ROOMS_SHEET);
  var values = sheet.getDataRange().getValues();
  var rooms = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    rooms.push({ roomId: values[i][0], name: values[i][1] });
  }
  return rooms;
}

/**
 * ルームを新規作成する。
 * @return {Object} 作成したルーム {roomId, name}
 */
function createRoom(name) {
  name = (name || '').toString().trim();
  if (!name) {
    throw new Error('ルーム名を入力してください');
  }
  var sheet = getSheet_(ROOMS_SHEET);
  var roomId = Utilities.getUuid();
  sheet.appendRow([roomId, name, new Date()]);
  return { roomId: roomId, name: name };
}

/**
 * 指定ルームのメッセージを取得する。
 * @param {string} roomId 対象ルームID
 * @return {Array<Object>} [{user, text, createdAt}]
 */
function getMessages(roomId) {
  var sheet = getSheet_(MESSAGES_SHEET);
  var values = sheet.getDataRange().getValues();
  var messages = [];
  for (var i = 1; i < values.length; i++) {
    if (values[i][1] !== roomId) continue;
    messages.push({
      user: values[i][2],
      text: values[i][3],
      createdAt: values[i][4] ? new Date(values[i][4]).getTime() : 0
    });
  }
  return messages;
}

/**
 * メッセージを投稿する。
 * @param {string} roomId 対象ルームID
 * @param {string} text   本文
 * @return {Object} 投稿したメッセージ
 */
function postMessage(roomId, text) {
  text = (text || '').toString().trim();
  if (!roomId) {
    throw new Error('ルームが選択されていません');
  }
  if (!text) {
    throw new Error('メッセージを入力してください');
  }
  var user = getCurrentUser();
  var now = new Date();
  var sheet = getSheet_(MESSAGES_SHEET);
  sheet.appendRow([Utilities.getUuid(), roomId, user, text, now]);
  return { user: user, text: text, createdAt: now.getTime() };
}

/**
 * シートを取得する内部ヘルパー。未生成なら setup() を促すエラーを投げる。
 */
function getSheet_(name) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
  if (!sheet) {
    throw new Error('シート「' + name + '」がありません。先に setup() を実行してください。');
  }
  return sheet;
}
