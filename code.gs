/**
 * code.gs
 * Webアプリのメイン動作。Googleチャット風のメッセージ送受信を担当する。
 * 定数（SPREADSHEET_ID やシート名・ヘッダー）は setup.gs を参照する。
 */

/**
 * Webアプリのエントリーポイント。
 * URLパラメータ ?app=ocr の場合は OCR画面（ocr.html）、
 * それ以外はチャット画面（index.html）を返す。
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.app === 'ocr') {
    return HtmlService.createHtmlOutputFromFile('ocr')
      .setTitle('PDF OCR')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
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

/* ============================================================
 * PDF OCR 機能（Gemini API）
 * ============================================================ */

/**
 * フェーズ0: APIキーをスクリプトプロパティに保存する。
 * 実行前に下の API_KEY に Google AI Studio のキーを貼り付けてください。
 */
function setupApiKey() {
  // ここに Gemini API キーを貼り付ける
  var API_KEY = 'YOUR_GEMINI_API_KEY';

  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', API_KEY);
  Logger.log('設定完了');
}

/**
 * フェーズ0: APIキーが設定済みか確認する。
 * 設定済みなら先頭10文字だけ、未設定なら「未設定です」とログ表示する。
 */
function checkApiKey() {
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (key) {
    Logger.log(key.substring(0, 10));
  } else {
    Logger.log('未設定です');
  }
}

/**
 * フェーズ1: PDFを Gemini API でOCRし、テキストを返す。
 * エラー時は例外をthrowせず、エラーメッセージ文字列を返す。
 * @param {string} fileId DriveのPDFファイルID
 * @return {string} OCR結果テキスト（または失敗時のエラーメッセージ）
 */
function ocrPdfWithGemini(fileId) {
  try {
    // APIキーをスクリプトプロパティから取得
    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) {
      return 'エラー: APIキーが未設定です。setupApiKey() を実行してください。';
    }

    // PDFを取得して Base64 エンコード
    var file = DriveApp.getFileById(fileId);
    var base64Pdf = Utilities.base64Encode(file.getBlob().getBytes());

    // Gemini API（gemini-1.5-flash）へのリクエストを組み立てる
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + apiKey;
    var payload = {
      contents: [
        {
          parts: [
            { text: 'このPDFの文字をすべてOCRして、テキストとして出力してください。' },
            {
              inline_data: {
                mime_type: 'application/pdf',
                data: base64Pdf
              }
            }
          ]
        }
      ]
    };

    var options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    // API呼び出し
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    var body = response.getContentText();

    if (code !== 200) {
      return 'エラー: APIリクエスト失敗 (HTTP ' + code + ') ' + body;
    }

    // レスポンスからテキストを取り出す
    var json = JSON.parse(body);
    if (!json.candidates || !json.candidates.length) {
      return 'エラー: OCR結果が取得できませんでした。' + body;
    }
    var text = json.candidates[0].content.parts[0].text;
    return text;

  } catch (e) {
    // 例外はthrowせずメッセージ文字列で返す
    return 'エラー: ' + e.message;
  }
}

/**
 * フェーズ3: OCR結果をスプレッドシートに保存する。
 * 保存先シート「OCR結果」が無ければ自動作成する。
 * 1行に [実行日時, fileId, ファイル名, OCRテキスト] を追記する。
 * @param {string} fileId  対象PDFのファイルID
 * @param {string} ocrText OCR結果テキスト
 * @return {string} 完了メッセージ
 */
function saveOcrResult(fileId, ocrText) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('OCR結果');

  // シートが無ければ作成し、ヘッダー行を設定する
  if (!sheet) {
    sheet = ss.insertSheet('OCR結果');
    sheet.appendRow(['実行日時', 'fileId', 'ファイル名', 'OCRテキスト']);
  }

  // ファイル名を取得（取得できない場合は空文字）
  var fileName = '';
  try {
    fileName = DriveApp.getFileById(fileId).getName();
  } catch (e) {
    fileName = '';
  }

  sheet.appendRow([new Date(), fileId, fileName, ocrText]);
  return '保存しました';
}
