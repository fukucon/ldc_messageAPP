/**
 * code.gs
 * Webアプリのメイン動作。Googleチャット風のメッセージ送受信を担当する。
 * 定数（SPREADSHEET_ID やシート名・ヘッダー）は setup.gs を参照する。
 */

/**
 * Webアプリのエントリーポイント。
 * デフォルトは就労規則チャットボット（bot.html）。
 * ?app=chat でチャット画面、?app=ocr で OCR画面（どちらも非表示の隠し機能として残す）。
 */
function doGet(e) {
  var app = (e && e.parameter && e.parameter.app) || '';

  if (app === 'ocr') {
    return HtmlService.createHtmlOutputFromFile('ocr')
      .setTitle('OCR (PDF / 画像)')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (app === 'settings') {
    return HtmlService.createHtmlOutputFromFile('settings')
      .setTitle('チャットボット設定')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (app === 'chat') {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('LDC Chat')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // デフォルト：就労規則チャットボット
  return HtmlService.createHtmlOutputFromFile('bot')
    .setTitle('就労規則チャットボット')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * WebアプリのURLを返す（画面遷移用）。
 */
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
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
 * 利用者向けエラーメッセージ
 * 専門用語は画面に出さず、詳細は Logger（実行ログ）に残す。
 * ============================================================ */

// 一時的な失敗：再送信を促す
var USER_ERROR_RETRY = 'エラーが発生しました。もう一度送信してください。';
// 設定不備など、利用者が再送しても直らない場合
var USER_ERROR_CONTACT = 'ただいまご利用いただけません。管理者にお問い合わせください。';

/* ============================================================
 * PDF OCR 機能（Gemini API）
 * ============================================================ */

// OCR時にPDF・画像と一緒に送る固定プロンプト文（ここを編集すれば全関数に反映される）
var OCR_PROMPT = 'この画像またはPDFの文字をすべてOCRして、テキストとして出力してください。';

/**
 * フェーズ0: APIキーをスクリプトプロパティに保存する。
 * 実行前に下の API_KEY に Google AI Studio のキーを貼り付けてください。
 */
function setupApiKey() {
  // ここに Gemini API キーを貼り付ける
  var API_KEY = 'YOUR_GEMINI_API_KEY';

  // プレースホルダのまま実行すると保存済みキーを消してしまうため、誤上書きを防ぐ
  if (!API_KEY || API_KEY === 'YOUR_GEMINI_API_KEY') {
    Logger.log('API_KEY が未入力です。実際のキーを貼り付けてから実行してください（上書きせず終了）。');
    return;
  }

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
      Logger.log('GEMINI_API_KEY が未設定です。setupApiKey() を実行してください。');
      return USER_ERROR_CONTACT;
    }

    // ファイルを取得して Base64 エンコード（PDF・画像どちらも対応）
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var base64Data = Utilities.base64Encode(blob.getBytes());
    var mimeType = blob.getContentType() || 'application/pdf';

    // Gemini API（gemini-2.5-flash）へのリクエストを組み立てる
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;
    var payload = {
      contents: [
        {
          parts: [
            { text: OCR_PROMPT },
            {
              inline_data: {
                mime_type: mimeType,
                data: base64Data
              }
            }
          ]
        }
      ],
      // OCR用途では思考(thinking)をオフにして高速化・空応答を防ぐ
      generationConfig: {
        temperature: 0,
        maxOutputTokens: 8192,
        thinkingConfig: { thinkingBudget: 0 }
      }
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
      Logger.log('Gemini APIリクエスト失敗 (HTTP ' + code + ') ' + body);
      return USER_ERROR_RETRY;
    }

    // レスポンスからテキストを取り出す
    var json = JSON.parse(body);
    if (!json.candidates || !json.candidates.length) {
      Logger.log('OCR結果が取得できませんでした: ' + body);
      return USER_ERROR_RETRY;
    }
    var text = json.candidates[0].content.parts[0].text;
    return text;

  } catch (e) {
    // 技術的な詳細はログにだけ残し、画面には平易な文言を返す
    Logger.log('ocrPdfWithGemini エラー: ' + e.message);
    return USER_ERROR_RETRY;
  }
}

/**
 * アップロードされたPDF・画像（Base64）をその場で Gemini OCR する。
 * Driveに保存せず、ブラウザから渡されたBase64データを直接送信する。
 * エラー時は例外をthrowせず、エラーメッセージ文字列を返す。
 * @param {string} base64Data Base64エンコード済みデータ
 * @param {string} mimeType   データのMIMEタイプ（例: application/pdf, image/png）
 * @return {string} OCR結果テキスト（または失敗時のエラーメッセージ）
 */
function ocrUploadedPdf(base64Data, mimeType) {
  try {
    // APIキーをスクリプトプロパティから取得
    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) {
      Logger.log('GEMINI_API_KEY が未設定です。setupApiKey() を実行してください。');
      return USER_ERROR_CONTACT;
    }
    if (!base64Data) {
      return 'ファイルが選択されていません。';
    }
    // MIMEタイプ未指定の場合はPDFとして扱う
    mimeType = mimeType || 'application/pdf';

    // Gemini API（gemini-2.5-flash）へのリクエストを組み立てる
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;
    var payload = {
      contents: [
        {
          parts: [
            { text: OCR_PROMPT },
            {
              inline_data: {
                mime_type: mimeType,
                data: base64Data
              }
            }
          ]
        }
      ],
      // OCR用途では思考(thinking)をオフにして高速化・空応答を防ぐ
      generationConfig: {
        temperature: 0,
        maxOutputTokens: 8192,
        thinkingConfig: { thinkingBudget: 0 }
      }
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
      Logger.log('Gemini APIリクエスト失敗 (HTTP ' + code + ') ' + body);
      return USER_ERROR_RETRY;
    }

    var json = JSON.parse(body);
    if (!json.candidates || !json.candidates.length) {
      Logger.log('OCR結果が取得できませんでした: ' + body);
      return USER_ERROR_RETRY;
    }
    return json.candidates[0].content.parts[0].text;

  } catch (e) {
    Logger.log('ocrUploadedPdf エラー: ' + e.message);
    return USER_ERROR_RETRY;
  }
}

/**
 * アップロードしたPDFのOCR結果をスプレッドシートに保存する。
 * 保存先シート「OCR結果」が無ければ自動作成する。
 * 1行に [実行日時, fileId(空), ファイル名, OCRテキスト] を追記する。
 * @param {string} fileName アップロードしたファイル名
 * @param {string} ocrText  OCR結果テキスト
 * @return {string} 完了メッセージ
 */
function saveUploadedOcrResult(fileName, ocrText) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('OCR結果');
  if (!sheet) {
    sheet = ss.insertSheet('OCR結果');
    sheet.appendRow(['実行日時', 'fileId', 'ファイル名', 'OCRテキスト']);
  }
  sheet.appendRow([new Date(), '', fileName || '', ocrText]);
  return '保存しました';
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

/* ============================================================
 * 就労規則チャットボット（Gemini API）
 * ============================================================ */

// ボットへの基本指示（就労規則の範囲だけで答えさせる）
var BOT_SYSTEM_INSTRUCTION =
  'あなたは会社の就労規則に関する社内アシスタントです。' +
  '以下に示す「就労規則」の内容だけに基づいて、社員の質問に日本語でわかりやすく回答してください。' +
  '就労規則に書かれていない事項については推測せず、「就労規則には記載がありません。担当部署にご確認ください。」と答えてください。' +
  '可能であれば、根拠となる条文や項目名も添えてください。';

/**
 * 「就労規則」シートから本文を読み込んで1つの文字列にする。
 * 1行目は説明行のため除外し、2行目以降を連結する。
 * @return {string} 就労規則の全文（未登録なら空文字）
 */
function getWorkRules_() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(RULES_SHEET);
  if (!sheet) {
    return '';
  }
  var values = sheet.getDataRange().getValues();
  var lines = [];
  // i=1 から（1行目の説明を除外）
  for (var i = 1; i < values.length; i++) {
    var row = values[i]
      .filter(function (c) { return c !== '' && c !== null; })
      .join(' ');
    if (row) {
      lines.push(row);
    }
  }
  return lines.join('\n');
}

/**
 * 就労規則についての質問に回答する。
 * 就労規則の全文＋質問を Gemini に渡し、回答テキストを返す。
 * エラー時は例外をthrowせず、エラーメッセージ文字列を返す。
 * @param {string} question 社員からの質問
 * @return {string} 回答テキスト（または失敗時のエラーメッセージ）
 */
function askWorkRules(question) {
  try {
    question = (question || '').toString().trim();
    if (!question) {
      return '質問を入力してください。';
    }

    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) {
      Logger.log('GEMINI_API_KEY が未設定です。setupApiKey() を実行してください。');
      return USER_ERROR_CONTACT;
    }

    var rules = getWorkRules_();
    // 設定画面で追加された指示・ルールを読み込む
    var settings = getBotSettings();

    if (!rules && !settings.extraRules) {
      Logger.log('「就労規則」シートと追加ルールがどちらも空です。');
      return '就労規則がまだ登録されていません。管理者にお問い合わせください。';
    }

    // プロンプトを組み立てる（指示＋追加指示＋就労規則＋追加ルール＋質問）
    var prompt = BOT_SYSTEM_INSTRUCTION;

    if (settings.extraInstruction) {
      prompt += '\n\n【追加の指示】\n' + settings.extraInstruction;
    }

    prompt += '\n\n===== 就労規則ここから =====\n' + rules;

    if (settings.extraRules) {
      prompt += '\n\n----- 追加ルール（就労規則を補足する社内ルール） -----\n' + settings.extraRules;
    }

    prompt += '\n===== 就労規則ここまで =====\n\n' +
      '【社員からの質問】\n' + question;

    var answer = callGeminiText_(prompt, apiKey);

    // 質問と回答をログシートに記録する
    logBotQa_(question, answer);

    return answer;

  } catch (e) {
    // 技術的な詳細はログにだけ残し、画面には平易な文言を返す
    Logger.log('askWorkRules エラー: ' + e.message);
    // 失敗した質問も履歴に残す
    logBotQa_(question, USER_ERROR_RETRY);
    return USER_ERROR_RETRY;
  }
}

/* ===== 設定（AIへの追加指示・就業規則への追記ルール） ===== */

// 設定画面で保存できる管理者のメールアドレス。
// 空配列のままなら、設定画面のURLを知っている人は誰でも保存できる。
// 例: var ADMIN_EMAILS = ['soumu@example.com', 'hr@example.com'];
var ADMIN_EMAILS = [];

/**
 * 現在のユーザーが設定を保存できるかどうか。
 */
function isAdmin_() {
  if (!ADMIN_EMAILS.length) {
    return true; // 未設定なら制限しない
  }
  var me = (getCurrentUser() || '').toLowerCase();
  for (var i = 0; i < ADMIN_EMAILS.length; i++) {
    if (String(ADMIN_EMAILS[i]).toLowerCase() === me) {
      return true;
    }
  }
  return false;
}

/**
 * 設定シートを取得する（無ければ作成する）。
 */
function getSettingsSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET);
    sheet.appendRow(['項目', '内容']);
    sheet.appendRow([SETTING_INSTRUCTION, '']);
    sheet.appendRow([SETTING_EXTRA_RULES, '']);
  }
  return sheet;
}

/**
 * 設定シートから項目名で値を読む。
 */
function readSetting_(sheet, key) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === key) {
      return (values[i][1] || '').toString();
    }
  }
  return '';
}

/**
 * 設定シートに項目名で値を書く（無ければ行を追加する）。
 */
function writeSetting_(sheet, key, value) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

/**
 * 設定画面用：現在の設定を取得する。
 * @return {Object} {extraInstruction, extraRules, rulesLength, canEdit}
 */
function getBotSettings() {
  try {
    var sheet = getSettingsSheet_();
    return {
      extraInstruction: readSetting_(sheet, SETTING_INSTRUCTION),
      extraRules: readSetting_(sheet, SETTING_EXTRA_RULES),
      rulesLength: getWorkRules_().length, // 就労規則シート本体の文字数（目安表示用）
      canEdit: isAdmin_()
    };
  } catch (e) {
    Logger.log('getBotSettings エラー: ' + e.message);
    return { extraInstruction: '', extraRules: '', rulesLength: 0, canEdit: false };
  }
}

/**
 * 設定画面用：設定を保存する。
 * @param {Object} settings {extraInstruction, extraRules}
 * @return {string} 画面に出すメッセージ
 */
function saveBotSettings(settings) {
  try {
    if (!isAdmin_()) {
      return '保存する権限がありません。管理者にお問い合わせください。';
    }
    settings = settings || {};
    var sheet = getSettingsSheet_();
    writeSetting_(sheet, SETTING_INSTRUCTION, (settings.extraInstruction || '').toString());
    writeSetting_(sheet, SETTING_EXTRA_RULES, (settings.extraRules || '').toString());
    return '保存しました';
  } catch (e) {
    Logger.log('saveBotSettings エラー: ' + e.message);
    return USER_ERROR_RETRY;
  }
}

/**
 * ログイン中の社員の質問履歴を新しい順に取得する。
 * 同じ内容の質問は重複を除き、最大50件返す。
 * @return {Array<Object>} [{question, at}] at はエポックミリ秒
 */
function getQuestionHistory() {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(BOT_LOG_SHEET);
    if (!sheet) {
      return [];
    }
    var values = sheet.getDataRange().getValues();
    var user = getCurrentUser();
    var seen = {};
    var list = [];

    // 新しい順に走査（1行目はヘッダーなので除外）
    for (var i = values.length - 1; i >= 1; i--) {
      var row = values[i];
      if (row[1] !== user) continue; // 自分の質問のみ

      var q = (row[2] || '').toString().trim();
      if (!q || seen[q]) continue;   // 空・重複は除外
      seen[q] = true;

      list.push({
        question: q,
        at: row[0] ? new Date(row[0]).getTime() : 0
      });
      if (list.length >= 50) break;
    }
    return list;

  } catch (e) {
    Logger.log('getQuestionHistory エラー: ' + e.message);
    return [];
  }
}

/**
 * Gemini にテキストプロンプトを送り、回答テキストを返す内部ヘルパー。
 * @param {string} promptText プロンプト全文
 * @param {string} apiKey     APIキー
 * @return {string} 回答テキスト
 */
function callGeminiText_(promptText, apiKey) {
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;
  var payload = {
    contents: [
      { parts: [{ text: promptText }] }
    ],
    generationConfig: {
      temperature: 0.2,
      maxOutputTokens: 2048,
      thinkingConfig: { thinkingBudget: 0 }
    }
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();
  var body = response.getContentText();
  if (code !== 200) {
    throw new Error('APIリクエスト失敗 (HTTP ' + code + ') ' + body);
  }

  var json = JSON.parse(body);
  if (!json.candidates || !json.candidates.length ||
      !json.candidates[0].content || !json.candidates[0].content.parts) {
    throw new Error('回答が取得できませんでした。' + body);
  }
  return json.candidates[0].content.parts[0].text;
}

/**
 * 質問と回答を「質問ログ」シートに記録する。失敗しても本処理は止めない。
 */
function logBotQa_(question, answer) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(BOT_LOG_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(BOT_LOG_SHEET);
      sheet.appendRow(['日時', '社員', '質問', '回答']);
    }
    sheet.appendRow([new Date(), getCurrentUser(), question, answer]);
  } catch (e) {
    // ログ失敗は無視
  }
}
