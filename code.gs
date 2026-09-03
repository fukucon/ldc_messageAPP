/**
 * code.gs
 * AIチャットボットのメイン動作。
 * カテゴリ（労務系・経理系・購買系）ごとに参照する規程シートを切り替える。
 * カテゴリ定義やシート名は setup.gs を参照。
 *
 * 退避したコード：
 * - PDF/画像OCR機能 … OCR.md
 * - 旧Googleチャット風アプリ … CHAT.md
 */

/* ============================================================
 * 画面の振り分け
 * ============================================================ */

/**
 * Webアプリのエントリーポイント。
 * - パラメータなし        … メニュー画面（index.html）
 * - ?app=bot&cat=<id>    … カテゴリ別チャット画面（bot.html）
 * - ?app=settings        … 設定画面（settings.html）
 */
function doGet(e) {
  var params = (e && e.parameter) || {};
  var app = params.app || '';

  if (app === 'settings') {
    return HtmlService.createHtmlOutputFromFile('settings')
      .setTitle('AIチャットボット設定')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (app === 'bot') {
    var category = findCategory_(params.cat);
    if (!category) {
      // 不正なカテゴリならメニューに戻す
      return renderMenu_();
    }
    // カテゴリ情報を埋め込んで返す（iframe内からURLを読めないため）
    var page = HtmlService.createTemplateFromFile('bot');
    page.categoryJson = JSON.stringify({ id: category.id, name: category.name });
    return page.evaluate()
      .setTitle('AIチャットボット｜' + category.name)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // デフォルト：メニュー画面
  return renderMenu_();
}

/**
 * メニュー画面を組み立てる。
 */
function renderMenu_() {
  var page = HtmlService.createTemplateFromFile('index');
  page.categoriesJson = JSON.stringify(CATEGORIES.map(function (c) {
    return { id: c.id, name: c.name, desc: c.desc, icon: c.icon };
  }));
  return page.evaluate()
    .setTitle('AIチャットボット')
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
 * カテゴリIDから定義を探す。見つからなければ null。
 */
function findCategory_(id) {
  for (var i = 0; i < CATEGORIES.length; i++) {
    if (CATEGORIES[i].id === id) {
      return CATEGORIES[i];
    }
  }
  return null;
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
 * APIキー（Gemini）
 * ============================================================ */

/**
 * APIキーをスクリプトプロパティに保存する。
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
 * APIキーが設定済みか確認する。
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

/* ============================================================
 * 規程の読み込みと回答生成
 * ============================================================ */

/**
 * ボットへの基本指示。{CATEGORY} はカテゴリ名に置き換わる。
 */
var BOT_SYSTEM_INSTRUCTION =
  'あなたは会社の{CATEGORY}の規程に関する社内アシスタントです。' +
  '以下に示す規程の内容だけに基づいて、社員の質問に日本語でわかりやすく回答してください。' +
  '規程に書かれていない事項については推測せず、「規程には記載がありません。担当部署にご確認ください。」と答えてください。' +
  '可能であれば、根拠となる条文や項目名も添えてください。';

/**
 * 指定カテゴリの規程シートから本文を読み込んで1つの文字列にする。
 * 1行目は説明行のため除外し、2行目以降を連結する。
 * @param {Object} category カテゴリ定義
 * @return {string} 規程の全文（未登録なら空文字）
 */
function getRulesText_(category) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(category.sheet);
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
 * 規程についての質問に回答する。
 * 規程の全文＋設定＋質問を Gemini に渡し、回答テキストを返す。
 * エラー時は例外をthrowせず、平易なエラーメッセージ文字列を返す。
 * @param {string} categoryId カテゴリID
 * @param {string} question   社員からの質問
 * @return {string} 回答テキスト（または失敗時のエラーメッセージ）
 */
function askBot(categoryId, question) {
  var category = findCategory_(categoryId);
  try {
    question = (question || '').toString().trim();
    if (!question) {
      return '質問を入力してください。';
    }
    if (!category) {
      Logger.log('不正なカテゴリID: ' + categoryId);
      return USER_ERROR_CONTACT;
    }

    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) {
      Logger.log('GEMINI_API_KEY が未設定です。setupApiKey() を実行してください。');
      return USER_ERROR_CONTACT;
    }

    var rules = getRulesText_(category);
    // 設定画面で追加された指示・ルールを読み込む
    var settings = readCategorySettings_(category.name);

    if (!rules && !settings.extraRules) {
      Logger.log('「' + category.sheet + '」シートと追加ルールがどちらも空です。');
      return category.name + 'の規程がまだ登録されていません。管理者にお問い合わせください。';
    }

    // プロンプトを組み立てる（基本指示＋追加指示＋規程本文＋追加ルール＋質問）
    var prompt = BOT_SYSTEM_INSTRUCTION.replace('{CATEGORY}', category.name);

    if (settings.extraInstruction) {
      prompt += '\n\n【追加の指示】\n' + settings.extraInstruction;
    }

    prompt += '\n\n===== ' + category.name + 'の規程ここから =====\n' + rules;

    if (settings.extraRules) {
      prompt += '\n\n----- 追加ルール（規程を補足する社内ルール） -----\n' + settings.extraRules;
    }

    prompt += '\n===== ' + category.name + 'の規程ここまで =====\n\n' +
      '【社員からの質問】\n' + question;

    var answer = callGeminiText_(prompt, apiKey);

    // 質問と回答をログシートに記録する
    logBotQa_(category.name, question, answer);

    return answer;

  } catch (e) {
    // 技術的な詳細はログにだけ残し、画面には平易な文言を返す
    Logger.log('askBot エラー: ' + e.message);
    // 失敗した質問も履歴に残す
    logBotQa_(category ? category.name : '', question, USER_ERROR_RETRY);
    return USER_ERROR_RETRY;
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

/* ============================================================
 * 質問ログ・履歴
 * ============================================================ */

/**
 * 質問と回答を「質問ログ」シートに記録する。失敗しても本処理は止めない。
 */
function logBotQa_(categoryName, question, answer) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(BOT_LOG_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(BOT_LOG_SHEET);
      sheet.appendRow(BOT_LOG_HEADER);
    }
    sheet.appendRow([new Date(), categoryName, getCurrentUser(), question, answer]);
  } catch (e) {
    // ログ失敗は無視
  }
}

/**
 * ログイン中の社員の質問履歴を新しい順に取得する。
 * 指定カテゴリの質問のみ、重複を除いて最大50件返す。
 * @param {string} categoryId カテゴリID
 * @return {Array<Object>} [{question, at}] at はエポックミリ秒
 */
function getQuestionHistory(categoryId) {
  try {
    var category = findCategory_(categoryId);
    if (!category) {
      return [];
    }
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(BOT_LOG_SHEET);
    if (!sheet) {
      return [];
    }
    var values = sheet.getDataRange().getValues();
    var user = getCurrentUser();
    var seen = {};
    var list = [];

    // 新しい順に走査（1行目はヘッダーなので除外）
    // 列: 0=日時 1=カテゴリ 2=社員 3=質問 4=回答
    for (var i = values.length - 1; i >= 1; i--) {
      var row = values[i];
      if (row[1] !== category.name) continue; // このカテゴリのみ
      if (row[2] !== user) continue;          // 自分の質問のみ

      var q = (row[3] || '').toString().trim();
      if (!q || seen[q]) continue;            // 空・重複は除外
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

/* ============================================================
 * 設定（カテゴリごとの追加指示・追加ルール）
 * ============================================================ */

// 設定画面を開くための簡易パスワード
var SETTINGS_PASSWORD = 'ldcpass';

/**
 * 設定画面のパスワードが正しいか判定する。
 * 照合はサーバー側で行うため、パスワードは画面のソースには出ない。
 * @param {string} password 入力されたパスワード
 * @return {boolean} 正しければ true
 */
function checkSettingsPassword(password) {
  return String(password || '') === SETTINGS_PASSWORD;
}

// 設定画面で保存できる管理者のメールアドレス。
// 空配列のままなら、パスワードを知っている人は誰でも保存できる。
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
 * 設定シートを取得する（無ければカテゴリ行付きで作成する）。
 */
function getSettingsSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET);
    sheet.appendRow(SETTINGS_HEADER);
    for (var i = 0; i < CATEGORIES.length; i++) {
      sheet.appendRow([CATEGORIES[i].name, '', '']);
    }
  }
  return sheet;
}

/**
 * 内部用：指定カテゴリの設定を読む（パスワード不要。回答生成から使う）。
 * @param {string} categoryName カテゴリ名
 * @return {Object} {extraInstruction, extraRules}
 */
function readCategorySettings_(categoryName) {
  try {
    var sheet = getSettingsSheet_();
    var values = sheet.getDataRange().getValues();
    // 列: 0=カテゴリ 1=追加指示 2=追加ルール
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === categoryName) {
        return {
          extraInstruction: (values[i][1] || '').toString(),
          extraRules: (values[i][2] || '').toString()
        };
      }
    }
    return { extraInstruction: '', extraRules: '' };
  } catch (e) {
    Logger.log('readCategorySettings_ エラー: ' + e.message);
    return { extraInstruction: '', extraRules: '' };
  }
}

/**
 * 設定画面用：全カテゴリの設定を取得する（パスワードが必要）。
 * @param {string} password 設定画面のパスワード
 * @return {Object|null} 正しければ {canEdit, categories:[...]}、違えば null
 */
function getBotSettings(password) {
  if (!checkSettingsPassword(password)) {
    return null;
  }
  try {
    var list = CATEGORIES.map(function (c) {
      var s = readCategorySettings_(c.name);
      return {
        id: c.id,
        name: c.name,
        desc: c.desc,
        sheet: c.sheet,
        extraInstruction: s.extraInstruction,
        extraRules: s.extraRules,
        rulesLength: getRulesText_(c).length // 規程本文の文字数（目安表示用）
      };
    });
    return { canEdit: isAdmin_(), categories: list };
  } catch (e) {
    Logger.log('getBotSettings エラー: ' + e.message);
    return { canEdit: false, categories: [] };
  }
}

/**
 * 設定画面用：全カテゴリの設定を保存する（パスワードが必要）。
 * @param {Array<Object>} items    [{id, extraInstruction, extraRules}]
 * @param {string}        password 設定画面のパスワード
 * @return {string} 画面に出すメッセージ
 */
function saveBotSettings(items, password) {
  try {
    if (!checkSettingsPassword(password)) {
      return 'パスワードが正しくありません。開き直してもう一度お試しください。';
    }
    if (!isAdmin_()) {
      return '保存する権限がありません。管理者にお問い合わせください。';
    }

    var sheet = getSettingsSheet_();
    var values = sheet.getDataRange().getValues();
    items = items || [];

    for (var n = 0; n < items.length; n++) {
      var category = findCategory_(items[n].id);
      if (!category) continue;

      var instruction = (items[n].extraInstruction || '').toString();
      var extraRules = (items[n].extraRules || '').toString();

      // 該当カテゴリの行を探して更新（無ければ追加）
      var rowIndex = -1;
      for (var i = 1; i < values.length; i++) {
        if (values[i][0] === category.name) { rowIndex = i + 1; break; }
      }
      if (rowIndex > 0) {
        sheet.getRange(rowIndex, 2, 1, 2).setValues([[instruction, extraRules]]);
      } else {
        sheet.appendRow([category.name, instruction, extraRules]);
      }
    }
    return '保存しました';

  } catch (e) {
    Logger.log('saveBotSettings エラー: ' + e.message);
    return USER_ERROR_RETRY;
  }
}
