/**
 * code.gs
 * AIチャットボットのメイン動作。
 * 分野（労務系・経理系・購買系）ごとに参照する規程シートを切り替える。
 * 分野の定義やシート名は setup.gs を参照。
 *
 * 用語：
 * - 分野   … 固定の大枠。メニューの選択肢で、規程シートに対応する。
 * - カテゴリ … トラブル・請求など、追加ルールを分類する小分類（自由に設定）。
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
 * - ?app=bot&cat=<id>    … 分野別チャット画面（bot.html）
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
    var area = findArea_(params.cat);
    if (!area) {
      // 不正な分野ならメニューに戻す
      return renderMenu_();
    }
    // 分野の情報を埋め込んで返す（iframe内からURLを読めないため）
    var page = HtmlService.createTemplateFromFile('bot');
    page.categoryJson = JSON.stringify({ id: area.id, name: area.name });
    return page.evaluate()
      .setTitle('AIチャットボット｜' + area.name)
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
  page.categoriesJson = JSON.stringify(AREAS.map(function (a) {
    return { id: a.id, name: a.name, desc: a.desc, icon: a.icon };
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
 * 分野IDから定義を探す。見つからなければ null。
 */
function findArea_(id) {
  for (var i = 0; i < AREAS.length; i++) {
    if (AREAS[i].id === id) {
      return AREAS[i];
    }
  }
  return null;
}

/**
 * 分野名から定義を探す。見つからなければ null。
 */
function findAreaByName_(name) {
  for (var i = 0; i < AREAS.length; i++) {
    if (AREAS[i].name === name) {
      return AREAS[i];
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
 * 指定分野の規程シートから本文を読み込んで1つの文字列にする。
 * 1行目は説明行のため除外し、2行目以降を連結する。
 * @param {Object} area 分野定義
 * @return {string} 規程の全文（未登録なら空文字）
 */
function getRulesText_(area) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(area.sheet);
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
 * @param {string} areaId   分野ID
 * @param {string} question 社員からの質問
 * @return {string} 回答テキスト（または失敗時のエラーメッセージ）
 */
function askBot(areaId, question) {
  var area = findArea_(areaId);
  try {
    question = (question || '').toString().trim();
    if (!question) {
      return '質問を入力してください。';
    }
    if (!area) {
      Logger.log('不正な分野ID: ' + areaId);
      return USER_ERROR_CONTACT;
    }

    var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) {
      Logger.log('GEMINI_API_KEY が未設定です。setupApiKey() を実行してください。');
      return USER_ERROR_CONTACT;
    }

    // 追加ルールは規程シートの末尾に書き込まれているので、シートを読むだけでよい
    var rules = getRulesText_(area);

    if (!rules) {
      Logger.log('「' + area.sheet + '」シートが空です。');
      return area.name + 'の規程がまだ登録されていません。管理者にお問い合わせください。';
    }

    // プロンプトを組み立てる（基本指示＋規程本文＋質問）
    var prompt = BOT_SYSTEM_INSTRUCTION.replace('{CATEGORY}', area.name) +
      '\n\n===== ' + area.name + 'の規程ここから =====\n' + rules +
      '\n===== ' + area.name + 'の規程ここまで =====\n\n' +
      '【社員からの質問】\n' + question;

    var answer = callGeminiText_(prompt, apiKey);

    // 質問と回答をログシートに記録する
    logBotQa_(area.name, question, answer);

    return answer;

  } catch (e) {
    // 技術的な詳細はログにだけ残し、画面には平易な文言を返す
    Logger.log('askBot エラー: ' + e.message);
    // 失敗した質問も履歴に残す
    logBotQa_(area ? area.name : '', question, USER_ERROR_RETRY);
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
function logBotQa_(areaName, question, answer) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(BOT_LOG_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(BOT_LOG_SHEET);
      sheet.appendRow(BOT_LOG_HEADER);
    }
    sheet.appendRow([new Date(), areaName, getCurrentUser(), question, answer]);
  } catch (e) {
    // ログ失敗は無視
  }
}

/**
 * ログイン中の社員の質問履歴を新しい順に取得する。
 * 指定分野の質問のみ、重複を除いて最大50件返す。
 * @param {string} areaId 分野ID
 * @return {Array<Object>} [{question, at}] at はエポックミリ秒
 */
function getQuestionHistory(areaId) {
  try {
    var area = findArea_(areaId);
    if (!area) {
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
    // 列: 0=日時 1=分野 2=社員 3=質問 4=回答
    for (var i = values.length - 1; i >= 1; i--) {
      var row = values[i];
      if (row[1] !== area.name) continue;     // この分野のみ
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
 * 設定シートを取得する（無ければ作成する）。
 * 形式は [ID, 分野, カテゴリ, 追加ルール]（1行 = 1ルール）。
 */
function getSettingsSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET);
    sheet.appendRow(SETTINGS_HEADER);
  }
  return sheet;
}

/**
 * 設定シートの全ルールを読む。
 * @return {Array<Object>} [{ruleId, area, category, text}]
 */
function readAllRules_() {
  try {
    var values = getSettingsSheet_().getDataRange().getValues();
    var list = [];
    // 列: 0=ID 1=分野 2=カテゴリ 3=追加ルール
    for (var i = 1; i < values.length; i++) {
      var area = (values[i][1] || '').toString().trim();
      var text = (values[i][3] || '').toString().trim();
      if (!area || !text) continue;
      list.push({
        ruleId: (values[i][0] || '').toString().trim(),
        area: area,
        category: (values[i][2] || '').toString().trim(),
        text: text
      });
    }
    return list;
  } catch (e) {
    Logger.log('readAllRules_ エラー: ' + e.message);
    return [];
  }
}

/**
 * 追加ルールを規程シートに書き出す1行分のテキストを組み立てる。
 * 形式：カテゴリ：ルール本文[ID]（カテゴリが空なら ルール本文[ID]）
 */
function buildRuleLine_(categoryName, ruleId, text) {
  // 改行を空白に潰して、必ずA列1行に収める
  var oneLine = String(text).replace(/[\r\n]+/g, ' ').trim();
  var head = (categoryName || '').toString().trim();
  return (head ? head + '：' : '') + oneLine + '[' + ruleId + ']';
}

/**
 * 規程シートの末尾に、その分野の追加ルールを書き出す。
 * 既に書かれている追加ルール行（末尾が [ID] の行）はいったん全部消してから入れ直す。
 * @param {Object} area 分野定義
 */
function syncRulesToSheet_(area) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(area.sheet);
  if (!sheet) {
    Logger.log('規程シートがありません: ' + area.sheet);
    return;
  }

  // 既存の追加ルール行を削除（下から消すと行番号がずれない）
  var marker = new RegExp('\\[' + area.prefix + '\\d+\\]\\s*$');
  var last = sheet.getLastRow();
  if (last >= 2) {
    var colA = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (var i = colA.length - 1; i >= 0; i--) {
      if (marker.test((colA[i][0] || '').toString())) {
        sheet.deleteRow(i + 2);
      }
    }
  }

  // 現在の追加ルールを末尾に追記
  var all = readAllRules_();
  for (var n = 0; n < all.length; n++) {
    if (all[n].area !== area.name) continue;
    sheet.appendRow([buildRuleLine_(all[n].category, all[n].ruleId, all[n].text)]);
  }
}

/**
 * 指定分野で既に使われている追加ルール番号の最大値を返す。
 * 設定シートの既存IDと、画面から送られてきたIDの両方を見る。
 * （削除された番号を再利用しないため、削除前の状態を渡すこと）
 * @param {Object} area   分野定義
 * @param {Array}  values 設定シートの全データ（削除前）
 * @param {Array}  rules  画面から送られてきたルール配列
 * @return {number} 最大番号（1件も無ければ 0）
 */
function maxRuleNumber_(area, values, rules) {
  var re = new RegExp('^' + area.prefix + '(\\d+)$');
  var max = 0;

  for (var i = 1; i < values.length; i++) {
    var m = re.exec((values[i][0] || '').toString().trim());
    if (m) {
      var n1 = parseInt(m[1], 10);
      if (n1 > max) max = n1;
    }
  }
  for (var k = 0; k < (rules || []).length; k++) {
    var m2 = re.exec((rules[k].ruleId || '').toString().trim());
    if (m2) {
      var n2 = parseInt(m2[1], 10);
      if (n2 > max) max = n2;
    }
  }
  return max;
}

/**
 * 設定画面用：分野一覧・既存カテゴリ・全追加ルールを取得する（パスワードが必要）。
 * @param {string} password 設定画面のパスワード
 * @return {Object|null} 正しければ {canEdit, areas, categories, rules}、違えば null
 */
function getBotSettings(password) {
  if (!checkSettingsPassword(password)) {
    return null;
  }
  try {
    var rules = readAllRules_();

    // 既にカテゴリとして使われている名前を集める（プルダウンの選択肢）
    var seen = {};
    var categories = [];
    for (var i = 0; i < rules.length; i++) {
      var c = rules[i].category;
      if (c && !seen[c]) {
        seen[c] = true;
        categories.push(c);
      }
    }
    categories.sort();

    var areas = AREAS.map(function (a) {
      return { id: a.id, name: a.name, sheet: a.sheet, prefix: a.prefix };
    });

    return { canEdit: isAdmin_(), areas: areas, categories: categories, rules: rules };
  } catch (e) {
    Logger.log('getBotSettings エラー: ' + e.message);
    return { canEdit: false, areas: [], categories: [], rules: [] };
  }
}

/**
 * 設定画面用：追加ルールを保存する（パスワードが必要）。
 * ルールは [{ruleId, area, category, text}] のフラットな配列で受け取る。
 * 保存後、各規程シートの末尾へ反映する。
 * @param {Object} payload  {rules:[...]}
 * @param {string} password 設定画面のパスワード
 * @return {string} 画面に出すメッセージ
 */
function saveBotSettings(payload, password) {
  try {
    if (!checkSettingsPassword(password)) {
      return 'パスワードが正しくありません。開き直してもう一度お試しください。';
    }
    if (!isAdmin_()) {
      return '保存する権限がありません。管理者にお問い合わせください。';
    }

    payload = payload || {};
    var incoming = payload.rules || [];

    var sheet = getSettingsSheet_();
    var original = sheet.getDataRange().getValues(); // 採番の基準（削除前の状態）

    // 分野ごとに採番カウンタを用意する
    var counters = {};
    for (var a = 0; a < AREAS.length; a++) {
      counters[AREAS[a].name] = maxRuleNumber_(AREAS[a], original, incoming);
    }

    // 保存する行を組み立てる
    var rows = [];
    var touched = {};  // 規程シートへの反映が必要な分野
    for (var i = 0; i < incoming.length; i++) {
      var area = findAreaByName_((incoming[i].area || '').toString().trim());
      if (!area) continue; // 未知の分野は無視

      var text = (incoming[i].text || '').toString().replace(/[\r\n]+/g, ' ').trim();
      if (!text) continue; // 空欄は登録しない

      var category = (incoming[i].category || '').toString().replace(/[\r\n]+/g, ' ').trim();

      // 既存IDはそのまま使い、新規は続き番号を振る
      var ruleId = (incoming[i].ruleId || '').toString().trim();
      if (!ruleId) {
        counters[area.name] = (counters[area.name] || 0) + 1;
        ruleId = area.prefix + counters[area.name];
      }
      rows.push([ruleId, area.name, category, text]);
      touched[area.name] = area;
    }

    // 元々ルールがあった分野も、消えた行を反映するため対象に入れる
    for (var r = 1; r < original.length; r++) {
      var oldName = (original[r][1] || '').toString().trim();
      if (!oldName || touched[oldName]) continue;
      var oldArea = findAreaByName_(oldName);
      if (oldArea) touched[oldName] = oldArea;
    }

    // 分野の並び順 → ID番号順に並べ替えて見やすくする
    rows.sort(function (x, y) {
      var ox = areaOrder_(x[1]);
      var oy = areaOrder_(y[1]);
      if (ox !== oy) return ox - oy;
      return ruleNumber_(x[0]) - ruleNumber_(y[0]);
    });

    // 設定シートを書き戻す（ヘッダーは残す）
    var last = sheet.getLastRow();
    if (last > 1) {
      sheet.getRange(2, 1, last - 1, SETTINGS_HEADER.length).clearContent();
    }
    if (rows.length) {
      sheet.getRange(2, 1, rows.length, SETTINGS_HEADER.length).setValues(rows);
    }
    SpreadsheetApp.flush();

    // 各規程シートの末尾へ反映
    for (var key in touched) {
      if (Object.prototype.hasOwnProperty.call(touched, key)) {
        syncRulesToSheet_(touched[key]);
      }
    }

    return '保存しました';

  } catch (e) {
    Logger.log('saveBotSettings エラー: ' + e.message);
    return USER_ERROR_RETRY;
  }
}

/**
 * 分野名の並び順（AREAS の定義順）。未知は最後。
 */
function areaOrder_(name) {
  for (var i = 0; i < AREAS.length; i++) {
    if (AREAS[i].name === name) return i;
  }
  return AREAS.length;
}

/**
 * 追加ルールIDから番号部分を取り出す（例 R12 → 12）。
 */
function ruleNumber_(ruleId) {
  var m = /(\d+)$/.exec((ruleId || '').toString());
  return m ? parseInt(m[1], 10) : 0;
}
