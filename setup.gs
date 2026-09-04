/**
 * setup.gs
 * スプレッドシートのシート生成・カラム設定を行うコード。
 * 初回に setupBot() を一度だけ手動実行してください。
 *
 * 用語：
 * - 分野   … 労務系・経理系・購買系（固定）。メニューの選択肢で、規程シートに対応する。
 * - カテゴリ … トラブル・請求など、追加ルールを分類する小分類（自由に設定）。
 */

// 利用するスプレッドシートのID（URLの /d/ と /edit の間の文字列）
var SPREADSHEET_ID = '1DIss4Wyo-TiA9tIgdboIoFXq0GlJM1Cx5jmyKNbhXIA';

/**
 * 分野の定義（固定）。
 * ここに1件足すだけで、メニュー・チャット画面・設定画面すべてに反映される。
 * - id     : URLで使う識別子（英数字）
 * - name   : 画面に出す名前
 * - desc   : メニューに出す補足説明
 * - sheet  : 規程本文を貼り付けるシート名
 * - prefix : 追加ルールのID接頭辞（例 R → R1, R2 ...）
 * - icon   : メニューに出すアイコン
 */
var AREAS = [
  { id: 'roumu', name: '労務系', desc: '就業規則・勤怠・休暇・給与など', sheet: '就労規則', prefix: 'R', icon: '👤' },
  { id: 'keiri', name: '経理系', desc: '経費精算・出張旅費・請求など',   sheet: '経理規程', prefix: 'K', icon: '💴' },
  { id: 'kobai', name: '購買系', desc: '発注・稟議・取引先・契約など',   sheet: '購買規程', prefix: 'B', icon: '📦' }
];

// 共通シート名
var BOT_LOG_SHEET = '質問ログ';
var SETTINGS_SHEET = '設定';

// 各シートのヘッダー
var BOT_LOG_HEADER = ['日時', '分野', '社員', '質問', '回答'];
var SETTINGS_HEADER = ['ID', '分野', 'カテゴリ', '追加ルール'];

// 規程シートの1行目に入れる説明文
var RULES_SHEET_NOTE = '▼この下(A2以降)に規程の本文を貼り付けてください';

// 分野が分からない旧データを引き継ぐときのカテゴリ名
var UNCATEGORIZED = '未分類';

/**
 * AIチャットボット用のシートをまとめて生成する。
 * 何度実行しても既存データは壊さない。
 */
function setupBot() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // 分野ごとの規程シート
  for (var i = 0; i < AREAS.length; i++) {
    setupRulesSheet_(ss, AREAS[i].sheet);
  }

  // 質問ログシート
  setupSheet_(ss, BOT_LOG_SHEET, BOT_LOG_HEADER);

  // 設定シート（1行 = 1つの追加ルール）
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
 * 設定シートを [ID, 分野, カテゴリ, 追加ルール] 形式で用意する。
 * 旧形式のデータがあれば、1行1ルールに分解して引き継ぐ。
 */
function setupSettingsSheet_(ss) {
  var sheet = ss.getSheetByName(SETTINGS_SHEET);
  var migrated = [];   // [分野名, カテゴリ, ルール本文]

  if (sheet) {
    var values = sheet.getDataRange().getValues();
    var header = values.length ? values[0].map(function (h) { return String(h); }) : [];

    if (header[0] === 'ID' && header[1] === 'カテゴリ') {
      // 旧形式C: [ID, カテゴリ(=分野名), 追加ルール]
      for (var a = 1; a < values.length; a++) {
        var areaName = String(values[a][1] || '');
        var text = String(values[a][2] || '').trim();
        if (areaName && text) {
          migrated.push([areaName, UNCATEGORIZED, text]);
        }
      }
      sheet.clear();

    } else if (header[0] === 'カテゴリ') {
      // 旧形式A: [カテゴリ(=分野名), 追加指示, 追加ルール]
      for (var b = 1; b < values.length; b++) {
        var nameB = String(values[b][0] || '');
        if (!nameB) continue;
        migrated = migrated.concat(splitToRules_(nameB, values[b][1]));
        migrated = migrated.concat(splitToRules_(nameB, values[b][2]));
      }
      sheet.clear();

    } else if (header[0] === '項目') {
      // 旧形式B: [項目, 内容] … すべて最初の分野へ引き継ぐ
      for (var c = 1; c < values.length; c++) {
        migrated = migrated.concat(splitToRules_(AREAS[0].name, values[c][1]));
      }
      sheet.clear();
    }
  }

  sheet = setupSheet_(ss, SETTINGS_SHEET, SETTINGS_HEADER);

  // 引き継いだルールにIDを振って書き込む
  if (migrated.length) {
    var counters = {};
    for (var k = 0; k < migrated.length; k++) {
      var prefix = prefixOfAreaName_(migrated[k][0]);
      counters[prefix] = (counters[prefix] || 0) + 1;
      sheet.appendRow([prefix + counters[prefix], migrated[k][0], migrated[k][1], migrated[k][2]]);
    }
    Logger.log('旧形式の設定を ' + migrated.length + ' 件引き継ぎました');
  }

  sheet.setColumnWidth(1, 70);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 560);
  return sheet;
}

/**
 * 複数行のテキストを、1行1ルールの配列に分解する。
 * @return {Array<Array>} [[分野名, カテゴリ, ルール本文], ...]
 */
function splitToRules_(areaName, text) {
  var out = [];
  var s = (text == null ? '' : String(text)).trim();
  if (!s) {
    return out;
  }
  var lines = s.split('\n');
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (line) {
      out.push([areaName, UNCATEGORIZED, line]);
    }
  }
  return out;
}

/**
 * 分野名からIDの接頭辞を返す。見つからなければ 'X'。
 */
function prefixOfAreaName_(name) {
  for (var i = 0; i < AREAS.length; i++) {
    if (AREAS[i].name === name) {
      return AREAS[i].prefix;
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
