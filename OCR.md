# 退避：PDF / 画像 OCR 機能

AIチャットボットへの作り替えに伴い、稼働コードから外した OCR 機能の控えです。
復活させたい場合は、下記を `code.gs` に貼り戻し、`ocr.html` を作成して
`doGet` に `?app=ocr` の分岐を追加してください。

> 依存：`SPREADSHEET_ID`（setup.gs）、`USER_ERROR_RETRY` / `USER_ERROR_CONTACT`、
> `setupApiKey()` / `checkApiKey()`（いずれも稼働側の code.gs に残っています）

## doGet に戻す分岐

```javascript
if (app === 'ocr') {
  return HtmlService.createHtmlOutputFromFile('ocr')
    .setTitle('OCR (PDF / 画像)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
```

## code.gs から外したコード

```javascript
// OCR時にPDF・画像と一緒に送る固定プロンプト文（ここを編集すれば全関数に反映される）
var OCR_PROMPT = 'この画像またはPDFの文字をすべてOCRして、テキストとして出力してください。';

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
```

## ocr.html

```html
<!DOCTYPE html>
<html lang="ja">
<head>
  <base target="_top">
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>OCR (PDF / 画像)</title>
  <style>
    body { font-family: sans-serif; max-width: 700px; margin: 20px auto; padding: 0 16px; }
    h1 { font-size: 20px; }
    .uploader { border: 2px dashed #b0c4f0; border-radius: 10px; padding: 20px; text-align: center; margin-bottom: 12px; background: #f7f9ff; }
    .uploader input[type="file"] { width: 100%; }
    #fileName { display: block; margin-top: 8px; font-size: 13px; color: #555; }
    button { padding: 10px 16px; font-size: 15px; border: none; border-radius: 6px; background: #1a73e8; color: #fff; cursor: pointer; }
    button:disabled { opacity: .5; cursor: default; }
    #status { color: #1a73e8; margin: 8px 0; min-height: 1.2em; }
    pre { background: #f5f5f5; border: 1px solid #ddd; border-radius: 6px; padding: 12px; white-space: pre-wrap; word-break: break-word; }
  </style>
</head>
<body>
  <h1>OCR (Gemini)</h1>
  <p><a href="#" id="toChat">← チャットに戻る</a></p>

  <!-- PDF・画像をその場でアップロードして解析する -->
  <div class="uploader">
    <!-- Safari(iOS)は accept に複数タイプ指定するとダイアログが開かない不具合があるため accept は付けない -->
    <input type="file" id="pdfFile">
    <span id="fileName">PDF または 画像（PNG/JPEG等）を選択</span>
  </div>
  <button id="run" onclick="runOcr()">OCR実行</button>

  <div id="status"></div>
  <pre id="result"></pre>

  <!-- OCR結果が出たら表示する保存ボタン -->
  <button id="save" onclick="saveResult()" style="display:none;">スプレッドシートに保存</button>
  <span id="saveStatus" style="margin-left:8px;"></span>

  <script>
    // チャット画面（パラメータなし）に戻る
    google.script.run.withSuccessHandler(function (url) {
      document.getElementById('toChat').onclick = function (ev) {
        ev.preventDefault();
        window.top.location.href = url;
      };
    }).getWebAppUrl();

    // 直近のファイル名とOCRテキストを保持する
    var lastFileName = '';
    var lastText = '';

    // 選択したファイル名を表示
    document.getElementById('pdfFile').addEventListener('change', function () {
      var f = this.files[0];
      document.getElementById('fileName').textContent = f ? f.name : '';
    });

    function runOcr() {
      var fileInput = document.getElementById('pdfFile');
      var file = fileInput.files[0];
      if (!file) { alert('PDFまたは画像ファイルを選択してください'); return; }
      // PDF・画像以外は弾く（accept未指定のため）
      var type = file.type || '';
      if (type !== 'application/pdf' && type.indexOf('image/') !== 0) {
        alert('PDFまたは画像ファイルを選択してください');
        return;
      }

      document.getElementById('run').disabled = true;
      document.getElementById('status').textContent = '読み取り中...';
      document.getElementById('result').textContent = '';
      document.getElementById('save').style.display = 'none';
      document.getElementById('saveStatus').textContent = '';

      // ブラウザでファイルをBase64化してサーバーへ送る
      var reader = new FileReader();
      reader.onload = function () {
        // data URL の "data:application/pdf;base64," 以降を取り出す
        var base64 = reader.result.split(',')[1];
        google.script.run
          .withSuccessHandler(function (text) {
            document.getElementById('status').textContent = '完了';
            document.getElementById('result').textContent = text;
            document.getElementById('run').disabled = false;
            lastFileName = file.name;
            lastText = text;
            document.getElementById('save').style.display = 'inline-block';
          })
          .withFailureHandler(function () {
            document.getElementById('status').textContent = '';
            document.getElementById('run').disabled = false;
            alert('エラーが発生しました。もう一度お試しください。');
          })
          .ocrUploadedPdf(base64, file.type || 'application/pdf');
      };
      reader.onerror = function () {
        document.getElementById('status').textContent = '';
        document.getElementById('run').disabled = false;
        alert('ファイルの読み込みに失敗しました');
      };
      reader.readAsDataURL(file);
    }

    function saveResult() {
      document.getElementById('save').disabled = true;
      document.getElementById('saveStatus').textContent = '保存中...';

      google.script.run
        .withSuccessHandler(function (msg) {
          document.getElementById('saveStatus').textContent = msg; // 「保存しました」
          document.getElementById('save').disabled = false;
        })
        .withFailureHandler(function () {
          document.getElementById('saveStatus').textContent = '';
          document.getElementById('save').disabled = false;
          alert('保存できませんでした。もう一度お試しください。');
        })
        .saveUploadedOcrResult(lastFileName, lastText);
    }
  </script>
</body>
</html>
```
