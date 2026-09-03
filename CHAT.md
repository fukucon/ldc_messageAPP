# 退避：旧 Googleチャット風アプリ

`index.html` をメニュー画面に作り替えたため、稼働コードから外した旧チャット機能の控えです。
復活させる場合は、下記を `code.gs` / `setup.gs` に貼り戻し、HTMLを別名（例 `chat.html`）で
作成して `doGet` に `?app=chat` の分岐を追加してください。

> 依存：`SPREADSHEET_ID`、`getCurrentUser()`、`setupSheet_()`（いずれも稼働側に残っています）

## doGet に戻す分岐

```javascript
if (app === 'chat') {
  return HtmlService.createHtmlOutputFromFile('chat')
    .setTitle('LDC Chat')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
```

## setup.gs から外したシート定義とセットアップ

```javascript
var ROOMS_SHEET = 'Rooms';
var MESSAGES_SHEET = 'Messages';
var ROOMS_HEADER = ['roomId', 'name', 'createdAt'];
var MESSAGES_HEADER = ['messageId', 'roomId', 'user', 'text', 'createdAt'];

function setup() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var rooms = setupSheet_(ss, ROOMS_SHEET, ROOMS_HEADER);
  setupSheet_(ss, MESSAGES_SHEET, MESSAGES_HEADER);
  if (rooms.getLastRow() < 2) {
    rooms.appendRow([Utilities.getUuid(), '雑談', new Date()]);
  }
  SpreadsheetApp.flush();
  Logger.log('setup 完了');
}
```

## code.gs から外したコード

```javascript
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
```

## 旧 index.html（チャット画面）

```html
<!DOCTYPE html>
<html lang="ja">
<head>
  <base target="_top">
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
  <title>LDC Chat</title>
  <style>
    :root {
      --primary: #1a73e8;
      --bg: #f1f3f4;
      --bubble-me: #1a73e8;
      --bubble-other: #ffffff;
    }
    * { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
    html, body {
      margin: 0; height: 100%;
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Helvetica Neue', 'Hiragino Kaku Gothic ProN', Meiryo, sans-serif;
      background: var(--bg);
    }
    body { display: flex; flex-direction: column; height: 100dvh; }

    /* ヘッダー */
    header {
      flex: 0 0 auto;
      background: var(--primary); color: #fff;
      display: flex; align-items: center; gap: 8px;
      padding: 12px 14px;
      box-shadow: 0 1px 3px rgba(0,0,0,.2);
    }
    header .back { font-size: 22px; cursor: pointer; display: none; }
    header .title { font-size: 17px; font-weight: 600; flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    header .user { font-size: 11px; opacity: .85; max-width: 40%; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    header .ocr-btn { flex: 0 0 auto; background: rgba(255,255,255,.2); color: #fff; border: none; border-radius: 16px; padding: 6px 12px; font-size: 13px; cursor: pointer; }
    header .ocr-btn:active { background: rgba(255,255,255,.35); }

    /* 画面切り替え */
    .screen { flex: 1 1 auto; display: none; flex-direction: column; min-height: 0; }
    .screen.active { display: flex; }

    /* ルーム一覧 */
    #rooms { overflow-y: auto; }
    .room {
      background: #fff; padding: 16px; border-bottom: 1px solid #e0e0e0;
      font-size: 16px; cursor: pointer; display: flex; align-items: center; gap: 12px;
    }
    .room:active { background: #eef3fd; }
    .room .avatar {
      width: 40px; height: 40px; border-radius: 50%; background: var(--primary);
      color: #fff; display: flex; align-items: center; justify-content: center; font-weight: 600;
    }
    .new-room {
      display: flex; gap: 8px; padding: 12px; background: #fff; border-bottom: 1px solid #e0e0e0;
    }
    .new-room input { flex: 1; padding: 10px; border: 1px solid #ccc; border-radius: 20px; font-size: 15px; }
    .new-room button { padding: 10px 16px; border: none; background: var(--primary); color: #fff; border-radius: 20px; font-size: 15px; }

    /* メッセージ一覧 */
    #messages { flex: 1 1 auto; overflow-y: auto; padding: 12px; display: flex; flex-direction: column; gap: 8px; }
    .msg { max-width: 78%; padding: 9px 13px; border-radius: 16px; font-size: 15px; line-height: 1.4; word-break: break-word; }
    .msg .name { font-size: 11px; color: #5f6368; margin-bottom: 2px; }
    .msg .time { font-size: 10px; opacity: .7; margin-top: 3px; text-align: right; }
    .msg.other { align-self: flex-start; background: var(--bubble-other); border: 1px solid #e0e0e0; border-bottom-left-radius: 4px; }
    .msg.me { align-self: flex-end; background: var(--bubble-me); color: #fff; border-bottom-right-radius: 4px; }
    .empty { text-align: center; color: #9aa0a6; margin: 24px 0; font-size: 14px; }

    /* 入力欄 */
    .composer { flex: 0 0 auto; display: flex; gap: 8px; padding: 8px; background: #fff; border-top: 1px solid #e0e0e0; }
    .composer textarea {
      flex: 1; resize: none; border: 1px solid #ccc; border-radius: 20px;
      padding: 10px 14px; font-size: 15px; max-height: 96px; font-family: inherit;
    }
    .composer button {
      flex: 0 0 auto; width: 44px; height: 44px; border: none; border-radius: 50%;
      background: var(--primary); color: #fff; font-size: 18px;
    }
    .composer button:disabled { opacity: .5; }
  </style>
</head>
<body>
  <header>
    <span class="back" id="back">‹</span>
    <span class="title" id="title">LDC Chat</span>
    <button class="ocr-btn" id="botBtn">規約</button>
    <button class="ocr-btn" id="ocrBtn">OCR</button>
    <span class="user" id="me"></span>
  </header>

  <!-- ルーム一覧画面 -->
  <section class="screen active" id="roomScreen">
    <div class="new-room">
      <input id="newRoomName" type="text" placeholder="新しいルーム名">
      <button id="addRoom">作成</button>
    </div>
    <div id="rooms"></div>
  </section>

  <!-- チャット画面 -->
  <section class="screen" id="chatScreen">
    <div id="messages"></div>
    <div class="composer">
      <textarea id="input" rows="1" placeholder="メッセージを入力"></textarea>
      <button id="send">➤</button>
    </div>
  </section>

  <script>
    var state = { rooms: [], roomId: null, roomName: '', me: 'ゲスト', timer: null, lastCount: -1 };

    function esc(s) {
      return String(s).replace(/[&<>"']/g, function (c) {
        return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
      });
    }
    function fmtTime(ms) {
      if (!ms) return '';
      var d = new Date(ms);
      var p = function (n) { return ('0' + n).slice(-2); };
      return p(d.getHours()) + ':' + p(d.getMinutes());
    }
    function $(id) { return document.getElementById(id); }

    // ---- ルーム一覧 ----
    function loadRooms() {
      google.script.run.withSuccessHandler(function (rooms) {
        state.rooms = rooms;
        renderRooms();
      }).getRooms();
    }
    function renderRooms() {
      var el = $('rooms');
      if (!state.rooms.length) {
        el.innerHTML = '<div class="empty">ルームがありません。作成してください。</div>';
        return;
      }
      el.innerHTML = state.rooms.map(function (r) {
        return '<div class="room" data-id="' + esc(r.roomId) + '" data-name="' + esc(r.name) + '">' +
          '<span class="avatar">' + esc(r.name.charAt(0)) + '</span>' +
          '<span>' + esc(r.name) + '</span></div>';
      }).join('');
      Array.prototype.forEach.call(el.querySelectorAll('.room'), function (node) {
        node.addEventListener('click', function () {
          openRoom(node.getAttribute('data-id'), node.getAttribute('data-name'));
        });
      });
    }
    $('addRoom').addEventListener('click', function () {
      var name = $('newRoomName').value;
      if (!name.trim()) return;
      google.script.run.withSuccessHandler(function (room) {
        $('newRoomName').value = '';
        loadRooms();
        openRoom(room.roomId, room.name);
      }).withFailureHandler(function (e) { alert(e.message); }).createRoom(name);
    });

    // ---- チャット ----
    function openRoom(roomId, name) {
      state.roomId = roomId;
      state.roomName = name;
      state.lastCount = -1;
      $('title').textContent = name;
      $('back').style.display = 'block';
      $('roomScreen').classList.remove('active');
      $('chatScreen').classList.add('active');
      loadMessages();
      startPolling();
    }
    function leaveRoom() {
      stopPolling();
      state.roomId = null;
      $('title').textContent = 'LDC Chat';
      $('back').style.display = 'none';
      $('chatScreen').classList.remove('active');
      $('roomScreen').classList.add('active');
      loadRooms();
    }
    $('back').addEventListener('click', leaveRoom);

    function loadMessages() {
      if (!state.roomId) return;
      google.script.run.withSuccessHandler(renderMessages).getMessages(state.roomId);
    }
    function renderMessages(msgs) {
      // 件数が変わらなければ再描画しない（無駄なちらつき防止）
      if (msgs.length === state.lastCount) return;
      state.lastCount = msgs.length;
      var el = $('messages');
      if (!msgs.length) {
        el.innerHTML = '<div class="empty">まだメッセージがありません</div>';
        return;
      }
      el.innerHTML = msgs.map(function (m) {
        var mine = (m.user === state.me);
        return '<div class="msg ' + (mine ? 'me' : 'other') + '">' +
          (mine ? '' : '<div class="name">' + esc(m.user) + '</div>') +
          '<div>' + esc(m.text) + '</div>' +
          '<div class="time">' + fmtTime(m.createdAt) + '</div></div>';
      }).join('');
      el.scrollTop = el.scrollHeight;
    }

    function send() {
      var text = $('input').value;
      if (!text.trim() || !state.roomId) return;
      $('send').disabled = true;
      google.script.run.withSuccessHandler(function () {
        $('input').value = '';
        $('input').style.height = 'auto';
        $('send').disabled = false;
        loadMessages();
      }).withFailureHandler(function (e) {
        $('send').disabled = false;
        alert(e.message);
      }).postMessage(state.roomId, text);
    }
    $('send').addEventListener('click', send);
    $('input').addEventListener('keydown', function (e) {
      if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); send(); }
    });
    $('input').addEventListener('input', function () {
      this.style.height = 'auto';
      this.style.height = Math.min(this.scrollHeight, 96) + 'px';
    });

    // ---- ポーリング（簡易リアルタイム） ----
    function startPolling() { stopPolling(); state.timer = setInterval(loadMessages, 3000); }
    function stopPolling() { if (state.timer) { clearInterval(state.timer); state.timer = null; } }

    // ---- OCR画面への遷移 ----
    // WebアプリURLを取得しておき、ボタン押下で ?app=ocr に遷移する
    var ocrUrl = '';
    google.script.run.withSuccessHandler(function (url) {
      ocrUrl = url;
    }).getWebAppUrl();
    $('ocrBtn').addEventListener('click', function () {
      if (!ocrUrl) return;
      window.top.location.href = ocrUrl + '?app=ocr';
    });
    // 就労規則ボットへの遷移
    $('botBtn').addEventListener('click', function () {
      if (!ocrUrl) return;
      window.top.location.href = ocrUrl + '?app=bot';
    });

    // ---- 初期化 ----
    google.script.run.withSuccessHandler(function (email) {
      state.me = email;
      $('me').textContent = email;
    }).getCurrentUser();
    loadRooms();
  </script>
</body>
</html>
```
