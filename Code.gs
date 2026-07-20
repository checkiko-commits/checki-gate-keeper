// =====================================================
// CHECKI GATE KEEPER — 案件データ共有バックエンド
// 対象：営業ジャッジ・ゲートアプリ（index.html）のデータ保存先
//
// 【初回セットアップ手順】
// 1. Google スプレッドシートを新規作成（例：「CHECKI Gate Keeper データ」）
// 2. 拡張機能 → Apps Script でこのコードを貼り付け（既存コードは削除）
// 3. 上部の API_TOKEN を、index.html 側の API_TOKEN と同じ値にする
//    （初期値のままでもよいが、他の用途と使い回さないこと）
// 4. スプレッドシートを開き直すとメニュー「🔒 Gate Keeper」が出るので
//    「📋 データシートを初期化する」を実行
// 5. 続けて「✉️ 通知テストを送信」を実行する
//    → 初回はGmail送信権限の許可を求められるので許可する
//    → CHAIRMAN_EMAIL 宛にテストメールが届けば設定完了
// 6. デプロイ → 新しいデプロイ → 種類の選択で「ウェブアプリ」を選択
//    - 説明: 任意
//    - 次のユーザーとして実行: 自分
//    - アクセスできるユーザー: 全員
// 7. 発行された「ウェブアプリ」の URL（.../exec で終わるもの）をコピーし、
//    index.html の API_URL にそのまま貼り付ける
// 8. 初回デプロイ後にこのコードを修正した場合は、
//    デプロイ → デプロイを管理 → 編集(鉛筆) → バージョン「新バージョン」→ デプロイ
//    をしないと変更が反映されないので注意
// =====================================================

// 他の用途と使い回さないこと。index.html 側の API_TOKEN と必ず一致させる。
const API_TOKEN = '6ad2c1ffcd6eea0125370b699b656146aa21e476';

const SHEET_NAME = '案件データ';
const HEADERS = ['id', 'name', 'createdAt', 'status', 'entryJson', 'draftDocOverride', 'submissionJson', 'judgeJson', 'updatedAt'];

const KNOWLEDGE_SHEET_NAME = 'ナレッジ';
const KNOWLEDGE_HEADERS = ['id', 'filename', 'content', 'uploadedAt'];

// 通知先メールアドレスと、通知メールに載せる公開URL
const CHAIRMAN_EMAIL = 'brands.masaki@gmail.com';
const NAKAGAWA_EMAIL = 'akiko.nakagawa@checki.jp';
const APP_URL = 'https://checkiko-commits.github.io/checki-gate-keeper/';

// ==============================
// メニュー
// ==============================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔒 Gate Keeper')
    .addItem('📋 データシートを初期化する', 'setupSheet')
    .addItem('✉️ 通知テストを送信', 'sendTestNotification')
    .addToUi();
}

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  sheet.clear();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  SpreadsheetApp.getUi().alert('「' + SHEET_NAME + '」シートを初期化しました。');
}

function sendTestNotification() {
  MailApp.sendEmail({
    to: CHAIRMAN_EMAIL,
    subject: '【CHECKI Gate Keeper】通知テスト',
    body: 'このメールが届いていれば、通知設定は正常に動作しています。\n\nアプリ: ' + APP_URL
  });
  SpreadsheetApp.getUi().alert(CHAIRMAN_EMAIL + ' 宛にテストメールを送信しました。');
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('「' + SHEET_NAME + '」シートがありません。メニューから初期化してください。');
  ensureCaseHeaders_(sheet);
  return sheet;
}

// 既存の案件データを消さずに、後から追加された列をヘッダー行にだけ追加する。
// setupSheet()の再実行（＝全消去）をせずに済ませるための移行処理。
function ensureCaseHeaders_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  HEADERS.forEach(function (h) {
    if (headerRow.indexOf(h) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h);
    }
  });
}

function getOrCreateKnowledgeSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(KNOWLEDGE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(KNOWLEDGE_SHEET_NAME);
    sheet.getRange(1, 1, 1, KNOWLEDGE_HEADERS.length).setValues([KNOWLEDGE_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ==============================
// 読み書き（案件データ）
// ==============================
function listCases_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = {};
  headerRow.forEach(function (h, i) { idx[h] = i; });
  const rows = sheet.getRange(2, 1, lastRow - 1, headerRow.length).getValues();
  return rows
    .filter(function (r) { return r[idx.id]; })
    .map(function (r) {
      return {
        id: r[idx.id],
        name: r[idx.name],
        createdAt: r[idx.createdAt],
        status: r[idx.status],
        entry: r[idx.entryJson] ? JSON.parse(r[idx.entryJson]) : {},
        draftDocOverride: r[idx.draftDocOverride] || '',
        submission: r[idx.submissionJson] ? JSON.parse(r[idx.submissionJson]) : null,
        judge: r[idx.judgeJson] ? JSON.parse(r[idx.judgeJson]) : null
      };
    });
}

// ==============================
// 読み書き（ナレッジ）
// ==============================
function listKnowledge_() {
  const sheet = getOrCreateKnowledgeSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const rows = sheet.getRange(2, 1, lastRow - 1, KNOWLEDGE_HEADERS.length).getValues();
  return rows
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return { id: r[0], filename: r[1], content: r[2], uploadedAt: r[3] };
    });
}

function addKnowledge_(filename, content) {
  const sheet = getOrCreateKnowledgeSheet_();
  const id = 'K-' + Date.now().toString(36).toUpperCase() + '-' + Math.floor(Math.random() * 900 + 100);
  sheet.appendRow([id, filename, content, new Date().toISOString()]);
  return id;
}

function deleteKnowledge_(id) {
  const sheet = getOrCreateKnowledgeSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) { sheet.deleteRow(i + 2); return; }
  }
}

function findRowById_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) return i + 2; // 1-indexed, +1 for header
  }
  return -1;
}

function getExistingStatus_(sheet, id) {
  const rowIndex = findRowById_(sheet, id);
  if (rowIndex === -1) return null;
  return sheet.getRange(rowIndex, 4, 1, 1).getValue();
}

function upsertCase_(sheet, c) {
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const valueMap = {
    id: c.id,
    name: c.name || '',
    createdAt: c.createdAt || new Date().toISOString(),
    status: c.status || 'draft',
    entryJson: JSON.stringify(c.entry || {}),
    draftDocOverride: c.draftDocOverride || '',
    submissionJson: c.submission ? JSON.stringify(c.submission) : '',
    judgeJson: c.judge ? JSON.stringify(c.judge) : '',
    updatedAt: new Date().toISOString()
  };
  const row = headerRow.map(function (h) { return valueMap.hasOwnProperty(h) ? valueMap[h] : ''; });
  const rowIndex = findRowById_(sheet, c.id);
  if (rowIndex === -1) {
    sheet.appendRow(row);
  } else {
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  }
}

// ==============================
// 通知（承認ステータスが実際に変化した時だけ送る）
// ==============================
function notifyOnTransition_(prevStatus, c) {
  try {
    if (c.status === 'pending' && prevStatus !== 'pending') {
      const score = (c.submission && c.submission.scoring) ? c.submission.scoring.total : '—';
      MailApp.sendEmail({
        to: CHAIRMAN_EMAIL,
        subject: '【CHECKI Gate Keeper】案件「' + c.name + '」の承認待ちです（AIスコア: ' + score + '点）',
        body: [
          '中川様より新規案件が提出されました。',
          '',
          '案件名: ' + c.name,
          'AI事前採点: ' + score + ' / 100',
          '提出日時: ' + (c.submission ? c.submission.submittedAt : ''),
          '',
          '確認・ご判断はこちらから:',
          APP_URL
        ].join('\n')
      });
    } else if ((c.status === 'approved' || c.status === 'rejected') && prevStatus !== c.status) {
      const isApproved = c.status === 'approved';
      const lines = [
        '服部会長により' + (isApproved ? '承認' : '却下') + 'されました。',
        '',
        '案件名: ' + c.name,
        '判定日時: ' + (c.judge ? c.judge.decidedAt : '')
      ];
      if (!isApproved && c.judge && c.judge.comment) lines.push('修正コメント: ' + c.judge.comment);
      lines.push('', 'アプリはこちら:', APP_URL);
      MailApp.sendEmail({
        to: NAKAGAWA_EMAIL,
        subject: '【CHECKI Gate Keeper】案件「' + c.name + '」が' + (isApproved ? '承認' : '却下') + 'されました',
        body: lines.join('\n')
      });
    }
  } catch (err) {
    Logger.log('notify error: ' + err);
  }
}

// ==============================
// HTTP エンドポイント
// ==============================
function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    if (params.token !== API_TOKEN) return jsonResponse_({ ok: false, error: 'unauthorized' });

    if (params.action === 'list') {
      const sheet = getSheet_();
      return jsonResponse_({ ok: true, cases: listCases_(sheet) });
    }
    if (params.action === 'listKnowledge') {
      return jsonResponse_({ ok: true, items: listKnowledge_() });
    }
    return jsonResponse_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse_({ ok: false, error: 'invalid request body' });
  }
  if (body.token !== API_TOKEN) return jsonResponse_({ ok: false, error: 'unauthorized' });

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (body.action === 'upsert') {
      const sheet = getSheet_();
      if (!body.case || !body.case.id) return jsonResponse_({ ok: false, error: 'case.id is required' });
      const prevStatus = getExistingStatus_(sheet, body.case.id);
      upsertCase_(sheet, body.case);
      notifyOnTransition_(prevStatus, body.case);
      return jsonResponse_({ ok: true, cases: listCases_(sheet) });
    }
    if (body.action === 'uploadKnowledge') {
      if (!body.filename || !body.content) return jsonResponse_({ ok: false, error: 'filename/content is required' });
      if (String(body.content).length > 200000) return jsonResponse_({ ok: false, error: 'ファイルが大きすぎます（20万文字が上限です）' });
      const id = addKnowledge_(body.filename, body.content);
      return jsonResponse_({ ok: true, id: id, items: listKnowledge_() });
    }
    if (body.action === 'deleteKnowledge') {
      if (!body.id) return jsonResponse_({ ok: false, error: 'id is required' });
      deleteKnowledge_(body.id);
      return jsonResponse_({ ok: true, items: listKnowledge_() });
    }
    return jsonResponse_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
