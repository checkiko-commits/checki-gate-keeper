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
// 5. デプロイ → 新しいデプロイ → 種類の選択で「ウェブアプリ」を選択
//    - 説明: 任意
//    - 次のユーザーとして実行: 自分
//    - アクセスできるユーザー: 全員
// 6. 発行された「ウェブアプリ」の URL（.../exec で終わるもの）をコピーし、
//    index.html の API_URL にそのまま貼り付ける
// 7. 初回デプロイ後にこのコードを修正した場合は、
//    デプロイ → デプロイを管理 → 編集(鉛筆) → バージョン「新バージョン」→ デプロイ
//    をしないと変更が反映されないので注意
// =====================================================

// 他の用途と使い回さないこと。index.html 側の API_TOKEN と必ず一致させる。
const API_TOKEN = '6ad2c1ffcd6eea0125370b699b656146aa21e476';

const SHEET_NAME = '案件データ';
const HEADERS = ['id', 'name', 'createdAt', 'status', 'entryJson', 'draftDocOverride', 'submissionJson', 'judgeJson', 'updatedAt'];

// ==============================
// メニュー
// ==============================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔒 Gate Keeper')
    .addItem('📋 データシートを初期化する', 'setupSheet')
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

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('「' + SHEET_NAME + '」シートがありません。メニューから初期化してください。');
  return sheet;
}

// ==============================
// 読み書き
// ==============================
function listCases_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const rows = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  return rows
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return {
        id: r[0],
        name: r[1],
        createdAt: r[2],
        status: r[3],
        entry: r[4] ? JSON.parse(r[4]) : {},
        draftDocOverride: r[5] || '',
        submission: r[6] ? JSON.parse(r[6]) : null,
        judge: r[7] ? JSON.parse(r[7]) : null
      };
    });
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

function upsertCase_(sheet, c) {
  const row = [
    c.id,
    c.name || '',
    c.createdAt || new Date().toISOString(),
    c.status || 'draft',
    JSON.stringify(c.entry || {}),
    c.draftDocOverride || '',
    c.submission ? JSON.stringify(c.submission) : '',
    c.judge ? JSON.stringify(c.judge) : '',
    new Date().toISOString()
  ];
  const rowIndex = findRowById_(sheet, c.id);
  if (rowIndex === -1) {
    sheet.appendRow(row);
  } else {
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
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
    return jsonResponse_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.token !== API_TOKEN) return jsonResponse_({ ok: false, error: 'unauthorized' });

    const sheet = getSheet_();
    if (body.action === 'upsert') {
      if (!body.case || !body.case.id) return jsonResponse_({ ok: false, error: 'case.id is required' });
      upsertCase_(sheet, body.case);
      return jsonResponse_({ ok: true, cases: listCases_(sheet) });
    }
    return jsonResponse_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
