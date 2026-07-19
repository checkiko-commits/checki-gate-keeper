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

function getExistingStatus_(sheet, id) {
  const rowIndex = findRowById_(sheet, id);
  if (rowIndex === -1) return null;
  return sheet.getRange(rowIndex, 4, 1, 1).getValue();
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
      const prevStatus = getExistingStatus_(sheet, body.case.id);
      upsertCase_(sheet, body.case);
      notifyOnTransition_(prevStatus, body.case);
      return jsonResponse_({ ok: true, cases: listCases_(sheet) });
    }
    return jsonResponse_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
