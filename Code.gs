/**
 * 三上グルコン 質問回答ボード - サーバーサイド
 *
 * - スプレッドシートをデータソースに、質問一覧を返す API
 * - 「回答済み」状態を書き戻す API
 * - HTML サービスでフロントエンドを配信
 */

// ===== 設定 =====
const SPREADSHEET_ID = '1WYvHNapbrnOU5_oe4FStzkAX6GQa1CbqEMaDVopeM_I';
const SHEET_NAME = ''; // 空文字 = 1枚目のシートを自動採用
const FILTER_FROM = new Date('2026-04-01T00:00:00+09:00'); // 2026/4/1 以降の質問のみ表示
const ANSWERED_COL_NAME = '回答済み';
const ANSWERED_AT_COL_NAME = '回答日時';

// ヘッダー名の候補（スプレッドシート側がどんな表記でもマッチするように）
const HEADER_PATTERNS = {
  timestamp: [/タイムスタンプ/, /timestamp/i, /回答日$/, /送信日時/, /日時/],
  studentName: [/受講生/, /回答者/, /ニックネーム/, /お名前\(.*ニック/, /表示名/],
  realName: [/本名/, /氏名/, /フルネーム/, /^名前$/, /お名前$/],
  question: [/質問内容/, /質問.*内容/, /三上.*質問/, /^質問/, /相談/],
  addnessConsent: [/[Aa]ddness/, /アドネス/]
};

// ===== Web App エントリーポイント =====
function doGet(e) {
  // パラメータが api=list なら JSON、それ以外は HTML を返す
  const action = e && e.parameter && e.parameter.action;
  if (action === 'list') {
    return jsonResponse({ ok: true, data: getQuestions() });
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('三上グルコン 質問回答ボード')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action;
    if (action === 'mark_answered') {
      const result = setAnswered(body.rowId, body.answered);
      return jsonResponse({ ok: true, data: result });
    }
    return jsonResponse({ ok: false, error: 'unknown action' });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

// HTML から CSS/JS をインクルードするヘルパ
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ===== データ取得 =====
function getQuestions() {
  const sheet = getSheet_();
  ensureExtraColumns_(sheet); // 「回答済み」「回答日時」列を必要に応じて追加

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const header = values[0];
  const colMap = mapColumns_(header);

  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const result = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const ts = row[colMap.timestamp];
    if (!(ts instanceof Date)) continue;
    if (ts < FILTER_FROM) continue;

    const answered = !!row[colMap.answered];
    const answeredAt = row[colMap.answeredAt];

    result.push({
      rowId: r + 1, // スプレッドシートの行番号（1始まり）
      timestamp: ts.toISOString(),
      timestampLabel: Utilities.formatDate(ts, tz, 'yyyy/MM/dd HH:mm'),
      studentName: safeStr_(row[colMap.studentName]),
      realName: safeStr_(row[colMap.realName]),
      question: safeStr_(row[colMap.question]),
      addnessConsent: normalizeConsent_(row[colMap.addnessConsent]),
      answered: answered,
      answeredAtLabel: answeredAt instanceof Date
        ? Utilities.formatDate(answeredAt, tz, 'yyyy/MM/dd HH:mm')
        : ''
    });
  }

  // 古いものが上、新しいものが下（最新が下に積まれる）
  result.sort(function(a, b) {
    return new Date(a.timestamp) - new Date(b.timestamp);
  });

  return result;
}

// ===== 回答済みフラグ更新 =====
function setAnswered(rowId, answered) {
  const sheet = getSheet_();
  ensureExtraColumns_(sheet);

  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colMap = mapColumns_(header);

  const flag = !!answered;
  sheet.getRange(rowId, colMap.answered + 1).setValue(flag);
  sheet.getRange(rowId, colMap.answeredAt + 1).setValue(flag ? new Date() : '');

  return { rowId: rowId, answered: flag };
}

// ===== 内部ユーティリティ =====
function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (SHEET_NAME) {
    const s = ss.getSheetByName(SHEET_NAME);
    if (!s) throw new Error('シートが見つかりません: ' + SHEET_NAME);
    return s;
  }
  return ss.getSheets()[0];
}

function ensureExtraColumns_(sheet) {
  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  let changed = false;
  if (header.indexOf(ANSWERED_COL_NAME) === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue(ANSWERED_COL_NAME);
    changed = true;
  }
  // ヘッダー再取得
  const header2 = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (header2.indexOf(ANSWERED_AT_COL_NAME) === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue(ANSWERED_AT_COL_NAME);
    changed = true;
  }
  return changed;
}

function mapColumns_(header) {
  const map = {
    timestamp: -1,
    studentName: -1,
    realName: -1,
    question: -1,
    addnessConsent: -1,
    answered: -1,
    answeredAt: -1
  };

  for (let i = 0; i < header.length; i++) {
    const h = String(header[i] || '').trim();
    if (!h) continue;

    if (h === ANSWERED_COL_NAME) { map.answered = i; continue; }
    if (h === ANSWERED_AT_COL_NAME) { map.answeredAt = i; continue; }

    if (map.timestamp === -1 && matchAny_(h, HEADER_PATTERNS.timestamp)) { map.timestamp = i; continue; }
    if (map.studentName === -1 && matchAny_(h, HEADER_PATTERNS.studentName)) { map.studentName = i; continue; }
    if (map.realName === -1 && matchAny_(h, HEADER_PATTERNS.realName)) { map.realName = i; continue; }
    if (map.question === -1 && matchAny_(h, HEADER_PATTERNS.question)) { map.question = i; continue; }
    if (map.addnessConsent === -1 && matchAny_(h, HEADER_PATTERNS.addnessConsent)) { map.addnessConsent = i; continue; }
  }

  // タイムスタンプが見つからなければ A列をフォールバック
  if (map.timestamp === -1) map.timestamp = 0;
  return map;
}

function matchAny_(text, patterns) {
  for (let i = 0; i < patterns.length; i++) {
    if (patterns[i].test(text)) return true;
  }
  return false;
}

function safeStr_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) {
    const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
    return Utilities.formatDate(v, tz, 'yyyy/MM/dd HH:mm');
  }
  return String(v).trim();
}

// アドネス使用可否を 'yes' / 'no' / '' に正規化
function normalizeConsent_(v) {
  const s = String(v || '').trim().toLowerCase();
  if (!s) return '';
  if (/(はい|yes|ok|可|許可|大文字|^a$|addness)/i.test(s)) return 'yes';
  if (/(いいえ|no|ng|不可|拒否|小文字)/i.test(s)) return 'no';
  return s;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== デバッグ用 =====
function debug_listFirstRows() {
  const sheet = getSheet_();
  const values = sheet.getRange(1, 1, Math.min(5, sheet.getLastRow()), sheet.getLastColumn()).getValues();
  Logger.log(values);
}

function debug_getQuestions() {
  Logger.log(JSON.stringify(getQuestions(), null, 2));
}
