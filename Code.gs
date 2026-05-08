/**
 * みかみグルコン 質問回答シート - サーバーサイド
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
const ANSWERED_AT_COL_NAME = '回答済み日時'; // 元データの「回答日時」と区別するため接頭辞を変更

// ヘッダー名の候補（スプレッドシート側がどんな表記でもマッチするように）
// ※ 「回答者ID」のような ID 系列にマッチしないよう、ネガティブも考慮した順序で書く。
const HEADER_PATTERNS = {
  timestamp: [/タイムスタンプ/, /timestamp/i, /回答日時/, /^回答日$/, /送信日時/, /日時/],
  studentName: [/回答者名/, /受講生(名|さん)?$/, /^受講生/, /ニックネーム/, /お名前.*ニック/, /表示名/],
  realName: [/本名/, /氏名/, /フルネーム/, /^名前$/, /お名前$/],
  question: [/質問.*相談/, /相談.*質問/, /質問内容/, /質問.*内容/, /(三上|みかみ).*質問/, /^質問/, /相談/],
  addnessConsent: [/[Aa]ddness/, /アドネス/]
};

// 明示的に除外したい列名（ID 系など）
const EXCLUDE_PATTERNS = [/ID$/i, /\bID\b/i];

// ===== Web App エントリーポイント =====
function doGet(e) {
  // パラメータが api=list なら JSON、それ以外は HTML を返す
  const action = e && e.parameter && e.parameter.action;
  if (action === 'list') {
    return jsonResponse({ ok: true, data: getQuestions() });
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('みかみグルコン 質問回答シート')
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
    const ts = parseDate_(row[colMap.timestamp]);
    if (!ts) continue;
    if (ts < FILTER_FROM) continue;

    const answered = !!row[colMap.answered];
    const answeredAt = parseDate_(row[colMap.answeredAt]);

    result.push({
      rowId: r + 1, // スプレッドシートの行番号（1始まり）
      timestamp: ts.toISOString(),
      timestampLabel: Utilities.formatDate(ts, tz, 'yyyy/MM/dd HH:mm'),
      studentName: safeStr_(colMap.studentName >= 0 ? row[colMap.studentName] : ''),
      realName: safeStr_(colMap.realName >= 0 ? row[colMap.realName] : ''),
      question: safeStr_(colMap.question >= 0 ? row[colMap.question] : ''),
      addnessConsent: normalizeConsent_(colMap.addnessConsent >= 0 ? row[colMap.addnessConsent] : ''),
      answered: answered,
      answeredAtLabel: answeredAt
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

    // ID 系の列はスキップ（「回答者ID」が「回答者名」より先に来てもマッチしないように）
    const isExcluded = matchAny_(h, EXCLUDE_PATTERNS);

    if (map.timestamp === -1 && matchAny_(h, HEADER_PATTERNS.timestamp)) { map.timestamp = i; continue; }
    if (!isExcluded && map.studentName === -1 && matchAny_(h, HEADER_PATTERNS.studentName)) { map.studentName = i; continue; }
    if (!isExcluded && map.realName === -1 && matchAny_(h, HEADER_PATTERNS.realName)) { map.realName = i; continue; }
    if (!isExcluded && map.question === -1 && matchAny_(h, HEADER_PATTERNS.question)) { map.question = i; continue; }
    if (!isExcluded && map.addnessConsent === -1 && matchAny_(h, HEADER_PATTERNS.addnessConsent)) { map.addnessConsent = i; continue; }
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

// 値を Date に変換。Date インスタンス、ISO文字列、yyyy/MM/dd 等に対応
function parseDate_(v) {
  if (!v && v !== 0) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (typeof v === 'number') {
    // Sheets のシリアル値（1900-01-01 起点）には今回は対応せず素直に new Date を試す
    const d = new Date(v);
    return isNaN(d.getTime()) ? null : d;
  }
  const s = String(v).trim();
  if (!s) return null;
  // 「2024-10-02」「2024/10/02 12:34」「2024-10-02T12:34:00」など幅広く対応
  const normalized = s.replace(/-/g, '/');
  const d = new Date(normalized);
  if (!isNaN(d.getTime())) return d;
  const d2 = new Date(s);
  return isNaN(d2.getTime()) ? null : d2;
}

// Addness 登録状態を 'yes'(登録済み) / 'no'(未登録) / '' に正規化
function normalizeConsent_(v) {
  const s = String(v || '').trim().toLowerCase();
  if (!s) return '';
  // 否定的表現を先に判定（「未登録」「登録なし」「登録していない」など）
  if (/(未登録|登録なし|登録して(い|お)?ない|登録してい?ません|いいえ|^no\b|ng|不可|拒否|小文字|使用していない|使ってない|なし$|未回答|未使用)/i.test(s)) return 'no';
  // 肯定的表現
  if (/(登録済|登録してい?ます|登録してい?る|登録あり|あり|^yes\b|はい|ok|可|許可|大文字|^a$|addness|使用してい?る|使ってい?る|使用済)/i.test(s)) return 'yes';
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

// 列の自動マッピング結果と、各列のヘッダーを表示する診断関数
function debug_inspectMapping() {
  const sheet = getSheet_();
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colMap = mapColumns_(header);
  const labelOf = function(idx) {
    if (idx == null || idx < 0) return '(検出されず)';
    return String.fromCharCode(65 + idx) + '列「' + header[idx] + '」';
  };
  Logger.log('--- ヘッダー ---');
  for (let i = 0; i < header.length; i++) {
    Logger.log('  ' + String.fromCharCode(65 + i) + ': ' + header[i]);
  }
  Logger.log('--- 自動マッピング ---');
  Logger.log('  timestamp     -> ' + labelOf(colMap.timestamp));
  Logger.log('  studentName   -> ' + labelOf(colMap.studentName));
  Logger.log('  realName      -> ' + labelOf(colMap.realName));
  Logger.log('  question      -> ' + labelOf(colMap.question));
  Logger.log('  addnessConsent-> ' + labelOf(colMap.addnessConsent));
  Logger.log('  answered      -> ' + labelOf(colMap.answered));
  Logger.log('  answeredAt    -> ' + labelOf(colMap.answeredAt));
  const items = getQuestions();
  Logger.log('--- 抽出件数: ' + items.length + ' 件（FILTER_FROM=' + FILTER_FROM.toISOString() + ' 以降） ---');
  for (let i = 0; i < Math.min(3, items.length); i++) {
    Logger.log('  ' + JSON.stringify(items[i]));
  }
}
