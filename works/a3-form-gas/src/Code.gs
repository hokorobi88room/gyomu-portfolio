/**
 * FormFlow v1.0 — Googleフォーム受付自動化(GAS)
 *
 * フォーム送信をトリガーに、以下を全自動で行う:
 *  1. 申込者へ確認メールの自動返信(テンプレート差し込み)
 *  2. 台帳シートへの整形転記(重複申込の検知つき)
 *  3. 管理者への通知(メール + 任意のWebhook)
 *  4. 定員到達時のフォーム自動クローズ
 *  5. 日次サマリの自動送信(時間トリガー)
 *
 * セットアップ手順は docs/SETUP.md を参照。
 * 設定はすべて「設定」シートで行う(コードの変更不要)。
 */

/* eslint-disable no-unused-vars */

const SHEET_CONFIG = '設定';
const SHEET_LEDGER = '台帳';
const SHEET_LOG = '実行ログ';

// ============================================================
// トリガー1: フォーム送信時(インストーラブルトリガーで設定する)
// ============================================================
function onFormSubmitHandler(e) {
  // 多重発火対策: 同時実行をロックで直列化する
  const lock = LockService.getScriptLock();
  lock.waitLock(30 * 1000);
  try {
    const cfg = loadConfig_();
    const entry = parseSubmission_(e);

    const duplicate = appendToLedger_(entry, cfg);
    sendConfirmationMail_(entry, cfg, duplicate);
    notifyAdmin_(entry, cfg, duplicate);

    const count = countEntries_();
    if (cfg.capacity > 0 && count >= cfg.capacity) {
      closeForm_(cfg, count);
    }
    log_('INFO', 'onFormSubmit', `受付 #${count} ${entry.email}${duplicate ? ' (重複)' : ''}`);
  } catch (err) {
    log_('ERROR', 'onFormSubmit', String(err && err.stack ? err.stack : err));
    // 失敗を管理者へ即時通知(申込者を待たせたまま黙らない)
    notifyError_(err);
    throw err; // GAS側のリトライ・失敗記録にも残す
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// トリガー2: 日次サマリ(時間主導型トリガー: 毎日 18-19時 推奨)
// ============================================================
function sendDailySummary() {
  const cfg = loadConfig_();
  const ledger = sheet_(SHEET_LEDGER);
  const values = ledger.getDataRange().getValues();
  if (values.length <= 1) return;

  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const todays = values.slice(1).filter((r) => r[0] instanceof Date && r[0] >= start);

  const total = values.length - 1;
  const capText = cfg.capacity > 0 ? ` / 定員 ${cfg.capacity}(残り ${Math.max(cfg.capacity - total, 0)})` : '';
  const lines = todays.map((r) => `・${formatTime_(r[0])} ${r[1]}(${r[2]})`).join('\n') || '(本日の新規申込はありません)';

  MailApp.sendEmail({
    to: cfg.adminEmail,
    subject: `【${cfg.eventName}】本日の申込 ${todays.length} 件 / 累計 ${total} 件${capText}`,
    body: `${lines}\n\n台帳: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`,
  });
  log_('INFO', 'dailySummary', `本日 ${todays.length} 件 / 累計 ${total} 件`);
}

// ============================================================
// 内部処理
// ============================================================

/** フォーム回答を {timestamp, name, email, answers{}} へ整形する */
function parseSubmission_(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const values = e && e.namedValues ? e.namedValues : {};
  const pick = (aliases) => {
    for (const key of Object.keys(values)) {
      const k = key.trim();
      if (aliases.some((a) => k === a || k.indexOf(a) === 0)) {
        return String(values[key][0] || '').trim();
      }
    }
    return '';
  };

  const entry = {
    timestamp: new Date(),
    name: pick(['お名前', '氏名', '名前']),
    email: pick(['メールアドレス', 'メール', 'Email', 'email']),
    answers: {},
  };
  Object.keys(values).forEach((key) => {
    entry.answers[key.trim()] = String(values[key].join(', ')).trim();
  });

  if (!entry.email) {
    throw new Error('メールアドレス項目が見つかりません。フォームの質問名に「メールアドレス」を含めてください。');
  }
  return entry;
}

/** 台帳へ転記し、同一メールの申込が既にあれば true を返す */
function appendToLedger_(entry, cfg) {
  const ledger = sheet_(SHEET_LEDGER);
  if (ledger.getLastRow() === 0) {
    ledger.appendRow(['受付日時', 'お名前', 'メールアドレス', '重複', '回答(全項目)']);
    ledger.getRange(1, 1, 1, 5).setFontWeight('bold');
  }

  const emails = ledger.getLastRow() > 1
    ? ledger.getRange(2, 3, ledger.getLastRow() - 1, 1).getValues().flat().map(String)
    : [];
  const duplicate = emails.some((m) => m.toLowerCase() === entry.email.toLowerCase());

  ledger.appendRow([
    entry.timestamp,
    entry.name,
    entry.email,
    duplicate ? '重複' : '',
    JSON.stringify(entry.answers),
  ]);
  if (duplicate) {
    ledger.getRange(ledger.getLastRow(), 1, 1, 5).setBackground('#fff3cd');
  }
  return duplicate;
}

/** 申込者への確認メール(設定シートのテンプレートに {{名前}} 等を差し込み) */
function sendConfirmationMail_(entry, cfg, duplicate) {
  const body = cfg.mailTemplate
    .replace(/{{名前}}/g, entry.name || 'お客様')
    .replace(/{{イベント名}}/g, cfg.eventName)
    .replace(/{{受付日時}}/g, formatDateTime_(entry.timestamp));

  MailApp.sendEmail({
    to: entry.email,
    subject: `【${cfg.eventName}】お申込みを受け付けました`,
    body: body + (duplicate ? '\n\n※ 同じメールアドレスでのお申込みが既にあります。変更のご連絡でしたらこのままで結構です。' : ''),
    noReply: true,
  });
}

function notifyAdmin_(entry, cfg, duplicate) {
  if (cfg.adminEmail) {
    MailApp.sendEmail({
      to: cfg.adminEmail,
      subject: `【${cfg.eventName}】新規申込: ${entry.name}${duplicate ? '(重複)' : ''}`,
      body: Object.entries(entry.answers).map(([k, v]) => `${k}: ${v}`).join('\n'),
    });
  }
  if (cfg.webhookUrl) {
    // Slack/Discord/LINE WORKS 等の Incoming Webhook(任意)
    UrlFetchApp.fetch(cfg.webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: `📩 新規申込: ${entry.name}(${entry.email})${duplicate ? ' ※重複' : ''}` }),
      muteHttpExceptions: true, // 通知失敗で受付処理全体を落とさない
    });
  }
}

/** 定員到達: フォームの受付を停止し、管理者へ知らせる */
function closeForm_(cfg, count) {
  const formUrl = SpreadsheetApp.getActiveSpreadsheet().getFormUrl();
  if (!formUrl) {
    log_('WARN', 'closeForm', 'フォームURLが取得できず自動クローズをスキップ');
    return;
  }
  const form = FormApp.openByUrl(formUrl);
  if (!form.isAcceptingResponses()) return;

  form.setAcceptingResponses(false);
  form.setCustomClosedFormMessage(`定員(${cfg.capacity}名)に達したため受付を終了しました。キャンセル待ちをご希望の方はお問い合わせください。`);
  MailApp.sendEmail({
    to: cfg.adminEmail,
    subject: `【${cfg.eventName}】定員到達(${count}件)— 受付を自動終了しました`,
    body: `定員 ${cfg.capacity} 件に達したため、フォームの受付を自動で締め切りました。`,
  });
  log_('INFO', 'closeForm', `定員 ${cfg.capacity} 到達で自動クローズ`);
}

// ============================================================
// 設定・ユーティリティ
// ============================================================

function loadConfig_() {
  const ws = sheet_(SHEET_CONFIG);
  const map = {};
  ws.getDataRange().getValues().forEach(([k, v]) => {
    if (k) map[String(k).trim()] = String(v == null ? '' : v).trim();
  });

  const required = ['イベント名', '管理者メール', '確認メール本文'];
  required.forEach((k) => {
    if (!map[k]) throw new Error(`設定シートに「${k}」がありません(A列にキー・B列に値)`);
  });

  return {
    eventName: map['イベント名'],
    adminEmail: map['管理者メール'],
    mailTemplate: map['確認メール本文'],
    capacity: Number(map['定員'] || 0),
    webhookUrl: map['Webhook URL'] || '',
  };
}

function countEntries_() {
  const ledger = sheet_(SHEET_LEDGER);
  return Math.max(ledger.getLastRow() - 1, 0);
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(name);
  if (!ws) ws = ss.insertSheet(name);
  return ws;
}

function log_(level, fn, message) {
  const ws = sheet_(SHEET_LOG);
  if (ws.getLastRow() === 0) ws.appendRow(['日時', 'レベル', '処理', '内容']);
  ws.appendRow([new Date(), level, fn, message]);
}

function notifyError_(err) {
  try {
    const cfg = loadConfig_();
    MailApp.sendEmail({
      to: cfg.adminEmail,
      subject: '【FormFlow】受付処理でエラーが発生しました',
      body: `内容:\n${err}\n\n実行ログシートを確認してください。`,
    });
  } catch (_) {
    // 設定自体が壊れている場合はログのみ(二次エラーを起こさない)
  }
}

function formatDateTime_(d) {
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
}

function formatTime_(d) {
  return Utilities.formatDate(d, 'Asia/Tokyo', 'HH:mm');
}

// ============================================================
// 動作テスト(フォームなしで転記・メール・定員判定を確認できる)
// ============================================================
function testWithDummySubmission() {
  const dummy = {
    namedValues: {
      'タイムスタンプ': [new Date().toString()],
      'お名前': ['試験 太郎'],
      'メールアドレス': [Session.getActiveUser().getEmail()], // 自分宛てに送る
      '参加日': ['2026/07/20'],
      '備考': ['動作テストです'],
    },
  };
  onFormSubmitHandler(dummy);
}
