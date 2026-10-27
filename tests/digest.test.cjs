const { test } = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');
const code = fs.readFileSync(path.join(__dirname, '../Code.gs'), 'utf8');

class Sheet {
  constructor(name) { this.name = name; this.rows = []; this.max = 1000; this.reads = []; this.formats = []; this.validation = null; }
  getName() { return this.name; }
  getLastRow() { return this.rows.length; }
  getLastColumn() { return Math.max(0, ...this.rows.map(r => r.length)); }
  getMaxRows() { return this.max; }
  insertRowsAfter(_, count) { this.max += count; }
  setFrozenRows() {}
  getRange(row, col, height = 1, width = 1) {
    const sh = this;
    return {
      getValues() {
        sh.reads.push({row, col, height, width});
        return Array.from({length: height}, (_, i) => Array.from({length: width}, (_, j) => sh.rows[row+i-1]?.[col+j-1] ?? ''));
      },
      setValues(values) {
        if (sh.failSent && values[0][0] === 'SENT') throw Error('sheet write failed');
        values.forEach((r, i) => {
          if (!sh.rows[row+i-1]) sh.rows[row+i-1] = [];
          r.forEach((v, j) => { sh.rows[row+i-1][col+j-1] = typeof v === 'string' && v.startsWith("'") ? v.slice(1) : v; });
        });
      },
      setNumberFormats(f) { sh.formats.push({row, col, f}); },
      setDataValidation(rule) { sh.validation = {row, col, height, rule}; },
      setValue(v) { this.setValues([[v]]); }
    };
  }
}
function element(name, value = '', children = [], attrs = {}) {
  return {getName: () => name, getValue: () => value, getChildren: () => children,
    getAttribute: key => key in attrs ? {getValue: () => attrs[key]} : null};
}
function rssItem(url, opts = {}) {
  return element('item', '', [
    element('title', opts.title || '=IMPORTXML("bad")'), element('link', url), element('description', '<b>テスト記事</b>'),
    element('pubDate', opts.date || 'Mon, 26 Oct 2026 10:00:00 GMT')
  ]);
}
function rssRoot(items = [rssItem('https://example.com/a')]) {
  return element('rss', '', [element('channel', '', items)]);
}
function geminiOk(text) {
  return {getResponseCode: () => 200, getContentText: () => JSON.stringify({candidates: [{finishReason: 'STOP', content: {parts: [{text}]}}]})};
}
function harness(options = {}) {
  const sheets = {}, sent = [], fetches = [], sleeps = [];
  const clock = {t: 1_000_000_000_000};
  let uuid = 0, locked = false;
  const book = {getSheetByName: n => sheets[n], insertSheet: n => (sheets[n] = new Sheet(n))};
  const props = {SPREADSHEET_ID: 'sheet_id', GEMINI_API_KEY: 'fake-key-for-test', GEMINI_MODEL: 'test-model', ...(options.props || {})};
  const validationBuilder = {requireValueInList(list) { this.list = list; return this; }, setAllowInvalid() { return this; },
    setHelpText() { return this; }, build() { return {list: this.list}; }};
  const ctx = vm.createContext({
    PropertiesService: {getScriptProperties: () => ({getProperty: k => props[k] ?? null})},
    SpreadsheetApp: {openById: () => book, getActiveSpreadsheet: () => options.noActive ? null : book, flush() {},
      newDataValidation: () => Object.create(validationBuilder)},
    LockService: {getScriptLock: () => ({tryLock() { if (locked) return false; locked = true; return true; }, releaseLock() { locked = false; }})},
    Utilities: {getUuid: () => 'test-id-' + (++uuid), formatDate: () => '2026-10-27', sleep(ms) { sleeps.push(ms); clock.t += ms; },
      newBlob: bytes => ({getDataAsString: cs => new TextDecoder(cs).decode(Uint8Array.from(bytes, b => b & 0xff))})},
    XmlService: {parse: () => ({getRootElement: () => options.root || rssRoot()})},
    UrlFetchApp: {fetch(url, init) {
      fetches.push({url, init});
      if (url.includes('generativelanguage')) {
        clock.t += options.geminiMs || 0;
        if (options.apiError) throw Error('secret=fake-key-for-test');
        if (options.gemini) return options.gemini(fetches.filter(f => f.url.includes('generativelanguage')).length);
        return {getResponseCode: () => 200, getContentText: () => JSON.stringify({candidates: [{finishReason: options.finishReason || 'STOP', content: {parts: [{text: options.summary || '抜粋に記載された内容の要約です。'}]}}]})};
      }
      return {getResponseCode: () => 200, getContent: () => [1], getContentText: () => '<rss/>', getHeaders: () => ({})};
    }},
    MailApp: {getRemainingDailyQuota: () => options.quota ?? 10, sendEmail(mail) {
      sent.push(mail);
      if (options.mailError && !/停止通知/.test(mail.subject)) throw Error('ambiguous');
      if (options.noticeMailError && /停止通知|記録警告/.test(mail.subject)) throw Error('ambiguous notice');
    }}
  });
  vm.runInContext(code, ctx);
  vm.runInContext('Date', ctx).now = () => clock.t;
  ctx.setupSheets();
  sheets['配信設定'].rows[1][1] = 'me@example.com';
  sheets['RSS設定'].rows.push([true, 'テストRSS', 'https://example.com/feed', 5]);
  return {ctx, sheets, sent, fetches, props, sleeps, clock};
}
const mailRows = h => h.sheets['メール履歴'];
const states = h => h.ctx.mailStates_(mailRows(h));
const digestMails = h => h.sent.filter(m => !/停止通知/.test(m.subject));
const notices = h => h.sent.filter(m => /停止通知/.test(m.subject));

// ---- carried over from 1.0.0 ----
test('initialization preserves existing settings and rows', () => {
  const h = harness(); h.ctx.setupSheets();
  assert.equal(h.sheets['配信設定'].rows[1][1], 'me@example.com');
  assert.equal(h.sheets['RSS設定'].rows.length, 2);
});
test('preview calls Gemini but sends no mail and does not mark as sent', () => {
  const h = harness(); const result = h.ctx.previewDigest();
  assert.equal(result.status, 'PREVIEW'); assert.equal(h.sent.length, 0);
  assert.equal(mailRows(h).rows[1][8], 'PREVIEW');
  assert.equal(h.ctx.sentArticleKeys_(states(h)).length, 0);
});
test('successful delivery archives exact bodies and skips already sent article', () => {
  const h = harness(); assert.equal(h.ctx.sendDailyDigest().status, 'SENT');
  const row = mailRows(h).rows[1];
  assert.equal(row[5], h.sent[0].body); assert.equal(row[6], h.sent[0].htmlBody);
  assert.equal(row[8], 'SENT'); assert.ok(row[2] instanceof vm.runInContext('Date', h.ctx));
  assert.equal(h.sheets['記事履歴'].rows.length, 2);
  assert.equal(h.ctx.sendDailyDigest().status, 'NO_NEW_ITEMS'); assert.equal(h.sent.length, 1);
});
test('ambiguous sending failure records UNKNOWN and blocks automatic retry', () => {
  const h = harness({mailError: true});
  assert.throws(() => h.ctx.sendDailyDigest(), /送信結果が不明/);
  assert.equal(mailRows(h).rows[1][8], 'UNKNOWN');
  assert.throws(() => h.ctx.sendDailyDigest(), /未確定/); assert.equal(h.sent.length, 1);
});
test('failure to record SENT preserves SENDING and blocks duplicate delivery', () => {
  const h = harness(); mailRows(h).failSent = true;
  assert.throws(() => h.ctx.sendDailyDigest());
  assert.equal(mailRows(h).rows[1][8], 'SENDING');
  assert.throws(() => h.ctx.sendDailyDigest(), /未確定/); assert.equal(h.sent.length, 1);
});
test('no quota sends nothing and leaves no pending message', () => {
  const h = harness({quota: 0}); assert.throws(() => h.ctx.sendDailyDigest(), /送信枠/);
  assert.equal(h.sent.length, 0); assert.equal(mailRows(h).rows.length, 1);
});
test('API failure hides raw exception and does not send or process articles', () => {
  const h = harness({apiError: true}); assert.throws(() => h.ctx.sendDailyDigest(), /要約がすべて失敗/);
  assert.equal(h.sent.length, 0); assert.equal(h.sheets['記事履歴'].rows.length, 1);
  assert.ok(!JSON.stringify(h.sheets['実行履歴'].rows).includes('fake-key-for-test'));
});
test('unfinished Gemini response is not sent', () => {
  const h = harness({finishReason: 'MAX_TOKENS'}); assert.throws(() => h.ctx.sendDailyDigest(), /要約がすべて失敗/);
  assert.equal(h.sent.length, 0);
});
test('Atom namespace-independent entries and alternate link are supported', () => {
  const root = element('feed', '', [element('entry', '', [element('title', 'Atom'),
    element('link', '', [], {rel: 'self', href: 'https://example.com/api'}),
    element('link', '', [], {rel: 'alternate', href: 'https://example.com/article'}),
    element('summary', '本文'), element('updated', '2026-10-26T00:00:00Z')])]);
  const h = harness({root}); const items = h.ctx.parseDigestFeed_('', 'Atom');
  assert.equal(items[0].url, 'https://example.com/article'); assert.equal(items[0].content, '本文');
});
test('rejects executable URL schemes and escapes HTML and sheet formulas', () => {
  const h = harness();
  assert.equal(h.ctx.validHttpUrl_('javascript:alert(1)'), '');
  assert.equal(h.ctx.validHttpUrl_('https://user:pass@example.com/a'), '');
  assert.equal(h.ctx.safeDigestCell_(' =IMPORTXML("x")'), "' =IMPORTXML(\"x\")");
  assert.equal(h.ctx.escapeDigestHtml_('<img onerror="x">'), '&lt;img onerror=&quot;x&quot;&gt;');
  h.ctx.sendDailyDigest();
  assert.equal(h.sheets['記事履歴'].rows[1][3], '=IMPORTXML("bad")');
});
test('deduplicates feed overlap, fragment variations and ignores PREVIEW', () => {
  const h = harness(); const items = h.ctx.parseDigestFeed_('', 'RSS');
  const selected = h.ctx.selectDigestItems_([...items, ...items], [], 5);
  assert.equal(selected.length, 1);
  assert.equal(h.ctx.selectDigestItems_(items, [items[0].key], 5).length, 0);
});
test('all persistent mail fields respect cell bound before sending', () => {
  const h = harness(); assert.throws(() => h.ctx.safeDigestCell_('x'.repeat(45001)), /保存上限/);
});
test('invalid model or recipient stops before network access', () => {
  const h = harness(); h.props.GEMINI_MODEL = '../invalid';
  assert.throws(() => h.ctx.sendDailyDigest(), /GEMINI_MODEL/); assert.equal(h.fetches.length, 0);
  const h2 = harness(); h2.sheets['配信設定'].rows[1][1] = 'a@example.com,b@example.com';
  assert.throws(() => h2.ctx.sendDailyDigest(), /1件/); assert.equal(h2.fetches.length, 0);
});
test('unknown or corrupt mail state fails closed', () => {
  const h = harness(); h.ctx.sendDailyDigest();
  mailRows(h).rows[1][8] = 'SNET';
  assert.throws(() => h.ctx.sendDailyDigest(), /状態が不正/); assert.equal(h.sent.length, 1);
  mailRows(h).rows[1][8] = 'SENT'; mailRows(h).rows[1][10] = 'broken';
  assert.throws(() => h.ctx.sendDailyDigest(), /壊れています/); assert.equal(h.sent.length, 1);
});
test('CANCELLED allows confirmed unsent articles to be retried', () => {
  const h = harness({mailError: true}); assert.throws(() => h.ctx.sendDailyDigest());
  mailRows(h).rows[1][8] = 'CANCELLED';
  assert.throws(() => h.ctx.sendDailyDigest(), /送信結果が不明/); assert.equal(h.sent.length, 2);
});
test('budget guard stops before new network operations', () => {
  const h = harness(); assert.throws(() => h.ctx.checkBudget_(0), /安全上限/);
  assert.equal(h.fetches.length, 0);
});

// ---- 1.1.0: operational fixes ----
test('pre-send failure is auto-CANCELLED, never sends, and does not block the next run', () => {
  const h = harness();
  const arch = h.sheets['記事履歴'], original = arch.getRange.bind(arch);
  arch.getRange = (row, ...args) => { if (row >= 2) throw Error('disk failure'); return original(row, ...args); };
  assert.throws(() => h.ctx.sendDailyDigest(), /自動で取消済み/);
  assert.equal(h.sent.length, 0); assert.equal(mailRows(h).rows[1][8], 'CANCELLED');
  arch.getRange = original;
  assert.equal(h.ctx.sendDailyDigest().status, 'SENT'); assert.equal(h.sent.length, 1);
});
test('if CANCELLED cannot be recorded either, PREPARED remains and blocks', () => {
  const h = harness();
  const arch = h.sheets['記事履歴'], aOrig = arch.getRange.bind(arch);
  arch.getRange = (row, ...args) => { if (row >= 2) throw Error('disk failure'); return aOrig(row, ...args); };
  const mail = mailRows(h), mOrig = mail.getRange.bind(mail);
  mail.getRange = (row, col, ...rest) => { if (col === 9) throw Error('disk failure'); return mOrig(row, col, ...rest); };
  assert.throws(() => h.ctx.sendDailyDigest(), /PREPAREDのまま/);
  mail.getRange = mOrig;
  assert.equal(mail.rows[1][8], 'PREPARED'); assert.equal(h.sent.length, 0);
  assert.throws(() => h.ctx.sendDailyDigest(), /未確定/);
});
test('time budget during summarization sends what is already summarized', () => {
  const items = ['a', 'b', 'c', 'd', 'e'].map((s, i) => rssItem('https://example.com/' + s, {title: s, date: `Mon, 26 Oct 2026 1${i}:00:00 GMT`}));
  const h = harness({root: rssRoot(items), geminiMs: 75000});
  const result = h.ctx.sendDailyDigest();
  assert.equal(result.status, 'SENT'); assert.equal(result.articles, 3);
  assert.ok(result.warnings.some(w => /残り2件は次回候補/.test(w)));
  assert.equal(JSON.parse(mailRows(h).rows[1][10]).length, 3);
  h.clock.t += 86_400_000;
  const h2result = h.ctx.sendDailyDigest(); // the 2 remaining are candidates next time
  assert.equal(h2result.articles, 2);
});
test('optional thinking config is passed to Gemini and validated before network', () => {
  const h = harness({props: {GEMINI_THINKING_BUDGET: '0'}}); h.ctx.previewDigest();
  const g = h.fetches.find(f => f.url.includes('generativelanguage'));
  assert.deepEqual(JSON.parse(g.init.payload).generationConfig, {thinkingConfig: {thinkingBudget: 0}});
  const h2 = harness({props: {GEMINI_THINKING_LEVEL: 'turbo'}});
  assert.throws(() => h2.ctx.previewDigest(), /THINKING_LEVEL/); assert.equal(h2.fetches.length, 0);
  const h3 = harness(); h3.ctx.previewDigest();
  assert.equal(JSON.parse(h3.fetches.find(f => f.url.includes('generativelanguage')).init.payload).generationConfig, undefined);
});
test('triggered stop sends a notice without the delivery ID; manual run does not', () => {
  const h = harness({apiError: true});
  assert.throws(() => h.ctx.sendDailyDigest());
  assert.equal(h.sent.length, 0);
  assert.throws(() => h.ctx.sendDailyDigest({triggerUid: 'tr-1'}), /要約がすべて失敗/);
  assert.equal(notices(h).length, 1); assert.equal(digestMails(h).length, 0);
  assert.ok(!/test-id-/.test(notices(h)[0].subject + notices(h)[0].body));
  assert.equal(notices(h)[0].to, 'me@example.com');
  assert.ok(h.sheets['実行履歴'].rows.some(r => r[2] === '停止通知送信'));
});
test('stop notice can be disabled and is sent daily while a delivery is unresolved', () => {
  const h = harness({mailError: true});
  assert.throws(() => h.ctx.sendDailyDigest({triggerUid: 't'}), /送信結果が不明/);
  assert.equal(notices(h).length, 1);
  assert.throws(() => h.ctx.sendDailyDigest({triggerUid: 't'}), /未確定/);
  assert.equal(notices(h).length, 2); assert.equal(digestMails(h).length, 1);
  const off = harness({apiError: true});
  off.sheets['配信設定'].rows.find(r => r[0] === '停止通知')[1] = 'FALSE';
  assert.throws(() => off.ctx.sendDailyDigest({triggerUid: 't'})); assert.equal(off.sent.length, 0);
});
test('429 is retried using RetryInfo, then succeeds', () => {
  const h = harness({gemini: n => n === 1
    ? {getResponseCode: () => 429, getContentText: () => JSON.stringify({error: {details: [{'@type': 'type.googleapis.com/google.rpc.RetryInfo', retryDelay: '21s'}]}})}
    : geminiOk('要約')});
  assert.equal(h.ctx.sendDailyDigest().status, 'SENT');
  assert.ok(h.sleeps.includes(21000));
});
test('5xx falls back to backoff and gives up after 3 attempts per article', () => {
  const h = harness({gemini: () => ({getResponseCode: () => 503, getContentText: () => 'oops'})});
  assert.throws(() => h.ctx.sendDailyDigest(), /GEMINI_HTTP_503/);
  assert.equal(h.fetches.filter(f => f.url.includes('generativelanguage')).length, 3);
  assert.deepEqual(h.sleeps.filter(ms => ms >= 5000), [5000, 15000]);
});
test('Gemini requests are paced by the 要約間隔秒 setting', () => {
  const items = ['a', 'b', 'c'].map(s => rssItem('https://example.com/' + s, {title: s}));
  const h = harness({root: rssRoot(items)});
  h.ctx.sendDailyDigest();
  assert.deepEqual(h.sleeps, [4000, 4000]);
  const h0 = harness({root: rssRoot(items)});
  h0.sheets['配信設定'].rows.find(r => r[0] === '要約間隔秒')[1] = 0;
  h0.ctx.sendDailyDigest(); assert.deepEqual(h0.sleeps, []);
});

// ---- 1.1.0: recommended improvements ----
test('empty namespaced elements (media:content, atom:link) do not hide real fields', () => {
  const root = rssRoot([element('item', '', [
    element('title', 'T'),
    element('link', '', [], {rel: 'self', href: 'https://example.com/self'}),
    element('link', 'https://example.com/a'),
    element('content', '', [], {url: 'https://example.com/img.jpg'}),
    element('description', '本文の抜粋です')])]);
  const h = harness({root}); const items = h.ctx.parseDigestFeed_('', 'S');
  assert.equal(items.length, 1);
  assert.equal(items[0].url, 'https://example.com/a'); assert.equal(items[0].content, '本文の抜粋です');
});
test('feed decoding honours BOM, HTTP charset and XML declaration (Shift_JIS / CP932)', () => {
  const h = harness();
  const sjisNihon = [0x93, 0xFA, 0x96, 0x7B].map(b => (b << 24) >> 24); // signed bytes like GAS
  const ascii = s => [...Buffer.from(s, 'ascii')];
  const declared = [...ascii('<?xml version="1.0" encoding="Shift_JIS"?><t>'), ...sjisNihon, ...ascii('</t>')];
  assert.equal(h.ctx.decodeDigestFeed_(declared, ''), '<?xml version="1.0"?><t>日本</t>');
  const bare = [...ascii('<t>'), ...sjisNihon, ...ascii('</t>')];
  assert.equal(h.ctx.decodeDigestFeed_(bare, 'application/rss+xml; charset=Shift_JIS'), '<t>日本</t>');
  const bom = [0xEF, 0xBB, 0xBF, ...Buffer.from('\n <rss>日本</rss>')].map(b => (b << 24) >> 24);
  assert.equal(h.ctx.decodeDigestFeed_(bom, 'text/xml; charset=ISO-8859-1'), '<rss>日本</rss>');
  assert.equal(h.ctx.normalizeCharset_('x-sjis'), 'windows-31j');
  assert.equal(h.ctx.headerValue_({'content-type': 'a'}, 'Content-Type'), 'a');
});
test('JST and numeric-offset dates are parsed', () => {
  const h = harness();
  assert.equal(h.ctx.parseFeedDate_('Mon, 26 Oct 2026 10:00:00 JST'), '2026-10-26T01:00:00.000Z');
  assert.equal(h.ctx.parseFeedDate_('Mon, 26 Oct 2026 10:00:00 +0900'), '2026-10-26T01:00:00.000Z');
  assert.equal(h.ctx.parseFeedDate_('not a date'), '');
});
test('duplicate check reads only status..keys columns, not mail bodies', () => {
  const h = harness(); h.ctx.sendDailyDigest(); mailRows(h).reads = [];
  h.ctx.sendDailyDigest();
  const bodyReads = mailRows(h).reads.filter(r => r.row >= 2);
  assert.ok(bodyReads.length > 0);
  assert.ok(bodyReads.every(r => r.col === 9 && r.width === 3));
});
test('text cells are written as plain text; dates and numbers keep their formats', () => {
  const h = harness(); h.ctx.sendDailyDigest();
  const rec = mailRows(h).formats.find(f => f.row === 2 && f.col === 1).f[0];
  assert.equal(rec[0], '@'); assert.equal(rec[1], 'yyyy-mm-dd hh:mm:ss'); assert.equal(rec[7], '0'); assert.equal(rec[10], '@');
  assert.ok(mailRows(h).formats.some(f => f.row === 2 && f.col === 3 && f.f[0][0] === 'yyyy-mm-dd hh:mm:ss'));
});
test('SPREADSHEET_ID is optional for a bound script', () => {
  const h = harness({props: {SPREADSHEET_ID: null}});
  assert.equal(h.ctx.previewDigest().status, 'PREVIEW');
  assert.throws(() => harness({props: {SPREADSHEET_ID: null}, noActive: true}), /SPREADSHEET_ID/);
});
test('setupSheets adds missing 1.1.0 settings to an old sheet and a status dropdown', () => {
  const h = harness();
  const s = h.sheets['配信設定'];
  s.rows = s.rows.filter(r => !['要約間隔秒', '停止通知'].includes(r[0]));
  const msg = h.ctx.setupSheets();
  assert.match(msg, /要約間隔秒、停止通知/);
  assert.equal(s.rows[1][1], 'me@example.com');
  assert.equal(s.rows.filter(r => r[0] === '停止通知').length, 1);
  assert.equal(h.ctx.setupSheets().includes('追加'), false);
  assert.deepEqual([...mailRows(h).validation.rule.list], ['PREVIEW', 'PREPARED', 'SENDING', 'UNKNOWN', 'SENT', 'CANCELLED']);
  assert.equal(mailRows(h).validation.col, 9);
});
test('old settings sheets without 1.1.0 rows still work with defaults', () => {
  const h = harness();
  h.sheets['配信設定'].rows = h.sheets['配信設定'].rows.filter(r => !['要約間隔秒', '停止通知'].includes(r[0]));
  assert.equal(h.ctx.sendDailyDigest().status, 'SENT');
});

// ---- 1.1.1: review regressions ----
test('stop notice exact recipient, subject and body are archived separately', () => {
  const h = harness({apiError: true});
  assert.throws(() => h.ctx.sendDailyDigest({triggerUid: 't'}));
  const row = h.sheets['通知履歴'].rows[1], mail = notices(h)[0];
  assert.equal(row[3], mail.to); assert.equal(row[4], mail.subject); assert.equal(row[5], mail.body);
  assert.equal(row[6], 'STOP'); assert.equal(row[7], 'SENT');
  assert.equal(mailRows(h).rows.length, 1);
  assert.ok(!/test-id-/.test(mail.body));
});
test('notice archive failure prevents unarchived notice from being sent', () => {
  const h = harness({apiError: true}), sh = h.sheets['通知履歴'], orig = sh.getRange.bind(sh);
  sh.getRange = (row, ...args) => { if (row >= 2) throw Error('unavailable'); return orig(row, ...args); };
  assert.throws(() => h.ctx.sendDailyDigest({triggerUid: 't'}));
  assert.equal(h.sent.length, 0);
  sh.getRange = orig; assert.equal(sh.rows.length, 1); // no orphan status-only row
});
test('post-send run log failure returns SENT with a warning and sends no extra mail', () => {
  const h = harness(), sh = h.sheets['実行履歴'], orig = sh.getRange.bind(sh);
  sh.getRange = (row, ...args) => { if (row >= 2) throw Error('log unavailable'); return orig(row, ...args); };
  const result = h.ctx.sendDailyDigest({triggerUid: 't'});
  assert.equal(result.status, 'SENT'); assert.equal(mailRows(h).rows[1][8], 'SENT');
  assert.equal(h.sent.length, 1); assert.equal(h.sheets['通知履歴'].rows.length, 1);
  assert.ok(result.warnings.some(w => /実行履歴の保存に失敗.*再送は不要/.test(w)));
  sh.getRange = orig;
  assert.equal(h.ctx.sendDailyDigest().status, 'NO_NEW_ITEMS'); assert.equal(h.sent.length, 1);
});
test('failed final mail history write reports accepted mail as record warning', () => {
  const h = harness(); mailRows(h).failSent = true;
  assert.throws(() => h.ctx.sendDailyDigest({triggerUid: 't'}), /送信処理は正常終了/);
  assert.equal(mailRows(h).rows[1][8], 'SENDING');
  assert.match(h.sent[1].subject, /記録警告/);
  assert.equal(h.sheets['通知履歴'].rows[1][6], 'RECORD_WARNING');
});
test('failed notice finalization keeps archived body and UNKNOWN', () => {
  const h = harness({apiError: true}); h.sheets['通知履歴'].failSent = true;
  assert.throws(() => h.ctx.sendDailyDigest({triggerUid: 't'}));
  const row = h.sheets['通知履歴'].rows[1];
  assert.equal(row[5], h.sent[0].body); assert.equal(row[7], 'UNKNOWN');
  assert.match(row[8], /通知送信処理は正常終了/); assert.equal(h.sent.length, 1);
});
test('RetryInfo 120s is preserved and no early retry or next-article request occurs', () => {
  const root = rssRoot(['a','b'].map(s => rssItem('https://example.com/' + s)));
  const response = {getResponseCode: () => 429, getContentText: () => JSON.stringify({error: {details: [{retryDelay: '120s'}]}})};
  const h = harness({root, gemini: () => response});
  assert.equal(h.ctx.geminiRetryDelayMs_(response, 0), 120000);
  assert.throws(() => h.ctx.sendDailyDigest(), /次回候補/);
  assert.equal(h.fetches.filter(f => f.url.includes('generativelanguage')).length, 1);
  assert.equal(h.sleeps.length, 0); assert.equal(h.sent.length, 0);
});
test('RetryInfo longer than remaining budget defers without shortened wait', () => {
  const h = harness({gemini: () => ({getResponseCode: () => 429,
    getContentText: () => JSON.stringify({error: {details: [{retryDelay: '21s'}]}})})});
  assert.throws(() => h.ctx.summarizeDigestArticle_({title:'T', source:'S', content:'C'}, h.ctx.loadGeminiConfig_(),
    h.clock.t + 10000, {last:0, intervalMs:0}), /RETRY_DEFERRED/);
  assert.equal(h.sleeps.length, 0);
});
test('long RetryInfo after one success sends only completed summaries', () => {
  const root = rssRoot(['a','b','c'].map(s => rssItem('https://example.com/' + s)));
  const h = harness({root, gemini: n => n === 1 ? geminiOk('要約') : {getResponseCode: () => 429,
    getContentText: () => JSON.stringify({error:{details:[{retryDelay:'120s'}]}})}});
  const result = h.ctx.sendDailyDigest(); assert.equal(result.status, 'SENT'); assert.equal(result.articles, 1);
  assert.equal(h.fetches.filter(f => f.url.includes('generativelanguage')).length, 2);
  assert.match(result.warnings[0], /API指定の待機/);
});

// ---- 1.1.2: notice staging, optional notice sheet, retry wording ----
test('transient failure archiving a notice leaves no row and sends nothing', () => {
  const h = harness({apiError: true}), sh = h.sheets['通知履歴'], orig = sh.getRange.bind(sh);
  let first = true;
  sh.getRange = (row, col, ...a) => { if (row >= 2 && col === 1 && first) { first = false; throw Error('transient'); } return orig(row, col, ...a); };
  assert.throws(() => h.ctx.sendDailyDigest({triggerUid: 't'}));
  assert.equal(h.sent.length, 0); assert.equal(sh.rows.length, 1);
  assert.ok(h.sheets['実行履歴'].rows.some(r => r[2] === '通知未送信' && /保存できなかった/.test(r[3])));
});
test('failure before sendEmail marks the archived notice CANCELLED, not UNKNOWN', () => {
  const h = harness({apiError: true}), sh = h.sheets['通知履歴'], orig = sh.getRange.bind(sh);
  sh.getRange = (row, col, ...a) => {
    const r = orig(row, col, ...a), set = r.setValues;
    r.setValues = v => { if (v[0][0] === 'SENDING') throw Error('write failed'); return set.call(r, v); };
    return r;
  };
  assert.throws(() => h.ctx.sendDailyDigest({triggerUid: 't'}));
  assert.equal(h.sent.length, 0);
  assert.equal(sh.rows[1][7], 'CANCELLED'); assert.match(sh.rows[1][8], /未送信/);
  assert.ok(sh.rows[1][5].length > 0); // body kept
});
test('exception from the notice sendEmail itself is UNKNOWN', () => {
  const h = harness({apiError: true, noticeMailError: true});
  assert.throws(() => h.ctx.sendDailyDigest({triggerUid: 't'}));
  const row = h.sheets['通知履歴'].rows[1];
  assert.equal(row[7], 'UNKNOWN'); assert.match(row[8], /通知送信で例外/);
  assert.ok(h.sheets['実行履歴'].rows.some(r => r[2] === '通知未確定'));
});
test('1.1.2 code on a sheet without 通知履歴 keeps delivering and warns', () => {
  const h = harness(); delete h.sheets['通知履歴'];
  const result = h.ctx.sendDailyDigest({triggerUid: 't'});
  assert.equal(result.status, 'SENT'); assert.equal(h.sent.length, 1);
  assert.ok(result.warnings.some(w => /通知履歴シートがない/.test(w)));
  const h2 = harness({apiError: true}); delete h2.sheets['通知履歴'];
  assert.throws(() => h2.ctx.sendDailyDigest({triggerUid: 't'}), /要約がすべて失敗/);
  assert.equal(h2.sent.length, 0);
  assert.ok(h2.sheets['実行履歴'].rows.some(r => r[2] === '通知未送信' && /setupSheets/.test(r[3])));
  h2.ctx.setupSheets(); assert.ok(h2.sheets['通知履歴']);
});
test('missing 通知履歴 warning does not mask an all-feeds failure', () => {
  const h = harness(); delete h.sheets['通知履歴'];
  h.ctx.XmlService.parse = () => { throw Error('bad xml'); };
  assert.throws(() => h.ctx.sendDailyDigest(), /全RSSの取得に失敗/);
});
test('deferral wording distinguishes API RetryInfo from local backoff', () => {
  const h = harness({gemini: () => ({getResponseCode: () => 503, getContentText: () => 'x'})});
  let err;
  try { h.ctx.summarizeDigestArticle_({title: 'T', source: 'S', content: 'C'}, h.ctx.loadGeminiConfig_(), h.clock.t + 3000, {last: 0, intervalMs: 0}); } catch (e) { err = e; }
  assert.ok(err.digestDeferred); assert.match(err.digestReason, /HTTP 503後の再試行待機/); assert.doesNotMatch(err.digestReason, /API指定/);
  const h2 = harness({gemini: () => ({getResponseCode: () => 429, getContentText: () => JSON.stringify({error: {details: [{retryDelay: '120s'}]}})})});
  try { h2.ctx.summarizeDigestArticle_({title: 'T', source: 'S', content: 'C'}, h2.ctx.loadGeminiConfig_(), h2.clock.t + 200000, {last: 0, intervalMs: 0}); } catch (e) { err = e; }
  assert.match(err.digestReason, /API指定の待機（120秒）/);
});

// ---- 1.1.3: RetryInfo on the last attempt still defers subsequent articles ----
for (const finalCode of [429, 503]) {
  for (const delay of ['120s', '2s']) {
    test(`last ${finalCode} with RetryInfo ${delay} defers later articles`, () => {
      const root = rssRoot(['a', 'b'].map(s => rssItem('https://example.com/' + s, {title:s})));
      const h = harness({root, gemini: n => n < 3
        ? {getResponseCode: () => 503, getContentText: () => '{}'}
        : n === 3 ? {getResponseCode: () => finalCode, getContentText: () => JSON.stringify({error:{details:[{retryDelay:delay}]}})}
        : geminiOk('should not be requested')});
      assert.throws(() => h.ctx.sendDailyDigest(), /最終試行の失敗.*次回候補/);
      assert.equal(h.fetches.filter(f => f.url.includes('generativelanguage')).length, 3);
      assert.deepEqual(h.sleeps, [5000, 15000]);
      assert.equal(h.sent.length, 0);
    });
  }
}
test('last RetryInfo failure retains already completed summaries without further requests', () => {
  const root = rssRoot(['a', 'b', 'c'].map(s => rssItem('https://example.com/' + s, {title:s})));
  const h = harness({root, gemini: n => n === 1 ? geminiOk('完了済み') : n < 4
    ? {getResponseCode: () => 503, getContentText: () => '{}'}
    : n === 4 ? {getResponseCode: () => 429, getContentText: () => JSON.stringify({error:{details:[{retryDelay:'120s'}]}})}
    : geminiOk('should not be requested')});
  const result = h.ctx.sendDailyDigest();
  assert.equal(result.status, 'SENT'); assert.equal(result.articles, 1);
  assert.equal(h.fetches.filter(f => f.url.includes('generativelanguage')).length, 4);
  assert.equal(h.sent.length, 1);
  assert.match(result.warnings[0], /最終試行の失敗.*残り2件/);
  assert.deepEqual(JSON.parse(mailRows(h).rows[1][10]), ['https://example.com/a']);
});
