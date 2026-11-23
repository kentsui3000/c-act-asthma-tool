const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const crypto = require('node:crypto');

const root = path.resolve(__dirname, '..');
const source = fs.readFileSync(path.join(root, 'GoogleAppsScript-Code.gs'), 'utf8');
const header = Array.from({ length: 16 }, (_, i) => 'Column ' + (i + 1));
const now = '2026-11-23T22:30:00.000Z'; // Nov 24 06:30 in Taipei.
const secret = 'synthetic-test-credential';
const record = (timestamp = '2026-11-23T17:00:00.000Z') =>
  [timestamp, 'QA-PATIENT', 'QA-123', 'Synthetic patient', '', 'adult', 'ACT', 0,
    'Synthetic level', 'good', '', '', '', '', '', ''];

function harness(options = {}) {
  const writes = [], appends = [], emails = [], logs = [], sheetReads = [];
  let flushes = 0, releases = 0, locked = false;
  function makeSheet(name, initial) {
    const data = initial.map(r => r.slice());
    return {
      data,
      getName: () => name,
      getMaxColumns: () => 26,
      insertColumnsAfter() {},
      getLastColumn: () => options.columns ?? 16,
      getLastRow: () => data.length,
      getDataRange: () => ({ getValues: () => data.map(r => r.slice()) }),
      appendRow(row) { data.push(row.slice()); appends.push({ sheet: name, row: row.slice() }); },
      setFrozenRows() {},
      getRange(row, col, count = 1, columns = 1) {
        return {
          getValues() {
            return Array.from({ length: count }, (_, i) => {
              const result = Array.from({ length: columns }, (_, j) => data[row - 1 + i]?.[col - 1 + j] ?? '');
              if (options.corruptReadback && writes.length && col === 1 && columns === 16) result[15] = 999;
              return result;
            });
          },
          setValue(value) {
            if (options.writeFailure) throw new Error('Synthetic write error');
            assert.equal(locked, true, 'PF write must hold script lock');
            data[row - 1][col - 1] = value;
            writes.push({ sheet: name, row, col, value });
          }
        };
      }
    };
  }
  const sheets = new Map();
  if (!options.missingSheet) sheets.set('ACT Records', makeSheet('ACT Records', options.data || [header, record()]));
  sheets.set('Email Log', makeSheet('Email Log', []));
  const ss = {
    getSheetByName(name) { sheetReads.push(name); return sheets.get(name) || null; },
    getSheets: () => [...sheets.values()],
    insertSheet(name) { const sheet = makeSheet(name, []); sheets.set(name, sheet); return sheet; }
  };
  const NativeDate = Date;
  function FixedDate(...args) { return new NativeDate(...(args.length ? args : [now])); }
  FixedDate.prototype = NativeDate.prototype;
  FixedDate.now = () => NativeDate.parse(now);
  FixedDate.parse = NativeDate.parse;
  FixedDate.UTC = NativeDate.UTC;
  const ctx = {
    Date: FixedDate,
    console: { log: value => logs.push(value) },
    SpreadsheetApp: {
      getActiveSpreadsheet() { if (options.sheetAuthFailure) throw Error('Synthetic auth failure'); return ss; },
      flush() { flushes++; if (options.flushFailure) throw Error('Synthetic flush failure'); }
    },
    GmailApp: {
      getAliases() { if (options.gmailAuthFailure) throw Error('Synthetic Gmail auth failure'); return ['never-return@example.invalid']; },
      sendEmail(...args) { if (options.mailFailure) throw Error('Synthetic mail failure'); emails.push(args); }
    },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => options.property === undefined ? secret : options.property }) },
    LockService: { getScriptLock: () => ({
      tryLock() { if (options.busy) return false; locked = true; return true; },
      releaseLock() { locked = false; releases++; }
    }) },
    Utilities: {
      DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
      computeDigest: (algorithm, value) => [...crypto.createHash(algorithm).update(value).digest()],
      formatDate(date) {
        return new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit' }).format(date);
      }
    },
    ContentService: { MimeType: { JSON: 'json', TEXT: 'text' }, createTextOutput: text => ({ text, setMimeType(type) { this.mime = type; return this; } }) }
  };
  vm.createContext(ctx);
  vm.runInContext(source, ctx, { filename: 'GoogleAppsScript-Code.gs' });
  const post = (mode, params = {}) => JSON.parse(ctx.doPost({ parameter: { mode, pin: secret, ...params } }).text);
  const list = () => post('nurse_list');
  const update = (params = {}) => {
    const row = list().rows[0];
    return post('update_pf', { rowIndex: String(row.rowIndex), rowKey: row.rowKey, peakFlow: '320', ...params });
  };
  return { ctx, post, list, update, writes, appends, emails, logs, sheetReads, sheets,
    get flushes() { return flushes; }, get releases() { return releases; } };
}

test('both deployed-path copies stay identical and script uses no new MailApp scope', () => {
  assert.equal(source, fs.readFileSync(path.join(root, 'c-act-asthma-tool', 'GoogleAppsScript-Code.gs'), 'utf8'));
  assert.doesNotMatch(source, /\bMailApp\s*\./);
});

test('public health does not touch spreadsheet, Gmail, or patient data', () => {
  const h = harness({ sheetAuthFailure: true, gmailAuthFailure: true });
  const output = JSON.parse(h.ctx.doGet({ parameter: { mode: 'health' } }).text);
  assert.equal(output.ok, true); assert.equal(output.apiVersion, 2); assert.equal(output.release, '2026-11-24-pf1');
  assert.equal(h.sheetReads.length + h.writes.length + h.appends.length + h.emails.length, 0);
  assert.equal(output.rows, undefined);
});

test('GET nurse modes fail closed without reads or writes', () => {
  const h = harness();
  for (const mode of ['nurse_list', 'update_pf']) {
    const response = JSON.parse(h.ctx.doGet({ parameter: { mode, pin: secret, rowIndex: '1', peakFlow: '999' } }).text);
    assert.equal(response.error.code, 'UPGRADE_REQUIRED');
  }
  assert.equal(h.sheetReads.length + h.writes.length + h.appends.length, 0);
});

test('missing and incorrect POST credentials cannot read or write', () => {
  const h = harness();
  for (const mode of ['nurse_list', 'update_pf']) for (const pin of ['', 'wrong']) {
    const response = h.post(mode, { pin });
    assert.equal(response.ok, false); assert.equal(response.error.code, 'AUTH_FAILED');
  }
  assert.equal(h.sheetReads.length + h.writes.length + h.appends.length, 0);
});

test('existing server compatibility PIN is used only when property absent', () => {
  const h = harness({ property: null });
  assert.equal(h.post('nurse_list', { pin: h.ctx.ACT_LEGACY_NURSE_PIN }).ok, true);
  const empty = harness({ property: '' });
  assert.equal(empty.post('nurse_list').error.code, 'AUTH_NOT_CONFIGURED');
});

test('Taipei list includes 00:00–07:59 UTC-previous-day records and keeps score zero', () => {
  const h = harness({ data: [header, record('2026-11-23T16:00:00.000Z'), record('2026-11-23T23:59:59.000Z'), record('2026-11-23T15:59:59.000Z'), record('2026-11-24T16:00:00.000Z')] });
  const response = h.list();
  assert.equal(response.date, '2026-11-24'); assert.equal(response.rows.length, 2);
  assert.equal(response.rows[0].score, 0); assert.match(response.rows[0].rowKey, /^[a-f0-9]{64}$/);
});

test('Date cells and explicit Taipei legacy timestamp strings are supported', () => {
  const h = harness({ data: [header, record(new Date('2026-11-23T17:00:00.000Z')), record('2026/11/24 01:00:00')] });
  assert.equal(h.list().rows.length, 2);
});

test('invalid timestamp rows are surfaced as a warning count', () => {
  const h = harness({ data: [header, record('not-a-date')] });
  assert.equal(h.list().invalidTimestampCount, 1); assert.equal(h.list().rows.length, 0);
});

test('sheet missing, schema mismatch, and service failures are distinct from valid empty list', () => {
  for (const [options, code] of [[{ missingSheet: true }, 'SHEET_NOT_FOUND'], [{ columns: 15 }, 'SCHEMA_MISMATCH'], [{ sheetAuthFailure: true }, 'BACKEND_UNAVAILABLE']]) {
    const response = harness(options).list(); assert.equal(response.ok, false); assert.equal(response.error.code, code); assert.equal(response.rows, undefined);
  }
  const response = harness({ data: [header] }).list(); assert.equal(response.ok, true); assert.equal(response.rows.length, 0);
});

test('only a confirmed integer PF write returns numeric acknowledgment and new version key', () => {
  const h = harness(); const before = h.list().rows[0]; const response = h.update();
  assert.equal(response.ok, true); assert.equal(response.peakFlow, 320); assert.equal(response.previousRowKey, before.rowKey);
  assert.notEqual(response.rowKey, before.rowKey); assert.match(response.rowKey, /^[a-f0-9]{64}$/);
  assert.deepEqual(h.writes, [{ sheet: 'ACT Records', row: 2, col: 16, value: 320 }]);
  assert.equal(h.flushes, 1); assert.equal(h.releases, 1); assert.equal(h.appends.length + h.emails.length, 0);
});

test('concurrent PF edits use compare-and-swap and reject a stale version', () => {
  const h = harness(); const row = h.list().rows[0];
  assert.equal(h.post('update_pf', { rowIndex: '2', rowKey: row.rowKey, peakFlow: '320' }).ok, true);
  const stale = h.post('update_pf', { rowIndex: '2', rowKey: row.rowKey, peakFlow: '340' });
  assert.equal(stale.error.code, 'RECORD_CHANGED'); assert.equal(h.writes.length, 1);
  assert.equal(h.sheets.get('ACT Records').data[1][15], 320);
});

test('unknown modes and empty legacy GETs do not append or mutate records', () => {
  const h = harness();
  assert.equal(JSON.parse(h.ctx.doGet({ parameter: { mode: 'not-supported' } }).text).ok, false);
  assert.equal(h.post('not-supported').ok, false);
  assert.equal(JSON.parse(h.ctx.doGet({ parameter: {} }).text).error.code, 'INVALID_SUBMISSION');
  assert.equal(h.appends.length + h.writes.length + h.emails.length, 0);
});

test('PF 50 and 900 boundaries are accepted', () => {
  for (const value of ['50', '900']) assert.equal(harness().update({ peakFlow: value }).peakFlow, Number(value));
});

test('invalid PF values never write', () => {
  for (const value of ['', '49', '901', '-1', '320.5', '3e2', '=1+1', '320bad', 'NaN', 'Infinity']) {
    const h = harness(); assert.equal(h.update({ peakFlow: value }).error.code, 'INVALID_PF'); assert.equal(h.writes.length, 0);
  }
});

test('header, malformed, negative and out-of-range row references never write', () => {
  for (const value of ['1', '0', '-1', '2x', '2.5', '02', '9007199254740992', '999']) {
    const h = harness(); const response = h.update({ rowIndex: value });
    assert.equal(response.ok, false); assert.equal(h.writes.length, 0);
  }
});

test('missing or forged key cannot write', () => {
  for (const key of ['', 'x', '0'.repeat(64)]) {
    const h = harness(); assert.equal(h.update({ rowKey: key }).ok, false); assert.equal(h.writes.length, 0);
  }
});

test('sorting or identity edits after listing cause RECORD_CHANGED', () => {
  const h = harness(); const row = h.list().rows[0]; h.sheets.get('ACT Records').data[1][1] = 'OTHER-PATIENT';
  const response = h.post('update_pf', { rowIndex: '2', rowKey: row.rowKey, peakFlow: '320' });
  assert.equal(response.error.code, 'RECORD_CHANGED'); assert.equal(h.writes.length, 0); assert.equal(h.releases, 1);
});

test('prior-date edits are rejected even with correct row identity', () => {
  const h = harness({ data: [header, record('2026-11-22T17:00:00.000Z')] });
  const key = h.ctx.nurseRowKey_(2, h.sheets.get('ACT Records').data[1]);
  const response = h.post('update_pf', { rowIndex: '2', rowKey: key, peakFlow: '320' });
  assert.equal(response.error.code, 'DATE_NOT_ALLOWED'); assert.equal(h.writes.length, 0);
});

test('busy lock returns retryable failure with no write', () => {
  const h = harness({ busy: true }); const response = h.update();
  assert.equal(response.error.code, 'BUSY'); assert.equal(response.error.retryable, true); assert.equal(h.writes.length, 0);
});

test('write or flush failure cannot produce a success acknowledgment and releases lock', () => {
  for (const options of [{ writeFailure: true }, { flushFailure: true }]) {
    const h = harness(options); assert.equal(h.update().ok, false); assert.equal(h.releases, 1);
  }
});

test('readback mismatch reports unconfirmed instead of success', () => {
  const h = harness({ corruptReadback: true }); const response = h.update();
  assert.equal(response.error.code, 'WRITE_NOT_CONFIRMED'); assert.equal(h.releases, 1);
});

test('legacy questionnaire preserves 16 columns and catches Gmail errors', () => {
  const h = harness({ mailFailure: true });
  const response = h.ctx.doGet({ parameter: { name: 'Synthetic', email: 'qa@example.invalid', timestamp: now, testType: 'ACT', score: '20', level: 'QA' } });
  assert.equal(response.text, 'ok');
  const submission = h.appends.find(row => row.sheet === 'ACT Records'); assert.equal(submission.row.length, 16);
  const emailLog = h.appends.find(row => row.sheet === 'Email Log'); assert.equal(emailLog.row[5], 'failed');
});

test('owner authorization check never writes, sends, or logs account aliases', () => {
  const h = harness(); const report = h.ctx.readOnlyAuthorizationCheck();
  assert.equal(report.sheets.ok, true); assert.equal(report.gmail.ok, true);
  assert.equal(h.writes.length + h.appends.length + h.emails.length, 0);
  assert.doesNotMatch(h.logs.join(''), /never-return|Synthetic patient|QA-123/);
});

test('manual cloud self-test writes synthetic fixture only in named QA sheet', () => {
  const h = harness(); const report = h.ctx.runIsolatedPFCloudSelfTest();
  assert.equal(report.ok, true); assert.equal(report.confirmedPeakFlow, 321);
  assert.equal(h.sheetReads.includes('ACT Records'), false);
  assert.equal(h.writes.every(write => write.sheet === 'ACT Repair QA 20261124'), true);
  assert.equal(h.appends.every(write => write.sheet === 'ACT Repair QA 20261124'), true);
  assert.equal(h.emails.length, 0); assert.equal(h.sheets.get('ACT Records').data[1][15], '');
});
