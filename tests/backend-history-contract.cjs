const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const crypto = require('node:crypto');

const root = path.resolve(__dirname, '..');
const source = fs.readFileSync(path.join(root, 'GoogleAppsScript-Code.gs'), 'utf8');
const privatePin = 'private-synthetic-pin-123';
const syncKey = 'private-synthetic-sync-' + 'a'.repeat(40);
const now = '2026-11-24T02:20:00.000Z';
const headers = ['event_id', 'chart_number', 'peak_flow', 'assessment_date', 'measured_at', 'recorded_at',
  'received_at', 'source', 'source_event_id', 'overlay_id', 'sync_status', 'committed_at', 'schema_version',
  'linked_act_identity', 'payload_hash', 'updated_at', 'error_code', 'error_message'];
const actHeader = Array.from({ length: 16 }, (_, i) => 'ACT column ' + i);
const actRow = (chart = '35292', stamp = '2026-11-23T17:00:00.000Z', pf = '') =>
  [stamp, 'CHART-' + chart, chart, 'Synthetic ACT name', '', 'adult', 'ACT', 0, 'QA', 'good', '', '', '', '', '', pf];

function harness(options = {}) {
  const writes = [], logs = [], emails = [], propertyWrites = [], reads = [];
  let activeLock = false, releases = 0, flushes = 0, writeCount = 0;
  const properties = new Map(Object.entries(options.properties ?? { ACT_NURSE_PIN: privatePin, ACT_PF_SYNC_KEY: syncKey }));
  function makeSheet(name, data = []) {
    const rows = data.map(row => row.slice());
    const sheet = {
      rows, getName: () => name, getMaxColumns: () => 26, getMaxRows: () => 10000,
      insertColumnsAfter() {}, insertRowsAfter() {},
      getLastColumn: () => Math.max(0, ...rows.map(row => row.length)),
      getLastRow: () => rows.length,
      getDataRange: () => ({ getValues: () => rows.map(row => row.slice()) }),
      appendRow(row) { rows.push(row.slice()); writes.push({ sheet: name, row: rows.length, values: row.slice() }); },
      getRange(row, col, count = 1, columns = 1) {
        const range = {
          setNumberFormat() { return range; },
          getValues() {
            const result = Array.from({ length: count }, (_, i) => Array.from({ length: columns }, (_, j) => rows[row - 1 + i]?.[col - 1 + j] ?? ''));
            if (options.corruptReadback && writeCount && name === 'PF Records' && row > 1) result[0][2] = 999;
            return result;
          },
          setValues(values) {
            if (name === 'ACT Records') throw Error('History feature must never write ACT Records');
            if (options.writeFailure && row > 1) throw Error('Synthetic registry write failure');
            if (row > 1) assert.equal(activeLock, true, 'event/ack writes hold the lock');
            values.forEach((cells, i) => {
              while (rows.length < row + i) rows.push([]);
              cells.forEach((value, j) => { rows[row - 1 + i][col - 1 + j] = value; });
            });
            writes.push({ sheet: name, row, col, values: values.map(r => r.slice()) }); writeCount++;
            return range;
          }
        }; return range;
      }
    }; return sheet;
  }
  const sheets = new Map();
  sheets.set('ACT Records', makeSheet('ACT Records', options.actRows || [actHeader]));
  if (!options.missingRegistry) sheets.set('PF Records', makeSheet('PF Records', options.pfRows || [headers]));
  const ss = {
    getSheetByName(name) { reads.push(name); return sheets.get(name) || null; },
    insertSheet(name) { const sheet = makeSheet(name); sheets.set(name, sheet); return sheet; }
  };
  const NativeDate = Date;
  function FixedDate(...args) { return new NativeDate(...(args.length ? args : [options.now || now])); }
  FixedDate.prototype = NativeDate.prototype; FixedDate.parse = NativeDate.parse; FixedDate.UTC = NativeDate.UTC;
  const ctx = {
    Date: FixedDate, console: { log: value => logs.push(value) },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: key => properties.has(key) ? properties.get(key) : null,
      setProperty(key, value) { properties.set(key, value); propertyWrites.push(key); }
    }) },
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush() { flushes++; if (options.flushFailure) throw Error('Synthetic flush failure'); } },
    GmailApp: { sendEmail: (...args) => emails.push(args), getAliases: () => [] },
    LockService: { getScriptLock: () => ({ tryLock() { if (options.busy) return false; activeLock = true; return true; }, releaseLock() { activeLock = false; releases++; } }) },
    Utilities: {
      DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' }, getUuid: () => crypto.randomUUID(),
      computeDigest: (algorithm, text) => [...crypto.createHash(algorithm).update(text).digest()],
      formatDate: date => new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit' }).format(date)
    },
    ContentService: { MimeType: { JSON: 'json', TEXT: 'text' }, createTextOutput: text => ({ text, setMimeType() { return this; } }) }
  };
  vm.createContext(ctx); vm.runInContext(source, ctx);
  const post = (mode, values = {}) => JSON.parse(ctx.doPost({ parameter: { mode, pin: privatePin, syncKey, ...values } }).text);
  const mobile = (values = {}) => post('pf_record', { eventId: 'mobile-event-00000001', chartNumber: '35292', peakFlow: '95', assessmentDate: '2026-11-23', ...values });
  const overlay = (values = {}) => post('pf_sync', { eventId: 'overlay:clinical_assessments:42', chartNumber: '0035292', peakFlow: '95.0',
    assessmentDate: '2026-11-23', recordedAt: '2026-11-24T09:00:00+08:00', sourceEventId: '42', ...values });
  const history = (values = {}) => post('nurse_history', { chartNumber: '35292', startDate: '2026-11-23', endDate: '2026-11-24', ...values });
  return { ctx, post, mobile, overlay, history, sheets, properties, writes, logs, emails, reads, propertyWrites,
    get releases() { return releases; }, get flushes() { return flushes; } };
}

module.exports = { harness, privatePin, syncKey, headers, actHeader, actRow, now };

if (require.main === module) {
test('history release is exact, both copies match, and no new MailApp scope', () => {
  assert.equal(source, fs.readFileSync(path.join(root, 'c-act-asthma-tool', 'GoogleAppsScript-Code.gs'), 'utf8'));
  const h = harness(); const health = JSON.parse(h.ctx.doGet({ parameter: { mode: 'health' } }).text);
  assert.equal(health.release, '2026-11-24-pf-history1'); assert.doesNotMatch(source, /\bMailApp\s*\./);
});

test('all new modes reject missing private credentials and public legacy fallback', () => {
  for (const properties of [{}, { ACT_NURSE_PIN: harness().ctx.ACT_LEGACY_NURSE_PIN }]) {
    const h = harness({ properties });
    for (const mode of ['nurse_history', 'pf_record', 'pf_sync', 'pf_events', 'pf_ack']) {
      assert.equal(h.post(mode).error.code, 'AUTH_NOT_CONFIGURED');
    }
    assert.equal(h.writes.length, 0);
  }
  const h = harness(); assert.equal(h.history({ pin: 'wrong' }).error.code, 'AUTH_FAILED');
  assert.equal(h.overlay({ syncKey: 'wrong' }).error.code, 'AUTH_FAILED');
});

test('exact chart normalization, Taipei midnight, and zero ACT score are preserved', () => {
  const h = harness({ actRows: [actHeader, actRow('35292'), actRow('0035292', '2026-11-23T10:00:00Z'), actRow('99999')] });
  const result = h.history(); assert.equal(result.ok, true); assert.equal(result.rows.length, 2);
  assert.equal(result.rows[0].assessmentDate, '2026-11-24'); assert.equal(result.rows[0].chartNumber, '0035292');
  assert.equal(result.rows[0].score, 0); assert.match(result.rows[0].rowKey, /^[a-f0-9]{64}$/);
  assert.equal(h.history({ chartNumber: '0035292' }).rows.length, 2); assert.equal(h.writes.length, 0);
});

test('history defaults today or last seven days and requires exact chart for past dates', () => {
  const h = harness();
  const today = h.post('nurse_history'); assert.equal(today.startDate, '2026-11-24'); assert.equal(today.endDate, '2026-11-24');
  const week = h.post('nurse_history', { chartNumber: '35292' }); assert.equal(week.startDate, '2026-11-18');
  assert.equal(h.history({ chartNumber: '' }).error.code, 'CHART_REQUIRED_FOR_HISTORY');
});

test('history rejects future, invalid dates, wildcard charts, and overlarge ranges', () => {
  const h = harness();
  for (const chartNumber of ['*', '35292 OR 1=1', '12345678', '0000000', 'NAME-test']) assert.equal(h.history({ chartNumber }).error.code, 'INVALID_CHART');
  assert.equal(h.history({ endDate: '2026-11-25' }).error.code, 'FUTURE_DATE');
  assert.equal(h.history({ startDate: '2026-02-30' }).error.code, 'INVALID_DATE');
  assert.equal(h.history({ startDate: '2026-10-01' }).error.code, 'DATE_RANGE_TOO_LARGE');
  assert.equal(h.history({ startDate: '2026-11-24', endDate: '2026-11-23' }).ok, false);
});

test('any past date can be selected within a bounded range; pagination does not fake completeness', () => {
  const h = harness({ actRows: [actHeader, actRow('35292', '2020-01-02T01:00:00Z'), actRow('35292', '2020-01-01T01:00:00Z')] });
  const first = h.history({ startDate: '2020-01-01', endDate: '2020-01-02', limit: '1' });
  assert.equal(first.rows.length, 1); assert.equal(first.hasMore, true); assert.match(first.nextCursor, /^v1\./);
  const second = h.history({ startDate: '2020-01-01', endDate: '2020-01-02', limit: '1', cursor: first.nextCursor });
  assert.equal(second.rows.length, 1); assert.equal(second.hasMore, false);
  assert.notEqual(second.rows[0].timestamp, first.rows[0].timestamp);
});

test('snapshot cursor excludes later appends and refuses changed ACT facts', () => {
  const h = harness({ actRows: [actHeader, actRow('35292', '2026-11-23T01:00:00Z'), actRow()] });
  const first = h.history({ limit: '1' });
  h.sheets.get('ACT Records').rows.push(actRow('35292', '2026-11-24T01:00:00Z'));
  const second = h.history({ limit: '1', cursor: first.nextCursor });
  assert.equal(second.ok, true); assert.equal(second.rows[0].timestamp, '2026-11-23T01:00:00Z'); assert.equal(second.hasMore, false);
  h.sheets.get('ACT Records').rows[1][15] = 100;
  assert.equal(h.history({ limit: '1', cursor: first.nextCursor }).error.code, 'HISTORY_CHANGED_RELOAD');
  assert.equal(h.history({ chartNumber: '99999', cursor: first.nextCursor }).error.code, 'INVALID_PAGINATION');
});

test('standalone mobile PF has no ACT projection, name, score or invented measurement timestamp', () => {
  const h = harness({ actRows: [actHeader, actRow()] }); const result = h.mobile();
  assert.equal(result.ok, true); assert.equal(result.record.peakFlow, 95); assert.equal(result.record.chartNumber, '0035292');
  assert.equal(result.record.source, 'mobile_pf'); assert.equal(result.record.measuredAt, '');
  assert.equal(result.record.syncStatus, 'pending'); assert.equal(result.record.overlayId, '');
  assert.equal(result.record.recordedAt, now); assert.equal(result.record.assessmentDate, '2026-11-23');
  assert.equal(result.record.name, undefined); assert.equal(result.record.score, undefined);
  assert.equal(h.sheets.get('ACT Records').rows[1][15], ''); assert.equal(h.emails.length, 0);
  assert.equal(h.writes.every(write => write.sheet === 'PF Records'), true);
  assert.equal(h.history().pfRecords[0].eventId, result.record.eventId);
});

test('same event retry returns existing acknowledgment; changed payload is a conflict', () => {
  const h = harness(); const first = h.mobile(); const count = h.writes.length;
  const retry = h.mobile(); assert.equal(retry.duplicate, true); assert.equal(retry.record.versionKey, first.record.versionKey);
  assert.equal(h.writes.length, count);
  assert.equal(h.mobile({ peakFlow: '96' }).error.code, 'IDEMPOTENCY_CONFLICT');
});

test('equivalent Overlay source id formats normalize to the same immutable event', () => {
  const h = harness(); assert.equal(h.overlay().ok, true);
  assert.equal(h.overlay({ sourceEventId: 'clinical_assessments:42' }).duplicate, true);
  assert.equal(h.history().pfRecords.length, 1);
});

test('registry fact corruption fails closed instead of returning altered clinical values', () => {
  const h = harness(); h.mobile(); h.sheets.get('PF Records').rows[1][2] = 200;
  assert.equal(h.history().error.code, 'PF_RECORD_INTEGRITY');
  assert.equal(h.post('pf_events').error.code, 'PF_RECORD_INTEGRITY');
  assert.equal(h.mobile().error.code, 'PF_RECORD_INTEGRITY');
});

test('next-day retry preserves original date and entry time without another append', () => {
  const first = harness(); const result = first.mobile();
  const rows = first.sheets.get('PF Records').rows;
  const next = harness({ now: '2026-11-25T02:20:00Z', pfRows: rows });
  const retry = next.mobile(); assert.equal(retry.duplicate, true); assert.equal(retry.record.recordedAt, result.record.recordedAt);
  assert.equal(retry.record.assessmentDate, '2026-11-23'); assert.equal(next.writes.length, 0);
});

test('an additional same-day measurement requires explicit confirmation; retry needs no reconfirmation', () => {
  const h = harness(); h.mobile();
  assert.equal(h.mobile({ eventId: 'mobile-event-00000002', peakFlow: '100' }).error.code, 'ADDITIONAL_CONFIRMATION_REQUIRED');
  const confirmed = h.mobile({ eventId: 'mobile-event-00000002', peakFlow: '100', confirmAdditional: 'yes' });
  assert.equal(confirmed.ok, true); assert.equal(h.mobile({ eventId: 'mobile-event-00000002', peakFlow: '100' }).duplicate, true);
});

test('existing ACT PF also triggers additional-measurement confirmation without modifying ACT', () => {
  const h = harness({ actRows: [actHeader, actRow('35292', '2026-11-23T01:00:00Z', 95)] });
  assert.equal(h.mobile().error.code, 'ADDITIONAL_CONFIRMATION_REQUIRED'); assert.equal(h.writes.length, 0);
  assert.equal(h.mobile({ confirmAdditional: 'yes' }).ok, true); assert.equal(h.sheets.get('ACT Records').rows[1][15], 95);
});

test('Overlay source requires stable committed ID and preserves actual recordedAt independently', () => {
  const h = harness(); const result = h.overlay();
  assert.equal(result.ok, true); assert.equal(result.record.overlayId, '42'); assert.equal(result.record.syncStatus, 'synced');
  assert.equal(result.record.source, 'overlay_manual'); assert.equal(result.record.recordedAt, '2026-11-24T01:00:00.000Z');
  assert.equal(result.record.assessmentDate, '2026-11-23'); assert.equal(result.record.measuredAt, '');
  assert.equal(h.overlay().duplicate, true); assert.equal(h.overlay({ sourceEventId: '43' }).error.code, 'INVALID_SOURCE_EVENT');
  assert.equal(h.overlay({ eventId: 'mobile-spoof-event' }).error.code, 'INVALID_SOURCE_EVENT');
  assert.equal(h.mobile({ eventId: 'overlay:clinical_assessments:1' }).error.code, 'INVALID_SOURCE_EVENT');
});

test('invalid PF, invalid or future dates, missing timestamps cannot create events', () => {
  for (const value of ['49', '901', '95.5', 'NaN', '=95', '9.5e1']) assert.equal(harness().mobile({ peakFlow: value }).error.code, 'INVALID_PF');
  const h = harness(); assert.equal(h.mobile({ assessmentDate: '2026-11-25' }).error.code, 'FUTURE_DATE');
  assert.equal(h.overlay({ recordedAt: '2026-11-24 09:00:00' }).error.code, 'INVALID_TIMESTAMP');
  assert.equal(h.mobile({ measuredAt: '2026-11-22T01:00:00Z' }).error.code, 'MEASUREMENT_DATE_MISMATCH');
  assert.equal(h.mobile({ measuredAt: '2026-02-30T01:00:00Z' }).error.code, 'INVALID_TIMESTAMP');
  assert.equal(h.writes.length, 0);
});

test('pending pull is service-only and cannot skip an unacknowledged event using cursor', () => {
  const h = harness(); h.mobile(); h.overlay();
  const pull = h.post('pf_events', { afterRow: '99999' }); assert.equal(pull.events.length, 1);
  assert.equal(pull.events[0].source, 'mobile_pf'); assert.equal(pull.events[0].syncStatus, 'pending');
  assert.equal(h.post('pf_events', { syncKey: '' }).error.code, 'AUTH_FAILED');
});

test('only durable Overlay acknowledgment marks synced, and repeats are idempotent', () => {
  const h = harness(); h.mobile();
  assert.equal(h.post('pf_ack', { eventId: 'mobile-event-00000001' }).error.code, 'INVALID_ACK');
  const args = { eventId: 'mobile-event-00000001', overlayId: '101', committedAt: '2026-11-24T02:19:00Z' };
  const ack = h.post('pf_ack', args); assert.equal(ack.record.syncStatus, 'synced'); assert.equal(ack.record.overlayId, '101');
  assert.equal(h.post('pf_events').events.length, 0); assert.equal(h.post('pf_ack', args).duplicate, true);
  assert.equal(h.post('pf_ack', { ...args, overlayId: '102' }).error.code, 'ACK_CONFLICT');
  assert.equal(h.mobile().record.syncStatus, 'synced');
});

test('unknown or ambiguous patient rejection never claims an Overlay write', () => {
  const h = harness(); h.mobile();
  const ack = h.post('pf_ack', { eventId: 'mobile-event-00000001', status: 'rejected_patient', errorCode: 'PATIENT_NOT_FOUND' });
  assert.equal(ack.ok, true); assert.equal(ack.record.syncStatus, 'rejected_patient');
  assert.equal(ack.record.overlayId, ''); assert.equal(ack.record.committedAt, ''); assert.equal(h.post('pf_events').events.length, 0);
  assert.equal(h.post('pf_ack', { eventId: 'mobile-event-00000001', status: 'rejected_patient', overlayId: '101' }).error.code, 'INVALID_ACK');
});

test('missing registry and reordered headers fail instead of pretending history is empty', () => {
  assert.equal(harness({ missingRegistry: true }).history().error.code, 'PF_REGISTRY_NOT_INITIALIZED');
  const bad = headers.slice(); [bad[1], bad[2]] = [bad[2], bad[1]];
  const h = harness({ pfRows: [bad] }); assert.equal(h.history().error.code, 'PF_SCHEMA_MISMATCH'); assert.equal(h.writes.length, 0);
});

test('write, readback or lock failure never returns a success acknowledgment', () => {
  for (const options of [{ writeFailure: true }, { corruptReadback: true }, { flushFailure: true }, { busy: true }]) {
    const h = harness(options); assert.equal(h.mobile().ok, false);
    if (!options.busy) assert.equal(h.releases, 1);
  }
});

test('owner initialization creates only missing secrets and header, without exposing credentials', () => {
  const h = harness({ properties: {}, missingRegistry: true }); const result = h.ctx.initializePFHistoryAccess();
  assert.equal(result.ok, true); assert.equal(result.nursePinCreated, true); assert.equal(result.syncKeyCreated, true);
  assert.match(h.properties.get('ACT_NURSE_PIN'), /^\d{12}$/); assert.ok(h.properties.get('ACT_PF_SYNC_KEY').length >= 32);
  assert.equal(h.writes.every(write => write.sheet === 'PF Records' && write.row === 1), true);
  const serialized = JSON.stringify(result) + h.logs.join('');
  assert.equal(serialized.includes(h.properties.get('ACT_NURSE_PIN')), false); assert.equal(serialized.includes(h.properties.get('ACT_PF_SYNC_KEY')), false);
  const writes = h.propertyWrites.length; h.ctx.initializePFHistoryAccess(); assert.equal(h.propertyWrites.length, writes);
});

test('owner initialization preserves existing private values and isolated QA never writes ACT or live registry', () => {
  const h = harness(); h.ctx.initializePFHistoryAccess(); assert.equal(h.propertyWrites.length, 0);
  h.writes.length = 0; h.reads.length = 0;
  const result = h.ctx.runIsolatedPFHistoryCloudSelfTest(); assert.equal(result.ok, true); assert.equal(result.syntheticOnly, true);
  assert.equal(h.writes.every(write => write.sheet === 'ACT PF History QA 20261124'), true);
  assert.equal(h.reads.includes('ACT Records'), false); assert.equal(h.reads.includes('PF Records'), false); assert.equal(h.emails.length, 0);
});

test('initializer preserves existing private values and reports unusable short configuration', () => {
  const h = harness({ properties: { ACT_NURSE_PIN: 'short', ACT_PF_SYNC_KEY: 'short-private-sync' } });
  const report = h.ctx.initializePFHistoryAccess();
  assert.equal(report.ok, false); assert.equal(h.propertyWrites.length, 0);
  assert.equal(h.history().error.code, 'AUTH_NOT_CONFIGURED');
});
}
