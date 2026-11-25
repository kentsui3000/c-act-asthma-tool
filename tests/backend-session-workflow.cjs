'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { harness, actHeader, actRow } = require('./backend-history-contract.cjs');
const today = '2026-11-25';
const fixedNow = '2026-11-25T01:20:00.000Z';
const rowAt = (clock, chart = '35292') => actRow(chart, today + 'T' + clock + '+08:00');
const fixture = rows => harness({ now: fixedNow, actRows: [actHeader, ...rows] });
const history = (h, values = {}) => h.post('nurse_history', {
  chartNumber: '35292', startDate: today, endDate: today, ...values
});
const record = (h, values = {}) => h.mobile({ assessmentDate: today, ...values });

test('workflow release is explicit and the session filter defaults to all', () => {
  const h = fixture([rowAt('09:00:00')]);
  const health = JSON.parse(h.ctx.doGet({ parameter: { mode: 'health' } }).text);
  assert.equal(health.release, '2026-11-25-pf-workflow1');
  assert.equal(history(h).clinicSession, 'all');
});

test('all exact Taipei session boundaries preserve off-hours as unassigned', () => {
  for (const [clock, expected] of [
    ['00:00:00', 'unassigned'], ['07:59:59', 'unassigned'], ['08:00:00', 'S'],
    ['15:29:59', 'S'], ['15:30:00', 'T'], ['18:29:59', 'T'],
    ['18:30:00', 'U'], ['23:29:59', 'U'], ['23:30:00', 'unassigned'], ['23:59:59', 'unassigned']
  ]) {
    const h = fixture([rowAt(clock)]);
    const result = history(h);
    assert.equal(result.ok, true, clock); assert.equal(result.rows.length, 1, clock);
    assert.equal(result.rows[0].clinicSession, expected, clock);
    assert.equal(result.rows[0].sessionSource, 'questionnaire_timestamp');
    assert.equal(result.rows[0].sessionFallback, true);
    assert.equal(result.rows[0].timestamp, today + 'T' + clock + '+08:00');
    assert.equal(result.rows[0].peakFlow, ''); assert.equal(h.writes.length, 0);
  }
});

test('UTC, offset, Date cell and unzoned legacy clock classify in Taipei', () => {
  const dates = ['2026-11-25T00:00:00Z', '2026-11-25T08:00:00+08:00',
    new Date('2026-11-25T00:00:00Z'), '2026/11/25 08:00:00'];
  const h = fixture(dates.map(value => actRow('35292', value)));
  const result = history(h, { clinicSession: 'S' });
  assert.equal(result.rows.length, dates.length);
  assert.equal(result.rows.every(row => row.clinicSession === 'S'), true);
  const midnight = fixture([actRow('35292', '2026-11-24T16:00:00Z')]);
  assert.equal(history(midnight).rows[0].clinicSession, 'unassigned');
});

test('date-only and malformed clocks do not invent a session or a native HIS code', () => {
  const h = fixture([actRow('35292', today)]);
  const result = history(h);
  assert.equal(result.rows[0].clinicSession, 'unassigned');
  assert.equal(result.rows[0].sessionSource, 'not_available');
  assert.equal(result.rows[0].sessionFallback, false);
  for (const value of ['2026-11-25T24:00:00+08:00', '2026/11/25 08:60:00', '2026-02-30T08:00:00+08:00', 'not-a-date']) {
    assert.equal(h.ctx.historySession_(value, 'questionnaire_timestamp').clinicSession, 'unassigned');
  }
  const misleading = rowAt('09:00:00'); misleading[9] = 'U';
  const originalLevel = fixture([misleading]);
  assert.equal(history(originalLevel).rows[0].clinicSession, 'S');
  assert.equal(history(originalLevel).rows[0].levelKey, 'U');
});

test('single past day supports all patients, while multi-day history still requires a chart', () => {
  const h = fixture([actRow('35292', '2026-11-24T09:00:00+08:00'), actRow('99', '2026-11-24T16:00:00+08:00')]);
  const past = h.post('nurse_history', { startDate: '2026-11-24', endDate: '2026-11-24', clinicSession: 'all' });
  assert.equal(past.ok, true); assert.equal(past.rows.length, 2); assert.equal(past.chartNumber, '');
  const afternoon = h.post('nurse_history', { startDate: '2026-11-24', endDate: '2026-11-24', clinicSession: 'T' });
  assert.equal(afternoon.rows.length, 1); assert.equal(afternoon.rows[0].chartNumber, '0000099');
  assert.equal(h.post('nurse_history', { startDate: '2026-11-24', endDate: today }).error.code, 'CHART_REQUIRED_FOR_HISTORY');
  assert.equal(h.post('nurse_history', { startDate: '2026-11-26', endDate: '2026-11-26' }).error.code, 'FUTURE_DATE');
});

test('session filtering precedes pagination and the cursor is bound to the selected session', () => {
  const h = fixture([rowAt('09:00:00'), rowAt('10:00:00'), rowAt('16:00:00'), rowAt('20:00:00'), rowAt('07:00:00')]);
  const first = history(h, { clinicSession: 'S', limit: '1' });
  assert.equal(first.rows.length, 1); assert.equal(first.rows[0].clinicSession, 'S'); assert.equal(first.hasMore, true);
  const second = history(h, { clinicSession: 'S', limit: '1', cursor: first.nextCursor });
  assert.equal(second.rows.length, 1); assert.equal(second.rows[0].clinicSession, 'S'); assert.equal(second.hasMore, false);
  assert.notEqual(second.rows[0].timestamp, first.rows[0].timestamp);
  assert.equal(history(h, { clinicSession: 'T', cursor: first.nextCursor }).error.code, 'INVALID_PAGINATION');
  assert.equal(history(h, { clinicSession: 'unassigned' }).rows.length, 1);
  for (const clinicSession of ['s', 'morning', 'native:S', 'S OR 1=1']) {
    assert.equal(history(h, { clinicSession }).error.code, 'INVALID_CLINIC_SESSION');
  }
  assert.equal(h.writes.length, 0);
});

test('independent PF with unknown measurement time is unassigned despite current entry clock', () => {
  const h = fixture([rowAt('09:00:00')]);
  const saved = record(h);
  assert.equal(saved.ok, true); assert.equal(saved.record.recordedAt, fixedNow); assert.equal(saved.record.measuredAt, '');
  const all = history(h);
  assert.equal(all.pfRecords[0].clinicSession, 'unassigned'); assert.equal(all.pfRecords[0].sessionSource, 'not_available');
  assert.equal(history(h, { clinicSession: 'S' }).pfRecords.length, 0);
  assert.equal(history(h, { clinicSession: 'unassigned' }).pfRecords.length, 1);
  assert.equal(h.sheets.get('ACT Records').rows[1][15], '');
});

test('an explicit validated measurement timestamp is distinct from entry-time fallback', () => {
  const h = fixture([]);
  assert.equal(record(h, { measuredAt: today + 'T09:00:00+08:00' }).ok, true);
  const event = history(h, { clinicSession: 'S' }).pfRecords[0];
  assert.equal(event.clinicSession, 'S'); assert.equal(event.sessionSource, 'measurement_timestamp');
  assert.equal(event.sessionFallback, false); assert.equal(event.measuredAt, '2026-11-25T01:00:00.000Z');
});

test('selected ACT anchors keep unknown-time PF in S/T/U context without inventing measurement time', () => {
  for (const [clock, session] of [['09:00:00', 'S'], ['16:00:00', 'T'], ['20:00:00', 'U']]) {
    const h = fixture([rowAt(clock)]); const selected = history(h).rows[0];
    const beforeACT = JSON.stringify(h.sheets.get('ACT Records').rows);
    const saved = record(h, { linkedActIdentity: selected.stableKey });
    assert.equal(saved.ok, true); assert.equal(saved.record.measuredAt, '');
    const result = history(h, { clinicSession: session });
    assert.equal(result.pfRecords.length, 1); assert.equal(result.pfRecords[0].clinicSession, session);
    assert.equal(result.pfRecords[0].sessionSource, 'linked_questionnaire_timestamp');
    assert.equal(result.pfRecords[0].sessionFallback, true);
    assert.equal(result.pfRecords[0].linkedActIdentity, selected.stableKey);
    assert.equal(result.pfRecords[0].assessmentDate, today);
    assert.equal(JSON.stringify(h.sheets.get('ACT Records').rows), beforeACT);
    assert.equal(h.emails.length + h.propertyWrites.length, 0);
  }
});

test('wrong-patient, wrong-date, stale and ambiguous new ACT anchors reject without a PF append', () => {
  for (const variant of ['chart', 'date', 'stale', 'ambiguous']) {
    const h = fixture([rowAt('09:00:00')]); const selected = history(h).rows[0];
    const values = { linkedActIdentity: selected.stableKey };
    if (variant === 'chart') values.chartNumber = '99';
    if (variant === 'date') values.assessmentDate = '2026-11-24';
    if (variant === 'stale') h.sheets.get('ACT Records').rows[1][3] = 'Changed synthetic name';
    if (variant === 'ambiguous') h.sheets.get('ACT Records').rows.push(h.sheets.get('ACT Records').rows[1].slice());
    assert.equal(record(h, values).error.code, 'ACT_LINK_CHANGED', variant);
    assert.equal(h.writes.length, 0, variant); assert.equal(h.sheets.get('PF Records').rows.length, 1);
  }
});

test('saved anchored retry remains idempotent after ACT changes or is removed', () => {
  for (const removed of [false, true]) {
    const h = fixture([rowAt('09:00:00')]); const selected = history(h).rows[0];
    const values = { linkedActIdentity: selected.stableKey };
    const first = record(h, values); const writes = h.writes.length;
    if (removed) h.sheets.get('ACT Records').rows.pop();
    else h.sheets.get('ACT Records').rows[1][3] = 'Changed synthetic name';
    const retry = record(h, values);
    assert.equal(retry.ok, true); assert.equal(retry.duplicate, true); assert.equal(retry.record.eventId, first.record.eventId);
    assert.equal(h.writes.length, writes);
    const retained = history(h, { clinicSession: 'unassigned' }).pfRecords[0];
    assert.equal(retained.linkedActIdentity, selected.stableKey); assert.equal(retained.measuredAt, '');
    assert.equal(retained.sessionSource, 'not_available');
  }
});

test('an actual measurement clock takes priority over a selected questionnaire group', () => {
  const h = fixture([rowAt('16:00:00')]); const selected = history(h).rows[0];
  const saved = record(h, { linkedActIdentity: selected.stableKey, measuredAt: today + 'T09:00:00+08:00' });
  assert.equal(saved.ok, true);
  assert.equal(history(h, { clinicSession: 'T' }).pfRecords.length, 0);
  assert.equal(history(h, { clinicSession: 'S' }).pfRecords[0].sessionSource, 'measurement_timestamp');
});
