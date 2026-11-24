'use strict';
// Only synthetic adapters. Never contacts the clinic, Google, or a database.
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
const { Bridge, adapters, exportSQL, committedSQL, patientSQL } = require(path.join(root, 'bridge', 'pf_bridge.cjs'));
const { harness } = require(path.join(root, 'tests', 'backend-history-contract.cjs'));
const copy = value => JSON.parse(JSON.stringify(value));
const probe = () => ({ kind: 'probe' });
const eventA = '11111111-1111-4111-8111-111111111111';
const eventB = '22222222-2222-4222-8222-222222222222';
const outgoing = id => ({ kind: 'event', id, kcstmr: '0000123', assessment_date: '2026-11-24',
  score: 95, source: 'manual', entered_by: 'tray_panel', created_at: '2026-11-24 09:00:00' });

function setup(options = {}) {
  const backend = harness();
  const pending = options.mobileEvents || [];
  for (const item of pending) {
    const result = backend.mobile({ eventId: item.id, chartNumber: item.chart || '123',
      assessmentDate: '2026-11-24', peakFlow: String(item.value || 95), confirmAdditional: 'yes' });
    assert.equal(result.ok, true);
  }
  const committed = new Map(), posts = [], cloudCalls = [], queries = [];
  let state = options.state || null, saveFailures = options.saveFailures || 0;
  let ackFailures = options.ackFailures || 0, nextId = 500;
  const overlay = {
    async query(sql) {
      queries.push(sql);
      if (sql.includes("entered_by='act_pf_sync:")) {
        if (options.missingCommitProbe) return { rows: [] };
        const id = sql.match(/entered_by='act_pf_sync:([^']+)'/)[1];
        return { rows: [probe(), ...(committed.has(id) ? [copy(committed.get(id))] : [])] };
      }
      if (sql.includes('FROM patients')) {
        if (options.missingPatientProbe) return { rows: [] };
        const chart = sql.match(/KCSTMR='(\d+)'/)[1];
        const patients = options.unknownPatient ? [] : [{ kind: 'patient', kcstmr: chart }];
        if (options.ambiguousPatient) patients.push({ kind: 'patient', kcstmr: chart });
        return { rows: [probe(), ...patients] };
      }
      if (options.missingExportProbe) return { rows: [] };
      const cursor = Number(sql.match(/AND id>(\d+)/)?.[1] || 0);
      return { rows: [probe(), ...(options.outgoing || []).filter(x => x.id > cursor).map(copy)] };
    },
    async postPF(payload) {
      posts.push(copy(payload));
      if (options.queuedOnly) return { ok: true, type: 'PF', score: payload.value_lpm };
      const id = payload.entered_by.slice('act_pf_sync:'.length);
      if (!committed.has(id)) committed.set(id, { kind: 'event', id: nextId++, kcstmr: payload.kcstmr,
        assessment_date: payload.assessment_date, score: payload.value_lpm, source: 'manual',
        entered_by: payload.entered_by, created_at: '2026-11-24 10:20:00' });
      if (options.lostEnqueueResponse) throw Error('synthetic response lost after commit');
      return { ok: true };
    }
  };
  const cloud = {
    async call(mode, fields) {
      cloudCalls.push({ mode, fields: copy(fields) });
      if (mode === 'pf_ack' && ackFailures) {
        ackFailures--;
        if (options.ackLostAfterCommit) backend.post(mode, fields);
        throw Error('synthetic ack response lost');
      }
      return backend.post(mode, fields);
    }
  };
  const store = {
    async load() { return state && copy(state); },
    async save(value) { if (saveFailures && value.outCursor > 0) { saveFailures--; throw Error('synthetic checkpoint failure'); } state = copy(value); }
  };
  const bridge = (overrides = {}) => new Bridge({ overlay, cloud, store, clock: () => new Date('2026-11-24T02:20:00Z'),
    sleep: async () => {}, commitAttempts: 2, ...overrides });
  return { backend, overlay, cloud, store, bridge, committed, posts, cloudCalls, queries,
    get state() { return copy(state); } };
}

test('actual backend and bridge agree on outbound synced and inbound pending contracts', async () => {
  const h = setup({ outgoing: [outgoing(42)], mobileEvents: [{ id: eventA }] });
  const result = await h.bridge().cycle();
  assert.equal(result.errorCode, null);
  assert.equal(result.exported, 1); assert.equal(result.imported, 1);
  const records = h.backend.sheets.get('PF Records').rows;
  assert.equal(records.length, 3);
  assert.equal(records.filter(r => r[7] === 'overlay_manual' && r[10] === 'synced').length, 1);
  assert.equal(records.filter(r => r[7] === 'mobile_pf' && r[10] === 'synced').length, 1);
  assert.equal(h.backend.sheets.get('ACT Records').rows.length, 1);
  assert.equal(h.backend.emails.length, 0);
});

test('queued-only HTTP success cannot acknowledge a clinical commit', async () => {
  const h = setup({ mobileEvents: [{ id: eventA }], queuedOnly: true });
  const result = await h.bridge().cycle();
  assert.equal(result.errorCode, 'OVERLAY_COMMIT_PENDING');
  assert.equal(result.imported, 0);
  assert.equal(h.cloudCalls.filter(x => x.mode === 'pf_ack').length, 0);
  assert.equal(h.backend.post('pf_events').events[0].syncStatus, 'pending');
});

test('restart after committed clinical write and failed ack reuses that committed event', async () => {
  const h = setup({ mobileEvents: [{ id: eventA }], ackFailures: 1 });
  assert.notEqual((await h.bridge().cycle()).errorCode, null);
  assert.equal(h.posts.length, 1);
  assert.equal(h.backend.post('pf_events').events.length, 1);
  assert.equal((await h.bridge().cycle()).errorCode, null);
  assert.equal(h.posts.length, 1, 'no second clinical insertion after restart');
  assert.equal(h.backend.post('pf_events').events.length, 0);
});

test('lost ack response after cloud commit is harmless on restart', async () => {
  const h = setup({ mobileEvents: [{ id: eventA }], ackFailures: 1, ackLostAfterCommit: true });
  assert.notEqual((await h.bridge().cycle()).errorCode, null);
  assert.equal(h.backend.post('pf_events').events.length, 0);
  assert.equal((await h.bridge().cycle()).errorCode, null);
  assert.equal(h.posts.length, 1);
});

test('lost enqueue response is resolved by readback instead of a duplicate write', async () => {
  const h = setup({ mobileEvents: [{ id: eventA }], lostEnqueueResponse: true });
  const result = await h.bridge().cycle();
  assert.equal(result.errorCode, null); assert.equal(result.imported, 1); assert.equal(h.posts.length, 1);
});

test('missing patient sentinel does not falsely reject patient or write a clinical record', async () => {
  const h = setup({ mobileEvents: [{ id: eventA }], missingPatientProbe: true });
  assert.equal((await h.bridge().cycle()).errorCode, 'DB_READ_UNCONFIRMED');
  assert.equal(h.posts.length, 0);
  assert.equal(h.cloudCalls.filter(x => x.mode === 'pf_ack').length, 0);
  assert.equal(h.backend.post('pf_events').events.length, 1);
});

test('confirmed unknown patient gets rejected_patient without clinical write or fake commit', async () => {
  const h = setup({ mobileEvents: [{ id: eventA }], unknownPatient: true });
  const result = await h.bridge().cycle();
  assert.equal(result.errorCode, null); assert.equal(result.rejected, 1); assert.equal(h.posts.length, 0);
  const row = h.backend.sheets.get('PF Records').rows[1];
  assert.equal(row[10], 'rejected_patient'); assert.equal(row[9], ''); assert.equal(row[11], '');
});

test('two exact canonical patient matches are terminally rejected without selecting either', async () => {
  const h = setup({ mobileEvents: [{ id: eventA }], ambiguousPatient: true });
  const result = await h.bridge().cycle();
  assert.equal(result.errorCode, null);
  assert.equal(result.rejected, 1);
  assert.equal(result.imported, 0);
  assert.equal(h.posts.length, 0, 'no clinical write when patient identity is ambiguous');
  const row = h.backend.sheets.get('PF Records').rows[1];
  assert.equal(row[10], 'rejected_patient');
  assert.equal(row[16], 'PATIENT_AMBIGUOUS');
  assert.equal(row[9], '');
  assert.equal(row[11], '');
  assert.equal(h.backend.post('pf_events').events.length, 0, 'terminal rejection does not loop forever');
});

test('an earlier failed inbound event cannot be skipped by a later event or cursor', async () => {
  const h = setup({ mobileEvents: [{ id: eventA }, { id: eventB, chart: '124' }], ackFailures: 1 });
  assert.notEqual((await h.bridge().cycle()).errorCode, null);
  assert.deepEqual(h.posts.map(x => x.entered_by), ['act_pf_sync:' + eventA]);
  assert.equal(h.backend.post('pf_events', { afterRow: '99999' }).events.length, 2);
  assert.equal((await h.bridge().cycle()).errorCode, null);
  assert.deepEqual(h.posts.map(x => x.entered_by), ['act_pf_sync:' + eventA, 'act_pf_sync:' + eventB]);
});

test('checkpoint failure after outbound cloud write retries stable source id without another row', async () => {
  const h = setup({ outgoing: [outgoing(42)], saveFailures: 1 });
  assert.notEqual((await h.bridge().cycle()).errorCode, null);
  assert.equal(h.state.outCursor, 0);
  assert.equal(h.backend.sheets.get('PF Records').rows.length, 2, 'cloud commit occurred before checkpoint failure');
  assert.equal((await h.bridge().cycle()).errorCode, null);
  assert.equal(h.state.outCursor, 42);
  assert.equal(h.backend.sheets.get('PF Records').rows.length, 2);
});

test('outbound page is validated before checkpoint can jump over an out-of-order event', async () => {
  const h = setup({ outgoing: [outgoing(43), outgoing(42)] });
  const result = await h.bridge().cycle();
  assert.equal(result.errorCode, 'OUTBOUND_ORDER_INVALID');
  assert.equal(h.state.outCursor, 0);
  assert.equal(h.cloudCalls.filter(x => x.mode === 'pf_sync').length, 0);
});

test('exported mobile-origin echo is rejected before cloud write', async () => {
  const echo = { ...outgoing(42), entered_by: 'act_pf_sync:' + eventA };
  const h = setup({ outgoing: [echo] });
  assert.equal((await h.bridge().cycle()).errorCode, 'OUTBOUND_SOURCE_INVALID');
  assert.equal(h.cloudCalls.length, 0); assert.equal(h.state.outCursor, 0);
});

test('SQL builder rejects injection in chart, event id, cursor and date', () => {
  assert.throws(() => patientSQL("123' OR 1=1--"));
  assert.throws(() => committedSQL("id';DROP_TABLE"));
  assert.throws(() => exportSQL('0 OR 1=1', '2026-11-24'));
  assert.throws(() => exportSQL(0, "2026-11-24';--"));
  assert.throws(() => exportSQL(0, '2026-02-30'));
});

test('missing export or commit sentinel cannot become success or an unknown-patient rejection', async () => {
  for (const option of ['missingExportProbe', 'missingCommitProbe']) {
    const h = setup({ mobileEvents: [{ id: eventA }], [option]: true });
    assert.equal((await h.bridge().cycle()).errorCode, 'DB_READ_UNCONFIRMED');
    assert.equal(h.posts.length, 0);
    assert.equal(h.cloudCalls.filter(x => x.mode === 'pf_ack').length, 0);
    assert.equal(h.state.outCursor, 0);
  }
});

test('changed outbound payload for an existing stable source event is not overwritten or checkpointed', async () => {
  const h = setup({ outgoing: [{ ...outgoing(42), score: 100 }] });
  assert.equal(h.backend.overlay({ chartNumber: '123', assessmentDate: '2026-11-24',
    recordedAt: '2026-11-24T09:00:00+08:00' }).ok, true);
  assert.notEqual((await h.bridge().cycle()).errorCode, null);
  assert.equal(h.state.outCursor, 0);
  assert.equal(h.backend.sheets.get('PF Records').rows.length, 2);
  assert.equal(h.backend.sheets.get('PF Records').rows[1][2], 95);
});

test('a committed event with a mismatched clinical value cannot be acknowledged or reposted', async () => {
  const h = setup({ mobileEvents: [{ id: eventA }] });
  h.committed.set(eventA, { ...outgoing(500), score: 100, entered_by: 'act_pf_sync:' + eventA });
  assert.equal((await h.bridge().cycle()).errorCode, 'COMMITTED_PAYLOAD_CONFLICT');
  assert.equal(h.posts.length, 0);
  assert.equal(h.cloudCalls.filter(x => x.mode === 'pf_ack').length, 0);
});

test('a malformed pending event id is rejected before any clinical SQL interpolation', async () => {
  const h = setup();
  const original = h.cloud.call;
  h.cloud.call = async (mode, fields) => mode === 'pf_events' ? { ok: true, events: [{
    source: 'mobile_pf', syncStatus: 'pending', eventId: "bad';DROP_X", chartNumber: '123',
    peakFlow: 95, assessmentDate: '2026-11-24'
  }] } : original(mode, fields);
  assert.equal((await h.bridge().cycle()).errorCode, 'INVALID_EVENT_ID');
  assert.equal(h.queries.length, 1, 'only the initial fixed export query ran');
  assert.equal(h.posts.length, 0);
});

test('corrupt checkpoint state fails closed before any query or automatic replay', async () => {
  const h = setup({ state: { schemaVersion: 1, outCursor: '0 OR 1=1' } });
  assert.equal((await h.bridge().cycle()).errorCode, 'STATE_INVALID');
  assert.equal(h.queries.length + h.posts.length + h.cloudCalls.length, 0);
});

test('worker configuration refuses a non-loopback Overlay or a missing private sync credential', () => {
  const config = { overlayBase: 'http://127.0.0.1:8080/',
    scriptUrl: 'https://script.google.com/macros/s/SYNTHETIC_ONLY/exec', syncKey: 'a'.repeat(40) };
  assert.throws(() => adapters({ ...config, overlayBase: 'http://example.invalid:8080/' }));
  assert.throws(() => adapters({ ...config, overlayBase: 'http://127.0.0.1:8080/elsewhere' }));
  assert.throws(() => adapters({ ...config, syncKey: '' }));
  assert.throws(() => adapters({ ...config, syncKey: 'short' }));
});

test('desktop bootstrap selects entry time, preserving a newly entered old measurement date', async () => {
  const sql = exportSQL(0, '2026-10-25');
  assert.match(sql, /created_at>='2026-10-25'::TIMESTAMP/);
  assert.doesNotMatch(sql, /assessment_date>=/);
  const h = setup({ outgoing: [{ ...outgoing(42), assessment_date: '2020-01-02' }] });
  assert.equal((await h.bridge().cycle()).errorCode, null);
  assert.equal(h.backend.sheets.get('PF Records').rows[1][3], '2020-01-02');
});

test('persisted bootstrap date does not slide forward after a long worker downtime', async () => {
  const h = setup();
  assert.equal((await h.bridge().cycle()).errorCode, null);
  assert.equal(h.state.bootstrapSince, '2026-10-25');
  assert.equal((await h.bridge({ clock: () => new Date('2027-02-24T02:20:00Z') }).cycle()).errorCode, null);
  assert.match(h.queries.at(-1), /created_at>='2026-10-25'::TIMESTAMP/);
  assert.equal(h.state.bootstrapSince, '2026-10-25');
});

test('same stable source id with changed facts is rejected and checkpoint stays before it', async () => {
  const h = setup({ outgoing: [outgoing(42)] });
  assert.equal((await h.bridge().cycle()).errorCode, null);
  h.store.load = async () => null; h.overlay.query = async sql => ({ rows: [probe(), { ...outgoing(42), score: 100 }] });
  const result = await h.bridge().cycle();
  assert.equal(result.errorCode, 'CLOUD_ACK_MISMATCH');
  assert.equal(h.backend.sheets.get('PF Records').rows.length, 2);
  assert.equal(h.backend.sheets.get('PF Records').rows[1][2], 95);
});
