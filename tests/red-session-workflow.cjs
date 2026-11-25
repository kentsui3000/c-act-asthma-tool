'use strict';
// Independent red-team synthetic source and full-UI regression checks.
{
'use strict';
// Actual Apps Script source; only in-memory Google adapters and synthetic records.
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
const { harness, actHeader, actRow } = require(path.join(root, 'tests', 'backend-history-contract.cjs'));
const date = '2026-11-25';
const fixture = rows => harness({ now: '2026-11-25T01:20:00Z', actRows: [actHeader, ...rows] });
const row = (clock, chart = '123') => actRow(chart, date + 'T' + clock + '+08:00');
const history = (h, fields = {}) => h.post('nurse_history', { chartNumber: '123', startDate: date, endDate: date, ...fields });
const save = (h, fields = {}) => h.mobile({ eventId: 'red-session-pf-0001', chartNumber: '123', assessmentDate: date, ...fields });

test('one-millisecond boundary transitions and different UTC offsets use Taipei time exactly', () => {
  const h = fixture([]);
  for (const [boundary, before, after] of [['08:00:00', 'unassigned', 'S'], ['15:30:00', 'S', 'T'],
    ['18:30:00', 'T', 'U'], ['23:30:00', 'U', 'unassigned']]) {
    const point = Date.parse(date + 'T' + boundary + '+08:00');
    for (const delta of [-1, 0, 1]) {
      const value = new Date(point + delta).toISOString();
      const actual = h.ctx.historySession_(value, 'questionnaire_timestamp');
      assert.equal(actual.clinicSession, delta < 0 ? before : after, boundary + '/' + delta);
      assert.equal(actual.sessionSource, 'questionnaire_timestamp');
      assert.equal(actual.sessionFallback, true);
    }
  }
  assert.equal(h.ctx.historySession_('2026-11-25T05:30:00+05:30', 'questionnaire_timestamp').clinicSession, 'S');
  assert.equal(h.ctx.historySession_('2026-11-24T18:30:00-05:00', 'questionnaire_timestamp').clinicSession, 'unassigned');
});

test('a stray extra sheet cell cannot masquerade as an authoritative native session field', () => {
  const entry = row('09:00:00'); entry.push('U');
  const h = fixture([entry]); const result = history(h).rows[0];
  assert.equal(result.clinicSession, 'S');
  assert.equal(result.sessionSource, 'questionnaire_timestamp');
  assert.equal(result.sessionFallback, true);
  for (const field of ['nativeCode', 'nativeSessionCode', 'hisSessionCode']) assert.equal(result[field], undefined);
  assert.equal(h.writes.length, 0);
});

test('a stable ACT anchor survives a P-only change and physical sorting without redirecting to another patient', () => {
  for (const change of ['P', 'sort']) {
    const h = fixture([row('09:00:00'), row('16:00:00', '999')]);
    const selected = history(h).rows[0];
    const cells = h.sheets.get('ACT Records').rows;
    if (change === 'P') cells[1][15] = 95;
    else [cells[1], cells[2]] = [cells[2], cells[1]];
    const before = JSON.stringify(cells);
    const saved = save(h, { linkedActIdentity: selected.stableKey, peakFlow: '100', confirmAdditional: 'yes' });
    assert.equal(saved.ok, true, change);
    assert.equal(saved.record.chartNumber, '0000123');
    assert.equal(saved.record.linkedActIdentity, selected.stableKey);
    const result = history(h, { clinicSession: 'S' }).pfRecords[0];
    assert.equal(result.sessionSource, 'linked_questionnaire_timestamp');
    assert.equal(result.measuredAt, '');
    assert.equal(JSON.stringify(cells), before, 'new independent PF does not write ACT P');
    assert.equal(h.emails.length + h.propertyWrites.length, 0);
  }
});

test('an explicit off-hours measurement does not fall back to the linked morning questionnaire', () => {
  const h = fixture([row('09:00:00')]); const selected = history(h).rows[0];
  const saved = save(h, { linkedActIdentity: selected.stableKey, measuredAt: date + 'T23:30:00+08:00' });
  assert.equal(saved.ok, true);
  assert.equal(history(h, { clinicSession: 'S' }).pfRecords.length, 0);
  const result = history(h, { clinicSession: 'unassigned' }).pfRecords[0];
  assert.equal(result.sessionSource, 'measurement_timestamp');
  assert.equal(result.sessionFallback, false);
  assert.notEqual(result.measuredAt, '');
});

test('yesterday PF entered this morning remains unassigned rather than borrowing its receipt clock', () => {
  const h = fixture([actRow('123', '2026-11-24T16:00:00+08:00')]);
  const saved = save(h, { assessmentDate: '2026-11-24' });
  assert.equal(saved.ok, true);
  assert.equal(saved.record.recordedAt, '2026-11-25T01:20:00.000Z');
  const result = history(h, { startDate: '2026-11-24', endDate: '2026-11-24', clinicSession: 'unassigned' });
  assert.equal(result.pfRecords.length, 1);
  assert.equal(result.pfRecords[0].sessionSource, 'not_available');
  assert.equal(history(h, { startDate: '2026-11-24', endDate: '2026-11-24', clinicSession: 'S' }).pfRecords.length, 0);
});

test('a saved event remains idempotent if its parent later becomes ambiguous, but loses the guessed group', () => {
  const h = fixture([row('09:00:00')]); const selected = history(h).rows[0];
  const fields = { linkedActIdentity: selected.stableKey };
  const first = save(h, fields); assert.equal(first.ok, true);
  const writes = h.writes.length;
  h.sheets.get('ACT Records').rows.push(h.sheets.get('ACT Records').rows[1].slice());
  const retry = save(h, fields);
  assert.equal(retry.ok, true); assert.equal(retry.duplicate, true); assert.equal(h.writes.length, writes);
  const result = history(h, { clinicSession: 'unassigned' }).pfRecords[0];
  assert.equal(result.sessionSource, 'not_available'); assert.equal(result.measuredAt, '');
  assert.equal(result.linkedActIdentity, selected.stableKey);
});

test('mixed ACT and PF filtering happens before pagination, with no hidden rows consuming the selected session', () => {
  const h = fixture([row('09:00:00'), row('10:00:00'), row('16:00:00'), row('20:00:00')]);
  for (const [i, clock] of ['08:10:00', '09:10:00', '16:10:00', ''].entries()) {
    const result = save(h, { eventId: 'red-mixed-pf-000' + i, measuredAt: clock ? date + 'T' + clock + '+08:00' : '', confirmAdditional: 'yes' });
    assert.equal(result.ok, true);
  }
  const before = h.writes.length; const seen = []; let cursor;
  for (let page = 0; page < 10; page++) {
    const response = history(h, { clinicSession: 'S', limit: '1', ...(cursor ? { cursor } : {}) });
    assert.equal(response.ok, true);
    for (const item of response.rows.concat(response.pfRecords)) {
      assert.equal(item.clinicSession, 'S'); seen.push(item.eventId || item.stableKey);
    }
    if (!response.hasMore) break;
    cursor = response.nextCursor;
  }
  assert.equal(seen.length, 4); assert.equal(new Set(seen).size, 4);
  assert.equal(h.writes.length, before);
});

}

{
'use strict';
// Full actual frontend + Apps Script source; every adapter is synthetic and local.
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
const { harness: backend, privatePin, actHeader, actRow } = require(path.join(root, 'tests', 'backend-history-contract.cjs'));
const { harness: frontend, response, history: fakeHistory, act } = require(path.join(root, 'tests', 'frontend-history-harness.cjs'));
const date = '2026-11-25';
const row = (chart, clock, pf = '') => actRow(chart, date + 'T' + clock + '+08:00', pf);
function linked(rows, fetcher) {
  const b = backend({ now: date + 'T01:20:00Z', actRows: [actHeader, ...rows] });
  const serve = params => response(b.post(params.get('mode'), Object.fromEntries(params)));
  const ui = frontend((params, options) => fetcher ? fetcher(params, options, b, serve) : Promise.resolve(serve(params)));
  ui.advance(86400000); ui.app.nurseState.pin = privatePin; ui.app.nurseState.session = 1; ui.app.initializeNurseDates();
  return { b, ui };
}
const buttonFor = (ui, chart) => ui.app.nurseState.quickButtons.find(button => button.textContent.includes(chart));
const pfCalls = ui => ui.calls.filter(call => call.params.get('mode') === 'pf_record');

test('ACT name button prefills its chart and original date, focuses PF, and never copies or writes the existing ACT PF', async () => {
  const { b, ui } = linked([row('123', '09:00:00', 95)]);
  assert.equal(await ui.app.loadNurseList(), true);
  const before = JSON.stringify(b.sheets.get('ACT Records').rows);
  await buttonFor(ui, '0000123').click();
  assert.equal(ui.nodes.get('pf-chart').value, '0000123');
  assert.equal(ui.nodes.get('pf-date').value, date);
  assert.equal(ui.nodes.get('pf-value').value, '');
  assert.equal(ui.nodes.get('pf-value').focused, true);
  assert.equal(pfCalls(ui).length, 0);
  assert.equal(JSON.stringify(b.sheets.get('ACT Records').rows), before);
});

test('an old quick-button closure cannot redirect the form after a new patient or session search', async () => {
  const { ui } = linked([row('123', '09:00:00'), row('999', '16:00:00')]);
  assert.equal(await ui.app.loadNurseList(), true);
  const stale = buttonFor(ui, '0000123');
  ui.nodes.get('nurse-history-chart').value = '999';
  ui.nodes.get('nurse-clinic-session').value = 'T';
  assert.equal(await ui.app.loadNurseList(), true);
  const before = [ui.nodes.get('pf-chart').value, ui.nodes.get('pf-date').value];
  await stale.click();
  assert.deepEqual([ui.nodes.get('pf-chart').value, ui.nodes.get('pf-date').value], before);
  assert.equal(ui.app.nurseState.selectedAct, null);
  assert.match(ui.nodes.get('pf-record-status').textContent, /名單已更新/);
  assert.equal(pfCalls(ui).length, 0);
});

test('double-select after an offline save preserves the draft and UUID, then actual retry appends exactly once', async () => {
  let fail = true;
  const { b, ui } = linked([row('123', '09:00:00')], (params, _options, _b, serve) => {
    if (params.get('mode') === 'pf_record' && fail) { fail = false; return Promise.reject(Error('Synthetic offline')); }
    return Promise.resolve(serve(params));
  });
  assert.equal(await ui.app.loadNurseList(), true);
  const button = buttonFor(ui, '0000123'); await button.click();
  ui.nodes.get('pf-value').value = '100';
  assert.equal(await ui.app.savePfRecord(), false);
  const id = ui.app.nurseState.attempt.eventId;
  await button.click(); await button.click();
  assert.equal(ui.nodes.get('pf-value').value, '100');
  assert.equal(ui.app.nurseState.attempt.eventId, id);
  assert.equal(await ui.app.savePfRecord(), true);
  assert.equal(pfCalls(ui).length, 2);
  assert.equal(pfCalls(ui)[1].params.get('eventId'), id);
  assert.equal(b.sheets.get('PF Records').rows.length, 2);
  assert.equal(b.sheets.get('ACT Records').rows[1][15], '');
});

test('pending save blocks another patient selection, and logout discards the late acknowledgment', async () => {
  let finish;
  const { b, ui } = linked([row('123', '09:00:00'), row('999', '16:00:00')], (params, _options, _b, serve) => {
    const result = serve(params);
    return params.get('mode') === 'pf_record' ? new Promise(resolve => { finish = () => resolve(result); }) : Promise.resolve(result);
  });
  assert.equal(await ui.app.loadNurseList(), true);
  await buttonFor(ui, '0000123').click();
  ui.nodes.get('pf-value').value = '100';
  const pending = ui.app.savePfRecord();
  await buttonFor(ui, '0000999').click();
  assert.equal(ui.nodes.get('pf-chart').value, '0000123');
  assert.equal(ui.app.nurseState.selectedAct.chartNumber, '0000123');
  await ui.nodes.get('btn-nurse-back').click(); finish(); await pending;
  assert.equal(ui.app.nurseState.pin, '');
  assert.equal(ui.app.nurseState.selectedAct, null);
  assert.equal(ui.nodes.get('pf-record-status').textContent, '');
  assert.equal(ui.nodes.get('nurse-list-container').children.length, 0);
  assert.equal(pfCalls(ui).length, 1);
  assert.equal(b.sheets.get('PF Records').rows[1][1], '0000123');
  assert.equal(b.emails.length, 0);
});

test('wrong outer or row session metadata is rejected instead of displaying another session as selected', async () => {
  for (const wrong of ['outer', 'row']) {
    const ui = frontend(params => Promise.resolve(response(fakeHistory(params, {
      clinicSession: wrong === 'outer' ? 'T' : 'S', rows: [{ ...act, clinicSession: wrong === 'row' ? 'T' : 'S' }]
    }))));
    ui.app.nurseState.pin = privatePin; ui.app.nurseState.session = 1; ui.app.initializeNurseDates();
    ui.nodes.get('nurse-clinic-session').value = 'S';
    assert.equal(await ui.app.loadNurseList(), false, wrong);
    assert.equal(ui.nodes.get('nurse-list-container').children[0].attributes.role, 'alert');
    assert.equal(pfCalls(ui).length, 0);
  }
});

test('manually changing the prefilled chart clears the old ACT anchor before constructing an independent PF', async () => {
  const { ui } = linked([row('123', '09:00:00'), row('999', '16:00:00')]);
  assert.equal(await ui.app.loadNurseList(), true); await buttonFor(ui, '0000123').click();
  ui.nodes.get('pf-chart').value = '999'; await ui.nodes.get('pf-chart').input();
  ui.nodes.get('pf-value').value = '100';
  assert.equal(ui.app.nurseState.selectedAct, null);
  const payload = ui.app.pfPayload();
  assert.equal(payload.chartNumber, '0000999'); assert.equal(payload.linkedActIdentity, undefined);
  assert.doesNotMatch(ui.nodes.get('pf-selected-context').textContent, /已選擇：/);
  assert.equal(pfCalls(ui).length, 0);
});

test('a late history response from the previous clinic selection cannot replace the newer session list', async () => {
  let releaseMorning;
  const { ui } = linked([row('123', '09:00:00'), row('999', '16:00:00')], (params, _options, _b, serve) => {
    const result = serve(params);
    return params.get('mode') === 'nurse_history' && params.get('clinicSession') === 'S' ?
      new Promise(resolve => { releaseMorning = () => resolve(result); }) : Promise.resolve(result);
  });
  ui.nodes.get('nurse-clinic-session').value = 'S'; const morning = ui.app.loadNurseList();
  ui.nodes.get('nurse-clinic-session').value = 'T'; assert.equal(await ui.app.loadNurseList(), true);
  releaseMorning(); await morning;
  assert.equal(ui.app.nurseState.history.query.clinicSession, 'T');
  assert.match(ui.nodes.get('nurse-list-container').text(), /0000999/);
  assert.doesNotMatch(ui.nodes.get('nurse-list-container').text(), /0000123/);
  assert.equal(pfCalls(ui).length, 0);
});

test('changing clinic during a quick-save keeps the new search when the old measurement acknowledgment arrives', async () => {
  let finish;
  const { ui } = linked([row('123', '16:00:00'), row('999', '09:00:00')], (params, _options, _b, serve) => {
    const result = serve(params);
    return params.get('mode') === 'pf_record' ? new Promise(resolve => { finish = () => resolve(result); }) : Promise.resolve(result);
  });
  ui.nodes.get('nurse-clinic-session').value = 'T'; assert.equal(await ui.app.loadNurseList(), true);
  await buttonFor(ui, '0000123').click(); ui.nodes.get('pf-value').value = '100';
  const saving = ui.app.savePfRecord();
  ui.nodes.get('nurse-clinic-session').value = 'S'; assert.equal(await ui.app.loadNurseList(), true);
  finish(); assert.equal(await saving, true);
  assert.equal(ui.nodes.get('nurse-clinic-session').value, 'S');
  assert.equal(ui.app.nurseState.history.query.clinicSession, 'S');
  assert.match(ui.nodes.get('nurse-list-container').text(), /0000999/);
  assert.doesNotMatch(ui.nodes.get('nurse-list-container').text(), /0000123/);
  assert.equal(pfCalls(ui).length, 1);
  assert.equal(pfCalls(ui)[0].params.get('chartNumber'), '0000123');
});

test('switching to another ACT does not discard an unsaved PF draft or silently change its patient', async () => {
  const { ui } = linked([row('123', '09:00:00'), row('999', '16:00:00')]);
  assert.equal(await ui.app.loadNurseList(), true);
  await buttonFor(ui, '0000123').click(); ui.nodes.get('pf-value').value = '100';
  const anchor = ui.app.nurseState.selectedAct.stableKey;
  await buttonFor(ui, '0000999').click();
  assert.equal(ui.nodes.get('pf-chart').value, '0000123');
  assert.equal(ui.nodes.get('pf-value').value, '100');
  assert.equal(ui.app.nurseState.selectedAct.stableKey, anchor);
  assert.match(ui.nodes.get('pf-record-status').textContent, /尚未儲存/);
  assert.equal(pfCalls(ui).length, 0);
});

}
