'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { harness: backend, syncKey, actHeader, actRow } = require('./backend-history-contract.cjs');
const { harness: frontend, response } = require('./frontend-history-harness.cjs');
const pin = '3613';
const day = '2026-11-24';
const at = (chart, date, clock) => actRow(chart, date + 'T' + clock + '+08:00');

function linked(rows, intercept) {
  const b = backend({ actRows: [actHeader, ...rows], properties: { ACT_NURSE_PIN: pin, ACT_PF_SYNC_KEY: syncKey } });
  const ui = frontend(async params => {
    const mode = params.get('mode'); const fields = Object.fromEntries(params);
    return intercept ? intercept({ mode, fields, b, response }) : response(b.post(mode, fields));
  });
  ui.app.nurseState.pin = pin; ui.app.nurseState.session = 1; ui.app.initializeNurseDates();
  return { b, ui };
}

function filter(ui, date, session = 'all') {
  ui.nodes.get('nurse-history-chart').value = '';
  ui.nodes.get('nurse-start-date').value = date;
  ui.nodes.get('nurse-end-date').value = date;
  ui.nodes.get('nurse-clinic-session').value = session;
}

async function clickPatient(ui, chart) {
  const button = ui.app.nurseState.quickButtons.find(item => item.textContent.includes(chart));
  assert.ok(button, 'ACT quick-entry button must exist for the displayed chart');
  await button.click(); return ui.app.nurseState.selectedAct;
}

test('real ACT button saves anchored PF and keeps the all-patient S overview, date, and clinical originals', async () => {
  const { b, ui } = linked([at('35292', day, '08:30:00'), at('99', day, '09:00:00'), at('88', day, '16:00:00')]);
  const originalACT = JSON.stringify(b.sheets.get('ACT Records').rows);
  filter(ui, day, 'S'); assert.equal(await ui.app.loadNurseList(), true);
  assert.equal(ui.app.nurseState.history.rows.length, 2);
  const selected = await clickPatient(ui, '0035292');
  assert.equal(ui.nodes.get('pf-chart').value, '0035292'); assert.equal(ui.nodes.get('pf-date').value, day);
  assert.equal(ui.nodes.get('pf-value').focused, true); assert.equal(ui.nodes.get('pf-entry-card').scrolled, true);
  assert.equal(ui.nodes.get('nurse-history-chart').value, '');
  ui.nodes.get('pf-value').value = '320';
  assert.equal(await ui.app.savePfRecord(), true);
  const sent = ui.calls.find(call => call.params.get('mode') === 'pf_record').params;
  assert.equal(sent.get('linkedActIdentity'), selected.stableKey);
  assert.equal(sent.get('assessmentDate'), day); assert.equal(sent.get('chartNumber'), '0035292');
  assert.equal(ui.nodes.get('nurse-history-chart').value, '');
  assert.equal(ui.nodes.get('nurse-start-date').value, day); assert.equal(ui.nodes.get('nurse-end-date').value, day);
  assert.equal(ui.nodes.get('nurse-clinic-session').value, 'S');
  assert.equal(ui.app.nurseState.history.query.chartNumber, '');
  assert.equal(ui.app.nurseState.history.rows.length, 2);
  const event = ui.app.nurseState.history.pfRecords[0];
  assert.equal(event.peakFlow, 320); assert.equal(event.clinicSession, 'S');
  assert.equal(event.sessionSource, 'linked_questionnaire_timestamp'); assert.equal(event.sessionFallback, true);
  assert.equal(event.measuredAt, ''); assert.equal(event.syncStatus, 'pending');
  assert.match(ui.nodes.get('nurse-list-container').text(), /此份 ACT 補登 PF：320/);
  assert.equal(JSON.stringify(b.sheets.get('ACT Records').rows), originalACT);
  assert.equal(b.writes.every(write => write.sheet === 'PF Records'), true);
  assert.equal(b.emails.length + b.propertyWrites.length, 0);
});

test('after a confirmed save the next ACT patient button clears the old draft without narrowing the overview', async () => {
  const { b, ui } = linked([at('35292', day, '08:30:00'), at('99', day, '09:00:00')]);
  filter(ui, day, 'S'); await ui.app.loadNurseList();
  await clickPatient(ui, '0035292'); ui.nodes.get('pf-value').value = '320';
  assert.equal(await ui.app.savePfRecord(), true);
  await clickPatient(ui, '0000099');
  assert.equal(ui.nodes.get('pf-value').value, ''); assert.equal(ui.nodes.get('pf-chart').value, '0000099');
  assert.equal(ui.nodes.get('nurse-history-chart').value, ''); assert.equal(ui.nodes.get('nurse-clinic-session').value, 'S');
  ui.nodes.get('pf-value').value = '300'; assert.equal(await ui.app.savePfRecord(), true);
  assert.equal(b.sheets.get('PF Records').rows.length, 3);
  assert.equal(ui.app.nurseState.history.pfRecords.length, 2);
  assert.equal(b.sheets.get('ACT Records').rows.slice(1).every(row => row[15] === ''), true);
  assert.equal(b.emails.length, 0);
});

test('a past single-day all-patient query uses real S/T/U filters and preserves the selected date', async () => {
  const past = '2026-11-23';
  const { b, ui } = linked([at('35292', past, '08:30:00'), at('99', past, '16:00:00'),
    at('88', past, '19:00:00'), at('77', past, '23:30:00')]);
  for (const [session, expectedChart] of [['S', '0035292'], ['T', '0000099'], ['U', '0000088'], ['unassigned', '0000077']]) {
    filter(ui, past, session); assert.equal(await ui.app.loadNurseList(), true);
    assert.equal(ui.app.nurseState.history.rows.length, 1);
    assert.equal(ui.app.nurseState.history.rows[0].chartNumber, expectedChart);
    assert.equal(ui.app.nurseState.history.rows[0].clinicSession, session);
    assert.equal(ui.nodes.get('nurse-start-date').value, past); assert.equal(ui.nodes.get('nurse-history-chart').value, '');
  }
  assert.equal(b.writes.length + b.emails.length, 0);
});

test('historical ACT quick entry keeps the past S context without substituting today or an invented measurement clock', async () => {
  const past = '2026-11-23';
  const { b, ui } = linked([at('35292', past, '08:30:00'), at('99', past, '09:00:00')]);
  filter(ui, past, 'S'); assert.equal(await ui.app.loadNurseList(), true);
  await clickPatient(ui, '0035292'); ui.nodes.get('pf-value').value = '95';
  assert.equal(await ui.app.savePfRecord(), true);
  const event = ui.app.nurseState.history.pfRecords[0];
  assert.equal(event.assessmentDate, past); assert.equal(event.recordedAt.slice(0, 10), day);
  assert.equal(event.measuredAt, ''); assert.equal(event.clinicSession, 'S');
  assert.equal(event.sessionSource, 'linked_questionnaire_timestamp');
  assert.equal(ui.nodes.get('nurse-start-date').value, past); assert.equal(ui.nodes.get('nurse-end-date').value, past);
  assert.equal(ui.nodes.get('nurse-history-chart').value, ''); assert.equal(ui.nodes.get('nurse-clinic-session').value, 'S');
  assert.equal(b.sheets.get('ACT Records').rows[1][15], ''); assert.equal(b.emails.length, 0);
});

test('a stale ACT anchor reaches the actual backend and fails without a PF row or a success message', async () => {
  const { b, ui } = linked([at('35292', day, '08:30:00')]);
  filter(ui, day, 'S'); await ui.app.loadNurseList(); await clickPatient(ui, '0035292');
  b.sheets.get('ACT Records').rows[1][3] = 'Changed synthetic fixture';
  ui.nodes.get('pf-value').value = '320';
  assert.equal(await ui.app.savePfRecord(), false);
  assert.equal(ui.app.nurseState.attempt.saved, false); assert.equal(b.sheets.get('PF Records').rows.length, 1);
  assert.match(ui.nodes.get('pf-record-status').textContent, /變更|重新選擇/);
  assert.equal(b.writes.length + b.emails.length, 0);
});

test('a patient/date mismatch on an anchored request is rejected end to end before a measurement is stored', async () => {
  for (const variant of ['chart', 'date']) {
    const { b, ui } = linked([at('35292', day, '08:30:00')], ({ mode, fields, b, response }) => {
      if (mode === 'pf_record') fields[variant === 'chart' ? 'chartNumber' : 'assessmentDate'] = variant === 'chart' ? '0000099' : '2026-11-23';
      return response(b.post(mode, fields));
    });
    filter(ui, day, 'S'); await ui.app.loadNurseList(); await clickPatient(ui, '0035292'); ui.nodes.get('pf-value').value = '320';
    assert.equal(await ui.app.savePfRecord(), false, variant);
    assert.equal(ui.app.nurseState.attempt.saved, false); assert.equal(b.writes.length, 0);
    assert.match(ui.nodes.get('pf-record-status').textContent, /變更|重新選擇/);
  }
});

test('a lost saved acknowledgment retries the same event after ACT changes and retains the original PF exactly once', async () => {
  let loseFirst = true;
  const { b, ui } = linked([at('35292', day, '08:30:00')], ({ mode, fields, b, response }) => {
    const result = b.post(mode, fields);
    if (mode === 'pf_record' && loseFirst) { loseFirst = false; throw Error('Synthetic lost acknowledgment'); }
    return response(result);
  });
  filter(ui, day, 'S'); await ui.app.loadNurseList(); await clickPatient(ui, '0035292'); ui.nodes.get('pf-value').value = '320';
  assert.equal(await ui.app.savePfRecord(), false);
  const firstId = ui.app.nurseState.attempt.eventId; assert.equal(b.sheets.get('PF Records').rows.length, 2);
  b.sheets.get('ACT Records').rows[1][3] = 'Changed after synthetic commit';
  assert.equal(await ui.app.savePfRecord(), true);
  assert.equal(ui.app.nurseState.attempt.eventId, firstId); assert.equal(b.sheets.get('PF Records').rows.length, 2);
  assert.equal(b.writes.length, 1); assert.equal(b.sheets.get('PF Records').rows[1][4], '');
  const posts = ui.calls.filter(call => call.params.get('mode') === 'pf_record');
  assert.equal(posts.length, 2); assert.equal(posts[0].params.get('eventId'), posts[1].params.get('eventId'));
  assert.equal(posts[0].params.get('linkedActIdentity'), posts[1].params.get('linkedActIdentity'));
  ui.nodes.get('nurse-clinic-session').value = 'unassigned'; assert.equal(await ui.app.loadNurseList(), true);
  assert.equal(ui.app.nurseState.history.pfRecords[0].eventId, firstId);
  assert.equal(ui.app.nurseState.history.pfRecords[0].sessionSource, 'not_available');
  assert.equal(b.sheets.get('ACT Records').rows[1][15], ''); assert.equal(b.emails.length, 0);
});
