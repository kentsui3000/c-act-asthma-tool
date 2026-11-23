'use strict';
// Cross-layer checks: actual Apps Script functions serve actual frontend code.
// The existing test harnesses supply synthetic Google/DOM adapters only.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.resolve(__dirname, '..');

function loadAdapter(filename) {
  const noTest = () => {};
  noTest.test = noTest;
  const ctx = {
    __dirname: path.join(root, 'tests'),
    require(name) { return name === 'node:test' ? noTest : require(name); },
    console: { log() {}, error() {}, warn() {} },
    AbortController, URLSearchParams, Promise, Date, setImmediate,
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(root, 'tests', filename), 'utf8') +
    '\nglobalThis.adapter = { harness, response: typeof response === "function" ? response : null };', ctx);
  return ctx.adapter;
}

const backend = loadAdapter('backend-pf.cjs');
const frontend = loadAdapter('frontend-nurse.test.cjs');

test('real frontend and Apps Script agree on changed-key PF acknowledgment, with no questionnaire append', async () => {
  const b = backend.harness({ property: 'test-pin' });
  const f = frontend.harness((_url, options) => {
    const parameter = Object.fromEntries(new URLSearchParams(options.body));
    const output = b.ctx.doPost({ parameter });
    return Promise.resolve(frontend.response(JSON.parse(output.text)));
  });
  f.app.nurseState.pin = 'test-pin';
  f.app.nurseState.session = 1;
  assert.equal(await f.app.loadNurseList(), true);
  const rendered = f.nodes.get('nurse-list-container').children[0];
  assert.equal(rendered.children[2].textContent, '0');
  await rendered.click();
  const oldKey = f.app.nurseState.editing.rowKey;
  f.nodes.get('pf-modal-input').value = '320';
  await f.nodes.get('pf-modal-save').click();
  assert.match(f.nodes.get('pf-modal-status').textContent, /已確認儲存/);
  assert.notEqual(f.app.nurseState.editing.rowKey, oldKey);
  assert.equal(f.app.nurseState.editing.oldValue, 320);
  assert.equal(b.writes.length, 1);
  assert.equal(b.writes[0].col, 16);
  assert.equal(b.appends.length, 0);
  assert.equal(b.emails.length, 0);

  // A second legitimate edit uses the new key and still updates one sheet row.
  f.nodes.get('pf-modal-input').value = '340';
  await f.nodes.get('pf-modal-save').click();
  assert.equal(f.app.nurseState.editing.oldValue, 340);
  assert.equal(b.writes.length, 2);
  assert.equal(b.sheets.get('ACT Records').data.length, 2);
  assert.equal(b.appends.length + b.emails.length, 0);
});

test('stale second editor cannot overwrite PF changed by another client', async () => {
  const b = backend.harness({ property: 'test-pin' });
  const initial = b.post('nurse_list', { pin: 'test-pin' }).rows[0];
  const first = b.post('update_pf', { pin: 'test-pin', rowIndex: '2', rowKey: initial.rowKey, peakFlow: '320' });
  assert.equal(first.ok, true);
  assert.equal(first.previousRowKey, initial.rowKey);
  assert.notEqual(first.rowKey, initial.rowKey);
  const stale = b.post('update_pf', { pin: 'test-pin', rowIndex: '2', rowKey: initial.rowKey, peakFlow: '340' });
  assert.equal(stale.ok, false);
  assert.equal(stale.error.code, 'RECORD_CHANGED');
  assert.equal(b.writes.length, 1);
  assert.equal(b.sheets.get('ACT Records').data[1][15], 320);
});
