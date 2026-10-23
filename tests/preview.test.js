'use strict';
/**
 * The local preview (npm run preview): one self-contained page rendered from the real
 * .gs views, and a server that listens on 127.0.0.1 only.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');
const { buildPreview, startServer } = require('../tools/preview/build');

const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), 'tt-preview-'));
let built;
const build = () => {
  if (!built) { const out = path.join(tmp(), 'timetable.html'); built = { out, ...buildPreview({ out }) }; built.html = fs.readFileSync(out, 'utf8'); }
  return built;
};

test('preview page carries every class, teacher and day view as the sheet renders them', () => {
  const { data } = build();
  assert.equal(data.classes.length, 28);
  assert.equal(data.dayView.length, 6);
  assert.equal(data.grid.length, 28);
  assert.equal(data.summary.rows, 1344);
  assert.deepEqual(data.periods[0], { label: 'Period 1', time: '9:30–10:20' });
  const c11 = data.classes.find(c => c.name === '11th C');
  const [subject, teacher, bg, note] = c11.days[0][0];
  assert.equal(subject, 'Hindi (E.M) / English (H.M)');
  assert.equal(teacher, 'Ekta Gour / Rajesh Kumar Bairagi');
  assert.equal(bg, '#FEF5E7', 'split colour from Class_View');
  assert.match(note, /Split class — 2 groups[\s\S]*Taught together:\n• Hindi \(E\.M\) — Ekta Gour: with 11th B, 11th D/);
  const neha = data.teachers.find(t => t.name === 'Neha Jamidar');
  assert.deepEqual(neha.cells[0][0], ['9th A\nSanskrit', '#FDFEFE']);
  assert.ok(neha.hours > neha.max, 'load and maximum come from the Teachers sheet');
  assert.equal(data.conflicts.length, data.summary.rowsFlagged);
  assert.match(data.sources.join(), /Timetable\.xlsx/);
});

test('preview page: data cannot break out of its block, script is CSP-hashed and inserts only text', () => {
  const { html } = build();
  const json = html.match(/<script type="application\/json" id="tt-data">([\s\S]*?)<\/script>/)[1];
  assert.ok(!json.includes('<'), 'no "<" inside the JSON block');
  assert.equal(JSON.parse(json).classes.length, 28);
  const script = html.match(/<script>([\s\S]*?)<\/script>/)[1];
  assert.doesNotThrow(() => new vm.Script(script));
  const hash = crypto.createHash('sha256').update(script).digest('base64');
  assert.ok(html.includes(`script-src 'sha256-${hash}'`), 'CSP allows exactly this script');
  assert.ok(html.includes("default-src 'none'"), 'nothing else may load');
  assert.ok(!/innerHTML|outerHTML|insertAdjacentHTML|document\.write|eval\(|new Function/.test(script), 'text-only DOM writes');
  assert.ok(!/__DATA__|__SCRIPT__|__SCRIPT_HASH__/.test(html), 'every placeholder filled');
});

test('preview server listens on 127.0.0.1 only and serves nothing but the page', async () => {
  const { out } = build();
  const { server, url } = await startServer({ port: 0, file: out });
  try {
    assert.equal(server.address().address, '127.0.0.1');
    const ok = await fetch(url);
    assert.equal(ok.status, 200);
    assert.match(ok.headers.get('content-type'), /^text\/html/);
    assert.match(await ok.text(), /<title>Timetable preview<\/title>/);
    assert.equal((await fetch(url + 'timetable.html')).status, 200);
    assert.equal((await fetch(url + 'package.json')).status, 404);
    assert.equal((await fetch(url + '..%2Fpackage.json')).status, 404);
    assert.equal((await fetch(url, { method: 'POST' })).status, 405);
  } finally {
    await new Promise(resolve => server.close(resolve));
  }
});
