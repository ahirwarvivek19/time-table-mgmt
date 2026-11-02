'use strict';
/**
 * What `clasp push` uploads must load in Apps Script. One file that fails at load time
 * (a Node script with require(), a leftover old file) stops every function, including
 * onOpen, so the Timetable menu disappears.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.resolve(__dirname, '..');

/** Root files clasp would push: not matched by .claspignore (simple globs only). */
function pushedFiles() {
  const globs = fs.readFileSync(path.join(ROOT, '.claspignore'), 'utf8').split('\n')
    .map(l => l.trim()).filter(l => l && !l.startsWith('#'))
    .map(g => new RegExp('^' + g.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*\*/g, '.*').replace(/(?<!\.)\*/g, '[^/]*') + '$'));
  return fs.readdirSync(ROOT, { withFileTypes: true })
    .filter(d => d.isFile() && !d.name.startsWith('.'))
    .map(d => d.name)
    .filter(n => !globs.some(g => g.test(n)));
}

test('clasp pushes only Apps Script files: .gs, .html and appsscript.json', () => {
  const odd = pushedFiles().filter(n => !/\.(gs|html)$/.test(n) && n !== 'appsscript.json');
  assert.deepEqual(odd, [], 'would be uploaded to Apps Script');
  assert.ok(!pushedFiles().includes('gen.js'));
});

test('every pushed .gs file loads together in a bare V8 context, and onOpen builds the menu', () => {
  const ctx = vm.createContext({});
  pushedFiles().filter(n => n.endsWith('.gs')).forEach(n => {
    assert.doesNotThrow(() => vm.runInContext(fs.readFileSync(path.join(ROOT, n), 'utf8'), ctx, { filename: n }), n);
  });
  assert.equal(typeof vm.runInContext('onOpen', ctx), 'function');
  // Removed files must stay removed: they defined menu functions of their own
  ['TeacherAvailabilityGrid.gs', 'GeneratorAlgorithm.gs', 'MultiSelectManager.gs'].forEach(n =>
    assert.ok(!fs.existsSync(path.join(ROOT, n)), n + ' is back'));
});
