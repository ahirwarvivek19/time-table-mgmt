/**
 * gas.js — loads the Apps Script files that define the timetable cell grammar
 * (Config.gs + ScheduleParser.gs) into Node, so the converter reads and writes
 * Subject / Teacher / Room cells with exactly the same rules as the sheet.
 */
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.resolve(__dirname, '..', '..', '..');

let cached = null;

/** @returns {{ScheduleParser: Object, CONFIG: Object}} */
function loadGrammar() {
  if (cached) return cached;
  const code = ['Config.gs', 'ScheduleParser.gs']
    .map(f => fs.readFileSync(path.join(ROOT, f), 'utf8'))
    .join('\n') + '\n;globalThis.__g = { ScheduleParser, CONFIG };';
  const ctx = vm.createContext({});
  vm.runInContext(code, ctx, { filename: 'grammar.gs' });
  // Plain copies of results: objects made inside the VM have foreign prototypes.
  const P = ctx.__g.ScheduleParser;
  const plain = x => JSON.parse(JSON.stringify(x));
  const wrap = {};
  Object.keys(P).forEach(k => {
    wrap[k] = typeof P[k] !== 'function' ? P[k] : (...a) => {
      const r = P[k].apply(P, a);
      return r === undefined ? undefined : plain(r);
    };
  });
  cached = { ScheduleParser: wrap, CONFIG: plain(ctx.__g.CONFIG) };
  return cached;
}

module.exports = { loadGrammar, ROOT };
