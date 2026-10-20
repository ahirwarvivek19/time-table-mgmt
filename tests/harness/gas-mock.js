/**
 * gas-mock.js — a small in-memory stand-in for the Apps Script services this
 * project uses (SpreadsheetApp, LockService, PropertiesService, ScriptApp,
 * HtmlService). It is NOT a full emulator: it implements the calls the .gs
 * files make, stores cell values, backgrounds and data-validation lists, and
 * treats other formatting calls as no-ops. Good enough to run real menu
 * functions and onEdit handlers end to end in Node.
 */
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

function a1ToRC(a1) {
  const m = String(a1).toUpperCase().match(/^([A-Z]+)(\d+)$/);
  if (!m) throw new Error('Unsupported A1 notation: ' + a1);
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return [Number(m[2]), col];
}

class Range {
  constructor(sheet, row, col, numRows = 1, numCols = 1) {
    if (row < 1 || col < 1 || numRows < 1 || numCols < 1) {
      throw new Error(`Range out of bounds: (${row},${col},${numRows},${numCols}) on ${sheet._name}`);
    }
    Object.assign(this, { sheet, row, col, numRows, numCols });
  }
  _cells(fn) { for (let r = 0; r < this.numRows; r++) for (let c = 0; c < this.numCols; c++) fn(this.row + r, this.col + c, r, c); }
  getSheet() { return this.sheet; }
  getRow() { return this.row; }
  getColumn() { return this.col; }
  getLastRow() { return this.row + this.numRows - 1; }
  getLastColumn() { return this.col + this.numCols - 1; }
  getNumRows() { return this.numRows; }
  getNumColumns() { return this.numCols; }
  getA1Notation() { return `R${this.row}C${this.col}`; }
  getValue() { return this.sheet._get(this.row, this.col); }
  getValues() { const out = []; for (let r = 0; r < this.numRows; r++) { const row = []; for (let c = 0; c < this.numCols; c++) row.push(this.sheet._get(this.row + r, this.col + c)); out.push(row); } return out; }
  getDisplayValues() { return this.getValues().map(r => r.map(v => String(v))); }
  setValue(v) { this._cells((r, c) => this.sheet._set(r, c, v)); return this; }
  setValues(vals) {
    if (vals.length !== this.numRows || vals.some(r => r.length !== this.numCols)) {
      throw new Error(`setValues dimension mismatch: range ${this.numRows}x${this.numCols}, data ${vals.length}x${(vals[0] || []).length}`);
    }
    this._cells((r, c, i, j) => this.sheet._set(r, c, vals[i][j]));
    return this;
  }
  clearContent() { this._cells((r, c) => this.sheet._set(r, c, '')); return this; }
  clear() { this._cells((r, c) => { this.sheet._set(r, c, ''); this.sheet._bg.delete(`${r},${c}`); this.sheet._dv.delete(`${r},${c}`); }); return this; }
  setBackground(color) { this._cells((r, c) => this.sheet._bg.set(`${r},${c}`, color)); return this; }
  setBackgrounds(grid) { this._cells((r, c, i, j) => this.sheet._bg.set(`${r},${c}`, grid[i][j])); return this; }
  getBackground() { return this.sheet._bg.get(`${this.row},${this.col}`) || '#ffffff'; }
  setDataValidation(rule) { this._cells((r, c) => this.sheet._dv.set(`${r},${c}`, rule)); return this; }
  getDataValidation() { return this.sheet._dv.get(`${this.row},${this.col}`) || null; }
  clearDataValidations() { this._cells((r, c) => this.sheet._dv.delete(`${r},${c}`)); return this; }
  merge() { this.sheet._merges.push([this.row, this.col, this.numRows, this.numCols]); return this; }
  breakApart() { this.sheet._merges = []; return this; }
  setNote(n) { this.sheet._notes.set(`${this.row},${this.col}`, n); return this; }
  setNotes(grid) { this._cells((r, c, i, j) => this.sheet._notes.set(`${r},${c}`, grid[i][j])); return this; }
  getNote() { return this.sheet._notes.get(`${this.row},${this.col}`) || ''; }
}
// Formatting calls the code makes but the tests don't inspect — chainable no-ops.
['setFontWeight', 'setFontWeights', 'setFontColor', 'setFontColors', 'setFontFamily', 'setFontSize', 'setHorizontalAlignment',
  'setVerticalAlignment', 'setWrap', 'setBorder', 'setNumberFormat', 'setFontStyle', 'setFontLine', 'activate']
  .forEach(m => { Range.prototype[m] = function () { return this; }; });

class Sheet {
  constructor(ss, name) {
    Object.assign(this, { _ss: ss, _name: name, _data: [], _bg: new Map(), _dv: new Map(), _notes: new Map(), _merges: [], _hidden: false, _frozenRows: 0 });
  }
  _get(r, c) { const row = this._data[r - 1]; return row && row[c - 1] !== undefined ? row[c - 1] : ''; }
  _set(r, c, v) { while (this._data.length < r) this._data.push([]); const row = this._data[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v === undefined || v === null ? '' : v; }
  getName() { return this._name; }
  setName(n) { this._name = n; return this; }
  getLastRow() { for (let r = this._data.length; r >= 1; r--) if (this._data[r - 1].some(v => v !== '')) return r; return 0; }
  getLastColumn() { let max = 0; this._data.forEach(row => { for (let c = row.length; c >= 1; c--) if (row[c - 1] !== '') { max = Math.max(max, c); break; } }); return max; }
  getMaxRows() { return Math.max(1000, this._data.length); }
  getMaxColumns() { return 26; }
  getRange(a, b, c, d) {
    if (typeof a === 'string') { const [r, col] = a1ToRC(a); return new Range(this, r, col, 1, 1); }
    return new Range(this, a, b, c || 1, d || 1);
  }
  getDataRange() { return new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
  appendRow(values) { const r = this.getLastRow() + 1; values.forEach((v, i) => this._set(r, i + 1, v)); return this; }
  clear() { this._data = []; this._bg.clear(); this._dv.clear(); this._notes.clear(); this._merges = []; return this; }
  clearContents() { this._data = []; return this; }
  hideSheet() { this._hidden = true; return this; }
  isSheetHidden() { return this._hidden; }
  setFrozenRows(n) { this._frozenRows = n; return this; }
  getFrozenRows() { return this._frozenRows; }
  getColumnWidth() { return 100; }
  activate() { this._ss._active = this; return this; }
}
['setFrozenColumns', 'setHiddenGridlines', 'setRowHeight', 'setRowHeights', 'setColumnWidth', 'setColumnWidths', 'autoResizeColumns', 'setTabColor']
  .forEach(m => { Sheet.prototype[m] = function () { return this; }; });

class Spreadsheet {
  constructor() { this._sheets = []; this._active = null; this.toasts = []; }
  getSheets() { return this._sheets.slice(); }
  getSheetByName(n) { return this._sheets.find(s => s._name === n) || null; }
  insertSheet(n) { if (this.getSheetByName(n)) throw new Error('Sheet exists: ' + n); const s = new Sheet(this, n); this._sheets.push(s); this._active = s; return s; }
  deleteSheet(s) { this._sheets = this._sheets.filter(x => x !== s); }
  getActiveSheet() { return this._active || this._sheets[0]; }
  setActiveSheet(s) { this._active = s; return s; }
  moveActiveSheet(pos) { const s = this._active; this._sheets = this._sheets.filter(x => x !== s); this._sheets.splice(pos - 1, 0, s); }
  toast(msg, title) { this.toasts.push({ msg, title }); }
  getId() { return 'mock-spreadsheet'; }
  getSpreadsheetTimeZone() { return 'Asia/Kolkata'; }
}

function createEnvironment() {
  const ss = new Spreadsheet();
  const alerts = [];
  const prompts = []; // queue of responses for ui.prompt
  const answers = []; // queue of button answers for ui.alert(title, msg, buttons)
  const properties = {};
  const triggers = [];
  let triggerSeq = 0;

  const menuStub = () => { const m = { addItem: () => m, addSeparator: () => m, addSubMenu: () => m, addToUi: () => m }; return m; };
  const ui = {
    // alert(msg) or alert(title, msg, buttons). Queue answers for YES/NO dialogs in `answers`.
    alert: (...args) => { alerts.push(args.length > 1 ? args[1] : args[0]); return args.length > 2 ? (answers.shift() || 'YES') : 'OK'; },
    prompt: () => {
      const next = prompts.shift();
      if (next === undefined) throw new Error('Unexpected ui.prompt (no queued response)');
      return { getSelectedButton: () => (next === null ? 'CANCEL' : 'OK'), getResponseText: () => next || '' };
    },
    createMenu: menuStub,
    showModalDialog: () => {}, showSidebar: () => {},
    ButtonSet: { OK: 'OK', OK_CANCEL: 'OK_CANCEL', YES_NO: 'YES_NO', YES_NO_CANCEL: 'YES_NO_CANCEL' },
    Button: { OK: 'OK', CANCEL: 'CANCEL', YES: 'YES', NO: 'NO' },
  };

  const dvBuilder = () => {
    const rule = { list: null, allowInvalid: true };
    const b = {
      requireValueInList: (list) => { rule.list = list.slice(); return b; },
      setAllowInvalid: (v) => { rule.allowInvalid = v; return b; },
      setHelpText: () => b,
      build: () => ({ ...rule, getCriteriaValues: () => [rule.list] }),
    };
    return b;
  };

  const props = {
    getProperty: k => (k in properties ? properties[k] : null),
    setProperty: (k, v) => { properties[k] = String(v); return props; },
    deleteProperty: k => { delete properties[k]; return props; },
  };

  const triggerBuilder = (fn) => {
    const t = { fn, type: null };
    const b = {
      forSpreadsheet: () => b, onEdit: () => { t.type = 'ON_EDIT'; return b; }, onOpen: () => { t.type = 'ON_OPEN'; return b; },
      create: () => { const id = 'uid-' + (++triggerSeq); const trig = { getHandlerFunction: () => fn, getEventType: () => t.type, getUniqueId: () => id }; triggers.push(trig); return trig; },
    };
    return b;
  };

  const globals = {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss, getActive: () => ss, getUi: () => ui, newDataValidation: dvBuilder, flush: () => {},
      BorderStyle: { SOLID: 'SOLID', SOLID_MEDIUM: 'SOLID_MEDIUM', SOLID_THICK: 'SOLID_THICK', DASHED: 'DASHED' },
    },
    LockService: { getDocumentLock: () => ({ waitLock: () => {}, tryLock: () => true, releaseLock: () => {} }) },
    PropertiesService: { getDocumentProperties: () => props, getScriptProperties: () => props },
    ScriptApp: {
      getProjectTriggers: () => triggers.slice(),
      deleteTrigger: (t) => { const i = triggers.indexOf(t); if (i >= 0) triggers.splice(i, 1); },
      newTrigger: triggerBuilder,
      EventType: { ON_EDIT: 'ON_EDIT', ON_OPEN: 'ON_OPEN' },
    },
    HtmlService: {
      createTemplateFromFile: () => ({ evaluate: () => ({ setWidth() { return this; }, setHeight() { return this; }, setTitle() { return this; }, addMetaTag() { return this; } }) }),
      // (doGet) template.evaluate().setTitle().addMetaTag()
      createHtmlOutputFromFile: () => ({ setWidth() { return this; }, setHeight() { return this; }, setTitle() { return this; } }),
    },
    Logger: { log: () => {} },
    Utilities: {
      formatDate: (d, tz, fmt) => {
        const pad = n => String(n).padStart(2, '0');
        if (fmt !== 'yyyy-MM-dd') throw new Error('mock formatDate supports yyyy-MM-dd only');
        return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
      },
    },
    console,
  };

  return { ss, ui, alerts, prompts, answers, properties, triggers, globals };
}

/** Loads every deployable .gs file into one VM context, like Apps Script does. */
function loadProject(root, env) {
  const files = fs.readdirSync(root).filter(f => f.endsWith('.gs')).sort();
  const ignore = (fs.readFileSync(path.join(root, '.claspignore'), 'utf8')).split('\n').map(s => s.trim()).filter(Boolean);
  const deployed = files.filter(f => !ignore.includes(f));
  let code = '';
  const names = new Set();
  deployed.forEach(f => {
    const src = fs.readFileSync(path.join(root, f), 'utf8');
    code += `\n// ── ${f} ──\n` + src + '\n';
    for (const m of src.matchAll(/^(?:function\s+([A-Za-z_$][\w$]*)|const\s+([A-Za-z_$][\w$]*)\s*=)/gm)) names.add(m[1] || m[2]);
  });
  code += `\n;globalThis.__project = { ${[...names].join(', ')} };`;
  const context = vm.createContext({ ...env.globals });
  vm.runInContext(code, context, { filename: 'project.gs' });
  return { api: context.__project, files: deployed };
}

/** Builds an onEdit event object for a user edit of `range` (values already written). */
function editEvent(env, range, extra = {}) {
  return { source: env.ss, range, value: range.getValue(), ...extra };
}

module.exports = { createEnvironment, loadProject, editEvent };
