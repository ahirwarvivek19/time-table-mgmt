(function () {
  'use strict';
  // Built by tools/preview/build.js. Every value from the data is inserted as text.
  var data = JSON.parse(document.getElementById('tt-data').textContent);
  var TABS = ['class', 'teacher', 'day', 'grid', 'conflicts'];
  var COLOUR = /^#[0-9a-f]{3,8}$/i;
  var CLASH_BG = '#fdedec', COMBINED_BG = '#f5eef8';

  function $(id) { return document.getElementById(id); }
  function el(tag, text, cls, attrs) {
    var n = document.createElement(tag);
    if (text !== undefined && text !== null) n.textContent = text;
    if (cls) n.className = cls;
    if (attrs) Object.keys(attrs).forEach(function (k) { n.setAttribute(k, attrs[k]); });
    return n;
  }
  function clear(n) { while (n.firstChild) n.removeChild(n.firstChild); }
  function paint(td, bg) { if (COLOUR.test(bg || '')) td.style.backgroundColor = bg; }
  function short(day) { return day.slice(0, 3); }

  function headRow(first) {
    var thead = el('thead'), tr = el('tr');
    tr.appendChild(el('th', first, null, { scope: 'col' }));
    data.periods.forEach(function (p) {
      var th = el('th', p.label, null, { scope: 'col' });
      if (p.time) th.appendChild(el('small', p.time));
      tr.appendChild(th);
    });
    thead.appendChild(tr);
    return thead;
  }

  function caption(table, title, detail) {
    var c = el('caption', title);
    if (detail) c.appendChild(el('small', detail));
    table.appendChild(c);
  }

  function tag(td, kind) { td.appendChild(el('span', kind, 'tag ' + kind)); }

  /** "Split class — …", "Taught together with …", "⚠ Clash: …" lines of a Class_View note. */
  function noteTags(note) {
    var out = [];
    if (/Split class/.test(note)) out.push('split');
    if (/Taught together/.test(note)) out.push('combined');
    if (/Clash/.test(note)) out.push('clash');
    return out;
  }

  function fill(select, names, groups) {
    clear(select);
    if (groups) {
      groups.forEach(function (g) {
        if (!g.names.length) return;
        var og = el('optgroup', null, null, { label: g.label });
        g.names.forEach(function (n) { og.appendChild(el('option', n, null, { value: n })); });
        select.appendChild(og);
      });
    } else {
      names.forEach(function (n) { select.appendChild(el('option', n, null, { value: n })); });
    }
  }

  // ── Class ────────────────────────────────────────────────────────────────
  function renderClass() {
    var name = $('sel-class').value;
    var c = data.classes.filter(function (x) { return x.name === name; })[0];
    var table = $('tbl-class');
    clear(table);
    if (!c) return;
    caption(table, 'Class ' + name);
    table.appendChild(headRow('Day'));
    var tbody = el('tbody'), notes = [];
    c.days.forEach(function (cells, i) {
      var tr = el('tr');
      tr.appendChild(el('th', data.days[i], null, { scope: 'row' }));
      cells.forEach(function (cell, p) {
        var subject = cell[0], teacher = cell[1], bg = cell[2], note = cell[3];
        var td = el('td');
        if (!subject && !teacher) { td.textContent = 'Free'; td.className = 'free'; }
        else {
          td.appendChild(el('div', subject, 's'));
          if (teacher) td.appendChild(el('div', teacher, 't'));
        }
        paint(td, bg);
        noteTags(note).forEach(function (k) { tag(td, k); });
        if (note) {
          td.title = note;
          notes.push(short(data.days[i]) + ' · ' + data.periods[p].label + ' — ' + note);
        }
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);

    var box = $('notes-class');
    clear(box);
    if (notes.length) {
      box.appendChild(el('h2', 'Details of split, combined and clashing periods'));
      var ul = el('ul');
      notes.forEach(function (n) { ul.appendChild(el('li', n)); });
      box.appendChild(ul);
    }
  }

  // ── Teacher ──────────────────────────────────────────────────────────────
  function isBusy(t) { return t.cells.some(function (r) { return r.some(function (c) { return c[0] && c[0] !== 'FREE'; }); }); }

  function teacherCell(td, cell, combinedTag) {
    var value = cell[0], bg = String(cell[1] || '').toLowerCase();
    if (!value || value === 'FREE') { td.textContent = 'FREE'; td.className = 'free'; }
    else td.textContent = value;
    paint(td, cell[1]);
    if (bg === CLASH_BG) tag(td, 'clash');
    else if (combinedTag && bg === COMBINED_BG) tag(td, 'combined');
  }

  function renderTeacher() {
    var name = $('sel-teacher').value;
    var t = data.teachers.filter(function (x) { return x.name === name; })[0];
    var table = $('tbl-teacher');
    clear(table);
    if (!t) return;
    var detail = t.hours + ' periods a week' + (t.max ? ' (max ' + t.max + ')' : '') + (t.max && t.hours > t.max ? ' — over the maximum' : '');
    caption(table, name, detail);
    table.appendChild(headRow('Day'));
    var tbody = el('tbody');
    t.cells.forEach(function (row, i) {
      var tr = el('tr');
      tr.appendChild(el('th', data.days[i], null, { scope: 'row' }));
      row.forEach(function (cell) { var td = el('td'); teacherCell(td, cell, false); tr.appendChild(td); });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
  }

  // ── Day ──────────────────────────────────────────────────────────────────
  function renderDay() {
    var day = $('sel-day').value;
    var d = data.dayView.filter(function (x) { return x.day === day; })[0];
    var table = $('tbl-day');
    clear(table);
    if (!d) return;
    var busy = {};
    data.teachers.forEach(function (t) { busy[t.name] = isBusy(t); });
    var rows = $('hide-idle').checked ? d.rows.filter(function (r) { return busy[r.teacher]; }) : d.rows;
    caption(table, day, rows.length + ' teachers');
    table.appendChild(headRow('Teacher'));
    var tbody = el('tbody');
    rows.forEach(function (r) {
      var tr = el('tr');
      tr.appendChild(el('th', r.teacher, null, { scope: 'row' }));
      r.cells.forEach(function (cell) { var td = el('td'); teacherCell(td, cell, true); tr.appendChild(td); });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
  }

  // ── All classes ──────────────────────────────────────────────────────────
  function renderGrid() {
    var table = $('tbl-grid');
    clear(table);
    caption(table, 'All classes', 'entries that do not run every day show their days in [ ]');
    table.appendChild(headRow('Class'));
    var tbody = el('tbody');
    data.grid.forEach(function (row) {
      var tr = el('tr');
      tr.appendChild(el('th', row[0], null, { scope: 'row' }));
      row.slice(1).forEach(function (v) { tr.appendChild(el('td', v)); });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
  }

  // ── Conflicts ────────────────────────────────────────────────────────────
  function renderConflicts() {
    var s = data.summary, box = $('conflict-summary');
    clear(box);
    var counts = Object.keys(s.counts).map(function (k) { return k + ': ' + s.counts[k]; });
    box.appendChild(el('p', s.rowsFlagged
      ? s.rowsFlagged + ' rows in Master_Schedule are flagged by Check For Conflicts' + (counts.length ? ' — ' + counts.join(', ') : '') + '.'
      : 'Check For Conflicts finds no problems.'));
    if (s.overloaded.length) box.appendChild(el('p', 'Over their weekly maximum: ' + s.overloaded.join(', ') + '.'));
    var table = $('tbl-conflicts');
    clear(table);
    if (!data.conflicts.length) return;
    caption(table, 'Flagged rows');
    var thead = el('thead'), tr = el('tr');
    ['Day', 'Period', 'Class', 'Subject', 'Teacher', 'Clash Status'].forEach(function (h) { tr.appendChild(el('th', h, null, { scope: 'col' })); });
    thead.appendChild(tr);
    table.appendChild(thead);
    var tbody = el('tbody');
    data.conflicts.forEach(function (r) {
      var row = el('tr');
      r.forEach(function (v) { row.appendChild(el('td', v)); });
      tbody.appendChild(row);
    });
    table.appendChild(tbody);
  }

  // ── Tabs, selection in the address bar, print ─────────────────────────────
  var state = { tab: 'class' };

  function saveHash() {
    var q = new URLSearchParams();
    q.set('tab', state.tab);
    ['class', 'teacher', 'day'].forEach(function (k) { var v = $('sel-' + k).value; if (v) q.set(k, v); });
    try { history.replaceState(null, '', '#' + q.toString()); } catch (e) { /* some browsers refuse this for file:// pages */ }
  }

  function show(tab, focus) {
    state.tab = TABS.indexOf(tab) === -1 ? 'class' : tab;
    TABS.forEach(function (t) {
      var on = t === state.tab, button = $('tab-' + t);
      button.setAttribute('aria-selected', String(on));
      button.tabIndex = on ? 0 : -1;
      $('panel-' + t).hidden = !on;
      if (on && focus) button.focus();
    });
    saveHash();
  }

  TABS.forEach(function (t, i) {
    var button = $('tab-' + t);
    button.addEventListener('click', function () { show(t); });
    button.addEventListener('keydown', function (e) {
      var step = e.key === 'ArrowRight' ? 1 : e.key === 'ArrowLeft' ? -1 : 0;
      if (!step) return;
      e.preventDefault();
      show(TABS[(i + step + TABS.length) % TABS.length], true);
    });
  });
  Array.prototype.forEach.call(document.querySelectorAll('[data-print]'), function (b) {
    b.addEventListener('click', function () { window.print(); });
  });

  // ── Start ────────────────────────────────────────────────────────────────
  var s = data.summary;
  $('meta').textContent = 'Built ' + data.generated + (data.sources.length ? ' from ' + data.sources.join(' · ') : '');
  [s.classes + ' classes', s.teachers + ' teachers', s.rows + ' periods', s.splitSlots + ' split periods',
    s.combinedLectures + ' combined lectures', s.rowsFlagged + ' rows with conflicts']
    .forEach(function (x) { $('summary').appendChild(el('li', x)); });
  $('tab-conflicts').textContent = 'Conflicts (' + s.rowsFlagged + ')';

  fill($('sel-class'), data.classes.map(function (c) { return c.name; }));
  fill($('sel-teacher'), null, [
    { label: 'Teachers with periods', names: data.teachers.filter(isBusy).map(function (t) { return t.name; }) },
    { label: 'No periods in these classes', names: data.teachers.filter(function (t) { return !isBusy(t); }).map(function (t) { return t.name; }) },
  ]);
  fill($('sel-day'), data.days);

  var start = new URLSearchParams(location.hash.slice(1));
  ['class', 'teacher', 'day'].forEach(function (k) {
    var v = start.get(k), sel = $('sel-' + k);
    if (v && Array.prototype.some.call(sel.options, function (o) { return o.value === v; })) sel.value = v;
  });
  $('sel-class').addEventListener('change', function () { renderClass(); saveHash(); });
  $('sel-teacher').addEventListener('change', function () { renderTeacher(); saveHash(); });
  $('sel-day').addEventListener('change', function () { renderDay(); saveHash(); });
  $('hide-idle').addEventListener('change', renderDay);

  renderClass(); renderTeacher(); renderDay(); renderGrid(); renderConflicts();
  show(start.get('tab') || 'class');
})();
