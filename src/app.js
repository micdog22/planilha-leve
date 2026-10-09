// Interface da planilha: grade, seleção, edição na célula e na barra de fórmulas,
// modo de apontar referências com o mouse ou as setas, área de transferência,
// formatos, CSV e salvamento automático no navegador.

import {
  Sheet, copyBlock, pasteChanges, fillDownChanges, fillRightChanges, clearChanges, selectionSummary, rowOf, colOf,
} from './engine.js';
import { cellName, colName, parseCellName, cycleReference, canInsertReference } from './formula.js';
import { alignmentOf, isError, formatValue, generalNumber } from './values.js';
import { parseCsv, toCsv, detectDelimiter, sheetToCsv } from './csv.js';
import { FUNCTION_LIST } from './functions.js';

const STORAGE_KEY = 'planilha-leve:dados';
const MIN_ROWS = 100;
const MIN_COLS = 26;
const MAX_GRID_ROWS = 2000;
const MAX_GRID_COLS = 52;
const ARROWS = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };

const $ = (id) => document.getElementById(id);
const grid = $('grid');
const wrap = $('grid-wrap');
const editor = $('cell-editor');
const bar = $('formula-input');
const nameBox = $('name-box');
const message = $('message');
const statusCell = $('status-cell');
const statusSummary = $('status-summary');
const fileInput = $('file-input');

const copyMarquee = document.createElement('div');
copyMarquee.className = 'marquee';
copyMarquee.hidden = true;
const refMarquee = document.createElement('div');
refMarquee.className = 'marquee ref';
refMarquee.hidden = true;
wrap.append(copyMarquee, refMarquee);

let sheet;
let tds = [];
let colHeads = [];
let rowHeads = [];
const state = {
  active: { row: 0, col: 0 },
  anchor: { row: 0, col: 0 },
  editing: null,
  clip: null,
  point: null,
  dragging: false,
  painted: null,
};

const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, v));

let messageTimer = 0;
function say(text, tone = 'ok') {
  message.textContent = text;
  message.dataset.tone = tone;
  clearTimeout(messageTimer);
  messageTimer = setTimeout(() => {
    message.textContent = '';
  }, 7000);
}

// ---------------------------------------------------------------------------
// Grade
// ---------------------------------------------------------------------------

function cellId(r, c) {
  return `c${r}_${c}`;
}

function makeRow(r) {
  const tr = document.createElement('tr');
  tr.setAttribute('role', 'row');
  const th = document.createElement('th');
  th.setAttribute('role', 'rowheader');
  th.scope = 'row';
  th.textContent = String(r + 1);
  th.dataset.r = String(r);
  tr.append(th);
  rowHeads[r] = th;
  tds[r] = [];
  for (let c = 0; c < sheet.cols; c++) {
    const td = document.createElement('td');
    td.id = cellId(r, c);
    td.setAttribute('role', 'gridcell');
    td.dataset.r = String(r);
    td.dataset.c = String(c);
    tr.append(td);
    tds[r][c] = td;
  }
  return tr;
}

function buildGrid() {
  tds = [];
  rowHeads = [];
  colHeads = [];
  const thead = document.createElement('thead');
  const hr = document.createElement('tr');
  hr.setAttribute('role', 'row');
  const corner = document.createElement('th');
  corner.className = 'corner';
  corner.setAttribute('role', 'columnheader');
  corner.innerHTML = '<span class="visually-hidden">Selecionar tudo</span>';
  hr.append(corner);
  for (let c = 0; c < sheet.cols; c++) {
    const th = document.createElement('th');
    th.setAttribute('role', 'columnheader');
    th.scope = 'col';
    th.textContent = colName(c);
    th.dataset.c = String(c);
    hr.append(th);
    colHeads[c] = th;
  }
  thead.append(hr);
  const tbody = document.createElement('tbody');
  for (let r = 0; r < sheet.rows; r++) tbody.append(makeRow(r));
  grid.replaceChildren(thead, tbody);
  grid.setAttribute('aria-rowcount', String(sheet.rows + 1));
  grid.setAttribute('aria-colcount', String(sheet.cols + 1));
  state.painted = null;
  for (const [k] of sheet.cells) paintCell(rowOf(k), colOf(k));
  paintSelection();
}

function appendRows(count) {
  const target = Math.min(MAX_GRID_ROWS, sheet.rows + count);
  const tbody = grid.tBodies[0];
  const frag = document.createDocumentFragment();
  for (let r = sheet.rows; r < target; r++) frag.append(makeRow(r));
  sheet.rows = target;
  tbody.append(frag);
  grid.setAttribute('aria-rowcount', String(sheet.rows + 1));
  for (const [k] of sheet.cells) if (rowOf(k) >= target - count) paintCell(rowOf(k), colOf(k));
}

/** Aumenta a grade para caber as células indicadas; devolve false se passar do limite. */
function ensureSize(rows, cols) {
  if (rows > MAX_GRID_ROWS || cols > MAX_GRID_COLS) return false;
  if (cols > sheet.cols) {
    sheet.cols = Math.min(MAX_GRID_COLS, Math.max(cols, sheet.cols + 5));
    if (rows > sheet.rows) sheet.rows = Math.min(MAX_GRID_ROWS, Math.ceil(rows / 50) * 50);
    buildGrid();
  } else if (rows > sheet.rows) {
    appendRows(Math.ceil((rows - sheet.rows) / 50) * 50);
  }
  return true;
}

const SPILL_MAX = 8;

/** Quantas células vazias à direita um texto pode ocupar (como no Excel). */
function spillWidth(r, c) {
  let k = 0;
  while (k < SPILL_MAX && c + k + 1 < sheet.cols && sheet.raw(r, c + k + 1) === '') k++;
  return k;
}

function paintCell(r, c) {
  const td = tds[r]?.[c];
  if (!td) return;
  const cell = sheet.get(r, c);
  const value = cell ? cell.value : null;
  const text = cell ? formatValue(value, cell.format) : '';
  const spill = typeof value === 'string' && text !== '' ? spillWidth(r, c) : 0;
  if (spill > 0) {
    const span = document.createElement('span');
    span.className = 'spill-text';
    span.style.width = `calc(${spill + 1} * var(--cell-w) - 0.8rem)`;
    span.textContent = text;
    td.replaceChildren(span);
  } else {
    td.textContent = text;
  }
  td.classList.toggle('spill', spill > 0);
  const align = alignmentOf(value);
  td.classList.toggle('num', value !== null && align === 'right');
  td.classList.toggle('center', align === 'center');
  td.classList.toggle('err', isError(value));
  if (isError(value) && value.detail) td.title = value.detail;
  else td.removeAttribute('title');
}

function paintKeys(keys) {
  const toPaint = new Set();
  for (const k of keys) {
    const r = rowOf(k);
    const c = colOf(k);
    toPaint.add(`${r},${c}`);
    // o vizinho da esquerda pode ganhar ou perder espaço para o texto transbordar
    for (let j = 1; j <= SPILL_MAX && c - j >= 0; j++) if (sheet.get(r, c - j)) toPaint.add(`${r},${c - j}`);
  }
  for (const item of toPaint) {
    const [r, c] = item.split(',').map(Number);
    paintCell(r, c);
  }
}

function selRect() {
  const { active: a, anchor: b } = state;
  return { r1: Math.min(a.row, b.row), c1: Math.min(a.col, b.col), r2: Math.max(a.row, b.row), c2: Math.max(a.col, b.col) };
}

function forRect(rect, fn) {
  for (let r = rect.r1; r <= rect.r2; r++) for (let c = rect.c1; c <= rect.c2; c++) fn(r, c);
}

function paintSelection() {
  const old = state.painted;
  if (old) {
    forRect(old, (r, c) => {
      const td = tds[r]?.[c];
      if (td) {
        td.classList.remove('sel', 'sel-multi', 'active');
        td.removeAttribute('aria-selected');
      }
    });
    for (let c = old.c1; c <= old.c2; c++) colHeads[c]?.classList.remove('hl');
    for (let r = old.r1; r <= old.r2; r++) rowHeads[r]?.classList.remove('hl');
  }
  const s = selRect();
  const multi = s.r1 !== s.r2 || s.c1 !== s.c2;
  forRect(s, (r, c) => {
    const td = tds[r]?.[c];
    if (!td) return;
    td.classList.add('sel');
    if (multi) td.classList.add('sel-multi');
    td.setAttribute('aria-selected', 'true');
  });
  for (let c = s.c1; c <= s.c2; c++) colHeads[c]?.classList.add('hl');
  for (let r = s.r1; r <= s.r2; r++) rowHeads[r]?.classList.add('hl');
  const { row, col } = state.active;
  tds[row]?.[col]?.classList.add('active');
  grid.setAttribute('aria-activedescendant', cellId(row, col));
  state.painted = s;
  nameBox.value = multi ? `${cellName(s.r1, s.c1)}:${cellName(s.r2, s.c2)}` : cellName(row, col);
  if (!state.editing) bar.value = sheet.raw(row, col);
  updateStatus();
}

function updateStatus() {
  const { row, col } = state.active;
  const v = sheet.value(row, col);
  statusCell.textContent = isError(v) && v.detail ? `${cellName(row, col)}: ${v.detail}` : '';
  const s = selRect();
  if (s.r1 === s.r2 && s.c1 === s.c2) {
    statusSummary.textContent = '';
    return;
  }
  const sum = selectionSummary(sheet, s.r1, s.c1, s.r2, s.c2);
  const parts = [];
  if (sum.numbers > 0) parts.push(`Soma: ${generalNumber(sum.sum)}`, `Média: ${generalNumber(sum.average)}`);
  parts.push(`Contagem: ${sum.count}`);
  statusSummary.textContent = parts.join(' · ');
}

function scrollToCell(r, c) {
  const td = tds[r]?.[c];
  if (!td) return;
  const headH = grid.tHead.offsetHeight;
  const headW = rowHeads[0]?.offsetWidth ?? 0;
  if (td.offsetTop - headH < wrap.scrollTop) wrap.scrollTop = td.offsetTop - headH;
  else if (td.offsetTop + td.offsetHeight > wrap.scrollTop + wrap.clientHeight) wrap.scrollTop = td.offsetTop + td.offsetHeight - wrap.clientHeight;
  if (td.offsetLeft - headW < wrap.scrollLeft) wrap.scrollLeft = td.offsetLeft - headW;
  else if (td.offsetLeft + td.offsetWidth > wrap.scrollLeft + wrap.clientWidth) wrap.scrollLeft = td.offsetLeft + td.offsetWidth - wrap.clientWidth;
}

function moveTo(row, col, extend = false, scroll = true) {
  if (row >= sheet.rows && sheet.rows < MAX_GRID_ROWS) appendRows(50);
  row = clamp(row, 0, sheet.rows - 1);
  col = clamp(col, 0, sheet.cols - 1);
  state.active = { row, col };
  if (!extend) state.anchor = { row, col };
  paintSelection();
  if (scroll) scrollToCell(row, col);
}

function moveBy(dr, dc, extend = false) {
  moveTo(state.active.row + dr, state.active.col + dc, extend);
}

/** Ctrl+seta: vai até a borda do bloco de dados, como no Excel. */
function jump(dr, dc, extend) {
  const inside = (r, c) => r >= 0 && c >= 0 && r < sheet.rows && c < sheet.cols;
  const filled = (r, c) => sheet.raw(r, c) !== '';
  const { row, col } = state.active;
  let r = row + dr;
  let c = col + dc;
  if (!inside(r, c)) return;
  if (filled(row, col) && filled(r, c)) {
    while (inside(r + dr, c + dc) && filled(r + dr, c + dc)) {
      r += dr;
      c += dc;
    }
  } else {
    while (inside(r, c) && !filled(r, c)) {
      r += dr;
      c += dc;
    }
    if (!inside(r, c)) {
      r -= dr;
      c -= dc;
    }
  }
  moveTo(r, c, extend);
}

function selectRect(r1, c1, r2, c2) {
  state.anchor = { row: r1, col: c1 };
  state.active = { row: r2, col: c2 };
  paintSelection();
}

function placeMarquee(el, rect) {
  const a = tds[rect.r1]?.[rect.c1];
  const b = tds[rect.r2]?.[rect.c2];
  if (!a || !b) {
    el.hidden = true;
    return;
  }
  el.style.top = `${a.offsetTop}px`;
  el.style.left = `${a.offsetLeft}px`;
  el.style.width = `${b.offsetLeft + b.offsetWidth - a.offsetLeft}px`;
  el.style.height = `${b.offsetTop + b.offsetHeight - a.offsetTop}px`;
  el.hidden = false;
}

// ---------------------------------------------------------------------------
// Alterações, histórico e salvamento
// ---------------------------------------------------------------------------

let saveTimer = 0;
function saveNow() {
  clearTimeout(saveTimer);
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(sheet.toJSON()));
  } catch {
    say('Não foi possível salvar no navegador (armazenamento cheio ou bloqueado).', 'error');
  }
}

function scheduleSave() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(saveNow, 400);
}

function afterChange() {
  $('btn-undo').disabled = !sheet.canUndo;
  $('btn-redo').disabled = !sheet.canRedo;
  if (!state.editing) bar.value = sheet.raw(state.active.row, state.active.col);
  updateStatus();
  scheduleSave();
}

function applyChanges(changes) {
  if (changes.length === 0) return;
  let maxRow = 0;
  let maxCol = 0;
  for (const ch of changes) {
    maxRow = Math.max(maxRow, ch.row + 1);
    maxCol = Math.max(maxCol, ch.col + 1);
  }
  if (!ensureSize(maxRow, maxCol)) {
    say(`A planilha vai até ${MAX_GRID_ROWS} linhas e ${MAX_GRID_COLS} colunas; o que passou disso ficou de fora.`, 'warn');
    changes = changes.filter((ch) => ch.row < MAX_GRID_ROWS && ch.col < MAX_GRID_COLS);
    ensureSize(Math.min(maxRow, MAX_GRID_ROWS), Math.min(maxCol, MAX_GRID_COLS));
  }
  paintKeys(sheet.apply(changes));
  afterChange();
}

function selectKeys(keys) {
  if (!keys.length) return;
  let r1 = Infinity;
  let c1 = Infinity;
  let r2 = -1;
  let c2 = -1;
  for (const k of keys) {
    r1 = Math.min(r1, rowOf(k));
    c1 = Math.min(c1, colOf(k));
    r2 = Math.max(r2, rowOf(k));
    c2 = Math.max(c2, colOf(k));
  }
  selectRect(r1, c1, r2, c2);
  scrollToCell(r1, c1);
}

function undo() {
  commitIfEditing();
  const r = sheet.undo();
  if (!r) return say('Nada para desfazer.', 'warn');
  paintKeys(r.keys);
  selectKeys(r.cells);
  afterChange();
}

function redo() {
  commitIfEditing();
  const r = sheet.redo();
  if (!r) return say('Nada para refazer.', 'warn');
  paintKeys(r.keys);
  selectKeys(r.cells);
  afterChange();
}

// ---------------------------------------------------------------------------
// Edição
// ---------------------------------------------------------------------------

const measureCtx = document.createElement('canvas').getContext('2d');

function activeInput() {
  return state.editing?.source === 'bar' ? bar : editor;
}

function resizeEditor() {
  const e = state.editing;
  if (!e) return;
  const td = tds[e.row]?.[e.col];
  if (!td) return;
  if (measureCtx) measureCtx.font = getComputedStyle(editor).font;
  const textWidth = measureCtx ? measureCtx.measureText(editor.value).width + 24 : editor.value.length * 9;
  editor.style.top = `${td.offsetTop}px`;
  editor.style.left = `${td.offsetLeft}px`;
  editor.style.height = `${td.offsetHeight}px`;
  editor.style.width = `${Math.min(Math.max(td.offsetWidth, textWidth), 520)}px`;
}

function startEdit(initial = null, mode = 'edit', source = 'cell') {
  const { row, col } = state.active;
  const original = sheet.raw(row, col);
  state.editing = { row, col, original, mode, source };
  state.point = null;
  const text = initial ?? original;
  editor.value = text;
  bar.value = text;
  editor.hidden = false;
  resizeEditor();
  if (source === 'cell') {
    editor.focus({ preventScroll: true });
    editor.setSelectionRange(text.length, text.length);
  }
}

function stopEdit() {
  state.editing = null;
  state.point = null;
  editor.hidden = true;
  refMarquee.hidden = true;
}

function commitEdit(dr = 0, dc = 0) {
  const e = state.editing;
  if (!e) return;
  const raw = activeInput().value;
  stopEdit();
  if (raw !== e.original) applyChanges([{ row: e.row, col: e.col, raw }]);
  grid.focus({ preventScroll: true });
  if (dr || dc) moveTo(e.row + dr, e.col + dc);
  else moveTo(e.row, e.col);
}

function commitIfEditing() {
  if (state.editing) commitEdit(0, 0);
}

function cancelEdit() {
  const e = state.editing;
  if (!e) return;
  stopEdit();
  bar.value = sheet.raw(e.row, e.col);
  grid.focus({ preventScroll: true });
}

function isFormulaEdit() {
  return state.editing && activeInput().value.startsWith('=');
}

/**
 * Modo de apontar: durante a edição de uma fórmula, clicar (ou usar as setas) numa
 * célula insere o endereço dela no cursor; arrastar ou usar Shift vira intervalo.
 */
function pointAt(row, col, extend) {
  const input = activeInput();
  const value = input.value;
  const caret = input.selectionStart;
  const p = state.point;
  let start;
  let end;
  if (p && caret === p.start + p.len && input.selectionEnd === caret) {
    start = p.start;
    end = p.start + p.len;
  } else if (input.selectionEnd === caret && canInsertReference(value, caret)) {
    start = caret;
    end = caret;
  } else {
    return false;
  }
  const anchor = extend && p ? p.anchor : { row, col };
  const rect = { r1: Math.min(anchor.row, row), c1: Math.min(anchor.col, col), r2: Math.max(anchor.row, row), c2: Math.max(anchor.col, col) };
  const text = rect.r1 === rect.r2 && rect.c1 === rect.c2 ? cellName(row, col) : `${cellName(rect.r1, rect.c1)}:${cellName(rect.r2, rect.c2)}`;
  input.value = value.slice(0, start) + text + value.slice(end);
  input.setSelectionRange(start + text.length, start + text.length);
  (input === bar ? editor : bar).value = input.value;
  state.point = { start, len: text.length, anchor, cursor: { row, col } };
  placeMarquee(refMarquee, rect);
  resizeEditor();
  return true;
}

function editKeydown(ev) {
  const e = state.editing;
  if (!e) return;
  const input = ev.currentTarget;
  if (ev.key === 'Enter') {
    ev.preventDefault();
    commitEdit(ev.shiftKey ? -1 : 1, 0);
  } else if (ev.key === 'Tab') {
    ev.preventDefault();
    commitEdit(0, ev.shiftKey ? -1 : 1);
  } else if (ev.key === 'Escape') {
    ev.preventDefault();
    cancelEdit();
  } else if (ev.key === 'F4') {
    ev.preventDefault();
    const r = cycleReference(input.value, input.selectionStart);
    if (r) {
      input.value = r.text;
      input.setSelectionRange(r.caret, r.caret);
      (input === bar ? editor : bar).value = r.text;
      state.point = null;
      resizeEditor();
    }
  } else if (ARROWS[ev.key] && e.mode === 'enter' && !ev.altKey && !ev.ctrlKey && !ev.metaKey) {
    const [dr, dc] = ARROWS[ev.key];
    ev.preventDefault();
    if (isFormulaEdit() && (state.point || canInsertReference(input.value, input.selectionStart))) {
      const base = state.point?.cursor ?? { row: e.row, col: e.col };
      const row = clamp(base.row + dr, 0, sheet.rows - 1);
      const col = clamp(base.col + dc, 0, sheet.cols - 1);
      pointAt(row, col, ev.shiftKey);
      scrollToCell(row, col);
    } else {
      commitEdit(dr, dc);
    }
  }
}

editor.addEventListener('keydown', editKeydown);
bar.addEventListener('keydown', (ev) => {
  if (!state.editing && (ev.key === 'Enter' || ev.key === 'Escape')) {
    ev.preventDefault();
    grid.focus();
    return;
  }
  editKeydown(ev);
});

editor.addEventListener('input', () => {
  bar.value = editor.value;
  state.point = null;
  refMarquee.hidden = true;
  resizeEditor();
});

bar.addEventListener('focus', () => {
  if (!state.editing) startEdit(null, 'edit', 'bar');
  else state.editing.source = 'bar';
});

bar.addEventListener('input', () => {
  if (!state.editing) startEdit(bar.value, 'edit', 'bar');
  editor.value = bar.value;
  state.point = null;
  refMarquee.hidden = true;
  resizeEditor();
});

editor.addEventListener('focus', () => {
  if (state.editing) state.editing.source = 'cell';
});

for (const input of [editor, bar]) {
  input.addEventListener('blur', (ev) => {
    const next = ev.relatedTarget;
    if (!state.editing || next === editor || next === bar || next === grid) return;
    commitEdit(0, 0);
  });
}

// ---------------------------------------------------------------------------
// Mouse
// ---------------------------------------------------------------------------

grid.addEventListener('mousedown', (ev) => {
  if (ev.button !== 0) return;
  const td = ev.target.closest('td');
  if (td) {
    const row = Number(td.dataset.r);
    const col = Number(td.dataset.c);
    if (isFormulaEdit() && pointAt(row, col, ev.shiftKey)) {
      ev.preventDefault();
      state.dragging = 'point';
      return;
    }
    if (state.editing) commitEdit(0, 0);
    ev.preventDefault();
    grid.focus({ preventScroll: true });
    moveTo(row, col, ev.shiftKey, false);
    state.dragging = 'select';
    return;
  }
  const th = ev.target.closest('th');
  if (!th) return;
  ev.preventDefault();
  commitIfEditing();
  grid.focus({ preventScroll: true });
  if (th.classList.contains('corner')) selectRect(0, 0, sheet.rows - 1, sheet.cols - 1);
  else if (th.dataset.c !== undefined) selectRect(0, Number(th.dataset.c), sheet.rows - 1, Number(th.dataset.c));
  else if (th.dataset.r !== undefined) selectRect(Number(th.dataset.r), 0, Number(th.dataset.r), sheet.cols - 1);
});

grid.addEventListener('mouseover', (ev) => {
  if (!state.dragging) return;
  const td = ev.target.closest('td');
  if (!td) return;
  const row = Number(td.dataset.r);
  const col = Number(td.dataset.c);
  if (state.dragging === 'point') pointAt(row, col, true);
  else moveTo(row, col, true, false);
});

window.addEventListener('mouseup', () => {
  state.dragging = false;
});

grid.addEventListener('dblclick', (ev) => {
  const td = ev.target.closest('td');
  if (!td || state.editing) return;
  startEdit(null, 'edit');
});

// ---------------------------------------------------------------------------
// Teclado na grade
// ---------------------------------------------------------------------------

grid.addEventListener('keydown', (ev) => {
  if (state.editing) return;
  const mod = ev.ctrlKey || ev.metaKey;
  const key = ev.key;
  if (ARROWS[key]) {
    ev.preventDefault();
    const [dr, dc] = ARROWS[key];
    if (mod) jump(dr, dc, ev.shiftKey);
    else moveBy(dr, dc, ev.shiftKey);
    return;
  }
  switch (key) {
    case 'Tab':
      ev.preventDefault();
      moveBy(0, ev.shiftKey ? -1 : 1);
      return;
    case 'Enter':
    case 'F2':
      ev.preventDefault();
      startEdit(null, 'edit');
      return;
    case 'Delete':
    case 'Backspace': {
      ev.preventDefault();
      const s = selRect();
      applyChanges(clearChanges(s.r1, s.c1, s.r2, s.c2));
      return;
    }
    case 'Escape':
      state.clip = null;
      copyMarquee.hidden = true;
      return;
    case 'Home':
      ev.preventDefault();
      moveTo(mod ? 0 : state.active.row, 0, ev.shiftKey);
      return;
    case 'End':
      if (mod) {
        ev.preventDefault();
        const b = sheet.usedBounds();
        moveTo(Math.max(0, b.maxRow), Math.max(0, b.maxCol), ev.shiftKey);
      }
      return;
    case 'PageDown':
      ev.preventDefault();
      moveBy(20, 0, ev.shiftKey);
      return;
    case 'PageUp':
      ev.preventDefault();
      moveBy(-20, 0, ev.shiftKey);
      return;
    default:
      break;
  }
  if (mod && !ev.altKey) {
    const k = key.toLowerCase();
    if (k === 'z' && !ev.shiftKey) {
      ev.preventDefault();
      undo();
    } else if (k === 'y' || (k === 'z' && ev.shiftKey)) {
      ev.preventDefault();
      redo();
    } else if (k === 'd') {
      ev.preventDefault();
      fill('down');
    } else if (k === 'r') {
      ev.preventDefault();
      fill('right');
    } else if (k === 'a') {
      ev.preventDefault();
      selectRect(0, 0, sheet.rows - 1, sheet.cols - 1);
    }
    return;
  }
  if (key.length === 1 && !ev.altKey) {
    ev.preventDefault();
    startEdit(key, 'enter');
  }
});

// ---------------------------------------------------------------------------
// Área de transferência
// ---------------------------------------------------------------------------

function gridFocused() {
  return document.activeElement === grid && !state.editing;
}

function copySelection(data, cut) {
  const s = selRect();
  const rows = [];
  for (let r = s.r1; r <= s.r2; r++) {
    const line = [];
    for (let c = s.c1; c <= s.c2; c++) line.push(sheet.display(r, c));
    rows.push(line);
  }
  const text = toCsv(rows, '\t');
  data.setData('text/plain', text);
  state.clip = { text, block: copyBlock(sheet, s.r1, s.c1, s.r2, s.c2), cut, rect: s };
  placeMarquee(copyMarquee, s);
  say(cut ? 'Recortado. Cole com Ctrl+V.' : 'Copiado. Cole com Ctrl+V (as fórmulas ajustam as referências).');
}

function pasteText(text) {
  const s = selRect();
  const targetRows = s.r2 - s.r1 + 1;
  const targetCols = s.c2 - s.c1 + 1;
  let changes;
  const clip = state.clip;
  if (clip && text.replace(/\r\n/g, '\n') === clip.text.replace(/\r\n/g, '\n')) {
    changes = pasteChanges(clip.block, s.r1, s.c1, { shift: !clip.cut, targetRows, targetCols });
    if (clip.cut) {
      const written = new Set(changes.map((ch) => `${ch.row},${ch.col}`));
      forRect(clip.rect, (r, c) => {
        if (!written.has(`${r},${c}`)) changes.push({ row: r, col: c, raw: '', format: null });
      });
      state.clip = null;
      copyMarquee.hidden = true;
    }
  } else {
    // texto vindo de fora (Excel, Google Planilhas, editor): colunas separadas por tabulação
    const rows = parseCsv(text.replace(/\r?\n$/, ''), '\t');
    changes = [];
    rows.forEach((line, r) => line.forEach((value, c) => changes.push({ row: s.r1 + r, col: s.c1 + c, raw: value })));
  }
  applyChanges(changes);
  if (changes.length) {
    const r2 = Math.max(...changes.map((ch) => ch.row));
    const c2 = Math.max(...changes.map((ch) => ch.col));
    selectRect(s.r1, s.c1, Math.min(r2, sheet.rows - 1), Math.min(c2, sheet.cols - 1));
  }
}

document.addEventListener('copy', (ev) => {
  if (!gridFocused()) return;
  ev.preventDefault();
  copySelection(ev.clipboardData, false);
});

document.addEventListener('cut', (ev) => {
  if (!gridFocused()) return;
  ev.preventDefault();
  copySelection(ev.clipboardData, true);
});

document.addEventListener('paste', (ev) => {
  if (!gridFocused()) return;
  ev.preventDefault();
  pasteText(ev.clipboardData.getData('text/plain'));
});

// ---------------------------------------------------------------------------
// Barra de ferramentas
// ---------------------------------------------------------------------------

function fill(direction) {
  commitIfEditing();
  const s = selRect();
  let changes;
  if (direction === 'down') {
    if (s.r1 === s.r2 && s.r1 === 0) return say('Selecione a célula de origem e as de baixo.', 'warn');
    changes = s.r1 === s.r2 ? fillDownChanges(sheet, s.r1 - 1, s.c1, s.r2, s.c2) : fillDownChanges(sheet, s.r1, s.c1, s.r2, s.c2);
  } else {
    if (s.c1 === s.c2 && s.c1 === 0) return say('Selecione a célula de origem e as da direita.', 'warn');
    changes = s.c1 === s.c2 ? fillRightChanges(sheet, s.r1, s.c1 - 1, s.r2, s.c2) : fillRightChanges(sheet, s.r1, s.c1, s.r2, s.c2);
  }
  applyChanges(changes);
}

const DEFAULT_DECIMALS = { numero: 2, moeda: 2, porcentagem: 0 };

function setFormat(type) {
  commitIfEditing();
  const changes = [];
  forRect(selRect(), (row, col) => {
    changes.push({ row, col, format: type === 'geral' ? null : { type, decimals: DEFAULT_DECIMALS[type] } });
  });
  applyChanges(changes);
}

function changeDecimals(delta) {
  commitIfEditing();
  const changes = [];
  forRect(selRect(), (row, col) => {
    const f = sheet.format(row, col);
    const type = f && f.type !== 'data' ? f.type : 'numero';
    if (f?.type === 'data') return;
    const base = f?.decimals ?? DEFAULT_DECIMALS[type] ?? 2;
    changes.push({ row, col, format: { type, decimals: clamp(base + delta, 0, 10) } });
  });
  applyChanges(changes);
}

for (const b of document.querySelectorAll('[data-format]')) b.addEventListener('click', () => setFormat(b.dataset.format));
for (const b of document.querySelectorAll('[data-decimals]')) b.addEventListener('click', () => changeDecimals(Number(b.dataset.decimals)));
$('btn-undo').addEventListener('click', undo);
$('btn-redo').addEventListener('click', redo);
$('btn-fill-down').addEventListener('click', () => fill('down'));
$('btn-fill-right').addEventListener('click', () => fill('right'));
$('btn-clear').addEventListener('click', () => {
  commitIfEditing();
  const s = selRect();
  applyChanges(clearChanges(s.r1, s.c1, s.r2, s.c2, { formats: true }));
});
$('btn-add-rows').addEventListener('click', () => {
  if (sheet.rows >= MAX_GRID_ROWS) return say(`O limite é de ${MAX_GRID_ROWS} linhas.`, 'warn');
  appendRows(50);
  say(`Agora são ${sheet.rows} linhas.`);
  scheduleSave();
});
$('btn-add-cols').addEventListener('click', () => {
  if (sheet.cols >= MAX_GRID_COLS) return say(`O limite é de ${MAX_GRID_COLS} colunas.`, 'warn');
  sheet.cols = Math.min(MAX_GRID_COLS, sheet.cols + 5);
  buildGrid();
  say(`Agora as colunas vão até ${colName(sheet.cols - 1)}.`);
  scheduleSave();
});

function download(name, content, type) {
  const url = URL.createObjectURL(new Blob([content], { type }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

$('btn-export').addEventListener('click', () => {
  commitIfEditing();
  const mode = $('csv-mode').value === 'formulas' ? 'formulas' : 'valores';
  download(mode === 'formulas' ? 'planilha-formulas.csv' : 'planilha.csv', sheetToCsv(sheet, { mode }), 'text/csv;charset=utf-8');
  say(mode === 'formulas' ? 'CSV exportado com as fórmulas.' : 'CSV exportado com os valores, separados por ponto e vírgula.');
});

$('btn-import').addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', async () => {
  const file = fileInput.files[0];
  fileInput.value = '';
  if (!file) return;
  if (file.size > 5 * 1024 * 1024) return say('Arquivo grande demais (limite de 5 MB).', 'error');
  const text = await file.text();
  const rows = parseCsv(text, detectDelimiter(text));
  const width = Math.max(0, ...rows.map((r) => r.length));
  if (rows.length > MAX_GRID_ROWS || width > MAX_GRID_COLS) {
    say(`O arquivo tem ${rows.length} linhas e ${width} colunas; o limite é ${MAX_GRID_ROWS} por ${MAX_GRID_COLS}.`, 'error');
    return;
  }
  replaceAll(rows.flatMap((line, r) => line.map((raw, c) => ({ row: r, col: c, raw, format: null }))));
  say(`Importado "${file.name}": ${rows.length} linhas e ${width} colunas. Ctrl+Z desfaz.`);
});

/** Troca todo o conteúdo numa única ação (que pode ser desfeita). */
function replaceAll(newCells) {
  commitIfEditing();
  const changes = [];
  const touched = new Set();
  for (const ch of newCells) {
    if (ch.raw === '' && !ch.format) continue;
    changes.push(ch);
    touched.add(`${ch.row},${ch.col}`);
  }
  for (const [k] of sheet.cells) {
    const r = rowOf(k);
    const c = colOf(k);
    if (!touched.has(`${r},${c}`)) changes.push({ row: r, col: c, raw: '', format: null });
  }
  state.clip = null;
  copyMarquee.hidden = true;
  applyChanges(changes);
  moveTo(0, 0);
}

// Ações que apagam a planilha pedem um segundo clique de confirmação.
function confirmable(button, label, action) {
  const original = button.textContent;
  let armedUntil = 0;
  let timer = 0;
  const reset = () => {
    armedUntil = 0;
    button.textContent = original;
    button.classList.remove('danger');
  };
  button.addEventListener('click', () => {
    if (sheet.cells.size > 0 && Date.now() > armedUntil) {
      armedUntil = Date.now() + 5000;
      button.textContent = label;
      button.classList.add('danger');
      say('Clique de novo para confirmar. Ctrl+Z desfaz depois, se precisar.', 'warn');
      clearTimeout(timer);
      timer = setTimeout(reset, 5000);
      return;
    }
    clearTimeout(timer);
    reset();
    action();
  });
}

confirmable($('btn-new'), 'Confirmar: nova', () => {
  replaceAll([]);
  say('Planilha vazia.');
});

confirmable($('btn-sample'), 'Confirmar: exemplo', () => {
  replaceAll(sampleCells());
  say('Exemplo carregado: um orçamento doméstico com várias fórmulas.');
});

nameBox.addEventListener('focus', () => nameBox.select());
nameBox.addEventListener('keydown', (ev) => {
  if (ev.key === 'Escape') {
    ev.preventDefault();
    paintSelection();
    grid.focus();
    return;
  }
  if (ev.key !== 'Enter') return;
  ev.preventDefault();
  const [a, b] = nameBox.value.split(':');
  const p = parseCellName(a ?? '');
  const q = b === undefined ? p : parseCellName(b);
  if (!p || !q || Math.max(p.row, q.row) >= MAX_GRID_ROWS || Math.max(p.col, q.col) >= MAX_GRID_COLS) {
    say(`Endereço inválido: "${nameBox.value}".`, 'error');
    return;
  }
  ensureSize(Math.max(p.row, q.row) + 1, Math.max(p.col, q.col) + 1);
  grid.focus({ preventScroll: true });
  selectRect(p.row, p.col, q.row, q.col);
  scrollToCell(q.row, q.col);
});

// ---------------------------------------------------------------------------
// Exemplo e início
// ---------------------------------------------------------------------------

function sampleCells() {
  const rows = [
    ['Orçamento de outubro'],
    [],
    ['Categoria', 'Previsto', 'Gasto', 'Diferença', 'Situação', 'Faixa', '', 'Gasto a partir de', 'Faixa'],
    ['Aluguel', '1500', '1500', '=B4-C4', '=SE(D4<0;"estourou";"ok")', '=PROCV(C4;$H$4:$I$6;2)', '', '0', 'Baixo'],
    ['Mercado', '900', '1043,27', '=B5-C5', '=SE(D5<0;"estourou";"ok")', '=PROCV(C5;$H$4:$I$6;2)', '', '500', 'Médio'],
    ['Transporte', '300', '212,4', '=B6-C6', '=SE(D6<0;"estourou";"ok")', '=PROCV(C6;$H$4:$I$6;2)', '', '1000', 'Alto'],
    ['Lazer', '250', '310', '=B7-C7', '=SE(D7<0;"estourou";"ok")', '=PROCV(C7;$H$4:$I$6;2)'],
    ['Contas', '400', '387,9', '=B8-C8', '=SE(D8<0;"estourou";"ok")', '=PROCV(C8;$H$4:$I$6;2)'],
    ['Total', '=SOMA(B4:B8)', '=SOMA(C4:C8)', '=B9-C9', '=SE(D9<0;"acima do previsto";"dentro do previsto")'],
    [],
    ['Maior gasto', '=MÁXIMO(C4:C8)'],
    ['Média por categoria', '=MÉDIA(C4:C8)'],
    ['Categorias estouradas', '=CONT.SE(D4:D8;"<0")'],
    ['Parte do previsto já gasta', '=C9/B9'],
    ['Resumo', '=CONCATENAR("Gasto total: R$ ";ARRED(C9;2))'],
    ['Atualizado em', '=HOJE()'],
  ];
  const money = { type: 'moeda', decimals: 2 };
  const cells = [];
  rows.forEach((line, r) => line.forEach((raw, c) => {
    let format = null;
    if ((r >= 3 && r <= 8 && c >= 1 && c <= 3) || r === 10 || r === 11) format = c === 1 || c === 2 || c === 3 ? money : null;
    if (r === 13 && c === 1) format = { type: 'porcentagem', decimals: 1 };
    if (r === 15 && c === 1) format = { type: 'data' };
    if (raw !== '') cells.push({ row: r, col: c, raw, format });
  }));
  return cells;
}

function renderFunctionList() {
  const dl = $('function-list');
  for (const f of FUNCTION_LIST) {
    const div = document.createElement('div');
    const dt = document.createElement('dt');
    dt.textContent = f.syntax;
    const dd = document.createElement('dd');
    dd.textContent = f.desc;
    div.append(dt, dd);
    dl.append(div);
  }
}

function load() {
  let data = null;
  try {
    data = JSON.parse(localStorage.getItem(STORAGE_KEY) ?? 'null');
  } catch {
    data = null;
  }
  if (data && Array.isArray(data.cells)) {
    sheet = Sheet.fromJSON(data, {});
    sheet.rows = clamp(Number(data.rows) || MIN_ROWS, MIN_ROWS, MAX_GRID_ROWS);
    sheet.cols = clamp(Number(data.cols) || MIN_COLS, MIN_COLS, MAX_GRID_COLS);
  } else {
    sheet = new Sheet({ rows: MIN_ROWS, cols: MIN_COLS });
    sheet.apply(sampleCells(), { record: false });
  }
}

load();
renderFunctionList();
buildGrid();
afterChange();
moveTo(0, 0);
document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'hidden') saveNow();
});
window.addEventListener('pagehide', saveNow);
