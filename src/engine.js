// Motor da planilha: guarda as células, mantém o grafo de dependências e
// recalcula só o que mudou, na ordem certa.
//
// Recalcular é ordenar topologicamente o subgrafo afetado. Aqui isso é feito com
// o algoritmo de Tarjan (componentes fortemente conexas): ele devolve os
// componentes já em ordem de avaliação (precedentes primeiro), e um componente com
// mais de uma célula, ou com uma célula que aponta para si mesma, é exatamente uma
// referência circular.

import {
  ERR, err, isError, parseInput, RangeValue, formatValue, todaySerial, intersect,
} from './values.js';
import {
  parseFormula, collectRefs, FormulaSyntaxError, shiftFormula, cellName, MAX_COLS,
} from './formula.js';
import { evaluate } from './evaluate.js';

const key = (row, col) => row * MAX_COLS + col;
const rowOf = (k) => Math.floor(k / MAX_COLS);
const colOf = (k) => k % MAX_COLS;
const MAX_RANGE_CELLS = 1_000_000;
const HISTORY_LIMIT = 200;

export { key as cellKey, rowOf, colOf };

function sameFormat(a, b) {
  return (a?.type ?? 'geral') === (b?.type ?? 'geral') && (a?.decimals ?? null) === (b?.decimals ?? null);
}

function normalizeFormat(f) {
  if (!f || f.type === 'geral') return null;
  return f.decimals === undefined || f.decimals === null ? { type: f.type } : { type: f.type, decimals: f.decimals };
}

export class Sheet {
  constructor({ rows = 100, cols = 26, today } = {}) {
    this.rows = rows;
    this.cols = cols;
    this.cells = new Map();
    this.dependents = new Map(); // célula -> fórmulas que a citam diretamente
    this.rangeUsers = new Map(); // fórmula -> intervalos que ela usa
    this.history = [];
    this.future = [];
    this.today = today ?? (() => todaySerial());
  }

  // ------------------------------------------------------------ leitura

  get(row, col) {
    return this.cells.get(key(row, col));
  }

  raw(row, col) {
    return this.get(row, col)?.raw ?? '';
  }

  value(row, col) {
    const cell = this.get(row, col);
    return cell === undefined ? null : cell.value;
  }

  format(row, col) {
    return this.get(row, col)?.format ?? null;
  }

  display(row, col) {
    const cell = this.get(row, col);
    return cell ? formatValue(cell.value, cell.format) : '';
  }

  /** Última linha e coluna com conteúdo (ou -1 se vazia). */
  usedBounds() {
    let maxRow = -1;
    let maxCol = -1;
    for (const [k, cell] of this.cells) {
      if (cell.raw === '') continue;
      maxRow = Math.max(maxRow, rowOf(k));
      maxCol = Math.max(maxCol, colOf(k));
    }
    return { maxRow, maxCol };
  }

  rangeValue(r1, c1, r2, c2) {
    const rows = r2 - r1 + 1;
    const cols = c2 - c1 + 1;
    if (rows * cols > MAX_RANGE_CELLS) return err(ERR.REF, 'Intervalo grande demais');
    const values = new Array(rows * cols).fill(null);
    if (rows * cols <= this.cells.size) {
      let i = 0;
      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++, i++) {
          const cell = this.cells.get(key(r, c));
          if (cell) values[i] = cell.value;
        }
      }
    } else {
      // intervalo maior que a quantidade de células: percorre só as que existem
      for (const [k, cell] of this.cells) {
        const r = rowOf(k);
        const c = colOf(k);
        if (r >= r1 && r <= r2 && c >= c1 && c <= c2) values[(r - r1) * cols + (c - c1)] = cell.value;
      }
    }
    const v = new RangeValue(rows, cols, values);
    v.origin = { row: r1, col: c1 };
    return v;
  }

  // ------------------------------------------------------------ escrita

  /**
   * Aplica alterações: [{ row, col, raw?, format? }]. raw ausente mantém o conteúdo;
   * format ausente mantém (ou adota a formatação sugerida pela digitação, como
   * "R$ 10" ou "31/12/2024"); format null volta ao Geral.
   * Devolve o conjunto de chaves cujo valor exibido pode ter mudado.
   */
  apply(changes, { record = true } = {}) {
    const entry = [];
    const seen = new Set();
    for (const ch of changes) {
      const k = key(ch.row, ch.col);
      if (seen.has(k)) continue;
      seen.add(k);
      const cell = this.cells.get(k);
      const before = { raw: cell?.raw ?? '', format: cell?.format ?? null };
      const raw = ch.raw === undefined ? before.raw : String(ch.raw);
      let format = ch.format === undefined ? before.format : normalizeFormat(ch.format);
      if (ch.raw !== undefined && ch.format === undefined && raw !== before.raw && !before.format) {
        format = normalizeFormat(this.suggestedFormat(raw)) ?? before.format;
      }
      if (raw === before.raw && sameFormat(format, before.format)) continue;
      entry.push({ key: k, before, after: { raw, format } });
    }
    if (entry.length === 0) return new Set();
    if (record) {
      this.history.push(entry);
      if (this.history.length > HISTORY_LIMIT) this.history.shift();
      this.future = [];
    }
    return this._commit(entry.map((e) => [e.key, e.after]));
  }

  suggestedFormat(raw) {
    const p = parseInput(raw);
    if (p.format) return p.format;
    if (p.kind === 'formula' && /\bHOJE\s*\(/i.test(raw)) return { type: 'data' };
    return null;
  }

  undo() {
    const entry = this.history.pop();
    if (!entry) return null;
    this.future.push(entry);
    return { keys: this._commit(entry.map((e) => [e.key, e.before])), cells: entry.map((e) => e.key) };
  }

  redo() {
    const entry = this.future.pop();
    if (!entry) return null;
    this.history.push(entry);
    return { keys: this._commit(entry.map((e) => [e.key, e.after])), cells: entry.map((e) => e.key) };
  }

  get canUndo() {
    return this.history.length > 0;
  }

  get canRedo() {
    return this.future.length > 0;
  }

  /** Recalcula todas as fórmulas (usado ao abrir, por causa de HOJE()). */
  recalcAll() {
    const formulas = new Set();
    for (const [k, cell] of this.cells) if (cell.kind === 'formula') formulas.add(k);
    this._evaluateSet(formulas);
    return formulas;
  }

  // ------------------------------------------------------------ internos

  _commit(pairs) {
    const dirty = [];
    for (const [k, { raw, format }] of pairs) {
      this._store(k, raw, format);
      dirty.push(k);
    }
    const affected = this._affectedBy(dirty);
    for (const k of dirty) if (this.cells.get(k)?.kind === 'formula') affected.add(k);
    this._evaluateSet(affected);
    for (const k of dirty) affected.add(k);
    return affected;
  }

  _store(k, raw, format) {
    const old = this.cells.get(k);
    if (old) this._unlink(k, old);
    if (raw === '' && !format) {
      this.cells.delete(k);
      return;
    }
    const cell = { raw, format: format ?? null, kind: 'blank', value: null, ast: null, refs: [], ranges: [], syntaxError: null };
    const p = parseInput(raw);
    cell.kind = p.kind;
    if (p.kind === 'formula') {
      try {
        cell.ast = parseFormula(raw);
        const { cells, ranges } = collectRefs(cell.ast);
        cell.refs = [...new Set(cells.map((c) => key(c.row, c.col)))];
        cell.ranges = ranges;
      } catch (e) {
        if (!(e instanceof FormulaSyntaxError)) throw e;
        cell.syntaxError = `${e.message} (posição ${e.position + 2})`;
        cell.value = err(ERR.SYNTAX, cell.syntaxError);
      }
    } else {
      cell.value = p.value;
    }
    this.cells.set(k, cell);
    this._link(k, cell);
  }

  _link(k, cell) {
    for (const r of cell.refs) {
      let set = this.dependents.get(r);
      if (!set) this.dependents.set(r, (set = new Set()));
      set.add(k);
    }
    if (cell.ranges.length) this.rangeUsers.set(k, cell.ranges);
  }

  _unlink(k, cell) {
    for (const r of cell.refs) {
      const set = this.dependents.get(r);
      if (set) {
        set.delete(k);
        if (set.size === 0) this.dependents.delete(r);
      }
    }
    this.rangeUsers.delete(k);
  }

  _dependentsOf(k) {
    const out = [];
    const direct = this.dependents.get(k);
    if (direct) for (const d of direct) out.push(d);
    const r = rowOf(k);
    const c = colOf(k);
    for (const [f, ranges] of this.rangeUsers) {
      for (const g of ranges) {
        if (r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2) {
          out.push(f);
          break;
        }
      }
    }
    return out;
  }

  /** Todas as fórmulas que dependem, direta ou indiretamente, das células alteradas. */
  _affectedBy(keys) {
    const affected = new Set();
    const stack = [...keys];
    while (stack.length) {
      const k = stack.pop();
      for (const d of this._dependentsOf(k)) {
        if (!affected.has(d)) {
          affected.add(d);
          stack.push(d);
        }
      }
    }
    return affected;
  }

  /** Precedentes de uma fórmula que também estão no conjunto a recalcular. */
  _precedentsIn(k, set) {
    const cell = this.cells.get(k);
    const out = [];
    if (!cell) return out;
    for (const r of cell.refs) if (set.has(r)) out.push(r);
    for (const g of cell.ranges) {
      const size = (g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1);
      if (size <= set.size) {
        for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
          const p = key(r, c);
          if (set.has(p)) out.push(p);
        }
      } else {
        for (const p of set) {
          const r = rowOf(p);
          const c = colOf(p);
          if (r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2) out.push(p);
        }
      }
    }
    return out;
  }

  /** Tarjan iterativo: componentes em ordem de avaliação; ciclos viram #REF!. */
  _evaluateSet(set) {
    if (set.size === 0) return;
    const index = new Map();
    const low = new Map();
    const onStack = new Set();
    const stack = [];
    const succ = new Map();
    let counter = 0;
    const visit = (v) => {
      index.set(v, counter);
      low.set(v, counter);
      counter += 1;
      stack.push(v);
      onStack.add(v);
      const s = this._precedentsIn(v, set);
      succ.set(v, s);
      return [v, s, 0];
    };
    for (const root of set) {
      if (index.has(root)) continue;
      const work = [visit(root)];
      while (work.length) {
        const frame = work[work.length - 1];
        const [v, s] = frame;
        if (frame[2] < s.length) {
          const w = s[frame[2]++];
          if (!index.has(w)) work.push(visit(w));
          else if (onStack.has(w)) low.set(v, Math.min(low.get(v), index.get(w)));
          continue;
        }
        work.pop();
        if (work.length) {
          const parent = work[work.length - 1][0];
          low.set(parent, Math.min(low.get(parent), low.get(v)));
        }
        if (low.get(v) !== index.get(v)) continue;
        const comp = [];
        let w;
        do {
          w = stack.pop();
          onStack.delete(w);
          comp.push(w);
        } while (w !== v);
        if (comp.length > 1 || succ.get(v).includes(v)) {
          const names = comp.map((c) => cellName(rowOf(c), colOf(c))).sort().join(', ');
          for (const c of comp) this.cells.get(c).value = err(ERR.REF, `Referência circular entre ${names}`);
        } else {
          this._evaluateCell(v);
        }
      }
    }
  }

  _evaluateCell(k) {
    const cell = this.cells.get(k);
    if (!cell || cell.kind !== 'formula') return;
    if (cell.syntaxError) {
      cell.value = err(ERR.SYNTAX, cell.syntaxError);
      return;
    }
    const ctx = {
      row: rowOf(k),
      col: colOf(k),
      cell: (r, c) => this.value(r, c),
      range: (r1, c1, r2, c2) => this.rangeValue(r1, c1, r2, c2),
      today: this.today,
    };
    let v = intersect(evaluate(cell.ast, ctx), ctx);
    if (v === null || v === undefined) v = 0;
    cell.value = v;
  }

  // ------------------------------------------------------------ persistência

  toJSON() {
    const cells = [];
    for (const [k, cell] of this.cells) {
      const item = [cellName(rowOf(k), colOf(k)), cell.raw];
      if (cell.format) item.push(cell.format);
      cells.push(item);
    }
    return { version: 1, rows: this.rows, cols: this.cols, cells };
  }

  static fromJSON(data, options = {}) {
    const sheet = new Sheet({ ...options, rows: data?.rows ?? options.rows, cols: data?.cols ?? options.cols });
    const changes = [];
    for (const item of data?.cells ?? []) {
      if (!Array.isArray(item) || typeof item[0] !== 'string' || typeof item[1] !== 'string') continue;
      const m = /^([A-Z]{1,3})(\d{1,7})$/.exec(item[0]);
      if (!m) continue;
      let col = 0;
      for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
      changes.push({ row: Number(m[2]) - 1, col: col - 1, raw: item[1], format: item[2] ?? null });
    }
    sheet.apply(changes, { record: false });
    return sheet;
  }
}

// ---------------------------------------------------------------------------
// Copiar, colar e preencher
// ---------------------------------------------------------------------------

/** Copia um bloco de células (conteúdo e formatação) guardando a origem. */
export function copyBlock(sheet, r1, c1, r2, c2) {
  const cells = [];
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) cells.push({ raw: sheet.raw(r, c), format: sheet.format(r, c) });
  }
  return { rows: r2 - r1 + 1, cols: c2 - c1 + 1, cells, origin: { row: r1, col: c1 } };
}

/**
 * Alterações para colar um bloco a partir de (row, col). Fórmulas têm as
 * referências relativas deslocadas; se a seleção de destino for múltipla do bloco,
 * o bloco é repetido para preenchê-la (como no Excel).
 */
export function pasteChanges(block, row, col, { shift = true, targetRows = 0, targetCols = 0 } = {}) {
  const repRows = targetRows > block.rows && targetRows % block.rows === 0 ? targetRows / block.rows : 1;
  const repCols = targetCols > block.cols && targetCols % block.cols === 0 ? targetCols / block.cols : 1;
  const changes = [];
  for (let i = 0; i < repRows; i++) {
    for (let j = 0; j < repCols; j++) {
      for (let r = 0; r < block.rows; r++) {
        for (let c = 0; c < block.cols; c++) {
          const src = block.cells[r * block.cols + c];
          const tr = row + i * block.rows + r;
          const tc = col + j * block.cols + c;
          const raw = shift && block.origin ? shiftFormula(src.raw, tr - (block.origin.row + r), tc - (block.origin.col + c)) : src.raw;
          changes.push({ row: tr, col: tc, raw, format: src.format });
        }
      }
    }
  }
  return changes;
}

/** Ctrl+D: repete a primeira linha da seleção nas de baixo, ajustando referências. */
export function fillDownChanges(sheet, r1, c1, r2, c2) {
  const changes = [];
  for (let c = c1; c <= c2; c++) {
    const raw = sheet.raw(r1, c);
    const format = sheet.format(r1, c);
    for (let r = r1 + 1; r <= r2; r++) changes.push({ row: r, col: c, raw: shiftFormula(raw, r - r1, 0), format });
  }
  return changes;
}

/** Ctrl+R: repete a primeira coluna da seleção nas da direita. */
export function fillRightChanges(sheet, r1, c1, r2, c2) {
  const changes = [];
  for (let r = r1; r <= r2; r++) {
    const raw = sheet.raw(r, c1);
    const format = sheet.format(r, c1);
    for (let c = c1 + 1; c <= c2; c++) changes.push({ row: r, col: c, raw: shiftFormula(raw, 0, c - c1), format });
  }
  return changes;
}

/** Alterações que limpam o conteúdo (e opcionalmente a formatação) de um bloco. */
export function clearChanges(r1, c1, r2, c2, { formats = false } = {}) {
  const changes = [];
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) changes.push(formats ? { row: r, col: c, raw: '', format: null } : { row: r, col: c, raw: '' });
  }
  return changes;
}

/** Resumo de uma seleção para a barra de status: soma, média e contagens. */
export function selectionSummary(sheet, r1, c1, r2, c2) {
  let count = 0;
  let numbers = 0;
  let total = 0;
  const visit = (cell) => {
    if (!cell || cell.raw === '') return;
    count++;
    if (typeof cell.value === 'number') {
      numbers++;
      total += cell.value;
    }
  };
  const size = (r2 - r1 + 1) * (c2 - c1 + 1);
  if (size <= sheet.cells.size) {
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) visit(sheet.get(r, c));
  } else {
    for (const [k, cell] of sheet.cells) {
      const r = rowOf(k);
      const c = colOf(k);
      if (r >= r1 && r <= r2 && c >= c1 && c <= c2) visit(cell);
    }
  }
  return { count, numbers, sum: total, average: numbers ? total / numbers : null };
}

export { isError };
