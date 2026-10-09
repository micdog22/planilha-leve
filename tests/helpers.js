// Utilitários dos testes: monta uma planilha a partir de { A1: '10', B2: '=A1*2' }.
import { Sheet } from '../src/engine.js';
import { parseCellName } from '../src/formula.js';

export const TODAY = 45657; // 31/12/2024

export function sheetWith(cells = {}, options = {}) {
  const sheet = new Sheet({ today: () => TODAY, ...options });
  const changes = Object.entries(cells).map(([name, raw]) => ({ ...parseCellName(name), raw }));
  sheet.apply(changes, { record: false });
  return sheet;
}

/** Calcula uma fórmula numa célula livre (Z1000) sobre os dados informados. */
export function calc(formula, cells = {}) {
  const sheet = sheetWith(cells);
  sheet.apply([{ row: 999, col: 25, raw: formula }], { record: false });
  return sheet.value(999, 25);
}

/** Valor como texto comparável: erros viram o código. */
export function show(v) {
  if (v && typeof v === 'object' && 'code' in v) return v.code;
  return v;
}

export function at(sheet, name) {
  const { row, col } = parseCellName(name);
  return show(sheet.value(row, col));
}

export function set(sheet, name, raw, extra = {}) {
  const { row, col } = parseCellName(name);
  return sheet.apply([{ row, col, raw, ...extra }]);
}
