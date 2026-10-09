// Valores da planilha: erros, conversões no estilo do Excel, comparação, datas,
// leitura do que o usuário digita (em pt-BR) e formatação para exibição.
//
// Tipos de valor: number, string, boolean, FormulaError e null (célula vazia).

export class FormulaError {
  constructor(code, detail = '') {
    this.code = code;
    this.detail = detail;
  }

  toString() {
    return this.code;
  }
}

export const ERR = {
  DIV0: '#DIV/0!',
  NAME: '#NOME?',
  VALUE: '#VALOR!',
  REF: '#REF!',
  NA: '#N/D',
  NUM: '#NÚM!',
  SYNTAX: '#ERRO!',
};

export const ERROR_CODES = Object.values(ERR);

const ERROR_DETAILS = {
  [ERR.DIV0]: 'Divisão por zero',
  [ERR.NAME]: 'Nome ou função desconhecida',
  [ERR.VALUE]: 'Tipo de valor inválido para a operação',
  [ERR.REF]: 'Referência inválida',
  [ERR.NA]: 'Valor não disponível',
  [ERR.NUM]: 'Número inválido para a operação',
  [ERR.SYNTAX]: 'Fórmula com erro de sintaxe',
};

export function err(code, detail) {
  return new FormulaError(code, detail ?? ERROR_DETAILS[code] ?? '');
}

export const isError = (v) => v instanceof FormulaError;

/** Bloco retangular de valores vindo de uma referência (A1) ou intervalo (A1:B3). */
export class RangeValue {
  constructor(rows, cols, values) {
    this.rows = rows;
    this.cols = cols;
    this.values = values; // linha a linha, tamanho rows * cols
  }

  get(r, c) {
    return this.values[r * this.cols + c];
  }

  *[Symbol.iterator]() {
    yield* this.values;
  }
}

/**
 * Intervalo usado onde se espera um valor só: como no Excel clássico, vale a
 * célula na mesma linha (intervalo vertical) ou coluna (horizontal) da fórmula.
 */
export function intersect(v, ctx) {
  if (!(v instanceof RangeValue)) return v;
  if (v.rows * v.cols === 1) return v.values[0];
  if (v.origin && ctx) {
    const { row, col } = v.origin;
    if (v.cols === 1 && ctx.row >= row && ctx.row < row + v.rows) return v.get(ctx.row - row, 0);
    if (v.rows === 1 && ctx.col >= col && ctx.col < col + v.cols) return v.get(0, ctx.col - col);
  }
  return err(ERR.VALUE, 'Intervalo usado onde se espera um único valor');
}

export function power(base, exp) {
  if (base === 0 && exp === 0) return err(ERR.NUM, '0 elevado a 0 é indefinido');
  if (base === 0 && exp < 0) return err(ERR.DIV0);
  const r = base ** exp;
  return Number.isFinite(r) ? r : err(ERR.NUM);
}

// ---------------------------------------------------------------------------
// Datas: número de série com época em 30/12/1899 (como o Excel a partir de 01/03/1900)
// ---------------------------------------------------------------------------

const EPOCH = Date.UTC(1899, 11, 30);
const DAY = 86400000;
export const MAX_SERIAL = 2958465; // 31/12/9999

export function dateToSerial(year, month, day) {
  return Math.round((Date.UTC(year, month - 1, day) - EPOCH) / DAY);
}

export function serialToDate(serial) {
  const d = new Date(EPOCH + Math.floor(serial) * DAY);
  return { year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate() };
}

export function todaySerial(now = new Date()) {
  return dateToSerial(now.getFullYear(), now.getMonth() + 1, now.getDate());
}

function validDate(y, m, d) {
  if (m < 1 || m > 12 || d < 1) return false;
  const dim = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return d <= dim;
}

// ---------------------------------------------------------------------------
// Números no formato brasileiro
// ---------------------------------------------------------------------------

const reNumberBR = /^([+-]?)((?:\d{1,3}(?:\.\d{3})+|\d+)(?:,(\d*))?|,(\d+))(?:[eE]([+-]?\d+))?$/;

/** Converte "1.234,56", "-12,5", ",5" ou "1e3" em número; devolve null se não for número. */
export function parseNumberBR(text) {
  const m = reNumberBR.exec(text);
  if (!m) return null;
  const intPart = m[4] !== undefined ? '0' : m[2].split(',')[0].replace(/\./g, '');
  const frac = m[3] ?? m[4] ?? '';
  const n = Number(`${m[1]}${intPart}.${frac || '0'}e${m[5] ?? '0'}`);
  return Number.isFinite(n) ? n : null;
}

function decimalsTyped(text) {
  const m = /,(\d+)/.exec(text);
  return m ? Math.min(m[1].length, 10) : 0;
}

/**
 * Interpreta o texto digitado numa célula:
 * fórmula (=...), texto forçado ('...), número, moeda (R$), porcentagem, data,
 * VERDADEIRO/FALSO, código de erro ou texto comum.
 * Devolve { kind, value, format? } onde format é a formatação sugerida.
 */
export function parseInput(raw) {
  if (raw === null || raw === undefined || raw === '') return { kind: 'blank', value: null };
  if (raw[0] === '=' && raw.length > 1) return { kind: 'formula', value: null };
  if (raw[0] === "'") return { kind: 'text', value: raw.slice(1) };
  const t = raw.trim();
  if (t === '') return { kind: 'text', value: raw };
  const upper = t.toUpperCase();
  if (upper === 'VERDADEIRO' || upper === 'FALSO') return { kind: 'bool', value: upper === 'VERDADEIRO' };
  const code = ERROR_CODES.find((c) => c === upper);
  if (code) return { kind: 'error', value: err(code) };
  const n = parseNumberBR(t);
  if (n !== null) return { kind: 'number', value: n };
  // os testes abaixo só rodam quando o texto tem o símbolo esperado (evita regex lenta em textos longos)
  let m = t.endsWith('%') ? /^(.*?)\s*%$/.exec(t) : null;
  if (m) {
    const p = parseNumberBR(m[1]);
    if (p !== null) return { kind: 'number', value: p / 100, format: { type: 'porcentagem', decimals: decimalsTyped(m[1]) } };
  }
  m = /R\$/i.test(t) && t.length < 64 ? /^([+-]?)\s*R\$\s*([+-]?)\s*(.+)$/i.exec(t) : null;
  if (m) {
    const p = parseNumberBR(m[3]);
    if (p !== null && !/^[+-]/.test(m[3])) {
      const negative = (m[1] === '-') !== (m[2] === '-');
      return { kind: 'number', value: negative ? -p : p, format: { type: 'moeda', decimals: 2 } };
    }
  }
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{4}|\d{2})$/.exec(t);
  if (m) {
    let year = Number(m[3]);
    // dois dígitos: 00 a 29 viram 2000 a 2029, 30 a 99 viram 1930 a 1999 (regra do Excel)
    if (m[3].length === 2) year += year < 30 ? 2000 : 1900;
    const month = Number(m[2]);
    const day = Number(m[1]);
    if (validDate(year, month, day) && year >= 1900) {
      return { kind: 'number', value: dateToSerial(year, month, day), format: { type: 'data' } };
    }
  }
  m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(t);
  if (m) {
    const [year, month, day] = [Number(m[1]), Number(m[2]), Number(m[3])];
    if (validDate(year, month, day) && year >= 1900) {
      return { kind: 'number', value: dateToSerial(year, month, day), format: { type: 'data' } };
    }
  }
  return { kind: 'text', value: raw };
}

// ---------------------------------------------------------------------------
// Conversões usadas pelos operadores e funções
// ---------------------------------------------------------------------------

/** Número para texto no formato "Geral": até 15 algarismos significativos, vírgula decimal. */
export function generalNumber(n) {
  if (n === 0 || Object.is(n, -0)) return '0';
  const abs = Math.abs(n);
  if (abs >= 1e15 || abs < 1e-9) {
    const [mant, exp] = n.toExponential(14).split('e');
    const m = mant.replace(/\.?0+$/, '').replace('.', ',');
    const e = Number(exp);
    return `${m}E${e < 0 ? '-' : '+'}${String(Math.abs(e)).padStart(2, '0')}`;
  }
  let s = n.toPrecision(15);
  if (s.includes('e')) s = String(Number(s));
  if (s.includes('.')) s = s.replace(/0+$/, '').replace(/\.$/, '');
  return s.replace('.', ',');
}

export function toNumber(v) {
  if (typeof v === 'number') return v;
  if (v === null || v === undefined) return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (isError(v)) return v;
  if (v instanceof RangeValue) return v.rows * v.cols === 1 ? toNumber(v.values[0]) : err(ERR.VALUE);
  const parsed = parseInput(String(v));
  if (parsed.kind === 'number') return parsed.value;
  return err(ERR.VALUE, `"${v}" não é um número`);
}

export function toText(v) {
  if (typeof v === 'string') return v;
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'VERDADEIRO' : 'FALSO';
  if (typeof v === 'number') return generalNumber(v);
  if (isError(v)) return v;
  if (v instanceof RangeValue) return v.rows * v.cols === 1 ? toText(v.values[0]) : err(ERR.VALUE);
  return String(v);
}

export function toBool(v) {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v === null || v === undefined) return false;
  if (isError(v)) return v;
  if (v instanceof RangeValue) return v.rows * v.cols === 1 ? toBool(v.values[0]) : err(ERR.VALUE);
  const upper = String(v).trim().toUpperCase();
  if (upper === 'VERDADEIRO') return true;
  if (upper === 'FALSO') return false;
  return err(ERR.VALUE, `"${v}" não é VERDADEIRO nem FALSO`);
}

const collator = new Intl.Collator('pt-BR', { sensitivity: 'accent' });
const TYPE_ORDER = { number: 0, string: 1, boolean: 2 };

/**
 * Compara como o Excel: vazio vale 0, "" ou FALSO conforme o outro lado;
 * tipos diferentes seguem a ordem número < texto < lógico; texto ignora maiúsculas.
 * Devolve -1, 0 ou 1.
 */
export function compareValues(a, b) {
  if (a === null || a === undefined) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null || b === undefined) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ta = TYPE_ORDER[typeof a];
  const tb = TYPE_ORDER[typeof b];
  if (ta !== tb) return ta < tb ? -1 : 1;
  if (typeof a === 'string') {
    const c = collator.compare(a, b);
    return c < 0 ? -1 : c > 0 ? 1 : 0;
  }
  if (typeof a === 'boolean') return a === b ? 0 : a ? 1 : -1;
  return a < b ? -1 : a > b ? 1 : 0;
}

// ---------------------------------------------------------------------------
// Arredondamento decimal (evita 2,675 virar 2,67 por causa do binário)
// ---------------------------------------------------------------------------

function shiftDecimal(n, places) {
  const [m, e] = String(n).split('e');
  return Number(`${m}e${Number(e ?? 0) + places}`);
}

/** Limita a 15 algarismos significativos, como o Excel faz antes de arredondar. */
export function to15(n) {
  return n === 0 || !Number.isFinite(n) ? n : Number(n.toPrecision(15));
}

export function roundHalfAway(n, places) {
  const x = to15(n);
  const r = shiftDecimal(Math.round(shiftDecimal(Math.abs(x), places)), -places);
  return x < 0 ? -r : r;
}

export function roundUp(n, places) {
  const x = to15(n);
  const r = shiftDecimal(Math.ceil(shiftDecimal(Math.abs(x), places)), -places);
  return x < 0 ? -r : r;
}

export function roundDown(n, places) {
  const x = to15(n);
  const r = shiftDecimal(Math.floor(shiftDecimal(Math.abs(x), places)), -places);
  return x < 0 ? -r : r;
}

// ---------------------------------------------------------------------------
// Formatação para exibição
// ---------------------------------------------------------------------------

export const FORMAT_TYPES = ['geral', 'numero', 'moeda', 'porcentagem', 'data'];

function groupThousands(intDigits) {
  let out = '';
  for (let i = 0; i < intDigits.length; i++) {
    if (i > 0 && (intDigits.length - i) % 3 === 0) out += '.';
    out += intDigits[i];
  }
  return out;
}

/** Número com casas fixas, separador de milhar e vírgula decimal: 1234.5 vira "1.234,50". */
export function fixedBR(n, decimals, thousands = true) {
  if (Math.abs(n) >= 1e15) return generalNumber(n);
  const r = roundHalfAway(n, decimals);
  const [intPart, frac = ''] = Math.abs(r).toFixed(decimals).split('.');
  const sign = r < 0 ? '-' : '';
  return sign + (thousands ? groupThousands(intPart) : intPart) + (decimals > 0 ? ',' + frac : '');
}

export function formatDate(serial) {
  const { year, month, day } = serialToDate(serial);
  return `${String(day).padStart(2, '0')}/${String(month).padStart(2, '0')}/${String(year).padStart(4, '0')}`;
}

/**
 * Texto exibido para um valor com uma formatação.
 * format = { type: 'geral' | 'numero' | 'moeda' | 'porcentagem' | 'data', decimals? }
 */
export function formatValue(value, format) {
  if (value === null || value === undefined) return '';
  if (isError(value)) return value.code;
  if (typeof value === 'boolean') return value ? 'VERDADEIRO' : 'FALSO';
  if (typeof value === 'string') return value;
  const type = format?.type ?? 'geral';
  switch (type) {
    case 'numero':
      return fixedBR(value, format.decimals ?? 2);
    case 'moeda': {
      const text = fixedBR(Math.abs(value), format.decimals ?? 2);
      return (roundHalfAway(value, format.decimals ?? 2) < 0 ? '-' : '') + 'R$ ' + text;
    }
    case 'porcentagem':
      // como no Excel, porcentagem não usa separador de milhar (1000%)
      return fixedBR(value * 100, format.decimals ?? 0, false) + '%';
    case 'data':
      return value >= 1 && value <= MAX_SERIAL ? formatDate(value) : generalNumber(value);
    default:
      return generalNumber(value);
  }
}

/** Alinhamento padrão: números à direita, texto à esquerda, lógicos e erros ao centro. */
export function alignmentOf(value) {
  if (typeof value === 'number') return 'right';
  if (typeof value === 'boolean' || isError(value)) return 'center';
  return 'left';
}
