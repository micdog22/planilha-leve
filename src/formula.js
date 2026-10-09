// Fórmulas no estilo do Excel em português: tokenizador, parser e utilidades de
// referência (endereços A1, deslocamento ao copiar e alternância de $ com F4).
//
// Sintaxe: "=" seguido de uma expressão. Ponto e vírgula separa argumentos, vírgula
// é o separador decimal. Operadores, do mais forte para o mais fraco:
//   - unário   %   ^   * /   + -   &   = <> < > <= >=
// Como no Excel, o menos unário vem antes da potência (=-2^2 dá 4) e ^ associa à esquerda.

import { ERROR_CODES } from './values.js';

export const MAX_ROWS = 1048576;
export const MAX_COLS = 16384;

// ---------------------------------------------------------------------------
// Endereços
// ---------------------------------------------------------------------------

/** Índice de coluna (base 0) para letras: 0 vira "A", 26 vira "AA". */
export function colName(col) {
  let n = col + 1;
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Letras para índice de coluna (base 0); devolve -1 se inválido. */
export function colIndex(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    const code = ch.charCodeAt(0);
    if (code < 65 || code > 90) return -1;
    n = n * 26 + (code - 64);
  }
  return n - 1;
}

export function cellName(row, col) {
  return colName(col) + (row + 1);
}

/** "B3" vira { row: 2, col: 1 }; devolve null se não for um endereço válido. */
export function parseCellName(text) {
  const m = /^\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})$/.exec(text.trim());
  if (!m) return null;
  const col = colIndex(m[1]);
  const row = Number(m[2]) - 1;
  if (col >= MAX_COLS || row >= MAX_ROWS) return null;
  return { row, col };
}

function refText(ref) {
  return (ref.absCol ? '$' : '') + colName(ref.col) + (ref.absRow ? '$' : '') + (ref.row + 1);
}

// ---------------------------------------------------------------------------
// Tokenizador
// ---------------------------------------------------------------------------

export class FormulaSyntaxError extends Error {
  constructor(message, position) {
    super(message);
    this.position = position;
  }
}

const reWhitespace = /\s+/y;
const reNumber = /(?:\d+(?:,\d*)?|,\d+)(?:[eE][+-]?\d+)?/y;
const reRef = /(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d{0,6})/y;
const reIdent = /[\p{L}_][\p{L}\p{N}_.]*/uy;
const reIdentChar = /[\p{L}\p{N}_.(]/u;
const OPERATORS = ['<>', '<=', '>=', '+', '-', '*', '/', '^', '&', '%', '=', '<', '>'];

function readRef(src, i) {
  reRef.lastIndex = i;
  const m = reRef.exec(src);
  if (!m) return null;
  const end = reRef.lastIndex;
  // "LOG10(" e "A1B" não são referências
  if (end < src.length && reIdentChar.test(src[end])) return null;
  const col = colIndex(m[2]);
  const row = Number(m[4]) - 1;
  if (col >= MAX_COLS || row >= MAX_ROWS) return null;
  return { ref: { row, col, absRow: m[3] === '$', absCol: m[1] === '$' }, end };
}

/**
 * Quebra o corpo da fórmula (sem o "=") em tokens com posição de início e fim.
 * Tipos: num, str, bool, err, ref, range, func, name, op, "(", ")", ";".
 */
export function tokenize(src) {
  const tokens = [];
  let i = 0;
  while (i < src.length) {
    reWhitespace.lastIndex = i;
    if (reWhitespace.exec(src)) {
      i = reWhitespace.lastIndex;
      continue;
    }
    const ch = src[i];
    const start = i;
    if ((ch >= '0' && ch <= '9') || (ch === ',' && /\d/.test(src[i + 1] ?? ''))) {
      reNumber.lastIndex = i;
      const m = reNumber.exec(src);
      const text = m[0].replace(',', '.');
      const value = Number(text.startsWith('.') ? '0' + text : text);
      if (!Number.isFinite(value)) throw new FormulaSyntaxError('Número fora do intervalo', start);
      i = reNumber.lastIndex;
      tokens.push({ t: 'num', v: value, s: start, e: i });
      continue;
    }
    if (ch === '"') {
      let value = '';
      i += 1;
      for (;;) {
        if (i >= src.length) throw new FormulaSyntaxError('Texto sem aspas de fechamento', start);
        if (src[i] === '"') {
          if (src[i + 1] === '"') {
            value += '"';
            i += 2;
            continue;
          }
          i += 1;
          break;
        }
        value += src[i];
        i += 1;
      }
      tokens.push({ t: 'str', v: value, s: start, e: i });
      continue;
    }
    if (ch === '#') {
      const upper = src.slice(i, i + 8).toUpperCase();
      const code = ERROR_CODES.find((c) => upper.startsWith(c));
      if (!code) throw new FormulaSyntaxError('Código de erro desconhecido', start);
      i += code.length;
      tokens.push({ t: 'err', v: code, s: start, e: i });
      continue;
    }
    if (ch === '$' || /[\p{L}_]/u.test(ch)) {
      const r1 = readRef(src, i);
      if (r1) {
        if (src[r1.end] === ':') {
          const r2 = readRef(src, r1.end + 1);
          if (!r2) throw new FormulaSyntaxError('Intervalo incompleto', start);
          tokens.push({ t: 'range', v: [r1.ref, r2.ref], s: start, e: r2.end });
          i = r2.end;
          continue;
        }
        tokens.push({ t: 'ref', v: r1.ref, s: start, e: r1.end });
        i = r1.end;
        continue;
      }
      if (ch === '$') throw new FormulaSyntaxError('Referência inválida', start);
      reIdent.lastIndex = i;
      const m = reIdent.exec(src);
      i = reIdent.lastIndex;
      let j = i;
      while (j < src.length && /\s/.test(src[j])) j++;
      if (src[j] === '(') {
        tokens.push({ t: 'func', v: m[0], s: start, e: j + 1 });
        i = j + 1;
        continue;
      }
      const upper = m[0].toUpperCase();
      if (upper === 'VERDADEIRO' || upper === 'FALSO') tokens.push({ t: 'bool', v: upper === 'VERDADEIRO', s: start, e: i });
      else tokens.push({ t: 'name', v: m[0], s: start, e: i });
      continue;
    }
    if (ch === '(' || ch === ')' || ch === ';') {
      tokens.push({ t: ch, v: ch, s: start, e: i + 1 });
      i += 1;
      continue;
    }
    const op = OPERATORS.find((o) => src.startsWith(o, i));
    if (op) {
      tokens.push({ t: 'op', v: op, s: start, e: i + op.length });
      i += op.length;
      continue;
    }
    throw new FormulaSyntaxError(`Caractere inesperado "${ch}"`, start);
  }
  return tokens;
}

// ---------------------------------------------------------------------------
// Parser (Pratt): cada operador tem uma força de ligação
// ---------------------------------------------------------------------------

const BINARY_BP = {
  '=': 10, '<>': 10, '<': 10, '>': 10, '<=': 10, '>=': 10,
  '&': 20,
  '+': 30, '-': 30,
  '*': 40, '/': 40,
  '^': 50,
};
const PERCENT_BP = 60;
const UNARY_BP = 70;

class Parser {
  constructor(tokens, length) {
    this.tokens = tokens;
    this.i = 0;
    this.length = length;
  }

  peek() {
    return this.tokens[this.i];
  }

  next() {
    return this.tokens[this.i++];
  }

  where() {
    const t = this.peek();
    return t ? t.s : this.length;
  }

  parse() {
    if (this.tokens.length === 0) throw new FormulaSyntaxError('Fórmula vazia', 0);
    const expr = this.expression(0);
    if (this.i < this.tokens.length) throw new FormulaSyntaxError('Sobrou texto depois da expressão', this.where());
    return expr;
  }

  expression(minBp) {
    let left = this.prefix();
    for (;;) {
      const t = this.peek();
      if (!t || t.t !== 'op') break;
      if (t.v === '%') {
        if (PERCENT_BP < minBp) break;
        this.next();
        left = { type: 'percent', arg: left };
        continue;
      }
      const bp = BINARY_BP[t.v];
      if (bp === undefined || bp < minBp) break;
      this.next();
      // bp + 1 deixa todos os binários associativos à esquerda (inclusive ^, como no Excel)
      const right = this.expression(bp + 1);
      left = { type: 'binary', op: t.v, left, right };
    }
    return left;
  }

  prefix() {
    const t = this.next();
    if (!t) throw new FormulaSyntaxError('Expressão incompleta', this.length);
    switch (t.t) {
      case 'num':
        return { type: 'num', value: t.v };
      case 'str':
        return { type: 'str', value: t.v };
      case 'bool':
        return { type: 'bool', value: t.v };
      case 'err':
        return { type: 'err', code: t.v };
      case 'ref':
        return { type: 'ref', ...t.v };
      case 'range': {
        const [a, b] = t.v;
        return {
          type: 'range',
          r1: Math.min(a.row, b.row), c1: Math.min(a.col, b.col),
          r2: Math.max(a.row, b.row), c2: Math.max(a.col, b.col),
        };
      }
      case 'name':
        return { type: 'name', name: t.v };
      case 'op':
        if (t.v === '-' || t.v === '+') return { type: 'unary', op: t.v, arg: this.expression(UNARY_BP) };
        throw new FormulaSyntaxError(`Operador "${t.v}" fora de lugar`, t.s);
      case '(': {
        const inner = this.expression(0);
        const close = this.next();
        if (!close || close.t !== ')') throw new FormulaSyntaxError('Faltou fechar parêntese', close ? close.s : this.length);
        return inner;
      }
      case 'func':
        return this.call(t);
      default:
        throw new FormulaSyntaxError(`"${t.v}" fora de lugar`, t.s);
    }
  }

  call(t) {
    const args = [];
    if (this.peek()?.t === ')') {
      this.next();
      return { type: 'call', name: t.v, args };
    }
    for (;;) {
      const p = this.peek();
      if (!p) throw new FormulaSyntaxError(`Faltou fechar parêntese de ${t.v}`, this.length);
      if (p.t === ';' || p.t === ')') args.push({ type: 'empty' });
      else args.push(this.expression(0));
      const sep = this.next();
      if (!sep) throw new FormulaSyntaxError(`Faltou fechar parêntese de ${t.v}`, this.length);
      if (sep.t === ')') break;
      if (sep.t !== ';') throw new FormulaSyntaxError('Esperava ";" ou ")" entre os argumentos', sep.s);
    }
    return { type: 'call', name: t.v, args };
  }
}

/** Analisa uma fórmula (com ou sem o "=" inicial) e devolve a árvore. */
export function parseFormula(text) {
  const body = text[0] === '=' ? text.slice(1) : text;
  return new Parser(tokenize(body), body.length).parse();
}

/** Referências usadas pela fórmula: células avulsas e intervalos. */
export function collectRefs(ast) {
  const cells = [];
  const ranges = [];
  const stack = [ast];
  while (stack.length) {
    const node = stack.pop();
    switch (node.type) {
      case 'ref':
        cells.push({ row: node.row, col: node.col });
        break;
      case 'range':
        ranges.push({ r1: node.r1, c1: node.c1, r2: node.r2, c2: node.c2 });
        break;
      case 'unary':
      case 'percent':
        stack.push(node.arg);
        break;
      case 'binary':
        stack.push(node.left, node.right);
        break;
      case 'call':
        for (const a of node.args) stack.push(a);
        break;
      default:
        break;
    }
  }
  return { cells, ranges };
}

// ---------------------------------------------------------------------------
// Copiar e colar: deslocamento de referências relativas
// ---------------------------------------------------------------------------

function shiftRef(ref, dRow, dCol) {
  const row = ref.absRow ? ref.row : ref.row + dRow;
  const col = ref.absCol ? ref.col : ref.col + dCol;
  if (row < 0 || col < 0 || row >= MAX_ROWS || col >= MAX_COLS) return null;
  return { ...ref, row, col };
}

/**
 * Ajusta as referências relativas de uma fórmula copiada dRow linhas e dCol colunas
 * adiante. Partes com $ ficam fixas; referências que saem da planilha viram #REF!.
 * O resto do texto (espaços, textos entre aspas) é mantido como foi digitado.
 */
export function shiftFormula(text, dRow, dCol) {
  if (text[0] !== '=' || (dRow === 0 && dCol === 0)) return text;
  const body = text.slice(1);
  let tokens;
  try {
    tokens = tokenize(body);
  } catch {
    return text;
  }
  let out = '=';
  let pos = 0;
  for (const t of tokens) {
    if (t.t !== 'ref' && t.t !== 'range') continue;
    out += body.slice(pos, t.s);
    if (t.t === 'ref') {
      const r = shiftRef(t.v, dRow, dCol);
      out += r ? refText(r) : '#REF!';
    } else {
      const a = shiftRef(t.v[0], dRow, dCol);
      const b = shiftRef(t.v[1], dRow, dCol);
      out += a && b ? `${refText(a)}:${refText(b)}` : '#REF!';
    }
    pos = t.e;
  }
  return out + body.slice(pos);
}

/**
 * F4: alterna a referência sob o cursor entre A1, $A$1, A$1 e $A1.
 * caret é a posição no texto completo (com o "="). Devolve { text, caret } ou null.
 */
export function cycleReference(text, caret) {
  if (text[0] !== '=') return null;
  const body = text.slice(1);
  let tokens;
  try {
    tokens = tokenize(body);
  } catch {
    return null;
  }
  const pos = caret - 1;
  const t = tokens.find((tok) => (tok.t === 'ref' || tok.t === 'range') && pos >= tok.s && pos <= tok.e);
  if (!t) return null;
  const first = t.t === 'ref' ? t.v : t.v[0];
  // A1 -> $A$1 -> A$1 -> $A1 -> A1
  const states = [[false, false], [true, true], [false, true], [true, false]];
  const k = states.findIndex(([c, r]) => c === first.absCol && r === first.absRow);
  const [absCol, absRow] = states[(k + 1) % states.length];
  const apply = (r) => ({ ...r, absCol, absRow });
  const replacement = t.t === 'ref' ? refText(apply(t.v)) : `${refText(apply(t.v[0]))}:${refText(apply(t.v[1]))}`;
  const newBody = body.slice(0, t.s) + replacement + body.slice(t.e);
  return { text: '=' + newBody, caret: 1 + t.s + replacement.length };
}

/** Diz se o cursor está num ponto da fórmula onde clicar numa célula insere a referência dela. */
export function canInsertReference(text, caret) {
  if (text[0] !== '=' || caret < 1) return false;
  let i = caret - 1;
  while (i > 0 && text[i] === ' ') i--;
  if (i <= 0) return true;
  return '=(;+-*/^&<>:'.includes(text[i]);
}
