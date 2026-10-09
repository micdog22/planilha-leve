// Avaliação da árvore de uma fórmula. Erros são valores (não exceções): cada
// operador devolve o primeiro erro que encontra, da esquerda para a direita.

import {
  ERR, err, isError, toNumber, toText, compareValues, intersect, power,
} from './values.js';
import { lookupFunction } from './functions.js';

function arithmetic(op, a, b) {
  const x = toNumber(a);
  if (isError(x)) return x;
  const y = toNumber(b);
  if (isError(y)) return y;
  let r;
  switch (op) {
    case '+': r = x + y; break;
    case '-': r = x - y; break;
    case '*': r = x * y; break;
    case '/':
      if (y === 0) return err(ERR.DIV0);
      r = x / y;
      break;
    default:
      return power(x, y);
  }
  return Number.isFinite(r) ? r : err(ERR.NUM);
}

const COMPARE = {
  '=': (c) => c === 0,
  '<>': (c) => c !== 0,
  '<': (c) => c < 0,
  '>': (c) => c > 0,
  '<=': (c) => c <= 0,
  '>=': (c) => c >= 0,
};

/** Avalia um argumento de função: referências e intervalos chegam como RangeValue. */
export function evaluateArg(node, ctx) {
  if (node.type === 'empty') return null;
  if (node.type === 'ref') return ctx.range(node.row, node.col, node.row, node.col);
  return evaluate(node, ctx);
}

export function evaluate(node, ctx) {
  switch (node.type) {
    case 'num':
    case 'str':
    case 'bool':
      return node.value;
    case 'err':
      return err(node.code);
    case 'empty':
      return null;
    case 'ref':
      return ctx.cell(node.row, node.col);
    case 'range':
      return ctx.range(node.r1, node.c1, node.r2, node.c2);
    case 'name':
      return err(ERR.NAME, `Nome desconhecido: ${node.name}`);
    case 'unary': {
      const v = intersect(evaluate(node.arg, ctx), ctx);
      if (isError(v)) return v;
      if (node.op === '+') return v;
      const n = toNumber(v);
      return isError(n) ? n : -n;
    }
    case 'percent': {
      const n = toNumber(intersect(evaluate(node.arg, ctx), ctx));
      return isError(n) ? n : n / 100;
    }
    case 'binary': {
      const a = intersect(evaluate(node.left, ctx), ctx);
      if (isError(a)) return a;
      const b = intersect(evaluate(node.right, ctx), ctx);
      if (isError(b)) return b;
      if (node.op === '&') {
        const ta = toText(a);
        if (isError(ta)) return ta;
        const tb = toText(b);
        return isError(tb) ? tb : ta + tb;
      }
      const cmp = COMPARE[node.op];
      if (cmp) return cmp(compareValues(a, b));
      return arithmetic(node.op, a, b);
    }
    case 'call': {
      const fn = lookupFunction(node.name);
      if (!fn) return err(ERR.NAME, `Função desconhecida: ${node.name}`);
      const n = node.args.length;
      if (n < fn.min || n > fn.max) {
        const expected = fn.min === fn.max ? `${fn.min}` : fn.max === Infinity ? `pelo menos ${fn.min}` : `de ${fn.min} a ${fn.max}`;
        return err(ERR.NA, `${fn.name} espera ${expected} argumento(s), recebeu ${n}`);
      }
      if (fn.lazy) return fn.run(node.args, ctx, (arg) => evaluateArg(arg, ctx));
      return fn.run(node.args.map((arg) => evaluateArg(arg, ctx)), ctx);
    }
    default:
      return err(ERR.VALUE);
  }
}
