// Biblioteca de funções com nomes em português, seguindo o comportamento do Excel
// nos casos comuns. Cada função declara a aridade, a sintaxe e uma descrição curta
// (usadas também na ajuda da interface).
//
// Regra importante do Excel reproduzida aqui: em SOMA, MÉDIA, MÍNIMO e MÁXIMO,
// texto e valores lógicos dentro de intervalos são ignorados, mas texto passado
// direto como argumento é convertido (e dá #VALOR! se não for número).

import {
  ERR, err, isError, RangeValue, toNumber, toText, toBool, compareValues, parseInput,
  parseNumberBR, roundHalfAway, roundUp, roundDown, intersect, power,
} from './values.js';

const collator = new Intl.Collator('pt-BR', { sensitivity: 'accent' });

function scalar(v, ctx) {
  return intersect(v, ctx);
}

function numberArg(v, ctx) {
  const s = scalar(v, ctx);
  return isError(s) ? s : toNumber(s);
}

function textArg(v, ctx) {
  const s = scalar(v, ctx);
  return isError(s) ? s : toText(s);
}

/** Números para SOMA, MÉDIA, MÍNIMO e MÁXIMO (ou o primeiro erro encontrado). */
function aggregateNumbers(args) {
  const out = [];
  for (const a of args) {
    if (a instanceof RangeValue) {
      for (const v of a) {
        if (isError(v)) return v;
        if (typeof v === 'number') out.push(v);
      }
    } else {
      const n = toNumber(a);
      if (isError(n)) return n;
      out.push(n);
    }
  }
  return out;
}

function sum(values) {
  // soma compensada (Neumaier): reduz o erro de arredondamento em listas longas
  let s = 0;
  let c = 0;
  for (const v of values) {
    const t = s + v;
    c += Math.abs(s) >= Math.abs(v) ? (s - t) + v : (v - t) + s;
    s = t;
  }
  return s + c;
}

function finite(n) {
  return Number.isFinite(n) ? n : err(ERR.NUM);
}

// --------------------------------------------------------------------------
// Critérios de CONT.SE e SOMASE: ">5", "<>0", "=abc", "ab*", 10, VERDADEIRO
// --------------------------------------------------------------------------

function wildcardRegex(pattern) {
  let re = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length && '*?~'.includes(pattern[i + 1])) {
      re += '\\' + pattern[++i];
    } else if (ch === '*') re += '[\\s\\S]*';
    else if (ch === '?') re += '[\\s\\S]';
    else re += ch.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&');
  }
  return new RegExp(`^${re}$`, 'iu');
}

const OPS = {
  '=': (c) => c === 0,
  '<>': (c) => c !== 0,
  '<': (c) => c < 0,
  '>': (c) => c > 0,
  '<=': (c) => c <= 0,
  '>=': (c) => c >= 0,
};

/** Transforma o critério numa função que testa cada célula. */
export function makeCriterion(criterion) {
  if (isError(criterion)) return (v) => isError(v) && v.code === criterion.code;
  if (criterion === null) criterion = 0;
  if (typeof criterion === 'boolean') return (v) => v === criterion;
  if (typeof criterion === 'number') return numericMatcher('=', criterion);
  const s = String(criterion);
  const m = /^(<=|>=|<>|=|<|>)/.exec(s);
  const op = m ? m[1] : '=';
  const rest = m ? s.slice(op.length) : s;
  if (rest === '') {
    if (op === '=') return (v) => v === null || v === '';
    if (op === '<>') return (v) => v !== null && v !== '';
    return () => false;
  }
  const parsed = parseInput(rest);
  if (parsed.kind === 'number') return numericMatcher(op, parsed.value);
  if (parsed.kind === 'bool') return (v) => (typeof v === 'boolean' ? OPS[op](compareValues(v, parsed.value)) : op === '<>');
  if (parsed.kind === 'error') return (v) => (isError(v) && v.code === parsed.value.code) === (op === '=');
  if (op === '=' || op === '<>') {
    const re = wildcardRegex(rest);
    const hit = (v) => typeof v === 'string' && re.test(v);
    return op === '=' ? hit : (v) => !hit(v);
  }
  return (v) => typeof v === 'string' && OPS[op](compareValues(v, rest));
}

function numericMatcher(op, x) {
  return (v) => {
    let n = null;
    if (typeof v === 'number') n = v;
    else if (typeof v === 'string') n = parseNumberBR(v.trim());
    if (n === null) return op === '<>';
    return OPS[op](n < x ? -1 : n > x ? 1 : 0);
  };
}

// --------------------------------------------------------------------------
// PROCV
// --------------------------------------------------------------------------

function exactMatch(v, lookup, re) {
  if (typeof lookup === 'string') {
    if (typeof v !== 'string') return false;
    return re ? re.test(v) : collator.compare(v, lookup) === 0;
  }
  return typeof v === typeof lookup && v === lookup;
}

function vlookup(lookup, table, colIndex, approx) {
  const rows = table.rows;
  if (!approx) {
    const re = typeof lookup === 'string' && /[*?~]/.test(lookup) ? wildcardRegex(lookup) : null;
    for (let r = 0; r < rows; r++) {
      if (exactMatch(table.get(r, 0), lookup, re)) return table.get(r, colIndex);
    }
    return err(ERR.NA, 'Valor procurado não encontrado');
  }
  // busca binária: a primeira coluna precisa estar em ordem crescente
  let lo = 0;
  let hi = rows - 1;
  let best = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const v = table.get(mid, 0);
    if (v === null || compareValues(v, lookup) <= 0) {
      if (v !== null && typeof v === typeof lookup) best = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  return best === -1 ? err(ERR.NA, 'Nenhum valor menor ou igual ao procurado') : table.get(best, colIndex);
}

// --------------------------------------------------------------------------
// Tabela de funções
// --------------------------------------------------------------------------

const defs = [
  {
    name: 'SOMA', min: 1, max: 255, syntax: 'SOMA(núm1; [núm2]; ...)', desc: 'Soma números e intervalos.',
    run(args) {
      const nums = aggregateNumbers(args);
      return isError(nums) ? nums : finite(sum(nums));
    },
  },
  {
    name: 'MÉDIA', min: 1, max: 255, syntax: 'MÉDIA(núm1; [núm2]; ...)', desc: 'Média aritmética dos números.',
    run(args) {
      const nums = aggregateNumbers(args);
      if (isError(nums)) return nums;
      return nums.length === 0 ? err(ERR.DIV0, 'Nenhum número para calcular a média') : finite(sum(nums) / nums.length);
    },
  },
  {
    name: 'MÍNIMO', min: 1, max: 255, syntax: 'MÍNIMO(núm1; [núm2]; ...)', desc: 'Menor número (0 se não houver números).',
    run(args) {
      const nums = aggregateNumbers(args);
      if (isError(nums)) return nums;
      return nums.length === 0 ? 0 : nums.reduce((a, b) => (b < a ? b : a));
    },
  },
  {
    name: 'MÁXIMO', min: 1, max: 255, syntax: 'MÁXIMO(núm1; [núm2]; ...)', desc: 'Maior número (0 se não houver números).',
    run(args) {
      const nums = aggregateNumbers(args);
      if (isError(nums)) return nums;
      return nums.length === 0 ? 0 : nums.reduce((a, b) => (b > a ? b : a));
    },
  },
  {
    name: 'CONT.NÚM', min: 1, max: 255, syntax: 'CONT.NÚM(valor1; [valor2]; ...)', desc: 'Conta quantos valores são números.',
    run(args) {
      let n = 0;
      for (const a of args) {
        if (a instanceof RangeValue) {
          for (const v of a) if (typeof v === 'number') n++;
        } else if (typeof a === 'number' || typeof a === 'boolean') n++;
        else if (typeof a === 'string' && parseInput(a).kind === 'number') n++;
      }
      return n;
    },
  },
  {
    name: 'CONT.VALORES', min: 1, max: 255, syntax: 'CONT.VALORES(valor1; [valor2]; ...)', desc: 'Conta células que não estão vazias.',
    run(args) {
      let n = 0;
      for (const a of args) {
        if (a instanceof RangeValue) {
          for (const v of a) if (v !== null) n++;
        } else if (a !== null) n++;
      }
      return n;
    },
  },
  {
    name: 'CONT.SE', min: 2, max: 2, syntax: 'CONT.SE(intervalo; critério)', desc: 'Conta as células que atendem ao critério (ex.: ">10", "ab*").',
    run([range, criterion], ctx) {
      if (!(range instanceof RangeValue)) return err(ERR.VALUE, 'O primeiro argumento precisa ser um intervalo');
      const c = scalar(criterion, ctx);
      const test = makeCriterion(c);
      let n = 0;
      for (const v of range) if (test(v)) n++;
      return n;
    },
  },
  {
    name: 'SOMASE', min: 2, max: 3, syntax: 'SOMASE(intervalo; critério; [intervalo_soma])', desc: 'Soma as células cujo vizinho atende ao critério.',
    run([range, criterion, sumRange], ctx) {
      if (!(range instanceof RangeValue)) return err(ERR.VALUE, 'O primeiro argumento precisa ser um intervalo');
      const target = sumRange === undefined || sumRange === null ? range : sumRange;
      if (!(target instanceof RangeValue)) return err(ERR.VALUE, 'O intervalo de soma precisa ser um intervalo');
      const test = makeCriterion(scalar(criterion, ctx));
      const values = [];
      for (let r = 0; r < range.rows; r++) {
        for (let c = 0; c < range.cols; c++) {
          if (!test(range.get(r, c))) continue;
          if (r >= target.rows || c >= target.cols) continue;
          const v = target.get(r, c);
          if (isError(v)) return v;
          if (typeof v === 'number') values.push(v);
        }
      }
      return finite(sum(values));
    },
  },
  {
    name: 'SE', min: 2, max: 3, lazy: true, syntax: 'SE(teste; valor_se_verdadeiro; [valor_se_falso])', desc: 'Escolhe um valor conforme o teste. Só o ramo escolhido é calculado.',
    run(nodes, ctx, ev) {
      const cond = toBool(scalar(ev(nodes[0]), ctx));
      if (isError(cond)) return cond;
      if (cond) return nodes[1].type === 'empty' ? 0 : scalar(ev(nodes[1]), ctx);
      if (nodes.length < 3) return false;
      return nodes[2].type === 'empty' ? 0 : scalar(ev(nodes[2]), ctx);
    },
  },
  {
    name: 'SEERRO', min: 2, max: 2, lazy: true, syntax: 'SEERRO(valor; valor_se_erro)', desc: 'Devolve o valor ou, se ele for um erro, a alternativa.',
    run(nodes, ctx, ev) {
      const v = scalar(ev(nodes[0]), ctx);
      if (!isError(v)) return v;
      return nodes[1].type === 'empty' ? 0 : scalar(ev(nodes[1]), ctx);
    },
  },
  ...['E', 'OU'].map((name) => ({
    name, min: 1, max: 255,
    syntax: `${name}(lógico1; [lógico2]; ...)`,
    desc: name === 'E' ? 'VERDADEIRO se todos os argumentos forem verdadeiros.' : 'VERDADEIRO se pelo menos um argumento for verdadeiro.',
    run(args) {
      const vals = [];
      for (const a of args) {
        if (a instanceof RangeValue) {
          for (const v of a) {
            if (isError(v)) return v;
            if (typeof v === 'boolean') vals.push(v);
            else if (typeof v === 'number') vals.push(v !== 0);
          }
        } else if (a !== null) {
          const b = toBool(a);
          if (isError(b)) return b;
          vals.push(b);
        }
      }
      if (vals.length === 0) return err(ERR.VALUE, 'Nenhum valor lógico nos argumentos');
      return name === 'E' ? vals.every(Boolean) : vals.some(Boolean);
    },
  })),
  {
    name: 'NÃO', min: 1, max: 1, syntax: 'NÃO(lógico)', desc: 'Inverte um valor lógico.',
    run([a], ctx) {
      const b = toBool(scalar(a, ctx));
      return isError(b) ? b : !b;
    },
  },
  ...[['ARRED', roundHalfAway, 'Arredonda (meio para longe do zero).'],
    ['ARREDONDAR.PARA.CIMA', roundUp, 'Arredonda para longe do zero.'],
    ['ARREDONDAR.PARA.BAIXO', roundDown, 'Arredonda em direção ao zero.']].map(([name, fn, desc]) => ({
    name, min: 2, max: 2, syntax: `${name}(número; núm_dígitos)`, desc,
    run([a, d], ctx) {
      const n = numberArg(a, ctx);
      if (isError(n)) return n;
      const places = numberArg(d, ctx);
      if (isError(places)) return places;
      return finite(fn(n, Math.trunc(places)));
    },
  })),
  {
    name: 'ABS', min: 1, max: 1, syntax: 'ABS(número)', desc: 'Valor absoluto.',
    run([a], ctx) {
      const n = numberArg(a, ctx);
      return isError(n) ? n : Math.abs(n);
    },
  },
  {
    name: 'RAIZ', min: 1, max: 1, syntax: 'RAIZ(número)', desc: 'Raiz quadrada (número negativo dá #NÚM!).',
    run([a], ctx) {
      const n = numberArg(a, ctx);
      if (isError(n)) return n;
      return n < 0 ? err(ERR.NUM, 'Raiz de número negativo') : Math.sqrt(n);
    },
  },
  {
    name: 'POTÊNCIA', min: 2, max: 2, syntax: 'POTÊNCIA(número; potência)', desc: 'Eleva o número à potência (o mesmo que ^).',
    run([a, b], ctx) {
      const x = numberArg(a, ctx);
      if (isError(x)) return x;
      const y = numberArg(b, ctx);
      return isError(y) ? y : power(x, y);
    },
  },
  {
    name: 'CONCATENAR', min: 1, max: 255, syntax: 'CONCATENAR(texto1; [texto2]; ...)', desc: 'Junta textos (o mesmo que &).',
    run(args) {
      let out = '';
      for (const a of args) {
        const items = a instanceof RangeValue ? a.values : [a];
        for (const v of items) {
          const t = toText(v);
          if (isError(t)) return t;
          out += t;
        }
      }
      return out;
    },
  },
  {
    name: 'ESQUERDA', min: 1, max: 2, syntax: 'ESQUERDA(texto; [núm_caract])', desc: 'Primeiros caracteres do texto.',
    run([a, n], ctx) {
      const t = textArg(a, ctx);
      if (isError(t)) return t;
      const k = n === undefined ? 1 : numberArg(n, ctx);
      if (isError(k)) return k;
      if (k < 0) return err(ERR.VALUE, 'Quantidade negativa');
      return Array.from(t).slice(0, Math.trunc(k)).join('');
    },
  },
  {
    name: 'DIREITA', min: 1, max: 2, syntax: 'DIREITA(texto; [núm_caract])', desc: 'Últimos caracteres do texto.',
    run([a, n], ctx) {
      const t = textArg(a, ctx);
      if (isError(t)) return t;
      const k = n === undefined ? 1 : numberArg(n, ctx);
      if (isError(k)) return k;
      if (k < 0) return err(ERR.VALUE, 'Quantidade negativa');
      const chars = Array.from(t);
      const count = Math.min(chars.length, Math.trunc(k));
      return chars.slice(chars.length - count).join('');
    },
  },
  {
    name: 'EXT.TEXTO', min: 3, max: 3, syntax: 'EXT.TEXTO(texto; início; núm_caract)', desc: 'Trecho do texto a partir de uma posição (a primeira é 1).',
    run([a, s, n], ctx) {
      const t = textArg(a, ctx);
      if (isError(t)) return t;
      const start = numberArg(s, ctx);
      if (isError(start)) return start;
      const k = numberArg(n, ctx);
      if (isError(k)) return k;
      if (start < 1) return err(ERR.VALUE, 'A posição inicial precisa ser pelo menos 1');
      if (k < 0) return err(ERR.VALUE, 'Quantidade negativa');
      const from = Math.trunc(start) - 1;
      return Array.from(t).slice(from, from + Math.trunc(k)).join('');
    },
  },
  {
    name: 'NÚM.CARACT', min: 1, max: 1, syntax: 'NÚM.CARACT(texto)', desc: 'Quantidade de caracteres.',
    run([a], ctx) {
      const t = textArg(a, ctx);
      return isError(t) ? t : Array.from(t).length;
    },
  },
  {
    name: 'MAIÚSCULA', min: 1, max: 1, syntax: 'MAIÚSCULA(texto)', desc: 'Texto em maiúsculas.',
    run([a], ctx) {
      const t = textArg(a, ctx);
      return isError(t) ? t : t.toLocaleUpperCase('pt-BR');
    },
  },
  {
    name: 'MINÚSCULA', min: 1, max: 1, syntax: 'MINÚSCULA(texto)', desc: 'Texto em minúsculas.',
    run([a], ctx) {
      const t = textArg(a, ctx);
      return isError(t) ? t : t.toLocaleLowerCase('pt-BR');
    },
  },
  {
    name: 'PROCV', min: 3, max: 4, syntax: 'PROCV(valor; tabela; núm_coluna; [aproximado])',
    desc: 'Procura na primeira coluna e devolve a coluna pedida. Use FALSO no último argumento para busca exata.',
    run([value, table, col, approx], ctx) {
      const lookup = scalar(value, ctx);
      if (isError(lookup)) return lookup;
      if (lookup === null) return err(ERR.NA, 'Valor procurado vazio');
      if (!(table instanceof RangeValue)) return err(ERR.VALUE, 'A tabela precisa ser um intervalo');
      const c = numberArg(col, ctx);
      if (isError(c)) return c;
      const index = Math.trunc(c);
      if (index < 1) return err(ERR.VALUE, 'Número de coluna menor que 1');
      if (index > table.cols) return err(ERR.REF, 'Número de coluna maior que a tabela');
      const mode = approx === undefined ? true : toBool(scalar(approx, ctx));
      if (isError(mode)) return mode;
      return vlookup(lookup, table, index - 1, mode);
    },
  },
  {
    name: 'HOJE', min: 0, max: 0, syntax: 'HOJE()', desc: 'Data de hoje (recalculada ao abrir a planilha).',
    run(args, ctx) {
      return ctx.today();
    },
  },
];

function nameKey(name) {
  return name.normalize('NFD').replace(/\p{M}/gu, '').toUpperCase();
}

const TABLE = new Map(defs.map((d) => [nameKey(d.name), d]));

/** Procura uma função pelo nome, sem diferenciar maiúsculas nem acentos (MEDIA vale MÉDIA). */
export function lookupFunction(name) {
  return TABLE.get(nameKey(name)) ?? null;
}

export const FUNCTION_LIST = defs.map(({ name, syntax, desc }) => ({ name, syntax, desc }));
