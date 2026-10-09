// Motor: ordem de recálculo, ciclos, histórico, formatos, persistência e cópia.
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  Sheet, copyBlock, pasteChanges, fillDownChanges, fillRightChanges, clearChanges, selectionSummary, cellKey,
} from '../src/engine.js';
import { cellName, parseCellName } from '../src/formula.js';
import { sheetWith, at, set, TODAY } from './helpers.js';

const names = (keys) => [...keys].map((k) => cellName(Math.floor(k / 16384), k % 16384)).sort();

test('cadeia simples e conjunto de células afetadas', () => {
  const s = sheetWith({ A1: '1', A2: '=A1+1', A3: '=A2*2', B1: '=7' });
  assert.equal(at(s, 'A3'), 4);
  const affected = set(s, 'A1', '10');
  assert.deepEqual(names(affected), ['A1', 'A2', 'A3']);
  assert.equal(at(s, 'A2'), 11);
  assert.equal(at(s, 'A3'), 22);
});

test('ordem topológica: cada fórmula é avaliada uma vez, depois dos precedentes', () => {
  const s = sheetWith({ A1: '1', B1: '=A1+1', C1: '=A1*2', D1: '=B1+C1', E1: '=D1+B1' });
  const order = [];
  const original = s._evaluateCell.bind(s);
  s._evaluateCell = (k) => {
    order.push(cellName(Math.floor(k / 16384), k % 16384));
    original(k);
  };
  set(s, 'A1', '5');
  assert.equal(order.length, 4);
  assert.equal(new Set(order).size, 4);
  assert.ok(order.indexOf('D1') > order.indexOf('B1') && order.indexOf('D1') > order.indexOf('C1'));
  assert.ok(order.indexOf('E1') > order.indexOf('D1'));
  assert.equal(at(s, 'E1'), 6 + 10 + 6);
});

test('fórmulas digitadas antes dos dados', () => {
  const s = new Sheet({ today: () => TODAY });
  set(s, 'C1', '=B1*2');
  set(s, 'B1', '=A1+1');
  assert.equal(at(s, 'C1'), 2);
  set(s, 'A1', '4');
  assert.equal(at(s, 'C1'), 10);
});

test('cadeia longa sem estouro de pilha', () => {
  const s = new Sheet();
  const changes = [{ row: 0, col: 0, raw: '1' }];
  for (let r = 1; r < 20000; r++) changes.push({ row: r, col: 0, raw: `=A${r}+1` });
  s.apply(changes, { record: false });
  assert.equal(s.value(19999, 0), 20000);
  s.apply([{ row: 0, col: 0, raw: '100' }]);
  assert.equal(s.value(19999, 0), 20099);
});

test('dependência por intervalo', () => {
  const s = sheetWith({ A1: '1', A2: '2', A3: '3', A10: '=SOMA(A1:A9)', B1: '=A10*2' });
  assert.equal(at(s, 'B1'), 12);
  set(s, 'A5', '10');
  assert.equal(at(s, 'A10'), 16);
  assert.equal(at(s, 'B1'), 32);
  set(s, 'A10', '0');
  set(s, 'A5', '99');
  assert.equal(at(s, 'B1'), 0); // A10 deixou de ser fórmula
});

test('referências circulares', () => {
  const s = sheetWith({ A1: '=B1', B1: '=A1', C1: '=A1+1', D1: '=D1+1', E1: '=SOMA(E1:E3)' });
  for (const name of ['A1', 'B1', 'C1', 'D1', 'E1']) assert.equal(at(s, name), '#REF!', name);
  const { row, col } = parseCellName('A1');
  assert.match(s.value(row, col).detail, /circular entre A1, B1/);
  assert.match(s.value(...Object.values(parseCellName('D1'))).detail, /circular/);
  // quebrar o ciclo recalcula tudo
  set(s, 'B1', '5');
  assert.equal(at(s, 'A1'), 5);
  assert.equal(at(s, 'C1'), 6);
});

test('ciclo de três células com cauda', () => {
  const s = sheetWith({ A1: '=B1', B1: '=C1', C1: '=A1', D1: '=C1*2', E1: '7' });
  assert.deepEqual(['A1', 'B1', 'C1', 'D1', 'E1'].map((n) => at(s, n)), ['#REF!', '#REF!', '#REF!', '#REF!', 7]);
  set(s, 'C1', '=E1');
  assert.deepEqual(['A1', 'B1', 'C1', 'D1'].map((n) => at(s, n)), [7, 7, 7, 14]);
});

test('erro de sintaxe e correção', () => {
  const s = sheetWith({ A1: '=1+', B1: '=A1*2' });
  assert.equal(at(s, 'A1'), '#ERRO!');
  assert.equal(at(s, 'B1'), '#ERRO!');
  assert.match(s.value(0, 0).detail, /posição/);
  set(s, 'A1', '=1+2');
  assert.equal(at(s, 'B1'), 6);
});

test('fórmula que aponta para vazio vale 0, mas &"" dá texto vazio', () => {
  const s = sheetWith({ A1: '=B1', A2: '=B1&""', A3: '=SE(B1="";"vazio";B1)' });
  assert.equal(at(s, 'A1'), 0);
  assert.equal(at(s, 'A2'), '');
  assert.equal(at(s, 'A3'), 'vazio');
});

test('desfazer e refazer', () => {
  const s = new Sheet();
  set(s, 'A1', '1');
  set(s, 'A2', '=A1*10');
  set(s, 'A1', '5');
  assert.equal(at(s, 'A2'), 50);
  s.undo();
  assert.equal(at(s, 'A1'), 1);
  assert.equal(at(s, 'A2'), 10);
  s.undo();
  assert.equal(s.raw(1, 0), '');
  assert.equal(s.canRedo, true);
  s.redo();
  s.redo();
  assert.equal(at(s, 'A2'), 50);
  assert.equal(s.redo(), null);
  s.undo();
  set(s, 'B1', 'novo');
  assert.equal(s.canRedo, false);
  // alteração sem efeito não entra no histórico
  const before = s.history.length;
  set(s, 'B1', 'novo');
  assert.equal(s.history.length, before);
});

test('histórico agrupa uma ação inteira (colar várias células)', () => {
  const s = new Sheet();
  s.apply([{ row: 0, col: 0, raw: '1' }, { row: 0, col: 1, raw: '2' }, { row: 0, col: 2, raw: '=A1+B1' }]);
  assert.equal(at(s, 'C1'), 3);
  s.undo();
  assert.equal(s.cells.size, 0);
});

test('formatação sugerida pela digitação e mantida depois', () => {
  const s = new Sheet({ today: () => TODAY });
  set(s, 'A1', 'R$ 1.234,50');
  assert.equal(s.format(0, 0).type, 'moeda');
  assert.equal(s.display(0, 0), 'R$ 1.234,50');
  set(s, 'A1', '99');
  assert.equal(s.display(0, 0), 'R$ 99,00');
  set(s, 'A2', '15%');
  assert.equal(s.display(1, 0), '15%');
  set(s, 'A3', '31/12/2024');
  assert.equal(s.value(2, 0), 45657);
  assert.equal(s.display(2, 0), '31/12/2024');
  set(s, 'A4', '=HOJE()');
  assert.equal(s.display(3, 0), '31/12/2024');
  s.apply([{ row: 0, col: 0, format: null }]);
  assert.equal(s.display(0, 0), '99');
  s.apply([{ row: 4, col: 0, raw: '0,5', format: { type: 'porcentagem', decimals: 1 } }]);
  assert.equal(s.display(4, 0), '50,0%');
  s.undo();
  assert.equal(s.display(4, 0), '');
});

test('célula só com formatação continua existindo', () => {
  const s = new Sheet();
  s.apply([{ row: 0, col: 0, format: { type: 'moeda' } }]);
  assert.equal(s.cells.size, 1);
  set(s, 'A1', '3');
  assert.equal(s.display(0, 0), 'R$ 3,00');
  s.apply([{ row: 0, col: 0, raw: '', format: null }]);
  assert.equal(s.cells.size, 0);
});

test('salvar e abrir (JSON) recalcula HOJE', () => {
  let today = TODAY;
  const s = sheetWith({ A1: '10', A2: '=A1*2', B1: 'R$ 5', C1: '=HOJE()' });
  const json = JSON.parse(JSON.stringify(s.toJSON()));
  assert.equal(json.version, 1);
  today = TODAY + 3;
  const t = Sheet.fromJSON(json, { today: () => today });
  assert.equal(at(t, 'A2'), 20);
  assert.equal(t.display(0, 1), 'R$ 5,00');
  assert.equal(at(t, 'C1'), TODAY + 3);
  assert.equal(t.history.length, 0);
  const bad = Sheet.fromJSON({ cells: [['??', 'x'], ['A1', 5], ['B2', 'ok']] });
  assert.equal(bad.cells.size, 1);
});

test('copiar e colar ajusta referências relativas', () => {
  const s = sheetWith({ A1: '1', A2: '2', B1: '=A1*10', B2: '=$A$1+A2' });
  const block = copyBlock(s, 0, 1, 1, 1);
  s.apply(pasteChanges(block, 0, 2));
  assert.equal(s.raw(0, 2), '=B1*10');
  assert.equal(s.raw(1, 2), '=$A$1+B2');
  assert.equal(at(s, 'C1'), 100);
  assert.equal(at(s, 'C2'), 4); // $A$1 + B2 = 1 + 3
  // colar sem deslocar (recortar)
  s.apply(pasteChanges(block, 5, 5, { shift: false }));
  assert.equal(s.raw(5, 5), '=A1*10');
  // repetição para preencher uma seleção maior
  const one = copyBlock(s, 0, 0, 0, 0);
  const changes = pasteChanges(one, 10, 0, { targetRows: 3, targetCols: 2 });
  assert.equal(changes.length, 6);
  // referência que sai da planilha vira #REF!
  s.apply(pasteChanges(copyBlock(s, 0, 1, 0, 1), 0, 0));
  assert.equal(s.raw(0, 0), '=#REF!*10');
  assert.equal(at(s, 'A1'), '#REF!');
});

test('preencher para baixo e para a direita', () => {
  const s = sheetWith({ A1: '1', A2: '2', A3: '3', B1: '=A1*2', C1: 'R$ 1' });
  s.apply(fillDownChanges(s, 0, 1, 2, 2));
  assert.equal(s.raw(1, 1), '=A2*2');
  assert.equal(s.raw(2, 1), '=A3*2');
  assert.equal(at(s, 'B3'), 6);
  assert.equal(s.display(2, 2), 'R$ 1,00');
  s.apply(fillRightChanges(s, 0, 1, 0, 3));
  assert.equal(s.raw(0, 2), '=B1*2');
  assert.equal(s.raw(0, 3), '=C1*2');
  assert.equal(at(s, 'D1'), 8);
  s.apply(clearChanges(0, 0, 2, 3));
  assert.equal(s.usedBounds().maxRow, -1);
});

test('resumo da seleção', () => {
  const s = sheetWith({ A1: '10', A2: '20', A3: 'x', A4: '=A1+A2' });
  assert.deepEqual(selectionSummary(s, 0, 0, 4, 0), { count: 4, numbers: 3, sum: 60, average: 20 });
  assert.deepEqual(selectionSummary(s, 5, 5, 6, 6), { count: 0, numbers: 0, sum: 0, average: null });
});

test('intervalo enorme não trava', () => {
  const s = sheetWith({ B2: '1', C3: '2', A1: '=SOMA(B1:Z5000)', A2: '=SOMA(B1:XFD1048576)' });
  assert.equal(at(s, 'A1'), 3);
  assert.equal(at(s, 'A2'), '#REF!');
  assert.match(s.value(1, 0).detail, /grande demais/);
});

test('planilha de orçamento (cenário completo)', () => {
  const s = sheetWith({
    A1: 'Item', B1: 'Qtd', C1: 'Preço', D1: 'Total',
    A2: 'Caneta', B2: '10', C2: 'R$ 2,50', D2: '=B2*C2',
    A3: 'Caderno', B3: '3', C3: 'R$ 15,90', D3: '=B3*C3',
    A4: 'Mochila', B4: '1', C4: 'R$ 120', D4: '=B4*C4',
    D5: '=SOMA(D2:D4)', D6: '=D5*10%', D7: '=D5-D6',
    E2: '=SE(D2>D$5*20%;"alto";"ok")',
    F1: '=PROCV("Caderno";A2:D4;4;FALSO)',
    G1: '=CONT.SE(D2:D4;">50")', H1: '=ARRED(MÉDIA(C2:C4);2)',
  });
  assert.equal(s.display(4, 3), '192,7');
  assert.equal(s.display(5, 3), '19,27');
  assert.equal(Math.round(at(s, 'D7') * 100) / 100, 173.43);
  assert.equal(at(s, 'E2'), 'ok');
  assert.ok(Math.abs(at(s, 'F1') - 47.7) < 1e-9);
  assert.equal(at(s, 'G1'), 1);
  assert.equal(at(s, 'H1'), 46.13);
  s.apply(fillDownChanges(s, 1, 4, 3, 4));
  assert.deepEqual(['E2', 'E3', 'E4'].map((n) => at(s, n)), ['ok', 'alto', 'alto']);
});

// gerador com semente fixa
function rng(seed) {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

test('propriedade: recálculo incremental igual ao recálculo do zero', () => {
  const next = rng(42);
  const pick = (n) => Math.floor(next() * n);
  const ref = () => cellName(pick(8), pick(4));
  const formulas = [
    () => `=${ref()}+${ref()}`,
    () => `=${ref()}*2-${pick(10)}`,
    () => `=SOMA(${ref()}:${ref()})`,
    () => `=SE(${ref()}>5;${ref()};${pick(100)})`,
    () => `=MÉDIA(${ref()};${ref()})`,
    () => `=${ref()}/${ref()}`,
    () => `=CONT.SE(A1:D8;">${pick(20)}")`,
    () => `=MÁXIMO(${ref()}:${ref()})&"x"`,
  ];
  for (let round = 0; round < 40; round++) {
    const s = new Sheet({ today: () => TODAY });
    for (let step = 0; step < 60; step++) {
      const r = pick(8);
      const c = pick(4);
      const roll = next();
      const raw = roll < 0.45 ? String(pick(20) - 5) : roll < 0.9 ? formulas[pick(formulas.length)]() : '';
      s.apply([{ row: r, col: c, raw }]);
      if (next() < 0.1 && s.canUndo) s.undo();
    }
    const fresh = Sheet.fromJSON(s.toJSON(), { today: () => TODAY });
    for (let r = 0; r < 8; r++) {
      for (let c = 0; c < 4; c++) {
        const a = s.value(r, c);
        const b = fresh.value(r, c);
        const norm = (v) => (v && typeof v === 'object' ? v.code : v);
        assert.deepEqual(norm(a), norm(b), `rodada ${round}, ${cellName(r, c)}: ${s.raw(r, c)}`);
      }
    }
  }
});

test('chave de célula', () => {
  assert.equal(cellKey(2, 3), 2 * 16384 + 3);
});
