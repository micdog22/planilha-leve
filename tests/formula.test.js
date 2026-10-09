// Tokenizador, parser (precedência, menos unário, porcentagem, &, comparações) e referências.
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  tokenize, parseFormula, FormulaSyntaxError, colName, colIndex, cellName, parseCellName,
  shiftFormula, cycleReference, canInsertReference, collectRefs,
} from '../src/formula.js';
import { calc, show } from './helpers.js';

test('endereços de colunas e células', () => {
  assert.equal(colName(0), 'A');
  assert.equal(colName(25), 'Z');
  assert.equal(colName(26), 'AA');
  assert.equal(colName(701), 'ZZ');
  assert.equal(colName(702), 'AAA');
  assert.equal(colName(16383), 'XFD');
  for (let c = 0; c < 20000; c += 7) assert.equal(colIndex(colName(c)), c);
  assert.equal(cellName(0, 0), 'A1');
  assert.deepEqual(parseCellName('$B$3'), { row: 2, col: 1 });
  assert.deepEqual(parseCellName('xfd1048576'), { row: 1048575, col: 16383 });
  assert.equal(parseCellName('XFE1'), null);
  assert.equal(parseCellName('A0'), null);
  assert.equal(parseCellName('A1048577'), null);
});

test('tokens', () => {
  const kinds = (src) => tokenize(src).map((t) => `${t.t}:${Array.isArray(t.v) ? 'R' : typeof t.v === 'object' ? 'ref' : t.v}`);
  assert.deepEqual(kinds('1,5+A1'), ['num:1.5', 'op:+', 'ref:ref']);
  assert.deepEqual(kinds('SOMA(A1:B2;3)'), ['func:SOMA', 'range:R', ';:;', 'num:3', '):)']);
  assert.deepEqual(kinds('"a""b"&x'), ['str:a"b', 'op:&', 'name:x']);
  assert.deepEqual(kinds('1<>2<=3>=4'), ['num:1', 'op:<>', 'num:2', 'op:<=', 'num:3', 'op:>=', 'num:4']);
  assert.deepEqual(kinds('verdadeiro;FALSO'), ['bool:true', ';:;', 'bool:false']);
  assert.deepEqual(kinds('#DIV/0!+#n/d'), ['err:#DIV/0!', 'op:+', 'err:#N/D']);
  assert.deepEqual(kinds('LOG10(1)'), ['func:LOG10', 'num:1', '):)']);
  assert.deepEqual(kinds('CONT.NÚM (A1)'), ['func:CONT.NÚM', 'ref:ref', '):)']);
  assert.deepEqual(kinds(',5*1E3'), ['num:0.5', 'op:*', 'num:1000']);
  assert.deepEqual(kinds('ABCD1'), ['name:ABCD1']);
  assert.deepEqual(kinds('A1B'), ['name:A1B']);
  const ref = tokenize('$C$7')[0];
  assert.deepEqual(ref.v, { row: 6, col: 2, absRow: true, absCol: true });
  assert.equal(ref.s, 0);
  assert.equal(ref.e, 4);
});

test('erros de sintaxe', () => {
  const bad = ['1+', '(1', '1)', 'SOMA(1;2', '"aberto', '1 2', '*1', 'A1:', '#XYZ', '1 @ 2', 'SOMA(1 2)', '$1', ''];
  for (const src of bad) assert.throws(() => parseFormula('=' + src), FormulaSyntaxError, src);
  assert.equal(show(calc('=1+')), '#ERRO!');
  assert.equal(show(calc('=SOMA(1;')), '#ERRO!');
});

test('precedência e associatividade', () => {
  const cases = [
    ['=1+2*3', 7],
    ['=(1+2)*3', 9],
    ['=2^3^2', 64],
    ['=-2^2', 4],
    ['=0-2^2', -4],
    ['=2^-2', 0.25],
    ['=-2^-2', 0.25],
    ['=10-2-3', 5],
    ['=100/10/5', 2],
    ['=2*3^2', 18],
    ['=1+2&3', '33'],
    ['="a"&1+1', 'a2'],
    ['=1+1=2', true],
    ['=1=1&""', false],
    ['="11"=1&1', true],
    ['=2>1+1', false],
    ['=--5', 5],
    ['=+-+5', -5],
    ['=-"3"', -3],
    ['=+"abc"', 'abc'],
  ];
  for (const [f, v] of cases) assert.equal(show(calc(f)), v, f);
});

test('porcentagem', () => {
  assert.equal(calc('=10%'), 0.1);
  assert.equal(calc('=50%*200'), 100);
  assert.equal(calc('=10/2%'), 500);
  assert.ok(Math.abs(calc('=10%^2') - 0.01) < 1e-15); // 0,1 ao quadrado: 0,01 exibido
  assert.equal(calc('=-5%'), -0.05);
  assert.equal(calc('=A1%', { A1: '30' }), 0.3);
  assert.equal(calc('=200%%'), 0.02);
});

test('concatenação com &', () => {
  assert.equal(calc('="a"&"b"'), 'ab');
  assert.equal(calc('=1,5&"x"'), '1,5x');
  assert.equal(calc('=VERDADEIRO&1'), 'VERDADEIRO1');
  assert.equal(calc('=A1&"|"', {}), '|');
  assert.equal(calc('=1/3&""'), '0,333333333333333');
  assert.equal(show(calc('=1/0&"x"')), '#DIV/0!');
});

test('comparações', () => {
  const cases = [
    ['=1<2', true], ['=2<=2', true], ['=3>=4', false], ['=1<>1', false],
    ['="abc"="ABC"', true], ['="a"<"b"', true], ['="B">"a"', true],
    ['=1<"a"', true], ['="a"<VERDADEIRO', true], ['=1="1"', false],
    ['=A1=0', true], ['=A1=""', true], ['=A1=FALSO', true],
    ['=VERDADEIRO>FALSO', true], ['=#N/D=1', '#N/D'],
  ];
  for (const [f, v] of cases) assert.equal(show(calc(f)), v, f);
});

test('nomes desconhecidos e erros literais', () => {
  assert.equal(show(calc('=ABC+1')), '#NOME?');
  assert.equal(show(calc('=FUNCAOX(1)')), '#NOME?');
  assert.equal(show(calc('=#REF!')), '#REF!');
  assert.equal(show(calc('="x"+1')), '#VALOR!');
  assert.equal(show(calc('=1/0')), '#DIV/0!');
  assert.equal(show(calc('=10^1000')), '#NÚM!');
  assert.equal(show(calc('=(-8)^(1/3)')), '#NÚM!');
  assert.equal(show(calc('=0^0')), '#NÚM!');
  assert.equal(show(calc('=0^-1')), '#DIV/0!');
});

test('interseção implícita de intervalos', () => {
  const sheet = { A1: '1', A2: '2', A3: '3', B1: '10', C1: '20' };
  // a fórmula está em Z1000: fora do intervalo vertical, então dá #VALOR!
  assert.equal(show(calc('=A1:A3*2', sheet)), '#VALOR!');
  assert.equal(calc('=A1:A1*2', sheet), 2);
});

test('referências usadas pela fórmula', () => {
  const { cells, ranges } = collectRefs(parseFormula('=A1+SOMA(B1:C3;$D$4)*SE(E5;F6;G7:G9)'));
  assert.deepEqual(cells.map((c) => cellName(c.row, c.col)).sort(), ['A1', 'D4', 'E5', 'F6']);
  assert.deepEqual(ranges.map((r) => `${cellName(r.r1, r.c1)}:${cellName(r.r2, r.c2)}`).sort(), ['B1:C3', 'G7:G9']);
  const inverted = parseFormula('=SOMA(C3:A1)').args[0];
  assert.deepEqual([inverted.r1, inverted.c1, inverted.r2, inverted.c2], [0, 0, 2, 2]);
});

test('deslocamento de referências ao copiar', () => {
  const cases = [
    ['=A1+B2', 1, 0, '=A2+B3'],
    ['=A1+B2', 0, 2, '=C1+D2'],
    ['=$A$1+A$1+$A1', 3, 3, '=$A$1+D$1+$A4'],
    ['=SOMA(A1:B2)', 2, 1, '=SOMA(B3:C4)'],
    ['=SOMA($A$1:B2)', 2, 1, '=SOMA($A$1:C4)'],
    ['=A1', -1, 0, '=#REF!'],
    ['=SOMA(A1:A3)', 0, -1, '=SOMA(#REF!)'],
    ['="A1"&A1', 1, 1, '="A1"&B2'],
    ['=  a1 +  b1', 1, 0, '=  A2 +  B2'],
    ['=SE(A1>0;A1;-A1)', 5, 0, '=SE(A6>0;A6;-A6)'],
    ['texto A1', 1, 1, 'texto A1'],
    ['=1+(', 1, 1, '=1+('],
  ];
  for (const [f, dr, dc, expected] of cases) assert.equal(shiftFormula(f, dr, dc), expected, f);
});

test('F4 alterna referências absolutas', () => {
  let text = '=A1+B2';
  const steps = [];
  let caret = 3;
  for (let i = 0; i < 4; i++) {
    const r = cycleReference(text, caret);
    text = r.text;
    caret = r.caret;
    steps.push(text);
  }
  assert.deepEqual(steps, ['=$A$1+B2', '=A$1+B2', '=$A1+B2', '=A1+B2']);
  assert.equal(cycleReference('=SOMA(A1:B2)', 8).text, '=SOMA($A$1:$B$2)');
  assert.equal(cycleReference('=1+2', 2), null);
  assert.equal(cycleReference('texto', 1), null);
});

test('onde um clique insere referência', () => {
  assert.equal(canInsertReference('=', 1), true);
  assert.equal(canInsertReference('=SOMA(', 6), true);
  assert.equal(canInsertReference('=A1+', 4), true);
  assert.equal(canInsertReference('=A1+ ', 5), true);
  assert.equal(canInsertReference('=SOMA(A1;', 9), true);
  assert.equal(canInsertReference('=A1', 3), false);
  assert.equal(canInsertReference('texto', 5), false);
  assert.equal(canInsertReference('=x', 0), false);
});
