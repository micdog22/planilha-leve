// Leitura do que se digita, formatação e conversões no padrão brasileiro.
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  parseInput, parseNumberBR, formatValue, generalNumber, roundHalfAway, roundUp, roundDown,
  compareValues, toNumber, toText, toBool, dateToSerial, serialToDate, formatDate, todaySerial, err, ERR,
} from '../src/values.js';

test('números no formato brasileiro', () => {
  const cases = [
    ['1234', 1234], ['1.234', 1234], ['1.234,56', 1234.56], ['-12,5', -12.5], ['+7', 7], [',5', 0.5],
    ['5,', 5], ['1e3', 1000], ['1,5E-2', 0.015], ['0', 0], ['1.234.567,8', 1234567.8],
  ];
  for (const [text, value] of cases) assert.equal(parseNumberBR(text), value, text);
  for (const text of ['1.5', '1.2345', '12.34,5', 'abc', '', '1,2,3', '--1', '1 000']) assert.equal(parseNumberBR(text), null, text);
});

test('o que o usuário digita', () => {
  const kind = (t) => {
    const p = parseInput(t);
    return [p.kind, p.value && typeof p.value === 'object' ? p.value.code : p.value, p.format?.type ?? null];
  };
  assert.deepEqual(kind(''), ['blank', null, null]);
  assert.deepEqual(kind('=1+1'), ['formula', null, null]);
  assert.deepEqual(kind('='), ['text', '=', null]);
  assert.deepEqual(kind("'=1+1"), ['text', '=1+1', null]);
  assert.deepEqual(kind("'123"), ['text', '123', null]);
  assert.deepEqual(kind(' 42 '), ['number', 42, null]);
  assert.deepEqual(kind('12,5%'), ['number', 0.125, 'porcentagem']);
  assert.deepEqual(kind('R$ 1.234,50'), ['number', 1234.5, 'moeda']);
  assert.deepEqual(kind('r$10'), ['number', 10, 'moeda']);
  assert.deepEqual(kind('-R$ 3,00'), ['number', -3, 'moeda']);
  assert.deepEqual(kind('31/12/2024'), ['number', 45657, 'data']);
  assert.deepEqual(kind('1/3/25'), ['number', dateToSerial(2025, 3, 1), 'data']);
  assert.deepEqual(kind('1/3/45'), ['number', dateToSerial(1945, 3, 1), 'data']);
  assert.deepEqual(kind('2024-02-29'), ['number', dateToSerial(2024, 2, 29), 'data']);
  assert.deepEqual(kind('29/02/2023'), ['text', '29/02/2023', null]);
  assert.deepEqual(kind('32/01/2024'), ['text', '32/01/2024', null]);
  assert.deepEqual(kind('verdadeiro'), ['bool', true, null]);
  assert.deepEqual(kind('FALSO'), ['bool', false, null]);
  assert.deepEqual(kind('#DIV/0!'), ['error', '#DIV/0!', null]);
  assert.deepEqual(kind('#n/d'), ['error', '#N/D', null]);
  assert.deepEqual(kind('texto qualquer'), ['text', 'texto qualquer', null]);
  assert.equal(parseInput('12,50%').format.decimals, 2);
});

test('formato Geral (15 algarismos, vírgula decimal)', () => {
  const cases = [
    [0, '0'], [1234.5, '1234,5'], [-0.25, '-0,25'], [1 / 3, '0,333333333333333'], [0.1 + 0.2, '0,3'],
    [100, '100'], [1e20, '1E+20'], [1.5e-12, '1,5E-12'], [123456789012345, '123456789012345'],
    [1e15, '1E+15'], [2 / 3, '0,666666666666667'], [-1e-10, '-1E-10'],
  ];
  for (const [n, text] of cases) assert.equal(generalNumber(n), text, String(n));
});

test('formatação de exibição', () => {
  assert.equal(formatValue(1234.5, { type: 'numero' }), '1.234,50');
  assert.equal(formatValue(1234.5, { type: 'numero', decimals: 0 }), '1.235');
  assert.equal(formatValue(-1234567.891, { type: 'numero', decimals: 1 }), '-1.234.567,9');
  assert.equal(formatValue(1.005, { type: 'numero' }), '1,01');
  assert.equal(formatValue(1234.5, { type: 'moeda' }), 'R$ 1.234,50');
  assert.equal(formatValue(-10, { type: 'moeda' }), '-R$ 10,00');
  assert.equal(formatValue(-0.001, { type: 'moeda' }), 'R$ 0,00');
  assert.equal(formatValue(0.153, { type: 'porcentagem' }), '15%');
  assert.equal(formatValue(0.1534, { type: 'porcentagem', decimals: 2 }), '15,34%');
  assert.equal(formatValue(10, { type: 'porcentagem' }), '1000%');
  assert.equal(formatValue(45657, { type: 'data' }), '31/12/2024');
  assert.equal(formatValue(-5, { type: 'data' }), '-5');
  assert.equal(formatValue(true, null), 'VERDADEIRO');
  assert.equal(formatValue('texto', { type: 'moeda' }), 'texto');
  assert.equal(formatValue(null, null), '');
  assert.equal(formatValue(err(ERR.NA), null), '#N/D');
});

test('arredondamento como o Excel', () => {
  assert.equal(roundHalfAway(2.675, 2), 2.68);
  assert.equal(roundHalfAway(2.5, 0), 3);
  assert.equal(roundHalfAway(-2.5, 0), -3);
  assert.equal(roundHalfAway(1234.5678, -2), 1200);
  assert.equal(roundHalfAway(1.005, 2), 1.01);
  assert.equal(roundUp(3.2, 0), 4);
  assert.equal(roundUp(-3.2, 0), -4);
  assert.equal(roundUp(0.1 + 0.2, 1), 0.3);
  assert.equal(roundDown(3.789, 1), 3.7);
  assert.equal(roundDown(-3.789, 1), -3.7);
  assert.equal(roundDown(31415.92654, -2), 31400);
});

test('comparação entre tipos', () => {
  assert.equal(compareValues(1, 2), -1);
  assert.equal(compareValues('a', 'A'), 0);
  assert.equal(compareValues('á', 'a'), 1);
  assert.equal(compareValues(100, 'a'), -1);
  assert.equal(compareValues('z', true), -1);
  assert.equal(compareValues(false, true), -1);
  assert.equal(compareValues(null, 0), 0);
  assert.equal(compareValues(null, ''), 0);
  assert.equal(compareValues(null, false), 0);
  assert.equal(compareValues('abc', 'abd'), -1);
});

test('conversões', () => {
  assert.equal(toNumber('1,5'), 1.5);
  assert.equal(toNumber('10%'), 0.1);
  assert.equal(toNumber('R$ 2,00'), 2);
  assert.equal(toNumber('01/01/2000'), 36526);
  assert.equal(toNumber(true), 1);
  assert.equal(toNumber(null), 0);
  assert.equal(toNumber('abc').code, '#VALOR!');
  assert.equal(toNumber('').code, '#VALOR!');
  assert.equal(toText(1.5), '1,5');
  assert.equal(toText(false), 'FALSO');
  assert.equal(toText(null), '');
  assert.equal(toBool(0), false);
  assert.equal(toBool('verdadeiro'), true);
  assert.equal(toBool('sim').code, '#VALOR!');
});

test('datas por número de série', () => {
  assert.equal(dateToSerial(2024, 12, 31), 45657);
  assert.equal(dateToSerial(1900, 3, 1), 61);
  assert.equal(dateToSerial(1970, 1, 1), 25569);
  assert.deepEqual(serialToDate(45658), { year: 2025, month: 1, day: 1 });
  assert.equal(formatDate(36526), '01/01/2000');
  assert.equal(todaySerial(new Date(2026, 9, 9, 23, 59)), dateToSerial(2026, 10, 9));
  for (let s = 61; s < 80000; s += 997) {
    const { year, month, day } = serialToDate(s);
    assert.equal(dateToSerial(year, month, day), s);
  }
});
