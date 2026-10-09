// CSV com ponto e vírgula: leitura, escrita, ida e volta e exportação da planilha.
import test from 'node:test';
import assert from 'node:assert/strict';
import { parseCsv, toCsv, detectDelimiter, sheetToCsv, csvToChanges } from '../src/csv.js';
import { Sheet } from '../src/engine.js';
import { sheetWith } from './helpers.js';

const BOM = String.fromCharCode(0xfeff);

test('leitura básica', () => {
  assert.deepEqual(parseCsv('a;b;c\n1;2;3'), [['a', 'b', 'c'], ['1', '2', '3']]);
  assert.deepEqual(parseCsv('a;b\r\n1;2\r\n'), [['a', 'b'], ['1', '2']]);
  assert.deepEqual(parseCsv('a;;c\n;;'), [['a', '', 'c'], ['', '', '']]);
  assert.deepEqual(parseCsv(BOM + 'x;y'), [['x', 'y']]);
  assert.deepEqual(parseCsv(''), []);
  assert.deepEqual(parseCsv('a\n\nb'), [['a'], [''], ['b']]);
  assert.deepEqual(parseCsv('a,b\n1,2', ','), [['a', 'b'], ['1', '2']]);
  assert.deepEqual(parseCsv('so\rcr'), [['so'], ['cr']]);
});

test('campos entre aspas', () => {
  assert.deepEqual(parseCsv('"a;b";"c""d"'), [['a;b', 'c"d']]);
  assert.deepEqual(parseCsv('"linha 1\nlinha 2";x'), [['linha 1\nlinha 2', 'x']]);
  assert.deepEqual(parseCsv('"";"  espaços  "'), [['', '  espaços  ']]);
  assert.deepEqual(parseCsv('"1.234,56";R$ 10'), [['1.234,56', 'R$ 10']]);
  assert.deepEqual(parseCsv('a"b;c'), [['a"b', 'c']]);
  assert.deepEqual(parseCsv('"sem fim'), [['sem fim']]);
});

test('escrita com aspas quando necessário', () => {
  assert.equal(toCsv([['a', 'b;c', 'd"e'], ['', 'x\ny', ' z']]), 'a;"b;c";"d""e"\r\n;"x\ny";" z"');
  assert.equal(toCsv([['1,5', 'texto']], ','), '"1,5",texto');
});

test('ida e volta com textos aleatórios', () => {
  let seed = 7;
  const next = () => {
    seed = (seed * 1103515245 + 12345) % 2147483648;
    return seed / 2147483648;
  };
  const alphabet = ['a', 'b', ';', '"', '\n', '\r\n', ' ', ',', 'ç', 'ã', '1'];
  for (let n = 0; n < 500; n++) {
    const rows = [];
    const width = 1 + Math.floor(next() * 4);
    const height = 1 + Math.floor(next() * 4);
    for (let r = 0; r < height; r++) {
      const row = [];
      for (let c = 0; c < width; c++) {
        let field = '';
        const len = Math.floor(next() * 6);
        for (let i = 0; i < len; i++) field += alphabet[Math.floor(next() * alphabet.length)];
        row.push(field);
      }
      rows.push(row);
    }
    rows[rows.length - 1][0] ||= 'fim';
    assert.deepEqual(parseCsv(toCsv(rows)), rows);
  }
});

test('detecção do separador', () => {
  assert.equal(detectDelimiter('a;b;c\n1;2;3'), ';');
  assert.equal(detectDelimiter('a,b,c\n1,2,3'), ',');
  assert.equal(detectDelimiter('a\tb\tc'), '\t');
  assert.equal(detectDelimiter('"x;y",z,w'), ',');
  assert.equal(detectDelimiter('valor;1,5'), ';');
  assert.equal(detectDelimiter('sozinho'), ';');
});

test('exportar valores formatados e importar de volta', () => {
  const s = sheetWith({
    A1: 'Produto', B1: 'Preço', C1: 'Data',
    A2: 'Café; forte', B2: 'R$ 12,50', C2: '31/12/2024',
    A3: '=1/0', B3: '15%', C3: '=B2*2',
    A4: "'=SOMA(1)", B4: '-5', C4: 'VERDADEIRO',
  });
  const csv = sheetToCsv(s);
  assert.ok(csv.startsWith(BOM));
  const lines = csv.slice(1).split('\r\n');
  assert.equal(lines[0], 'Produto;Preço;Data');
  assert.equal(lines[1], '"Café; forte";R$ 12,50;31/12/2024');
  assert.equal(lines[2], '#DIV/0!;15%;25');
  assert.equal(lines[3], "'=SOMA(1);-5;VERDADEIRO");
  const back = new Sheet();
  back.apply(csvToChanges(parseCsv(csv)), { record: false });
  assert.equal(back.display(1, 1), 'R$ 12,50');
  assert.equal(back.value(1, 2), 45657);
  assert.equal(back.value(3, 0), '=SOMA(1)');
  assert.equal(back.value(3, 1), -5);
  assert.equal(back.value(2, 1), 0.15);
});

test('exportar fórmulas preserva o conteúdo digitado', () => {
  const s = sheetWith({ A1: '2', A2: '=A1*3', B1: 'texto "citado"' });
  const csv = sheetToCsv(s, { mode: 'formulas', bom: false });
  assert.equal(csv, '2;"texto ""citado"""\r\n=A1*3;\r\n');
  const back = new Sheet();
  back.apply(csvToChanges(parseCsv(csv)), { record: false });
  assert.equal(back.value(1, 0), 6);
  assert.equal(back.value(0, 1), 'texto "citado"');
});

test('planilha vazia gera CSV vazio', () => {
  assert.equal(sheetToCsv(new Sheet(), { bom: false }), '');
});

test('texto que começa com apóstrofo sobrevive à ida e volta', () => {
  const s = sheetWith({ A1: "''citação" });
  assert.equal(s.value(0, 0), "'citação");
  const back = new Sheet();
  back.apply(csvToChanges(parseCsv(sheetToCsv(s))), { record: false });
  assert.equal(back.value(0, 0), "'citação");
});
