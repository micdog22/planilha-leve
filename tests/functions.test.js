// Cada função comparada com o comportamento do Excel nos casos comuns.
import test from 'node:test';
import assert from 'node:assert/strict';
import { calc, show, TODAY } from './helpers.js';
import { lookupFunction, FUNCTION_LIST, makeCriterion } from '../src/functions.js';

// Dados usados por vários testes:
//   A: 10, 20, "abc", (vazio), VERDADEIRO, 5
//   B: números com erro em B3
const DATA = {
  A1: '10', A2: '20', A3: 'abc', A5: 'VERDADEIRO', A6: '5',
  B1: '1', B2: '2', B3: '=1/0',
  C1: '0', C2: '-3,5', C3: "'7",
};

function check(name, cases, data = DATA) {
  test(name, () => {
    for (const [formula, expected] of cases) {
      const got = show(calc(formula, data));
      if (typeof expected === 'number' && typeof got === 'number') {
        assert.ok(Math.abs(got - expected) < 1e-9, `${formula}: esperado ${expected}, veio ${got}`);
      } else {
        assert.equal(got, expected, formula);
      }
    }
  });
}

check('SOMA', [
  ['=SOMA(1;2;3)', 6],
  ['=SOMA(A1:A6)', 35], // texto, vazio e lógico no intervalo são ignorados
  ['=SOMA(A1;A3)', 10], // referência direta a texto também é ignorada
  ['=SOMA("5";1)', 6], // texto numérico digitado direto é convertido
  ['=SOMA("abc")', '#VALOR!'],
  ['=SOMA(VERDADEIRO;1)', 2], // lógico digitado direto conta como 1
  ['=SOMA(B1:B3)', '#DIV/0!'], // erro no intervalo se propaga
  ['=SOMA(1;;2)', 3],
  ['=SOMA(C3)', 0], // '7 é texto
  ['=SOMA(0,1;0,2)', 0.3],
  ['=soma(a1:a2)', 30],
]);

check('MÉDIA', [
  ['=MÉDIA(A1:A6)', 35 / 3],
  ['=MEDIA(1;2;3;4)', 2.5],
  ['=MÉDIA(A3)', '#DIV/0!'],
  ['=MÉDIA(1;;2)', 1],
  ['=MÉDIA(B1:B3)', '#DIV/0!'],
  ['=MÉDIA(C1:C2)', -1.75],
]);

check('MÍNIMO e MÁXIMO', [
  ['=MÍNIMO(A1:A6)', 5],
  ['=MÁXIMO(A1:A6)', 20],
  ['=MÍNIMO(C1:C3)', -3.5],
  ['=MÁXIMO(A3)', 0],
  ['=MINIMO(3;"1";2)', 1],
  ['=MAXIMO(-1;-5)', -1],
  ['=MÁXIMO(B1:B3)', '#DIV/0!'],
]);

check('CONT.NÚM e CONT.VALORES', [
  ['=CONT.NÚM(A1:A6)', 3],
  ['=CONT.NÚM(B1:B3)', 2], // erros não contam e não se propagam
  ['=CONT.NÚM(1;"2";"x";VERDADEIRO)', 3],
  ['=CONT.NÚM(C1:C3)', 2],
  ['=CONT.VALORES(A1:A6)', 5],
  ['=CONT.VALORES(B1:B3)', 3],
  ['=CONT.VALORES(A4)', 0],
  ['=CONT.VALORES(1;"";FALSO)', 3],
  ['=CONT.NUM(A1:C3)', 6],
]);

check('CONT.SE', [
  ['=CONT.SE(A1:A6;">5")', 2],
  ['=CONT.SE(A1:A6;">=5")', 3],
  ['=CONT.SE(A1:A6;10)', 1],
  ['=CONT.SE(A1:A6;"10")', 1],
  ['=CONT.SE(A1:A6;"<>10")', 5],
  ['=CONT.SE(A1:A6;"abc")', 1],
  ['=CONT.SE(A1:A6;"ABC")', 1],
  ['=CONT.SE(A1:A6;"a*")', 1],
  ['=CONT.SE(A1:A6;"?b?")', 1],
  ['=CONT.SE(A1:A6;"*")', 1],
  ['=CONT.SE(A1:A6;"")', 1],
  ['=CONT.SE(A1:A6;"<>")', 5],
  ['=CONT.SE(A1:A6;VERDADEIRO)', 1],
  ['=CONT.SE(A1:A6;"<15")', 2],
  ['=CONT.SE(C1:C3;7)', 1], // texto que parece número conta no critério numérico
  ['=CONT.SE(B1:B3;"#DIV/0!")', 1],
  ['=CONT.SE(5;">1")', '#VALOR!'],
]);

check('SOMASE', [
  ['=SOMASE(A1:A6;">5")', 30],
  ['=SOMASE(A1:A6;">5";B1:B6)', 3],
  ['=SOMASE(A1:A6;"abc";C1:C6)', 0], // C3 é texto
  ['=SOMASE(B1:B2;">1";A1:A2)', 20],
  ['=SOMASE(A1:A6;"<>abc";B1:B6)', 3],
  ['=SOMASE(A1:A6;"abc";B1:B6)', '#DIV/0!'], // a linha que casa tem erro na soma
  ['=SOMASE(A1:A2;"x")', 0],
]);

check('SE, E, OU, NÃO', [
  ['=SE(1>0;"sim";"não")', 'sim'],
  ['=SE(0;"sim";"não")', 'não'],
  ['=SE(FALSO;1)', false],
  ['=SE(VERDADEIRO;;2)', 0],
  ['=SE(C1=0;0;1/C1)', 0], // só o ramo escolhido é calculado
  ['=SE("x";1;2)', '#VALOR!'],
  ['=SE(A4;1;2)', 2], // vazio é falso
  ['=SE(B3;1;2)', '#DIV/0!'],
  ['=SE(A1>5;A1*2;A1/2)', 20],
  ['=E(1;VERDADEIRO)', true],
  ['=E(1;0)', false],
  ['=E(A1:A6)', true], // texto e vazio são ignorados no intervalo
  ['=E(A3)', '#VALOR!'], // só texto: nenhum lógico
  ['=E("x")', '#VALOR!'],
  ['=OU(0;FALSO)', false],
  ['=OU(0;1)', true],
  ['=OU(B1:B3)', '#DIV/0!'],
  ['=NÃO(VERDADEIRO)', false],
  ['=NAO(0)', true],
  ['=NÃO("x")', '#VALOR!'],
]);

check('arredondamentos', [
  ['=ARRED(2,675;2)', 2.68],
  ['=ARRED(2,5;0)', 3],
  ['=ARRED(-2,5;0)', -3],
  ['=ARRED(1234,5678;-2)', 1200],
  ['=ARRED(1,23456;3,9)', 1.235],
  ['=ARRED("x";1)', '#VALOR!'],
  ['=ARREDONDAR.PARA.CIMA(3,2;0)', 4],
  ['=ARREDONDAR.PARA.CIMA(-3,2;0)', -4],
  ['=ARREDONDAR.PARA.CIMA(76,9;0)', 77],
  ['=ARREDONDAR.PARA.CIMA(3,14159;3)', 3.142],
  ['=ARREDONDAR.PARA.CIMA(31415,92654;-2)', 31500],
  ['=ARREDONDAR.PARA.BAIXO(3,2;0)', 3],
  ['=ARREDONDAR.PARA.BAIXO(76,9;0)', 76],
  ['=ARREDONDAR.PARA.BAIXO(3,14159;3)', 3.141],
  ['=ARREDONDAR.PARA.BAIXO(-3,14159;1)', -3.1],
  ['=ARREDONDAR.PARA.BAIXO(31415,92654;-2)', 31400],
  ['=ARRED(1)', '#N/D'],
]);

check('ABS, RAIZ e POTÊNCIA', [
  ['=ABS(-4,5)', 4.5],
  ['=ABS(C2)', 3.5],
  ['=RAIZ(16)', 4],
  ['=RAIZ(2)', Math.SQRT2],
  ['=RAIZ(-1)', '#NÚM!'],
  ['=POTÊNCIA(2;10)', 1024],
  ['=POTENCIA(5;-1)', 0.2],
  ['=POTÊNCIA(-8;1/3)', '#NÚM!'],
  ['=POTÊNCIA(0;-1)', '#DIV/0!'],
  ['=POTÊNCIA(4;0,5)', 2],
]);

check('texto', [
  ['=CONCATENAR("Olá";", ";"mundo")', 'Olá, mundo'],
  ['=CONCATENAR(1,5;"x";VERDADEIRO)', '1,5xVERDADEIRO'],
  ['=CONCATENAR(A1:A2)', '1020'],
  ['=ESQUERDA("planilha";4)', 'plan'],
  ['=ESQUERDA("planilha")', 'p'],
  ['=ESQUERDA("abc";10)', 'abc'],
  ['=ESQUERDA("abc";-1)', '#VALOR!'],
  ['=ESQUERDA(1234,5;3)', '123'],
  ['=DIREITA("planilha";4)', 'ilha'],
  ['=DIREITA("abc";0)', ''],
  ['=DIREITA("ação")', 'o'],
  ['=EXT.TEXTO("planilha leve";10;4)', 'leve'],
  ['=EXT.TEXTO("abc";5;2)', ''],
  ['=EXT.TEXTO("abc";0;1)', '#VALOR!'],
  ['=EXT.TEXTO("abc";2;-1)', '#VALOR!'],
  ['=NÚM.CARACT("coração")', 7],
  ['=NUM.CARACT(1,5)', 3],
  ['=NÚM.CARACT(A4)', 0],
  ['=MAIÚSCULA("ação é")', 'AÇÃO É'],
  ['=MINUSCULA("ÁGUA")', 'água'],
  ['=MAIÚSCULA(VERDADEIRO)', 'VERDADEIRO'],
]);

const TABLE = {
  A1: 'Código', B1: 'Produto', C1: 'Preço',
  A2: '10', B2: 'Caneta', C2: '2,5',
  A3: '20', B3: 'Caderno', C3: '15',
  A4: '30', B4: 'Mochila', C4: '120',
  E1: 'banana', F1: '1',
  E2: 'Maçã', F2: '2',
  E3: 'uva', F3: '3',
};

check('PROCV exato', [
  ['=PROCV(20;A2:C4;2;FALSO)', 'Caderno'],
  ['=PROCV(30;A2:C4;3;0)', 120],
  ['=PROCV(25;A2:C4;2;FALSO)', '#N/D'],
  ['=PROCV("maçã";E1:F3;2;FALSO)', 2],
  ['=PROCV("ba*";E1:F3;2;FALSO)', 1],
  ['=PROCV("u?a";E1:F3;2;FALSO)', 3],
  ['=PROCV("20";A2:C4;2;FALSO)', '#N/D'], // texto não casa com número
  ['=PROCV(20;A2:C4;4;FALSO)', '#REF!'],
  ['=PROCV(20;A2:C4;0;FALSO)', '#VALOR!'],
  ['=PROCV(10;A2:C4;2;)', 'Caneta'], // argumento vazio vale FALSO (busca exata)
  ['=PROCV(A9;A2:C4;2;FALSO)', '#N/D'],
], TABLE);

check('PROCV aproximado', [
  ['=PROCV(25;A2:C4;2)', 'Caderno'],
  ['=PROCV(25;A2:C4;2;VERDADEIRO)', 'Caderno'],
  ['=PROCV(30;A2:C4;2;1)', 'Mochila'],
  ['=PROCV(1000;A2:C4;3)', 120],
  ['=PROCV(5;A2:C4;2)', '#N/D'],
  ['=PROCV(10;A2:C4;2)', 'Caneta'],
], TABLE);

check('PROCV em faixas de imposto', [
  ['=PROCV(1500;A1:B4;2)', 0],
  ['=PROCV(2500;A1:B4;2)', 0.075],
  ['=PROCV(3000;A1:B4;2)', 0.15],
  ['=PROCV(9000;A1:B4;2)', 0.225],
], { A1: '0', B1: '0', A2: '2259,21', B2: '7,5%', A3: '2826,66', B3: '15%', A4: '3751,06', B4: '22,5%' });

check('HOJE e SEERRO', [
  ['=HOJE()', TODAY],
  ['=HOJE()+1', TODAY + 1],
  ['=HOJE(1)', '#N/D'],
  ['=SEERRO(1/0;"erro")', 'erro'],
  ['=SEERRO(10/2;"erro")', 5],
  ['=SEERRO(PROCV(99;A1:A2;1;FALSO);"não achei")', 'não achei'],
]);

check('aridade e nomes', [
  ['=SE(1)', '#N/D'],
  ['=ABS(1;2)', '#N/D'],
  ['=SOMA()', '#N/D'],
  ['=PROCV(1;A1:A2)', '#N/D'],
  ['=NAOEXISTE(1)', '#NOME?'],
]);

test('nomes sem acento e em minúsculas são aceitos', () => {
  for (const alias of ['media', 'MINIMO', 'maximo', 'cont.num', 'nao', 'potencia', 'num.caract', 'maiuscula', 'minuscula']) {
    assert.ok(lookupFunction(alias), alias);
  }
  assert.equal(lookupFunction('MEDIAX'), null);
  assert.equal(FUNCTION_LIST.length, 28);
  for (const f of FUNCTION_LIST) assert.ok(f.syntax.startsWith(f.name + '(') && f.desc.length > 5, f.name);
});

test('critérios isolados', () => {
  const c = makeCriterion('>=10');
  assert.deepEqual([5, 10, 15, '12', 'x', null].map(c), [false, true, true, true, false, false]);
  const t = makeCriterion('<>Maçã');
  assert.deepEqual(['maçã', 'uva', null, 3].map(t), [false, true, true, true]);
  const w = makeCriterion('~*');
  assert.deepEqual(['*', 'a'].map(w), [true, false]);
});
