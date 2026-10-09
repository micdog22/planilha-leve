# Planilha Leve

Planilha que roda no navegador com fórmulas em português, no padrão do Excel brasileiro: ponto e vírgula entre os argumentos, vírgula decimal, `SOMA`, `MÉDIA`, `SE`, `PROCV`, `CONT.SE` e companhia. Por dentro tem um motor de fórmulas completo escrito do zero, com tokenizador, parser, grafo de dependências e recálculo em ordem topológica.

**Acesse online:** https://micdog22.github.io/planilha-leve/

## Recursos

- Grade de 100 linhas por 26 colunas, expansível até 2000 por 52, com navegação de planilha (setas, Tab, Enter, Ctrl+setas, Home, PageDown) e seleção com Shift ou arrastando.
- Fórmulas no estilo do Excel pt-BR: `=SOMA(A1:A10)`, `=SE(B2>0;"ok";"atenção")`, `=1,5*C3`, referências relativas e absolutas (`A1`, `$A$1`, `A$1`), intervalos e os operadores `+ - * / ^ % &` e `= <> < > <= >=`.
- 28 funções: SOMA, MÉDIA, MÍNIMO, MÁXIMO, CONT.NÚM, CONT.VALORES, CONT.SE, SOMASE, SE, SEERRO, E, OU, NÃO, ARRED, ARREDONDAR.PARA.CIMA, ARREDONDAR.PARA.BAIXO, ABS, RAIZ, POTÊNCIA, CONCATENAR, ESQUERDA, DIREITA, EXT.TEXTO, NÚM.CARACT, MAIÚSCULA, MINÚSCULA, PROCV (exato e aproximado) e HOJE.
- Erros como no Excel: `#DIV/0!`, `#NOME?`, `#VALOR!`, `#REF!`, `#N/D` e `#NÚM!`, com a explicação na barra de status.
- Referências circulares detectadas e marcadas, sem travar a página.
- Clique numa célula enquanto escreve uma fórmula para inserir a referência (ou use as setas); arraste para inserir um intervalo; F4 alterna o `$`.
- Copiar e colar (também de e para o Excel e o Google Planilhas), recortar, preencher para baixo (Ctrl+D) e para a direita (Ctrl+R) com ajuste das referências relativas.
- Desfazer e refazer de tudo, inclusive colagens e importações.
- Formatos: geral, número, moeda (R$), porcentagem e data, com controle de casas decimais. Digitar `R$ 10`, `15%` ou `31/12/2024` já aplica o formato certo.
- Importação e exportação de CSV com ponto e vírgula (com valores ou com fórmulas).
- Salvamento automático no navegador.

## Como usar

Clique numa célula e digite. Comece com `=` para uma fórmula. Nos exemplos abaixo, A1, A2 e A3 têm 10, 20 e 30, e C1 está vazia:

| Você digita | A célula mostra | Por quê |
| --- | --- | --- |
| `1.234,56` | 1234,56 | número no formato brasileiro |
| `R$ 1.500` | R$ 1.500,00 | vira número com formato de moeda |
| `=2^3^2` | 64 | `^` associa à esquerda, como no Excel |
| `=-2^2` | 4 | o menos unário vem antes da potência |
| `=10/2%` | 500 | `%` divide por 100 antes da divisão |
| `="Total: "&SOMA(A1:A3)` | Total: 60 | `&` junta textos |
| `=SE(C1=0;0;A1/C1)` | 0 | só o ramo escolhido é calculado |
| `=PROCV(25;A1:A3;1)` | 20 | busca aproximada: maior valor até 25 |
| `=PROCV(25;A1:A3;1;FALSO)` | #N/D | busca exata não encontra 25 |
| `=A1/0` | #DIV/0! | divisão por zero |

Os nomes das funções aceitam minúsculas e a falta de acento: `=media(a1:a3)` funciona igual a `=MÉDIA(A1:A3)`.

Atalhos principais: Enter ou F2 edita, Enter confirma e desce, Tab confirma e vai para a direita, Esc cancela, Delete apaga, Ctrl+Z e Ctrl+Y desfazem e refazem, Ctrl+D e Ctrl+R preenchem. A lista completa está na própria página.

### Diferenças em relação ao Excel

- Uma referência circular vira `#REF!` em todas as células do ciclo (como no Google Planilhas), em vez de um aviso.
- Fórmula com erro de sintaxe é aceita e mostra `#ERRO!` com a posição do problema, para você corrigir depois.
- Intervalo usado onde se espera um valor só (`=A1:A3*2`) segue a interseção implícita do Excel clássico: vale a célula da mesma linha ou coluna; fora disso, `#VALOR!`.
- Datas usam número de série com início em 30/12/1899, iguais às do Excel a partir de 01/03/1900.
- `CONCATENAR` aceita intervalos inteiros. Em `SOMASE`, o intervalo de soma deve ter o mesmo tamanho do intervalo testado.
- Não há várias abas, nomes definidos nem referências de coluna inteira (`A:A`).

## Como rodar localmente

Sem build e sem dependências. Como a página usa módulos ES, sirva a pasta por HTTP:

```bash
python3 -m http.server 8000
```

E abra http://localhost:8000.

## Testes

```bash
npm test
```

(equivale a `node --test`; precisa do Node 18 ou mais novo). Os testes cobrem:

- `tests/formula.test.js`: tokenizador, precedência, menos unário, porcentagem, `&`, comparações entre tipos, erros de sintaxe, deslocamento de referências ao copiar e o F4;
- `tests/functions.test.js`: cada função contra o comportamento do Excel (texto em intervalo ignorado, texto direto convertido, erros propagados, critérios com curinga, PROCV exato e aproximado e outros casos);
- `tests/engine.test.js`: ordem de recálculo, cadeias de 20 mil fórmulas, ciclos, desfazer e refazer, formatos, copiar e colar, preencher, e um teste de propriedade com semente fixa que faz centenas de edições aleatórias e confere que o recálculo incremental chega aos mesmos valores de uma planilha recalculada do zero;
- `tests/values.test.js` e `tests/csv.test.js`: números e datas em pt-BR, formatação, arredondamento e CSV (inclusive ida e volta com textos aleatórios).

## Como funciona

### Da fórmula à árvore

O texto `=SOMA(A1:B2;3)*2` passa por um **tokenizador**, que reconhece números com vírgula decimal, textos entre aspas (com `""` para aspas dentro), referências (`$A$1`), intervalos (`A1:B2`), nomes de função seguidos de `(`, operadores e o `;`. Cada token guarda a posição de início e fim no texto original, o que depois permite reescrever só as referências ao copiar uma fórmula.

O **parser** é do tipo Pratt: cada operador tem uma força de ligação e a expressão é montada consumindo operadores enquanto eles forem mais fortes que o contexto atual. A tabela segue o Excel:

| Força | Operadores |
| --- | --- |
| 70 | `-` e `+` unários |
| 60 | `%` (pós-fixo) |
| 50 | `^` |
| 40 | `*` `/` |
| 30 | `+` `-` |
| 20 | `&` |
| 10 | `=` `<>` `<` `>` `<=` `>=` |

Dois detalhes que pegam muita gente: o menos unário é mais forte que a potência (`=-2^2` dá 4) e a potência associa à esquerda (`=2^3^2` dá 64).

### Avaliação no estilo do Excel

Erros são valores, não exceções: `#DIV/0!` dentro de uma soma vira o resultado da soma. As conversões seguem regras específicas: célula vazia vale 0 numa conta e texto vazio numa concatenação; texto que parece número (`"1,5"`, `"10%"`, `"R$ 2"`) é convertido; comparações entre tipos diferentes seguem a ordem número < texto < lógico, e texto ignora maiúsculas. Nas funções de agregação, o Excel trata diferente o que vem de uma referência e o que é digitado direto: `=SOMA(A1)` ignora texto em A1, mas `=SOMA("abc")` dá `#VALOR!`. Para reproduzir isso, referências chegam às funções como blocos de valores, e o resto como valor simples.

`SE` e `SEERRO` recebem os argumentos sem avaliar e calculam só o ramo necessário, por isso `=SE(C1=0;0;A1/C1)` nunca mostra `#DIV/0!`. O `ARRED` usa deslocamento decimal pela representação em texto do número (com 15 algarismos significativos, como o Excel), evitando o clássico `2,675` virar `2,67` por causa do binário.

### Grafo de dependências e recálculo

Cada fórmula registra as células e intervalos que usa. O motor mantém o caminho inverso: para cada célula, quem depende dela (e uma lista separada das fórmulas que usam intervalos). Quando você edita uma célula:

1. o motor coleta todas as fórmulas afetadas, direta ou indiretamente (uma busca no grafo inverso);
2. ordena esse subgrafo com o **algoritmo de Tarjan** de componentes fortemente conexas, numa versão iterativa (sem recursão, para aguentar cadeias de milhares de fórmulas);
3. avalia os componentes na ordem em que o Tarjan os entrega, que é justamente a ordem topológica: os precedentes primeiro.

O bônus do Tarjan é a detecção de ciclos de graça: um componente com mais de uma célula, ou uma célula que aponta para si mesma (inclusive por um intervalo, como `=SOMA(A1:A3)` em A1), é uma referência circular. Essas células recebem `#REF!` e quem depende delas herda o erro. Quando o ciclo é quebrado, tudo é recalculado normalmente.

Só o que mudou é recalculado, e cada fórmula afetada é avaliada uma única vez por edição.

### Copiar, colar e preencher

Ao colar uma fórmula N linhas abaixo e M colunas à direita, cada referência relativa anda N e M; a parte com `$` fica parada. O texto é reescrito a partir das posições dos tokens, então espaços e textos entre aspas ficam intactos. Uma referência que sairia da planilha vira `#REF!`, como no Excel. O histórico guarda, para cada ação, o antes e o depois das células tocadas, e desfazer é aplicar o "antes" de volta.

### CSV

O CSV usa ponto e vírgula, porque a vírgula já é o separador decimal no Brasil, e é gravado com BOM para o Excel reconhecer os acentos. Campos com `;`, aspas ou quebras de linha vão entre aspas. Na exportação com valores, um texto que começa com `=`, `+`, `-` ou `@` ganha um apóstrofo na frente: isso evita que outro programa interprete o conteúdo como fórmula (a chamada injeção de fórmulas em CSV), e a própria Planilha Leve remove o apóstrofo ao importar.

### Organização do código

```
src/values.js      tipos de valor, erros, conversões, datas, leitura e formatação pt-BR
src/formula.js     tokenizador, parser, endereços, deslocamento de referências e F4
src/evaluate.js    avaliação da árvore e operadores
src/functions.js   biblioteca de funções
src/engine.js      células, grafo de dependências, recálculo, histórico, copiar e preencher
src/csv.js         leitura e escrita de CSV
src/app.js         interface
```

## Contribuindo

Issues e pull requests são bem-vindos. Encontrou uma fórmula que dá resultado diferente do Excel? Abra uma issue com a fórmula, os dados e o resultado esperado; ela vira um caso de teste.

## Licença

MIT. Veja [LICENSE](LICENSE).
