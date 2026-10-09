// CSV no padrão brasileiro: ponto e vírgula entre campos (a vírgula já é o
// separador decimal), aspas duplas em volta de campos especiais e "" para aspas
// dentro do campo. Leitura tolerante a CRLF, BOM e quebras de linha dentro de aspas.

import { formatValue, isError } from './values.js';

const BOM = String.fromCharCode(0xfeff);

/** Escolhe o separador mais provável olhando a primeira linha (fora das aspas). */
export function detectDelimiter(text) {
  const counts = { ';': 0, ',': 0, '\t': 0 };
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') quoted = !quoted;
    else if (!quoted && (ch === '\n' || ch === '\r')) break;
    else if (!quoted && ch in counts) counts[ch]++;
  }
  if (counts['\t'] > counts[';'] && counts['\t'] >= counts[',']) return '\t';
  if (counts[','] > counts[';']) return ',';
  return ';';
}

/** Lê o texto CSV e devolve uma lista de linhas, cada uma uma lista de campos. */
export function parseCsv(text, delimiter = ';') {
  if (text.startsWith(BOM)) text = text.slice(1);
  const rows = [];
  let row = [];
  let field = '';
  let quoted = false;
  let i = 0;
  const n = text.length;
  while (i < n) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        quoted = false;
        i += 1;
        continue;
      }
      field += ch;
      i += 1;
      continue;
    }
    if (ch === '"' && field === '') {
      quoted = true;
      i += 1;
    } else if (ch === delimiter) {
      row.push(field);
      field = '';
      i += 1;
    } else if (ch === '\r' || ch === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
      i += ch === '\r' && text[i + 1] === '\n' ? 2 : 1;
    } else {
      field += ch;
      i += 1;
    }
  }
  if (field !== '' || row.length > 0 || quoted) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

function quoteField(value, delimiter) {
  if (value === '') return '';
  const needs = value.includes(delimiter) || /["\r\n]/.test(value) || value !== value.trim();
  return needs ? `"${value.replace(/"/g, '""')}"` : value;
}

/** Monta o CSV (linhas separadas por CRLF, como o Excel espera). */
export function toCsv(rows, delimiter = ';') {
  return rows.map((r) => r.map((v) => quoteField(v, delimiter)).join(delimiter)).join('\r\n');
}

/**
 * Exporta a área usada da planilha.
 * mode "valores": o texto exibido (com formatação), pronto para abrir no Excel.
 *   Texto que começa com = + - @ ganha um apóstrofo para não virar fórmula ao
 *   ser aberto em outro programa (proteção contra injeção de fórmulas).
 * mode "formulas": o conteúdo digitado, inclusive fórmulas.
 */
export function sheetToCsv(sheet, { mode = 'valores', delimiter = ';', bom = true } = {}) {
  const { maxRow, maxCol } = sheet.usedBounds();
  const rows = [];
  for (let r = 0; r <= maxRow; r++) {
    const line = [];
    for (let c = 0; c <= maxCol; c++) {
      const cell = sheet.get(r, c);
      if (!cell || cell.raw === '') {
        line.push('');
      } else if (mode === 'formulas') {
        line.push(cell.raw);
      } else {
        const v = cell.value;
        let text = isError(v) ? v.code : formatValue(v, cell.format);
        if (typeof v === 'string' && /^[=+\-@\t\r]/.test(text)) text = "'" + text;
        else if (typeof v === 'string' && text.startsWith("'")) text = "'" + text;
        line.push(text);
      }
    }
    rows.push(line);
  }
  return (bom ? BOM : '') + toCsv(rows, delimiter) + (rows.length ? '\r\n' : '');
}

/** Converte linhas lidas de um CSV em alterações a partir de (row, col). */
export function csvToChanges(rows, startRow = 0, startCol = 0) {
  const changes = [];
  rows.forEach((line, r) => {
    line.forEach((value, c) => {
      if (value !== '') changes.push({ row: startRow + r, col: startCol + c, raw: value });
    });
  });
  return changes;
}
