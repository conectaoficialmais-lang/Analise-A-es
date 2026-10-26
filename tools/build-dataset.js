#!/usr/bin/env node
/**
 * Gera data/empresas.json a partir da planilha "PREÇO JUSTO.xlsx".
 *
 * Uso:  node tools/build-dataset.js "caminho/para/PREÇO JUSTO.xlsx"
 *
 * O que ele faz:
 *  - le as duas abas da planilha (Planilha1 e ENERGIA) com tools/xlsx-reader.js
 *  - normaliza codigo/empresa (varias linhas da planilha estao com as duas colunas trocadas)
 *  - converte valores em texto no formato "R$ 1.234.567,00" para numero
 *  - marca como incompleta toda linha que nao tem TOTAL DE PAPEIS (na planilha elas dao #DIV/0!)
 *  - confere o PRECO JUSTO recalculado contra o valor que a planilha tinha em cache
 */

const fs = require('fs');
const path = require('path');
const { readWorkbook } = require('./xlsx-reader.js');

const xlsxPath = process.argv[2] || path.join(process.env.USERPROFILE || process.env.HOME, 'Downloads', 'PREÇO JUSTO.xlsx');
const outPath = path.join(__dirname, '..', 'data', 'empresas.json');

/** Um ticker da B3: 4 caracteres (comecando por letra, ex. B3SA) + 1-2 digitos + sufixo opcional. */
const TICKER = /^[A-Za-z][A-Za-z0-9]{3}[0-9]{1,2}[A-Za-z]?$/;

function toNumber(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  const s = String(v).trim();
  if (!s || s === '#N/A' || s === '#DIV/0!' || s === '#REF!' || s === '#VALUE!') return null;
  // "R$ 1.234.567,89"  ->  1234567.89      /     "10,13,%" -> 10.13
  const limpo = s
    .replace(/R\$/gi, '')
    .replace(/%/g, '')
    .replace(/\s| /g, '')
    .replace(/\.(?=\d{3}\b)/g, '')
    .replace(/,(?=\D*$)/, '')
    .replace(',', '.');
  const n = Number(limpo);
  return Number.isFinite(n) ? n : null;
}

/** Algumas linhas da planilha tem CODIGO e EMPRESA invertidos (BRASKEN|BRKM3, CIELO|CIEL3...). */
function normalizarIdentidade(a, b) {
  const ca = String(a || '').trim();
  const cb = String(b || '').trim();
  if (TICKER.test(ca)) return { codigo: ca.toUpperCase(), empresa: cb || ca.toUpperCase() };
  if (TICKER.test(cb)) return { codigo: cb.toUpperCase(), empresa: ca || cb.toUpperCase() };
  return { codigo: ca.toUpperCase(), empresa: cb || ca.toUpperCase() };
}

function slug(codigo, fonte) {
  return (codigo + '-' + fonte).toLowerCase().replace(/[^a-z0-9]+/g, '-');
}

const wb = readWorkbook(xlsxPath);
const empresas = [];
const avisos = [];

// ---------------------------------------------------------------- Planilha1
// A=CODIGO B=EMPRESA C=PL D=ATIVO CIRC. E=CAIXA F=DIVIDA LIQ. G=PAPEIS H=VALOR REAL I=PRECO JUSTO
const vistosGeral = new Map();
for (const { r, cells } of wb.Planilha1) {
  if (r === 1) continue;
  const { codigo, empresa } = normalizarIdentidade(cells[1], cells[2]);
  if (!codigo || !TICKER.test(codigo)) continue;

  const reg = {
    id: slug(codigo, 'geral'),
    codigo,
    empresa,
    setor: 'Geral',
    fonte: 'Planilha1',
    linhaPlanilha: r,
    patrimonioLiquido: toNumber(cells[3]),
    ativoCirculante: toNumber(cells[4]),
    caixa: toNumber(cells[5]),
    dividaLiquida: toNumber(cells[6]),
    totalPapeis: toNumber(cells[7]),
    preco: null,
    dy: null,
    // LPA, P/L, VPA e P/VP eram digitados a mao na planilha. Ficam separados dos
    // calculados pelo app, que tem os mesmos nomes e sobrescreveriam estes.
    anotado: null,
    obs: null,
    planilha: { valorReal: toNumber(cells[8]), precoJusto: toNumber(cells[9]) },
  };

  const anterior = vistosGeral.get(codigo);
  if (anterior) {
    // TRPL3 aparece duas vezes com os MESMOS numeros -> fica so um registro.
    const igual = ['patrimonioLiquido', 'ativoCirculante', 'caixa', 'dividaLiquida', 'totalPapeis']
      .every((k) => anterior[k] === reg[k]);
    if (igual) {
      avisos.push(`${codigo}: linha ${r} duplica a linha ${anterior.linhaPlanilha} (mesmos números) — ignorada.`);
      continue;
    }
    // BBAS3 aparece duas vezes com numeros DIFERENTES -> fica o balanco maior (mais recente)
    // e o outro vira observacao, para nao duplicar o ticker no ranking.
    const maisNovo = (reg.patrimonioLiquido || 0) > (anterior.patrimonioLiquido || 0) ? reg : anterior;
    const maisVelho = maisNovo === reg ? anterior : reg;
    maisNovo.obs =
      `A planilha tinha dois balanços para ${codigo}. Em uso: linha ${maisNovo.linhaPlanilha}. ` +
      `O outro (linha ${maisVelho.linhaPlanilha}, PL ${fmtBi(maisVelho.patrimonioLiquido)}) dava preço justo ` +
      `${(maisVelho.planilha.precoJusto || 0).toFixed(2)}.`;
    avisos.push(`${codigo}: dois balanços diferentes (linhas ${maisVelho.linhaPlanilha} e ${maisNovo.linhaPlanilha}) — mantido o de PL maior.`);
    const i = empresas.indexOf(maisVelho);
    if (i >= 0) empresas.splice(i, 1, maisNovo);
    vistosGeral.set(codigo, maisNovo);
    continue;
  }

  vistosGeral.set(codigo, reg);
  empresas.push(reg);
}

// ------------------------------------------------------------------ ENERGIA
// A=CODIGO B=(vazia) C=LPA D=P/L E=VPA F=P/VP G=PRECO H=DY
// I=PL J=ATIVO CIRC. K=CAIXA L=DIVIDA LIQ. M=PAPEIS N=VALOR REAL O=PRECO JUSTO
for (const { r, cells } of wb.ENERGIA) {
  if (r === 1) continue;
  const codigo = String(cells[1] || '').trim().toUpperCase();
  if (!codigo || !TICKER.test(codigo)) continue;

  const dyBruto = cells[8];
  const dy = toNumber(dyBruto);
  const obs = [];
  if (typeof dyBruto === 'string' && !/^[\d.,]+%?$/.test(dyBruto.trim())) {
    obs.push(`O DY estava escrito como "${dyBruto}" na planilha; interpretado como ${(dy * 100).toFixed(2)}%.`);
  }
  if (dy !== null && dy > 0.5) {
    obs.push(`DY de ${(dy * 100).toFixed(2)}% veio assim da planilha — vale conferir.`);
  }

  empresas.push({
    id: slug(codigo, 'energia'),
    codigo,
    empresa: codigo,
    setor: 'Energia',
    fonte: 'ENERGIA',
    linhaPlanilha: r,
    patrimonioLiquido: toNumber(cells[9]),
    ativoCirculante: toNumber(cells[10]),
    caixa: toNumber(cells[11]),
    dividaLiquida: toNumber(cells[12]),
    totalPapeis: toNumber(cells[13]),
    preco: toNumber(cells[7]) || null, // EMAE3 veio 0 e POWE3/ENMA3B vieram #N/A
    dy,
    anotado: (() => {
      const a = {
        lpa: toNumber(cells[3]),
        pl: toNumber(cells[4]),
        vpa: toNumber(cells[5]),
        pvp: toNumber(cells[6]),
      };
      return Object.values(a).some((v) => v !== null) ? a : null;
    })(),
    obs: obs.length ? obs.join(' ') : null,
    planilha: { valorReal: toNumber(cells[14]), precoJusto: toNumber(cells[15]) },
  });
}

function fmtBi(v) {
  if (v === null || v === undefined) return '-';
  return 'R$ ' + (v / 1e9).toFixed(2) + ' bi';
}

// --------------------------------------------- recalcula e confere com o cache
let conferidas = 0;
const divergencias = [];
for (const e of empresas) {
  const temBalanco = [e.patrimonioLiquido, e.ativoCirculante, e.caixa, e.dividaLiquida]
    .every((v) => v !== null);
  e.valorReal = temBalanco
    ? e.patrimonioLiquido + e.ativoCirculante + e.caixa - e.dividaLiquida
    : e.patrimonioLiquido !== null
      ? e.patrimonioLiquido // as linhas incompletas da aba ENERGIA tinham so o PL
      : null;
  e.precoJusto = e.valorReal !== null && e.totalPapeis ? e.valorReal / e.totalPapeis : null;
  e.completa = temBalanco && !!e.totalPapeis;

  if (e.planilha.precoJusto !== null && e.precoJusto !== null) {
    conferidas++;
    const dif = Math.abs(e.planilha.precoJusto - e.precoJusto);
    if (dif > Math.max(1e-6, Math.abs(e.precoJusto) * 1e-8)) {
      divergencias.push(`${e.codigo} (${e.fonte}): planilha ${e.planilha.precoJusto} x recalculado ${e.precoJusto}`);
    }
  }
}

empresas.sort((a, b) => a.setor.localeCompare(b.setor) || a.codigo.localeCompare(b.codigo));

const dataset = {
  gerado: new Date().toISOString().slice(0, 10),
  origem: 'PREÇO JUSTO.xlsx',
  periodoAproximado: 'balanços de aproximadamente 2021/2022 (nomes como B2W, Via Varejo, BR Malls e SulAmérica datam a planilha)',
  formulas: {
    valorReal: 'PATRIMÔNIO LÍQUIDO + ATIVO CIRCULANTE + DISPONIBILIDADE (CAIXA) − DÍVIDA LÍQUIDA',
    precoJusto: 'VALOR REAL ÷ TOTAL DE PAPÉIS',
  },
  avisos,
  empresas,
};

fs.writeFileSync(outPath, JSON.stringify(dataset, null, 1));

console.log(`empresas: ${empresas.length} (Geral ${empresas.filter((e) => e.setor === 'Geral').length}, Energia ${empresas.filter((e) => e.setor === 'Energia').length})`);
console.log(`completas: ${empresas.filter((e) => e.completa).length}`);
console.log(`precos justos conferidos contra a planilha: ${conferidas}`);
console.log(divergencias.length ? 'DIVERGENCIAS:\n  ' + divergencias.join('\n  ') : 'nenhuma divergencia');
avisos.forEach((a) => console.log('aviso: ' + a));
console.log('escrito em ' + outPath);
if (divergencias.length) process.exit(1);
