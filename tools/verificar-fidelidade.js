#!/usr/bin/env node
/**
 * Teste de fidelidade: a conta do aplicativo tem que dar exatamente o mesmo
 * numero que a planilha dava.
 *
 * Para cada empresa de data/empresas.json que a planilha conseguia calcular,
 * recalcula com a mesma funcao usada no navegador e compara com o valor que o
 * Excel tinha gravado em cache na propria celula.
 *
 * Uso: node tools/verificar-fidelidade.js
 */

const fs = require('fs');
const path = require('path');

const dataset = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'data', 'empresas.json'), 'utf8'));

/* --- a mesma conta que roda no navegador (assets/app.js -> calcular) ------ */
function calcular(e) {
  const contas = [e.patrimonioLiquido, e.ativoCirculante, e.caixa, e.dividaLiquida];
  const temBalanco = contas.every((v) => typeof v === 'number' && Number.isFinite(v));
  const valorReal = temBalanco
    ? e.patrimonioLiquido + e.ativoCirculante + e.caixa - e.dividaLiquida
    : null;
  const papeis = typeof e.totalPapeis === 'number' && e.totalPapeis > 0 ? e.totalPapeis : null;
  const precoJusto = valorReal !== null && papeis ? valorReal / papeis : null;
  return { valorReal, precoJusto, completa: temBalanco && !!papeis };
}

/* --- amostras anotadas a mao direto das celulas da planilha --------------- */
const AMOSTRAS = [
  { codigo: 'ABEV3', fonte: 'Planilha1', precoJusto: 8.130745489 },
  { codigo: 'ELET3', fonte: 'Planilha1', precoJusto: 73.39944561 },
  { codigo: 'LLIS3', fonte: 'Planilha1', precoJusto: -12.53783015 },
  { codigo: 'BBRK3', fonte: 'Planilha1', precoJusto: 0.4077468159 },
  { codigo: 'PETR3', fonte: 'Planilha1', precoJusto: 42.56423249 },
  { codigo: 'B3SA3', fonte: 'Planilha1', precoJusto: 13.41459517 },
  { codigo: 'CLSC3', fonte: 'ENERGIA', precoJusto: 156.9328058 },
  { codigo: 'LIPR3', fonte: 'ENERGIA', precoJusto: 38.17919574 },
  { codigo: 'ENMT3', fonte: 'ENERGIA', precoJusto: 10.40293153 },
  { codigo: 'CEPE3', fonte: 'ENERGIA', precoJusto: 9.180781079 },
];

let ok = 0;
const falhas = [];

/* 1. toda linha calculavel tem que bater com o cache do Excel */
for (const e of dataset.empresas) {
  const r = calcular(e);
  const esperado = e.planilha && e.planilha.precoJusto;
  if (esperado === null || esperado === undefined) continue;
  if (r.precoJusto === null) {
    falhas.push(`${e.codigo} (${e.fonte}): planilha tinha ${esperado}, o app nao calcula nada.`);
    continue;
  }
  // a planilha guarda ~10 digitos significativos
  const tolerancia = Math.max(1e-7, Math.abs(esperado) * 1e-9);
  if (Math.abs(r.precoJusto - esperado) > tolerancia) {
    falhas.push(`${e.codigo} (${e.fonte}): planilha ${esperado} x app ${r.precoJusto}`);
  } else {
    ok++;
  }
}

/* 2. as amostras conferidas a mao */
let okAmostras = 0;
for (const a of AMOSTRAS) {
  const e = dataset.empresas.find((x) => x.codigo === a.codigo && x.fonte === a.fonte);
  if (!e) { falhas.push(`amostra ${a.codigo} (${a.fonte}) nao esta no dataset.`); continue; }
  const r = calcular(e);
  if (r.precoJusto === null || Math.abs(r.precoJusto - a.precoJusto) > 1e-7) {
    falhas.push(`amostra ${a.codigo}: esperado ${a.precoJusto}, obtido ${r.precoJusto}`);
  } else {
    okAmostras++;
  }
}

/* 3. sanidade do dataset */
const porChave = new Map();
for (const e of dataset.empresas) {
  const k = e.codigo + '|' + e.fonte;
  if (porChave.has(k)) falhas.push(`duplicata no dataset: ${k}`);
  porChave.set(k, e);
  if (!e.id || !e.codigo) falhas.push(`registro sem id/codigo: ${JSON.stringify(e).slice(0, 90)}`);
}

/* ------------------------------------------------------------------ saida */
console.log('empresas no dataset ........ ' + dataset.empresas.length);
console.log('linhas conferidas .......... ' + ok);
console.log('amostras manuais ........... ' + okAmostras + '/' + AMOSTRAS.length);

if (falhas.length) {
  console.log('\nFALHAS (' + falhas.length + '):');
  falhas.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('\nOK: a conta do aplicativo reproduz a planilha em todas as linhas calculaveis.');
