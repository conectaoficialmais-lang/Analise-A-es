#!/usr/bin/env node
/**
 * Testes das funcoes puras de assets/app.js.
 * Uso: node tools/testar.js
 */

const { calcular, lerNumero } = require('../assets/app.js');

let passou = 0;
const falhou = [];

function conferir(nome, obtido, esperado) {
  const igual =
    typeof esperado === 'number' && typeof obtido === 'number'
      ? Math.abs(obtido - esperado) < 1e-9
      : obtido === esperado;
  if (igual) passou++;
  else falhou.push(`${nome}: esperado ${esperado}, obtido ${obtido}`);
}

/* ------------------------------------------------- leitura de numeros --- */

conferir('inteiro simples', lerNumero('410545000000'), 410545000000);
conferir('milhar com ponto', lerNumero('410.545.000.000'), 410545000000);
conferir('decimal com virgula', lerNumero('1234,56'), 1234.56);
conferir('formato BR cheio', lerNumero('1.234.567,89'), 1234567.89);
conferir('decimal com ponto', lerNumero('1.5'), 1.5);
conferir('com R$', lerNumero('R$ 737.734.000,00'), 737734000);
conferir('negativo com milhar', lerNumero('-13.132'), -13132);
conferir('negativo simples', lerNumero('-151425000'), -151425000);
conferir('sufixo bi', lerNumero('410,5 bi'), 410500000000);
conferir('sufixo mi', lerNumero('2,5mi'), 2500000);
conferir('sufixo mil', lerNumero('30 mil'), 30000);
conferir('percentual', lerNumero('6,67%'), 6.67);
conferir('vazio', lerNumero(''), null);
conferir('so espaco', lerNumero('   '), null);
conferir('texto', lerNumero('#N/A'), null);
conferir('nulo', lerNumero(null), null);
conferir('zero', lerNumero('0'), 0);

/* ------------------------------------------------------ as duas contas --- */

// ABEV3, direto da linha 2 da planilha
const abev = calcular({
  patrimonioLiquido: 86771210000,
  ativoCirculante: 38197082000,
  caixa: 20000745000,
  dividaLiquida: 16973032000,
  totalPapeis: 15742222552,
});
conferir('ABEV3 valor real', abev.valorReal, 127996005000);
conferir('ABEV3 preco justo', Number(abev.precoJusto.toFixed(9)), 8.130745489);
conferir('ABEV3 completa', abev.completa, true);

// LLIS3: divida maior que o resto, preco justo negativo
const llis = calcular({
  patrimonioLiquido: 182944000,
  ativoCirculante: 513063000,
  caixa: 60413000,
  dividaLiquida: 1619660000,
  totalPapeis: 68850829,
});
conferir('LLIS3 valor real negativo', llis.valorReal, -863240000);
conferir('LLIS3 preco justo', Number(llis.precoJusto.toFixed(8)), -12.53783015);

// Divida negativa (caixa liquido) entra somando, como na planilha
const aura = calcular({
  patrimonioLiquido: 1501160000,
  ativoCirculante: 1458828000,
  caixa: 898496000,
  dividaLiquida: -51889000,
  totalPapeis: 70742460,
});
conferir('AURA33 divida negativa', aura.valorReal, 3910373000);

/* ----------------------------------------------------- casos de borda --- */

const semPapeis = calcular({
  patrimonioLiquido: 737734000,
  ativoCirculante: null,
  caixa: null,
  dividaLiquida: null,
  totalPapeis: null,
});
conferir('sem balanco: valor real nulo', semPapeis.valorReal, null);
conferir('sem balanco: preco justo nulo', semPapeis.precoJusto, null);
conferir('sem balanco: incompleta', semPapeis.completa, false);

const papeisZero = calcular({
  patrimonioLiquido: 100, ativoCirculante: 0, caixa: 0, dividaLiquida: 0, totalPapeis: 0,
});
conferir('papeis zero nao divide', papeisZero.precoJusto, null);

/* ------------------------------------------------- margem e multiplos --- */

const comPreco = calcular({
  patrimonioLiquido: 1000, ativoCirculante: 0, caixa: 0, dividaLiquida: 0,
  totalPapeis: 100, preco: 5, lucroLiquido: 200,
});
conferir('preco justo 10', comPreco.precoJusto, 10);
conferir('margem de 100%', comPreco.margem, 1);
conferir('VPA', comPreco.vpa, 10);
conferir('P/VP', comPreco.pvp, 0.5);
conferir('LPA', comPreco.lpa, 2);
conferir('P/L', comPreco.pl, 2.5);
conferir('valor de mercado', comPreco.valorMercado, 500);

const semPreco = calcular({
  patrimonioLiquido: 1000, ativoCirculante: 0, caixa: 0, dividaLiquida: 0, totalPapeis: 100,
});
conferir('sem preco: margem nula', semPreco.margem, null);
conferir('sem preco: P/VP nulo', semPreco.pvp, null);

/* -------------------------------------------------------------- saida --- */

console.log('testes ok ....... ' + passou);
if (falhou.length) {
  console.log('\nFALHAS (' + falhou.length + '):');
  falhou.forEach((f) => console.log('  - ' + f));
  process.exit(1);
}
console.log('\nTodos os testes passaram.');
