/* ===========================================================================
   Preço Justo — calculadora de valuation patrimonial
   Reproduz as duas fórmulas da planilha PREÇO JUSTO.xlsx:
     valorReal  = patrimonioLiquido + ativoCirculante + caixa - dividaLiquida
     precoJusto = valorReal / totalPapeis
   =========================================================================== */

'use strict';

const CHAVE_DADOS = 'precojusto:v1:empresas';
const CHAVE_TEMA = 'precojusto:v1:tema';

/* ------------------------------------------------------------------ núcleo */

/** As duas contas da planilha. Única fonte de verdade do cálculo. */
function calcular(e) {
  const contas = [e.patrimonioLiquido, e.ativoCirculante, e.caixa, e.dividaLiquida];
  const temBalanco = contas.every((v) => typeof v === 'number' && Number.isFinite(v));

  const valorReal = temBalanco
    ? e.patrimonioLiquido + e.ativoCirculante + e.caixa - e.dividaLiquida
    : null;

  const papeis = typeof e.totalPapeis === 'number' && e.totalPapeis > 0 ? e.totalPapeis : null;
  const precoJusto = valorReal !== null && papeis ? valorReal / papeis : null;

  const preco = typeof e.preco === 'number' && e.preco > 0 ? e.preco : null;
  const margem = precoJusto !== null && preco ? precoJusto / preco - 1 : null;

  const vpa = typeof e.patrimonioLiquido === 'number' && papeis ? e.patrimonioLiquido / papeis : null;
  const pvp = vpa && vpa > 0 && preco ? preco / vpa : null;
  const lpa = typeof e.lucroLiquido === 'number' && papeis ? e.lucroLiquido / papeis : null;
  const pl = lpa && lpa !== 0 && preco ? preco / lpa : null;
  const valorMercado = preco && papeis ? preco * papeis : null;

  return { valorReal, precoJusto, preco, margem, vpa, pvp, lpa, pl, valorMercado, completa: temBalanco && !!papeis };
}

/* ----------------------------------------------------------------- formato */

const fmtMoeda = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' });
const fmtNum2 = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtInt = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 0 });

function moeda(v) {
  return v === null || v === undefined || !Number.isFinite(v) ? '—' : fmtMoeda.format(v);
}

/** Valores de balanço em escala legível: R$ 410,55 bi. */
function moedaCurta(v) {
  if (v === null || v === undefined || !Number.isFinite(v)) return '—';
  const s = v < 0 ? '-' : '';
  const a = Math.abs(v);
  if (a >= 1e9) return s + 'R$ ' + fmtNum2.format(a / 1e9) + ' bi';
  if (a >= 1e6) return s + 'R$ ' + fmtNum2.format(a / 1e6) + ' mi';
  if (a >= 1e3) return s + 'R$ ' + fmtNum2.format(a / 1e3) + ' mil';
  return s + 'R$ ' + fmtNum2.format(a);
}

function numeroCurto(v) {
  if (v === null || v === undefined || !Number.isFinite(v)) return '—';
  const a = Math.abs(v);
  if (a >= 1e9) return fmtNum2.format(v / 1e9) + ' bi';
  if (a >= 1e6) return fmtNum2.format(v / 1e6) + ' mi';
  if (a >= 1e3) return fmtNum2.format(v / 1e3) + ' mil';
  return fmtInt.format(v);
}

function pct(v) {
  if (v === null || v === undefined || !Number.isFinite(v)) return '—';
  return (v > 0 ? '+' : '') + fmtNum2.format(v * 100) + '%';
}

function pctSimples(v) {
  if (v === null || v === undefined || !Number.isFinite(v)) return '—';
  return fmtNum2.format(v * 100) + '%';
}

function mult(v) {
  return v === null || v === undefined || !Number.isFinite(v) ? '—' : fmtNum2.format(v) + '×';
}

/**
 * Lê número digitado em português, tolerando separadores e sufixos.
 *  "410.545.000.000" -> 410545000000     "410,5 bi" -> 410500000000
 *  "1234,56"         -> 1234.56          "-13.132"  -> -13132
 */
function lerNumero(txt) {
  if (txt === null || txt === undefined) return null;
  let s = String(txt).trim().toLowerCase();
  if (!s) return null;

  let escala = 1;
  const sufixo = s.match(/(bi|bilh(?:ao|ão|oes|ões)|mi|milh(?:ao|ão|oes|ões)|mil|k)\.?$/);
  if (sufixo) {
    const u = sufixo[1];
    escala = u.startsWith('bi') ? 1e9 : u.startsWith('mi') && u !== 'mil' ? 1e6 : 1e3;
    s = s.slice(0, sufixo.index).trim();
  }

  s = s.replace(/r\$/g, '').replace(/%/g, '').replace(/[\s ]/g, '');
  if (!s || s === '-') return null;

  const temPonto = s.includes('.');
  const temVirgula = s.includes(',');

  if (temPonto && temVirgula) {
    // formato brasileiro cheio: ponto é milhar, vírgula é decimal
    s = s.split('.').join('').replace(',', '.');
  } else if (temVirgula) {
    s = s.replace(',', '.');
  } else if (temPonto) {
    // só ponto: milhar se vier em grupos de 3 (1.234.567), senão decimal (1.5)
    const partes = s.replace('-', '').split('.');
    const milhar = partes.length > 1 && partes.slice(1).every((p) => p.length === 3);
    if (milhar) s = partes.join('') * (s.startsWith('-') ? -1 : 1) + '';
  }

  const n = Number(s);
  return Number.isFinite(n) ? n * escala : null;
}

/* ------------------------------------------------------------------- dados */

const estado = {
  base: [],       // o que veio da planilha
  meta: {},
  empresas: [],   // base + edições do usuário
  ordem: { campo: 'margem', dir: 'desc' },
  idAtual: null,
};

function lerLocal(chave, padrao) {
  try {
    const bruto = localStorage.getItem(chave);
    return bruto === null ? padrao : JSON.parse(bruto);
  } catch (e) {
    return padrao;
  }
}

function gravarLocal(chave, valor) {
  try {
    localStorage.setItem(chave, JSON.stringify(valor));
    return true;
  } catch (e) {
    return false;
  }
}

/** Edições do usuário: { editadas: { id: {...} }, novas: [ {...} ], removidas: [ id ] } */
function lerEdicoes() {
  const d = lerLocal(CHAVE_DADOS, null);
  return {
    editadas: (d && d.editadas) || {},
    novas: (d && d.novas) || [],
    removidas: (d && d.removidas) || [],
  };
}

function gravarEdicoes(ed) {
  if (!gravarLocal(CHAVE_DADOS, ed)) {
    aviso('Não consegui gravar no navegador (armazenamento bloqueado). A edição vale só nesta aba.');
  }
}

function montarEmpresas() {
  const ed = lerEdicoes();
  const removidas = new Set(ed.removidas);
  const lista = [];

  for (const base of estado.base) {
    if (removidas.has(base.id)) continue;
    const over = ed.editadas[base.id];
    lista.push(over ? Object.assign({}, base, over, { editada: true }) : base);
  }
  for (const nova of ed.novas) {
    if (!removidas.has(nova.id)) lista.push(Object.assign({ minha: true }, nova));
  }

  estado.empresas = lista.map((e) => Object.assign({}, e, calcular(e)));
}

/* -------------------------------------------------------------------- tema */

function aplicarTema(t) {
  document.documentElement.dataset.tema = t;
  const icone = document.querySelector('[data-tema-icone]');
  if (icone) icone.textContent = t === 'escuro' ? '☀' : '☾';
}

function iniciarTema() {
  const salvo = lerLocal(CHAVE_TEMA, null);
  const escuroPadrao = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
  aplicarTema(salvo || (escuroPadrao ? 'escuro' : 'claro'));

  document.getElementById('btn-tema').addEventListener('click', () => {
    const novo = document.documentElement.dataset.tema === 'escuro' ? 'claro' : 'escuro';
    aplicarTema(novo);
    gravarLocal(CHAVE_TEMA, novo);
  });
}

/* -------------------------------------------------------------------- abas */

function iniciarAbas() {
  const botoes = document.querySelectorAll('.aba');
  botoes.forEach((b) => {
    b.addEventListener('click', () => {
      botoes.forEach((o) => o.classList.toggle('ativa', o === b));
      document.querySelectorAll('.painel').forEach((p) => {
        p.classList.toggle('ativo', p.id === 'painel-' + b.dataset.aba);
      });
      if (b.dataset.aba === 'ranking') desenharTabela();
    });
  });
}

/* ------------------------------------------------------------- calculadora */

const campo = (id) => document.getElementById(id);

function lerFormulario() {
  return {
    codigo: campo('in-codigo').value.trim().toUpperCase(),
    empresa: campo('in-empresa').value.trim(),
    setor: campo('in-setor').value,
    patrimonioLiquido: lerNumero(campo('in-pl').value),
    ativoCirculante: lerNumero(campo('in-ac').value),
    caixa: lerNumero(campo('in-caixa').value),
    dividaLiquida: lerNumero(campo('in-divida').value),
    totalPapeis: lerNumero(campo('in-papeis').value),
    preco: lerNumero(campo('in-preco').value),
    lucroLiquido: lerNumero(campo('in-lucro').value),
    dy: (() => {
      const v = lerNumero(campo('in-dy').value);
      return v === null ? null : v / 100;
    })(),
  };
}

function escreverFormulario(e) {
  campo('in-codigo').value = e.codigo || '';
  campo('in-empresa').value = e.empresa || '';
  campo('in-setor').value = ['Geral', 'Energia'].includes(e.setor) ? e.setor : 'Outro';
  // escreve em pt-BR; "12.345" relido daria 12345, então o decimal vai com vírgula
  const posNum = (id, v) => {
    campo(id).value =
      v === null || v === undefined || !Number.isFinite(v) ? '' : String(v).replace('.', ',');
  };
  posNum('in-pl', e.patrimonioLiquido);
  posNum('in-ac', e.ativoCirculante);
  posNum('in-caixa', e.caixa);
  posNum('in-divida', e.dividaLiquida);
  posNum('in-papeis', e.totalPapeis);
  posNum('in-preco', e.preco);
  posNum('in-lucro', e.lucroLiquido);
  campo('in-dy').value = e.dy === null || e.dy === undefined ? '' : fmtNum2.format(e.dy * 100);
  atualizar();
}

function aviso(msg, ok) {
  const el = campo('aviso-calc');
  if (!msg) { el.hidden = true; return; }
  el.textContent = msg;
  el.className = 'aviso' + (ok ? ' ok' : '');
  el.hidden = false;
}

function atualizarEcos() {
  document.querySelectorAll('.eco[data-eco-de]').forEach((eco) => {
    const alvo = campo(eco.dataset.ecoDe);
    const v = lerNumero(alvo.value);
    const papeis = eco.dataset.ecoDe === 'in-papeis';
    eco.textContent = v === null ? '—' : papeis ? numeroCurto(v) + ' papéis' : moedaCurta(v);
    eco.classList.toggle('ativo', v !== null);
  });
}

function desenharComposicao(e, r) {
  const alvo = document.getElementById('grafico-composicao');
  if (r.valorReal === null) {
    alvo.innerHTML = '<p class="vazio">Preencha as quatro contas do balanço para ver a composição.</p>';
    return;
  }

  const partes = [
    { nome: 'Patrimônio líquido', v: e.patrimonioLiquido, sinal: 1 },
    { nome: 'Ativo circulante', v: e.ativoCirculante, sinal: 1 },
    { nome: 'Disponibilidade (caixa)', v: e.caixa, sinal: 1 },
    { nome: 'Dívida líquida', v: e.dividaLiquida, sinal: -1 },
  ];

  const maior = Math.max(...partes.map((p) => Math.abs(p.v)), 1);

  alvo.innerHTML =
    partes
      .map((p) => {
        const larg = Math.min(100, (Math.abs(p.v) / maior) * 100);
        // uma conta negativa que entra somando (ou positiva subtraindo) inverte a cor
        const subtrai = p.sinal * Math.sign(p.v || 1) < 0;
        return `
      <div class="comp-linha">
        <div class="comp-topo">
          <span class="comp-nome">${p.sinal < 0 ? '−' : '+'} ${p.nome}</span>
          <span class="comp-valor">${moedaCurta(p.v)}</span>
        </div>
        <div class="comp-trilha"><div class="comp-barra${subtrai ? ' menos' : ''}" style="width:${larg}%"></div></div>
      </div>`;
      })
      .join('') +
    `<div class="comp-total">
       <span>Valor real da empresa</span>
       <strong>${moedaCurta(r.valorReal)}</strong>
     </div>`;
}

function atualizar() {
  atualizarEcos();

  const e = lerFormulario();
  const r = calcular(e);

  const gigante = campo('out-preco-justo');
  gigante.textContent = r.precoJusto === null ? '—' : moeda(r.precoJusto);
  gigante.classList.toggle('vazio-val', r.precoJusto === null);
  gigante.classList.toggle('neg', r.precoJusto !== null && r.precoJusto < 0);

  const ver = campo('out-veredito');
  if (r.margem === null) {
    ver.hidden = true;
  } else {
    const barato = r.margem > 0.05;
    const caro = r.margem < -0.05;
    ver.hidden = false;
    ver.className = 'veredito ' + (barato ? 'barato' : caro ? 'caro' : 'justo');
    ver.textContent = barato
      ? `Negociando ${pctSimples(Math.abs(r.margem))} abaixo do preço justo`
      : caro
        ? `Negociando ${pctSimples(Math.abs(r.margem))} acima do preço justo`
        : 'Perto do preço justo';
  }

  const por = (id, texto, classe) => {
    const el = campo(id);
    el.textContent = texto;
    el.className = classe || (texto === '—' ? 'sem' : '');
  };

  por('out-valor-real', moedaCurta(r.valorReal));
  por('out-preco', moeda(r.preco));
  por('out-margem', r.margem === null ? '—' : pct(r.margem), r.margem === null ? 'sem' : r.margem > 0 ? 'pos' : 'neg');
  por('out-vpa', moeda(r.vpa));
  por('out-pvp', mult(r.pvp));
  por('out-lpa', moeda(r.lpa));
  por('out-pl', mult(r.pl));
  por('out-dy', e.dy === null ? '—' : pctSimples(e.dy));
  por('out-mercado', moedaCurta(r.valorMercado));

  desenharComposicao(e, r);
  campo('btn-excluir').hidden = !estado.idAtual;
}

/**
 * Procura uma empresa pelo que foi digitado.
 * Vários códigos aparecem nas duas abas da planilha (ELET3, CPLE3, TAEE3,
 * ALUP3, ENGI3, TRPL3) — entre duas com o mesmo código, vale a que tem
 * balanço completo, senão a busca traz a versão sem preço justo.
 */
function procurarEmpresa(texto) {
  const busca = String(texto || '').trim().toLowerCase();
  if (!busca) return null;

  const melhor = (lista) => lista.find((e) => e.completa) || lista[0] || null;

  const porCodigo = estado.empresas.filter((e) => e.codigo.toLowerCase() === busca);
  if (porCodigo.length) return melhor(porCodigo);

  const porPrefixo = estado.empresas.filter((e) => e.codigo.toLowerCase().startsWith(busca));
  if (porPrefixo.length) return melhor(porPrefixo);

  const porNome = estado.empresas.filter((e) => (e.empresa || '').toLowerCase().includes(busca));
  return melhor(porNome);
}

/** Põe a empresa no formulário e diz o que há de ressalva nela. */
/**
 * A aba ENERGIA tinha alguns múltiplos digitados à mão. Mostra o que estava
 * anotado — vem de `e.anotado`, não dos campos calculados de mesmo nome.
 */
function mostrarAnotados(e) {
  const el = campo('anotado-planilha');
  const a = e.anotado || {};
  const itens = [
    ['LPA', a.lpa, moeda],
    ['P/L', a.pl, mult],
    ['VPA', a.vpa, moeda],
    ['P/VP', a.pvp, mult],
  ].filter(([, v]) => typeof v === 'number' && Number.isFinite(v));

  if (!itens.length) { el.hidden = true; return; }
  el.innerHTML =
    `<strong>Anotado à mão na planilha para ${e.codigo}:</strong> ` +
    itens.map(([nome, v, f]) => `${nome} ${f(v)}`).join(' · ');
  el.hidden = false;
}

function abrirEmpresa(e) {
  estado.idAtual = e.id;
  escreverFormulario(e);
  mostrarAnotados(e);

  const partes = [];
  if (!e.completa) partes.push('faltam dados do balanço nessa linha da planilha');
  if (e.fonte === 'ENERGIA' && e.preco && !e.editada && !e.minha) {
    partes.push('o preço é a cotação antiga em cache do GOOGLEFINANCE');
  }
  if (e.obs) partes.push(e.obs);
  aviso(partes.length ? `${e.codigo}: ${partes.join('; ')}.` : `${e.codigo} carregada.`, !partes.length);
}

function carregarEmpresa(texto) {
  if (!String(texto || '').trim()) return;
  const achada = procurarEmpresa(texto);
  if (!achada) {
    aviso(`Não encontrei "${texto}" na lista. Você pode digitar as contas e salvar como empresa nova.`);
    return;
  }
  abrirEmpresa(achada);
}

function salvar() {
  const e = lerFormulario();
  if (!e.codigo) { aviso('Informe pelo menos o código da ação para salvar.'); return; }

  const ed = lerEdicoes();
  let existente = estado.base.find((b) => b.id === estado.idAtual);

  // Trocar o código quer dizer "outra empresa", não renomear a da planilha:
  // vira um registro novo e o original fica intacto.
  if (existente && existente.codigo !== e.codigo) {
    existente = null;
    estado.idAtual = null;
  }

  if (existente) {
    ed.editadas[existente.id] = e;
  } else {
    const id = estado.idAtual || 'minha-' + e.codigo.toLowerCase().replace(/[^a-z0-9]+/g, '-');
    const i = ed.novas.findIndex((n) => n.id === id);
    const reg = Object.assign({ id, fonte: 'minha lista', linhaPlanilha: null, obs: null }, e);
    if (i >= 0) ed.novas[i] = reg; else ed.novas.push(reg);
    ed.removidas = ed.removidas.filter((r) => r !== id);
    estado.idAtual = id;
  }

  gravarEdicoes(ed);
  montarEmpresas();
  preencherDatalist();
  desenharTabela();
  atualizar();
  aviso(`${e.codigo} salva. Aparece no ranking com o selo “minha”.`, true);
}

function excluir() {
  if (!estado.idAtual) return;
  const alvo = estado.empresas.find((e) => e.id === estado.idAtual);
  if (!alvo) return;
  if (!window.confirm(`Remover ${alvo.codigo} da lista?`)) return;

  const ed = lerEdicoes();
  ed.novas = ed.novas.filter((n) => n.id !== alvo.id);
  delete ed.editadas[alvo.id];
  if (estado.base.some((b) => b.id === alvo.id)) ed.removidas.push(alvo.id);
  gravarEdicoes(ed);

  estado.idAtual = null;
  montarEmpresas();
  preencherDatalist();
  desenharTabela();
  limpar();
  aviso(`${alvo.codigo} removida. “Restaurar dados da planilha” traz de volta.`, true);
}

function limpar() {
  estado.idAtual = null;
  document.getElementById('form-calc').reset();
  campo('in-setor').value = 'Geral';
  campo('anotado-planilha').hidden = true;
  aviso(null);
  atualizar();
}

/* ----------------------------------------------------------------- ranking */

function preencherDatalist() {
  // um código por sugestão, mesmo que ele apareça nas duas abas da planilha
  const porCodigo = new Map();
  for (const e of estado.empresas) {
    const atual = porCodigo.get(e.codigo);
    if (!atual || (e.completa && !atual.completa)) porCodigo.set(e.codigo, e);
  }

  const dl = document.getElementById('lista-empresas');
  dl.innerHTML = Array.from(porCodigo.values())
    .sort((a, b) => a.codigo.localeCompare(b.codigo))
    .map((e) => {
      const nome = e.empresa && e.empresa !== e.codigo ? e.empresa : e.setor;
      return `<option value="${e.codigo}">${nome.replace(/"/g, '')}</option>`;
    })
    .join('');

  const sel = document.getElementById('filtro-setor');
  const setores = Array.from(new Set(estado.empresas.map((e) => e.setor))).sort();
  const atual = sel.value;
  sel.innerHTML = '<option value="">Todos</option>' + setores.map((s) => `<option value="${s}">${s}</option>`).join('');
  sel.value = atual;
}

function empresasFiltradas() {
  const busca = document.getElementById('filtro-busca').value.trim().toLowerCase();
  const setor = document.getElementById('filtro-setor').value;
  const dados = document.getElementById('filtro-dados').value;

  return estado.empresas.filter((e) => {
    if (setor && e.setor !== setor) return false;
    if (dados === 'completas' && !e.completa) return false;
    if (dados === 'incompletas' && e.completa) return false;
    if (busca && !(e.codigo + ' ' + (e.empresa || '')).toLowerCase().includes(busca)) return false;
    return true;
  });
}

function ordenar(lista) {
  const { campo: c, dir } = estado.ordem;
  const fator = dir === 'asc' ? 1 : -1;
  return lista.slice().sort((a, b) => {
    const va = a[c];
    const vb = b[c];
    const na = va === null || va === undefined;
    const nb = vb === null || vb === undefined;
    if (na && nb) {
      // empate por falta de dado: desempata pelo preço justo, depois pelo código
      const pa = a.precoJusto;
      const pb = b.precoJusto;
      if (c !== 'precoJusto' && pa !== null && pb !== null && pa !== pb) return pb - pa;
      return a.codigo.localeCompare(b.codigo);
    }
    if (na) return 1;   // sem valor sempre no fim
    if (nb) return -1;
    if (typeof va === 'string') return va.localeCompare(vb) * fator;
    return (va - vb) * fator;
  });
}

function desenharTabela() {
  const corpo = document.querySelector('#tabela tbody');
  const lista = ordenar(empresasFiltradas());

  document.querySelectorAll('.th-ord').forEach((th) => {
    if (th.dataset.ordenar === estado.ordem.campo) th.dataset.dir = estado.ordem.dir;
    else delete th.dataset.dir;
  });

  const comMargem = lista.filter((e) => e.margem !== null).length;
  document.getElementById('resumo-ranking').innerHTML =
    `<b>${lista.length}</b> empresa(s) na lista · <b>${lista.filter((e) => e.completa).length}</b> com balanço completo` +
    (comMargem ? ` · <b>${comMargem}</b> com preço para comparar` : '');

  if (!lista.length) {
    corpo.innerHTML = '<tr class="linha-vazia"><td colspan="8">Nada encontrado com esses filtros.</td></tr>';
    return;
  }

  const maiorMargem = Math.max(1, ...lista.map((e) => Math.abs(e.margem || 0)));

  corpo.innerHTML = lista
    .map((e) => {
      const selos =
        (e.minha ? '<span class="selo minha">minha</span>' : '') +
        (e.editada ? '<span class="selo minha">editada</span>' : '') +
        (!e.completa ? '<span class="selo falta">sem papéis</span>' : '') +
        (e.fonte === 'ENERGIA' && e.preco && !e.editada && !e.minha
          ? '<span class="selo falta" title="Cotação antiga, em cache do GOOGLEFINANCE na planilha">preço antigo</span>'
          : '');

      const nome = e.empresa && e.empresa !== e.codigo ? e.empresa : '—';

      let celMargem = '<td class="num sem">—</td>';
      if (e.margem !== null) {
        const larg = Math.min(100, (Math.abs(e.margem) / maiorMargem) * 100);
        celMargem = `<td class="num ${e.margem > 0 ? 'pos' : 'neg'}">
          <span class="celula-margem">
            <span class="margem-trilha"><span class="margem-barra${e.margem < 0 ? ' neg' : ''}" style="width:${larg}%"></span></span>
            ${pct(e.margem)}
          </span></td>`;
      }

      return `<tr data-id="${e.id}" tabindex="0">
        <td><span class="cod-nome"><span class="cod">${e.codigo}</span>${selos}</span></td>
        <td${nome === '—' ? ' class="sem"' : ''}>${nome}</td>
        <td class="num ${e.precoJusto === null ? 'sem' : e.precoJusto < 0 ? 'neg' : ''}">${e.precoJusto === null ? '—' : fmtNum2.format(e.precoJusto)}</td>
        <td class="num ${e.preco === null ? 'sem' : ''}">${e.preco === null ? '—' : fmtNum2.format(e.preco)}</td>
        ${celMargem}
        <td class="num ${e.valorReal === null ? 'sem' : ''}">${moedaCurta(e.valorReal)}</td>
        <td class="num ${e.totalPapeis ? '' : 'sem'}">${numeroCurto(e.totalPapeis)}</td>
        <td>${e.setor}</td>
      </tr>`;
    })
    .join('');
}

/* --------------------------------------------------------------------- csv */

const COLUNAS_CSV = [
  ['codigo', 'CODIGO'],
  ['empresa', 'EMPRESA'],
  ['setor', 'GRUPO'],
  ['patrimonioLiquido', 'PATRIMONIO LIQUIDO'],
  ['ativoCirculante', 'ATIVO CIRCULANTE'],
  ['caixa', 'DISPONIBILIDADE (CAIXA)'],
  ['dividaLiquida', 'DIVIDA LIQUIDA'],
  ['totalPapeis', 'TOTAL DE PAPEIS'],
  ['valorReal', 'VALOR REAL'],
  ['precoJusto', 'PRECO JUSTO'],
  ['preco', 'PRECO HOJE'],
  ['margem', 'MARGEM DE SEGURANCA'],
  ['fonte', 'FONTE'],
];

function exportarCsv() {
  const lista = ordenar(empresasFiltradas());
  const celula = (v) => {
    if (v === null || v === undefined) return '';
    if (typeof v === 'number') return String(v).replace('.', ',');
    return /[;"\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
  };

  const csv = [COLUNAS_CSV.map((c) => c[1]).join(';')]
    .concat(lista.map((e) => COLUNAS_CSV.map((c) => celula(e[c[0]])).join(';')))
    .join('\r\n');

  // BOM para o Excel brasileiro abrir com acento e ponto-e-vírgula certos
  const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'preco-justo.csv';
  a.click();
  URL.revokeObjectURL(url);
}

function dividirLinhaCsv(linha, sep) {
  const saida = [];
  let atual = '';
  let dentro = false;
  for (let i = 0; i < linha.length; i++) {
    const c = linha[i];
    if (dentro) {
      if (c === '"' && linha[i + 1] === '"') { atual += '"'; i++; }
      else if (c === '"') dentro = false;
      else atual += c;
    } else if (c === '"') dentro = true;
    else if (c === sep) { saida.push(atual); atual = ''; }
    else atual += c;
  }
  saida.push(atual);
  return saida;
}

function importarCsv(texto) {
  const limpo = texto.replace(/^﻿/, '').trim();
  const linhas = limpo.split(/\r?\n/).filter((l) => l.trim());
  if (linhas.length < 2) { aviso('O CSV não tem linhas de dados.'); return; }

  const sep = (linhas[0].match(/;/g) || []).length >= (linhas[0].match(/,/g) || []).length ? ';' : ',';
  const cabeca = dividirLinhaCsv(linhas[0], sep).map((h) =>
    h.trim().toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
  );

  const achar = (...nomes) => {
    for (const n of nomes) {
      const i = cabeca.findIndex((h) => h === n || h.startsWith(n));
      if (i >= 0) return i;
    }
    return -1;
  };

  const idx = {
    codigo: achar('CODIGO', 'TICKER', 'ATIVO'),
    empresa: achar('EMPRESA', 'NOME'),
    setor: achar('GRUPO', 'SETOR'),
    pl: achar('PATRIMONIO'),
    ac: achar('ATIVO CIRCULANTE'),
    caixa: achar('DISPONIBILIDADE', 'CAIXA'),
    divida: achar('DIVIDA'),
    papeis: achar('TOTAL DE PAPEIS', 'PAPEIS', 'ACOES'),
    preco: achar('PRECO HOJE', 'PRECO ATUAL', 'COTACAO'),
  };

  if (idx.codigo < 0) { aviso('Não achei a coluna CODIGO no CSV.'); return; }

  const ed = lerEdicoes();
  let gravadas = 0;

  for (const linha of linhas.slice(1)) {
    const c = dividirLinhaCsv(linha, sep);
    const pega = (i) => (i >= 0 && i < c.length ? c[i].trim() : '');
    const codigo = pega(idx.codigo).toUpperCase();
    if (!codigo) continue;

    const reg = {
      codigo,
      empresa: pega(idx.empresa) || codigo,
      setor: pega(idx.setor) || 'Outro',
      patrimonioLiquido: lerNumero(pega(idx.pl)),
      ativoCirculante: lerNumero(pega(idx.ac)),
      caixa: lerNumero(pega(idx.caixa)),
      dividaLiquida: lerNumero(pega(idx.divida)),
      totalPapeis: lerNumero(pega(idx.papeis)),
      preco: lerNumero(pega(idx.preco)),
      lucroLiquido: null,
      dy: null,
    };

    // casa por código + grupo antes de por código só: o próprio export leva a
    // coluna GRUPO, e sem ela um ELET3 da aba Geral sobrescreveria o da Energia
    const grupo = pega(idx.setor);
    const base =
      (grupo && estado.base.find((b) => b.codigo === codigo && b.setor === grupo)) ||
      estado.base.find((b) => b.codigo === codigo && b.completa) ||
      estado.base.find((b) => b.codigo === codigo);

    if (base) {
      ed.editadas[base.id] = Object.assign({}, base, reg);
      ed.removidas = ed.removidas.filter((r) => r !== base.id);
    } else {
      const id = 'csv-' + codigo.toLowerCase().replace(/[^a-z0-9]+/g, '-');
      const i = ed.novas.findIndex((n) => n.id === id);
      const novo = Object.assign({ id, fonte: 'CSV importado', linhaPlanilha: null, obs: null }, reg);
      if (i >= 0) ed.novas[i] = novo; else ed.novas.push(novo);
      ed.removidas = ed.removidas.filter((r) => r !== id);
    }
    gravadas++;
  }

  gravarEdicoes(ed);
  montarEmpresas();
  preencherDatalist();
  desenharTabela();
  aviso(`${gravadas} linha(s) importada(s) do CSV.`, true);
}

function restaurar() {
  if (!window.confirm('Isso apaga todas as suas edições e empresas adicionadas, voltando aos dados da planilha. Continuar?')) return;
  try { localStorage.removeItem(CHAVE_DADOS); } catch (e) { /* ignora */ }
  montarEmpresas();
  preencherDatalist();
  desenharTabela();
  limpar();
  aviso('Dados da planilha restaurados.', true);
}

/* ------------------------------------------------------------- metodologia */

function preencherMetodologia() {
  const total = estado.base.length;
  const completas = estado.base.filter((e) => calcular(e).completa).length;
  const geral = estado.base.filter((e) => e.setor === 'Geral').length;
  const energia = estado.base.filter((e) => e.setor === 'Energia').length;

  document.getElementById('meta-dados').textContent =
    `${total} empresas vieram da planilha — ${geral} da aba principal e ${energia} da aba ENERGIA. ` +
    `Dessas, ${completas} têm balanço completo e total de papéis, então têm preço justo. ` +
    `As outras ${total - completas} entram marcadas como incompletas. Os números são ` +
    `${estado.meta.periodoAproximado || 'os que estavam na planilha'}.`;

  document.getElementById('meta-conferidas').textContent = String(completas);

  const ul = document.getElementById('meta-avisos');
  ul.innerHTML = (estado.meta.avisos || []).map((a) => `<li>${a}</li>`).join('');
}

/* ------------------------------------------------------------------ início */

function ligarEventos() {
  document.getElementById('form-calc').addEventListener('input', (ev) => {
    if (ev.target.id !== 'busca-empresa') atualizar();
  });
  document.getElementById('form-calc').addEventListener('submit', (ev) => ev.preventDefault());

  document.getElementById('btn-carregar').addEventListener('click', () =>
    carregarEmpresa(campo('busca-empresa').value)
  );
  campo('busca-empresa').addEventListener('change', (ev) => carregarEmpresa(ev.target.value));
  campo('busca-empresa').addEventListener('keydown', (ev) => {
    if (ev.key === 'Enter') { ev.preventDefault(); carregarEmpresa(ev.target.value); }
  });

  document.getElementById('btn-salvar').addEventListener('click', salvar);
  document.getElementById('btn-limpar').addEventListener('click', limpar);
  document.getElementById('btn-excluir').addEventListener('click', excluir);
  document.getElementById('btn-restaurar').addEventListener('click', restaurar);

  ['filtro-busca', 'filtro-setor', 'filtro-dados'].forEach((id) =>
    document.getElementById(id).addEventListener('input', desenharTabela)
  );

  document.querySelectorAll('.th-ord').forEach((th) => {
    th.addEventListener('click', () => {
      const c = th.dataset.ordenar;
      if (estado.ordem.campo === c) {
        estado.ordem.dir = estado.ordem.dir === 'asc' ? 'desc' : 'asc';
      } else {
        estado.ordem.campo = c;
        estado.ordem.dir = ['codigo', 'empresa', 'setor'].includes(c) ? 'asc' : 'desc';
      }
      desenharTabela();
    });
  });

  const abrirLinha = (tr) => {
    if (!tr) return;
    // Abre exatamente a linha clicada: procurar de novo pelo código traria
    // a outra empresa de mesmo ticker que existe na outra aba da planilha.
    const e = estado.empresas.find((x) => x.id === tr.dataset.id);
    if (!e) return;
    document.querySelector('.aba[data-aba="calculadora"]').click();
    abrirEmpresa(e);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const corpo = document.querySelector('#tabela tbody');
  corpo.addEventListener('click', (ev) => abrirLinha(ev.target.closest('tr[data-id]')));
  corpo.addEventListener('keydown', (ev) => {
    if (ev.key === 'Enter' || ev.key === ' ') {
      ev.preventDefault();
      abrirLinha(ev.target.closest('tr[data-id]'));
    }
  });

  document.getElementById('btn-exportar').addEventListener('click', exportarCsv);
  document.getElementById('arq-importar').addEventListener('change', (ev) => {
    const arq = ev.target.files && ev.target.files[0];
    if (!arq) return;
    const leitor = new FileReader();
    leitor.onload = () => importarCsv(String(leitor.result));
    leitor.onerror = () => aviso('Não consegui ler esse arquivo.');
    leitor.readAsText(arq, 'utf-8');
    ev.target.value = '';
  });
}

async function iniciar() {
  iniciarTema();
  iniciarAbas();
  ligarEventos();
  atualizar();

  try {
    const resp = await fetch('data/empresas.json', { cache: 'no-cache' });
    if (!resp.ok) throw new Error('HTTP ' + resp.status);
    const dados = await resp.json();
    estado.base = dados.empresas || [];
    estado.meta = dados;
  } catch (erro) {
    aviso('Não consegui carregar data/empresas.json. A calculadora funciona, mas sem a lista pronta. (Abra o projeto por um servidor, não pelo file://)');
    estado.base = [];
    estado.meta = {};
  }

  montarEmpresas();
  preencherDatalist();
  preencherMetodologia();
  desenharTabela();
}

// No navegador, sobe o app. Em Node (tools/testar.js) só expõe as funções puras.
if (typeof document !== 'undefined') {
  iniciar();
} else if (typeof module !== 'undefined' && module.exports) {
  module.exports = { calcular, lerNumero, moedaCurta, numeroCurto, pct, mult };
}
