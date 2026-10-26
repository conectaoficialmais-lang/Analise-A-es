# Preço Justo

Calculadora de **preço justo de ações da B3 pelo método patrimonial**, feita a partir da
planilha `PREÇO JUSTO.xlsx`. As fórmulas são exatamente as da planilha — nada foi
"melhorado" no modelo.

Site estático, sem build e sem dependências: HTML, CSS e um arquivo de JavaScript.

![status](https://img.shields.io/badge/testes-36%20passando-brightgreen)
![sem dependências](https://img.shields.io/badge/depend%C3%AAncias-0-blue)

---

## As duas contas

```
valor real da empresa = patrimônio líquido + ativo circulante + disponibilidade (caixa) − dívida líquida

preço justo por ação  = valor real ÷ total de papéis
```

Essas eram as colunas **H** e **I** da aba principal (`C+D+E-F` e `H/G`) e as colunas
**N** e **O** da aba `ENERGIA` (`I+J+K-L` e `N/M`). O aplicativo reproduz as duas
em `calcular()`, em [`assets/app.js`](assets/app.js).

A única conta que o aplicativo **acrescenta** é a margem de segurança, porque a aba
principal da planilha não tinha coluna de preço:

```
margem de segurança = (preço justo ÷ preço de hoje) − 1
```

### Ressalva honesta sobre a fórmula

A disponibilidade (caixa) normalmente já está dentro do ativo circulante, e o
patrimônio líquido já é ativo menos passivo. Somar as três contas e subtrair a dívida
conta o caixa duas vezes e mistura estoque com fluxo — o resultado é um teto generoso,
não um valor patrimonial no sentido contábil estrito.

A conta foi mantida como está na planilha, de propósito: é esse o critério em uso.
Só não trate o número como valor contábil.

---

## O que o aplicativo faz

- **Calculadora** — digite as cinco contas e veja valor real, preço justo, margem de
  segurança e os múltiplos derivados (VPA, P/VP, LPA, P/L, DY, valor de mercado), com um
  gráfico de como o valor real se forma.
- **96 empresas já carregadas** da planilha — 64 da aba principal e 32 da aba `ENERGIA`.
  68 delas têm balanço completo e, portanto, preço justo.
- **Ranking** ordenável e filtrável por grupo, busca e completude dos dados, com barra
  de margem de segurança.
- **Edição** — altere qualquer empresa, adicione novas; fica salvo no navegador
  (`localStorage`). Um botão restaura os dados originais da planilha.
- **CSV** — exporta o que está filtrado e importa planilhas suas (aceita `;` ou `,`,
  números em formato brasileiro e BOM do Excel).
- Tema claro/escuro, funciona em celular.

Os campos numéricos aceitam `410545000000`, `410.545.000.000`, `410,5 bi`, `2,5 mi`,
`R$ 1.234.567,89` e negativos.

---

## Rodando

```bash
node tools/servidor.js
```

Abra <http://localhost:4173>. É preciso um servidor: abrir o `index.html` direto pelo
`file://` bloqueia o `fetch` do `data/empresas.json` e a lista pronta não carrega
(a calculadora continua funcionando).

### Testes

```bash
npm test
```

Roda duas coisas:

| arquivo | o que faz |
| --- | --- |
| `tools/testar.js` | 36 testes das funções puras: leitura de números em formato BR, as duas contas, casos de borda (sem papéis, dívida negativa, preço justo negativo) |
| `tools/verificar-fidelidade.js` | recalcula as **68** linhas que a planilha conseguia calcular e compara com o valor que o Excel tinha em cache na célula, mais 10 amostras conferidas à mão |

O segundo é o que importa: é a prova de que a conta do aplicativo dá o mesmo número
que a planilha dava. Hoje passa nas 68 linhas, sem divergência.

### Regerando os dados a partir do Excel

```bash
node tools/build-dataset.js "caminho/para/PREÇO JUSTO.xlsx"
```

Reescreve `data/empresas.json`. O leitor de `.xlsx`
([`tools/xlsx-reader.js`](tools/xlsx-reader.js)) é próprio, sem dependências — abre o
zip pelo *central directory* porque arquivos exportados do Google Sheets deixam os
tamanhos zerados no cabeçalho local.

---

## O que estava furado na planilha

O importador corrige o que podia e sinaliza o resto, sem inventar número nenhum:

| problema | o que foi feito |
| --- | --- |
| Código e empresa trocados em 7 linhas (BRASKEN/BRKM3, CIELO/CIEL3, COPEL/CPLE3, ENERGISA/ENGI3, ENEVA/ENEV3, BRMALLS/BRML3, RUMO/RAIL3) | desvirado por reconhecimento de ticker; `Mils3` virou `MILS3` |
| 28 das 32 empresas da aba `ENERGIA` só tinham patrimônio líquido, sem total de papéis — a coluna de preço justo dava `#DIV/0!` | entram marcadas como incompletas, **sem** preço justo inventado |
| Patrimônio líquido gravado como texto (`"R$ 737.734.000,00"`) nessas mesmas linhas | convertido para número |
| Preço da aba `ENERGIA` vinha de `GOOGLEFINANCE` | ficou o último valor em cache; é cotação antiga, não a de hoje — está avisado na interface |
| DY da TAEE3 escrito `10,13,%`; DY da CEBR3 em 105,35% | importados assim mesmo, com observação na empresa |
| Preços `#N/A` (POWE3, ENMA3B) e `0` (EMAE3) | virariam ruído na margem; ficaram vazios |
| BBAS3 duas vezes, com balanços **diferentes** (PL de R$ 136,7 bi e de R$ 151,9 bi) | ficou o mais recente, com nota apontando o outro |
| TRPL3 duas vezes, com os **mesmos** números | ficou um |
| `B3SA3` não casa com o padrão comum de ticker (4 letras + dígito) | regex ajustada para aceitar dígito no meio |

Os números são balanços de aproximadamente **2021/2022** — nomes como B2W, Via Varejo,
BR Malls e SulAmérica datam a planilha. Atualize as contas antes de decidir qualquer coisa.

---

## Estrutura

```
index.html                     uma página, três abas
assets/style.css               tema claro/escuro
assets/app.js                  cálculo, estado, tabela, CSV
data/empresas.json             96 empresas extraídas da planilha
tools/xlsx-reader.js           leitor de .xlsx sem dependências
tools/build-dataset.js         Excel -> data/empresas.json
tools/verificar-fidelidade.js  prova que a conta bate com a planilha
tools/testar.js                testes das funções puras
tools/servidor.js              servidor estático local
```

---

## Publicando no GitHub Pages

O repositório já traz [`.github/workflows/pages.yml`](.github/workflows/pages.yml), que
roda `npm test` antes de publicar — se a conta do aplicativo deixar de bater com a
planilha, o site não sobe.

Falta um passo manual, uma vez só: em **Settings → Pages**, escolha
*Source: **GitHub Actions***. Sem isso o job `publicar` falha com `Get Pages site failed`;
é só refazer a execução depois de configurar.

O site sai em <https://conectaoficialmais-lang.github.io/Analise-A-es/>.
Pages gratuito exige repositório público.

---

## Aviso

Conteúdo de estudo. Não é recomendação de compra ou venda de ativos, nem análise de
investimento. Confira os balanços na fonte antes de usar qualquer número daqui.

## Licença

[MIT](LICENSE)
