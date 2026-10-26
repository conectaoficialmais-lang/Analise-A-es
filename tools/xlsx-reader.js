/**
 * Leitor minimo de .xlsx, sem dependencia nenhuma.
 * So faz o necessario para esse projeto: abrir o zip, ler sharedStrings e as
 * duas planilhas, devolvendo { NomeDaAba: [{ r, cells: { 1: valor, ... } }] }.
 *
 * As colunas vem numeradas a partir de 1 (A=1, B=2, ...).
 * Celulas com formula devolvem o ULTIMO valor calculado que ficou em cache no
 * arquivo - que e exatamente o que a gente quer conferir.
 */

const fs = require('fs');
const zlib = require('zlib');

const BARRA = String.fromCharCode(92);

/* ------------------------------------------------------------------ unzip */

function lerZip(arquivo) {
  const buf = fs.readFileSync(arquivo);
  const entradas = {};

  // Le pelo "central directory", porque arquivos exportados do Google Sheets
  // deixam os tamanhos zerados no local header (usam data descriptor).
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('zip sem central directory: ' + arquivo);

  const total = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);

  for (let n = 0; n < total; n++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) break;
    const metodo = buf.readUInt16LE(p + 10);
    const comprimido = buf.readUInt32LE(p + 20);
    const tamNome = buf.readUInt16LE(p + 28);
    const tamExtra = buf.readUInt16LE(p + 30);
    const tamComentario = buf.readUInt16LE(p + 32);
    const offsetLocal = buf.readUInt32LE(p + 42);
    const nome = buf.slice(p + 46, p + 46 + tamNome).toString('utf8');
    p += 46 + tamNome + tamExtra + tamComentario;

    // pula o local header para chegar nos bytes
    const nomeLocal = buf.readUInt16LE(offsetLocal + 26);
    const extraLocal = buf.readUInt16LE(offsetLocal + 28);
    const inicio = offsetLocal + 30 + nomeLocal + extraLocal;
    const dados = buf.slice(inicio, inicio + comprimido);
    try {
      entradas[nome] = metodo === 0 ? dados : zlib.inflateRawSync(dados);
    } catch (e) {
      /* entrada ilegivel: ignora */
    }
  }

  if (!Object.keys(entradas).length) {
    throw new Error('nao consegui ler o zip do xlsx: ' + arquivo);
  }
  return entradas;
}

/* ------------------------------------------------------------------- xml */

function desescapar(s) {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(+d))
    .replace(/&amp;/g, '&');
}

function lerSharedStrings(xml) {
  const out = [];
  if (!xml) return out;
  for (const si of xml.matchAll(/<si>([\s\S]*?)<\/si>/g)) {
    let txt = '';
    for (const t of si[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)) txt += t[1];
    out.push(desescapar(txt));
  }
  return out;
}

function indiceColuna(ref) {
  const letras = ref.match(/^([A-Z]+)/)[1];
  let n = 0;
  for (const c of letras) n = n * 26 + (c.charCodeAt(0) - 64);
  return n;
}

function lerPlanilha(xml, strings) {
  const linhas = [];
  for (const lm of xml.matchAll(/<row([^>]*)>([\s\S]*?)<\/row>/g)) {
    const r = Number((lm[1].match(/r="(\d+)"/) || [])[1]);
    const corpo = lm[2];
    const cells = {};
    // Atencao: ha celulas auto-fechadas (<c r="B3" s="101"/>) misturadas com
    // celulas com corpo. Precisa consumir uma de cada vez, na mao.
    const re = /<c\s+([^>]*?)(\/>|>)/g;
    let cm;
    while ((cm = re.exec(corpo)) !== null) {
      const attrs = cm[1];
      const ref = (attrs.match(/r="([A-Z]+\d+)"/) || [])[1];
      if (!ref) continue;
      let dentro = '';
      if (cm[2] === '>') {
        const fim = corpo.indexOf('</c>', re.lastIndex);
        dentro = corpo.slice(re.lastIndex, fim);
        re.lastIndex = fim + 4;
      }
      const tipo = (attrs.match(/t="([^"]+)"/) || [])[1];
      const vm = dentro.match(/<v>([\s\S]*?)<\/v>/);
      if (!vm) continue;
      let valor = desescapar(vm[1]);
      if (tipo === 's') valor = strings[Number(valor)];
      else if (tipo !== 'str' && tipo !== 'inlineStr' && tipo !== 'e') {
        const n = Number(valor);
        if (!Number.isNaN(n)) valor = n;
      }
      cells[indiceColuna(ref)] = valor;
    }
    linhas.push({ r, cells });
  }
  return linhas;
}

/* ----------------------------------------------------------------- api */

function readWorkbook(arquivo) {
  const zip = lerZip(arquivo);
  const texto = (nome) => (zip[nome] ? zip[nome].toString('utf8') : null);

  const strings = lerSharedStrings(texto('xl/sharedStrings.xml'));

  // nome da aba -> rId -> arquivo da planilha
  const wbXml = texto('xl/workbook.xml');
  const relsXml = texto('xl/_rels/workbook.xml.rels') || '';
  const rels = {};
  for (const m of relsXml.matchAll(/<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)) {
    rels[m[1]] = m[2].replace(/^\/?xl\//, '').replace(/^\.\//, '');
  }

  const out = {};
  for (const m of wbXml.matchAll(/<sheet[^>]*\/>/g)) {
    const tag = m[0];
    const nome = desescapar((tag.match(/name="([^"]*)"/) || [])[1] || '');
    const rid = (tag.match(/r:id="([^"]*)"/) || [])[1];
    const alvo = rels[rid];
    if (!alvo) continue;
    const xml = texto('xl/' + alvo.split(BARRA).join('/'));
    if (!xml) continue;
    out[nome] = lerPlanilha(xml, strings);
  }
  return out;
}

module.exports = { readWorkbook };

if (require.main === module) {
  const wb = readWorkbook(process.argv[2]);
  for (const [nome, linhas] of Object.entries(wb)) {
    console.log('=== ' + nome + ' (' + linhas.length + ' linhas) ===');
    for (const l of linhas.slice(0, 5)) console.log(l.r, JSON.stringify(l.cells));
  }
}
