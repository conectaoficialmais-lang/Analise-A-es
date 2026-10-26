#!/usr/bin/env node
/**
 * Servidor estatico minimo, so para abrir o app localmente.
 * O app usa fetch() para ler data/empresas.json, e isso nao funciona por file://.
 *
 * Uso: node tools/servidor.js [porta]
 */

const http = require('http');
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const PORTA = Number(process.argv[2]) || 4173;

const TIPOS = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.png': 'image/png',
  '.md': 'text/plain; charset=utf-8',
};

http
  .createServer((req, res) => {
    const url = decodeURIComponent(req.url.split('?')[0]);
    const rel = url === '/' ? 'index.html' : url.replace(/^\/+/, '');
    const alvo = path.resolve(RAIZ, rel);

    // nao deixa sair da pasta do projeto
    if (!alvo.startsWith(RAIZ)) {
      res.writeHead(403).end('403');
      return;
    }

    fs.readFile(alvo, (erro, dados) => {
      if (erro) {
        res.writeHead(404, { 'content-type': 'text/plain; charset=utf-8' }).end('404 — ' + rel);
        return;
      }
      res.writeHead(200, {
        'content-type': TIPOS[path.extname(alvo).toLowerCase()] || 'application/octet-stream',
        'cache-control': 'no-cache',
      });
      res.end(dados);
    });
  })
  .listen(PORTA, () => {
    console.log('Preço Justo rodando em http://localhost:' + PORTA);
  });
