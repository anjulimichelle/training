'use strict';
// Builds the single-file claude.ai artifact.
// Usage: node artifact/build.js  ->  artifact/dist/ticket-assessment-hub.html
const fs = require('fs');
const path = require('path');
const read = (p) => fs.readFileSync(path.join(__dirname, p), 'utf8');
const maker = fs.readFileSync(path.join(__dirname, '../apps-script/FormMaker.gs'), 'utf8');
const app = read('app.src.js').replace("/*__FORM_MAKER__*/''", () => JSON.stringify(maker)).replace(/<\/script/gi, '<\\/script');
const html = read('page.src.html').replace('/*__APP__*/', () => app);
fs.mkdirSync(path.join(__dirname, 'dist'), { recursive: true });
const out = path.join(__dirname, 'dist', 'ticket-assessment-hub.html');
fs.writeFileSync(out, html);
console.log(`Wrote ${out} (${(html.length / 1024).toFixed(1)} KB)`);
