'use strict';
// Builds the single-file claude.ai artifact from the shared core, the demo data and the page source.
// Usage: node artifact/build.js  ->  artifact/dist/ticket-assessment-hub.html
const fs = require('fs');
const path = require('path');
const dir = __dirname;
const read = (p) => fs.readFileSync(path.join(dir, p), 'utf8');
const core = read('../server/core.js');
const demo = read('../server/demo.js');
const app = read('app.src.js');
const demoShim = `(function () {\n  var module = { exports: {} };\n  var require = function () { return window.HubCore; };\n${demo}\n  window.__HUB_DEMO_QUIZ__ = function (trainees) { return module.exports.sampleQuiz(trainees); };\n  window.__HUB_DEMO_ANSWERS__ = module.exports.sampleAnswers;\n})();`;
const guard = (s) => s.replace(/<\/script/gi, '<\\/script');
const html = read('page.src.html')
  .replace('/*__CORE__*/', () => guard(core + '\n' + demoShim))
  .replace('/*__APP__*/', () => guard(app));
fs.mkdirSync(path.join(dir, 'dist'), { recursive: true });
const out = path.join(dir, 'dist', 'ticket-assessment-hub.html');
fs.writeFileSync(out, html);
console.log(`Wrote ${out} (${(html.length / 1024).toFixed(1)} KB)`);
