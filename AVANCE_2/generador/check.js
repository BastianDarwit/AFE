// Uso: node check.js dir_html
// Toma captura de cada lámina y reporta textos que desbordan su caja y
// pares de cajas de texto que se superponen.
const fs = require('fs');
const path = require('path');
const { chromium } = require('/root/.nvm/versions/node/v22.23.3/lib/node_modules/@playwright/mcp/node_modules/playwright');
(async () => {
  const dir = process.argv[2];
  const b = await chromium.launch({
    executablePath: '/inner_container/opt/playwright/chromium-1232/chrome-linux64/chrome',
    args: ['--disable-crash-reporter', '--no-zygote', '--disable-gpu', '--crash-dumps-dir=/tmp/cd'],
    env: { ...process.env, HOME: '/tmp' },
  });
  const p = await (await b.newContext({ deviceScaleFactor: 1 })).newPage();
  const files = fs.readdirSync(dir).filter(f => f.endsWith('.html')).sort();
  let total = 0;
  for (const f of files) {
    await p.goto('file://' + path.resolve(dir, f));
    const el = await p.$('.slide');
    const box = await el.boundingBox();
    await p.setViewportSize({ width: Math.ceil(box.width), height: Math.ceil(box.height) });
    await el.screenshot({ path: path.join(dir, f.replace('.html', '.png')) });
    const issues = await p.evaluate(() => {
      const out = [];
      const S = document.querySelector('.slide').getBoundingClientRect();
      const boxes = [];
      document.querySelectorAll('.sp').forEach(sp => {
        const name = sp.dataset.name || '';
        if (name.startsWith('_')) return; // elementos de plantilla
        const tx = sp.querySelector(':scope > .tx');
        if (!tx) return;
        const r = sp.getBoundingClientRect();
        let top = Infinity, bot = -Infinity, left = Infinity, right = -Infinity, has = false;
        tx.querySelectorAll('span').forEach(s => {
          if (!s.textContent.trim()) return;
          for (const rr of s.getClientRects()) {
            has = true;
            top = Math.min(top, rr.top); bot = Math.max(bot, rr.bottom);
            left = Math.min(left, rr.left); right = Math.max(right, rr.right);
          }
        });
        if (!has) return;
        const tol = 1.5;
        if (bot > r.bottom + tol || top < r.top - tol || right > r.right + tol || left < r.left - tol)
          out.push('DESBORDA ' + name + ' texto[' + Math.round(top - S.top) + '-' + Math.round(bot - S.top) + '] caja[' + Math.round(r.top - S.top) + '-' + Math.round(r.bottom - S.top) + '] ancho ' + Math.round(right - left) + '/' + Math.round(r.width));
        if (bot > S.bottom || right > S.right || left < S.left || top < S.top) out.push('FUERA ' + name);
        boxes.push({ name, t: top, b: bot, l: left, r: right });
      });
      for (let i = 0; i < boxes.length; i++) for (let j = i + 1; j < boxes.length; j++) {
        const a = boxes[i], c = boxes[j];
        const ox = Math.min(a.r, c.r) - Math.max(a.l, c.l);
        const oy = Math.min(a.b, c.b) - Math.max(a.t, c.t);
        if (ox > 2 && oy > 2) out.push('SOLAPE ' + a.name + ' <-> ' + c.name + ' (' + Math.round(ox) + 'x' + Math.round(oy) + ')');
      }
      // tablas: texto que supera la fila
      document.querySelectorAll('table.tb').forEach((t, k) => {
        const declared = parseFloat(t.dataset.h || '0');
        const real = t.getBoundingClientRect().height;
        if (declared && real > declared + 2) out.push('TABLA ' + (t.dataset.name || k) + ' crece ' + Math.round(declared) + '->' + Math.round(real));
      });
      return out;
    });
    if (issues.length) { console.log(f); issues.forEach(i => console.log('   ' + i)); total += issues.length; }
  }
  await b.close();
  console.log('problemas:', total);
})();
