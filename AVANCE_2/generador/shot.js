// Uso: node shot.js dir_html [dir_png] [deviceScale]
const path = require('path');
const fs = require('fs');
const { chromium } = require('/root/.nvm/versions/node/v22.23.3/lib/node_modules/@playwright/mcp/node_modules/playwright');
(async () => {
  const dir = process.argv[2];
  const out = process.argv[3] || dir;
  const dsf = parseFloat(process.argv[4] || '1');
  fs.mkdirSync(out, { recursive: true });
  const b = await chromium.launch({
    executablePath: '/inner_container/opt/playwright/chromium-1232/chrome-linux64/chrome',
    args: ['--disable-crash-reporter', '--no-zygote', '--disable-gpu', '--crash-dumps-dir=/tmp/cd'],
    env: { ...process.env, HOME: '/tmp' },
  });
  const files = fs.readdirSync(dir).filter(f => f.endsWith('.html')).sort();
  const ctx = await b.newContext({ deviceScaleFactor: dsf });
  const p = await ctx.newPage();
  for (const f of files) {
    await p.goto('file://' + path.resolve(dir, f));
    const el = await p.$('.slide');
    const box = await el.boundingBox();
    await p.setViewportSize({ width: Math.ceil(box.width), height: Math.ceil(box.height) });
    await el.screenshot({ path: path.join(out, f.replace('.html', '.png')) });
  }
  await b.close();
  console.log('shots', files.length);
})();
