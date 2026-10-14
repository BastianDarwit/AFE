// Uso: node clip.js jobs.json [deviceScale]
const fs = require('fs');
const path = require('path');
const { chromium } = require('/root/.nvm/versions/node/v22.23.3/lib/node_modules/@playwright/mcp/node_modules/playwright');
(async () => {
  const jobs = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
  const dsf = parseFloat(process.argv[3] || '2');
  const b = await chromium.launch({
    executablePath: '/inner_container/opt/playwright/chromium-1232/chrome-linux64/chrome',
    args: ['--disable-crash-reporter', '--no-zygote', '--disable-gpu', '--crash-dumps-dir=/tmp/cd'],
    env: { ...process.env, HOME: '/tmp' },
  });
  const ctx = await b.newContext({ deviceScaleFactor: dsf, viewport: { width: 1280, height: 720 } });
  const p = await ctx.newPage();
  for (const j of jobs) {
    await p.goto('file://' + path.resolve(j.html));
    await p.screenshot({ path: j.png, clip: j.clip });
  }
  await b.close();
  console.log('clips', jobs.length);
})();
