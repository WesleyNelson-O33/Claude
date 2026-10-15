// Screenshot each chart page with Playwright, at 3x so they stay crisp in the deck.
const { chromium } = require('playwright');
const fs = require('fs'), path = require('path');

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ deviceScaleFactor: 3 });
  const dir = path.join(__dirname, 'html');
  for (const f of fs.readdirSync(dir).filter(x => x.endsWith('.html'))) {
    const name = f.replace('.html', '');
    await page.goto('file://' + path.join(dir, f));
    await page.waitForTimeout(150);
    const el = await page.$('.wrap');
    const out = path.join(__dirname, 'img', name + '.png');
    await el.screenshot({ path: out });
    const { width, height } = await el.boundingBox();
    console.log(`${name}.png  ${Math.round(width)}x${Math.round(height)} css px`);
  }
  await browser.close();
})();
