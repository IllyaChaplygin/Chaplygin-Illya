/* Browser scraper for sites that need real JS + a browser fingerprint.
 *
 * The sandbox terminates TLS with an internal CA that Chromium's own root store
 * does not carry, so instead of loosening TLS in the browser we let Chromium
 * make no network connections at all: every request is intercepted and
 * performed by Node's fetch, which verifies against /root/.ccr/ca-bundle.crt
 * and goes through HTTPS_PROXY. Chromium still runs the page, keeps the cookie
 * jar and solves JS challenges. */
import { chromium } from '/opt/node22/lib/node_modules/playwright/index.mjs';
import fs from 'fs';

const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 ' +
           '(KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36';
const HOP = new Set(['connection', 'keep-alive', 'transfer-encoding', 'upgrade',
                     'proxy-authenticate', 'proxy-authorization', 'te', 'trailer',
                     'content-encoding', 'content-length']);

const targets = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const outPath = process.argv[3];
const shotDir = process.argv[4] || null;
if (shotDir) fs.mkdirSync(shotDir, { recursive: true });

const browser = await chromium.launch({
  executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome',
  headless: false,                       // headful under Xvfb: passes bot checks
  args: ['--no-sandbox', '--disable-blink-features=AutomationControlled',
         '--disable-gpu', '--window-size=1440,1000'],
});
await (async () => {})();
const ctx = await browser.newContext({
  locale: 'uk-UA', timezoneId: 'Europe/Kyiv', userAgent: UA,
  viewport: { width: 1440, height: 2000 },
  extraHTTPHeaders: { 'Accept-Language': 'uk-UA,uk;q=0.9,en;q=0.8' },
});

await ctx.route('**/*', async (route) => {
  const req = route.request();
  const url = req.url();
  if (!/^https?:/.test(url)) return route.continue();
  const headers = { ...req.headers() };
  for (const k of Object.keys(headers)) if (HOP.has(k.toLowerCase())) delete headers[k];
  try {
    const res = await fetch(url, {
      method: req.method(), headers,
      body: ['GET', 'HEAD'].includes(req.method()) ? undefined : req.postDataBuffer(),
      redirect: 'manual',
    });
    const buf = Buffer.from(await res.arrayBuffer());
    const out = {};
    res.headers.forEach((v, k) => { if (!HOP.has(k.toLowerCase())) out[k] = v; });
    const sc = res.headers.getSetCookie ? res.headers.getSetCookie() : [];
    await route.fulfill({ status: res.status, headers: out, body: buf,
                          ...(sc.length ? { headers: { ...out, 'set-cookie': sc.join('\n') } } : {}) });
  } catch (e) {
    await route.abort();
  }
});

const out = {};
for (const [name, url] of Object.entries(targets)) {
  const page = await ctx.newPage();
  try {
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 90000 });
    await page.waitForTimeout(7000);
    // ride out a JS challenge if one is showing
    for (let i = 0; i < 4; i++) {
      const t = await page.title();
      if (!/just a moment|momentik|хвилин/i.test(t)) break;
      await page.waitForTimeout(5000);
    }
    for (const dy of [2500, 3500, 4000, 4000]) {
      await page.mouse.wheel(0, dy); await page.waitForTimeout(1800);
    }
    const txt = await page.evaluate(() => document.body.innerText.replace(/\n{3,}/g, '\n\n'));
    out[name] = { url, title: await page.title(), text: txt.slice(0, 30000) };
    if (shotDir) await page.screenshot({ path: `${shotDir}/${name}.png`, fullPage: false });
    console.error(`OK  ${name.padEnd(22)} ${txt.length} chars`);
  } catch (e) {
    out[name] = { url, error: String(e).split('\n')[0].slice(0, 160) };
    console.error(`ERR ${name.padEnd(22)} ${String(e).split('\n')[0].slice(0, 90)}`);
  }
  await page.close();
}
await browser.close();
fs.writeFileSync(outPath, JSON.stringify(out, null, 1));
