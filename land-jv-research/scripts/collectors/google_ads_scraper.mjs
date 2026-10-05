import { chromium } from '/opt/node22/lib/node_modules/playwright/index.mjs';
import fs from 'fs';
const domains = fs.readFileSync(process.argv[2],'utf8').split('\n').map(s=>s.trim()).filter(Boolean);
const outPath = process.argv[3];
const b = await chromium.launch({ executablePath:'/opt/pw-browsers/chromium', args:['--no-sandbox','--disable-dev-shm-usage','--disable-blink-features=AutomationControlled'] });
const ctx = await b.newContext({ userAgent:'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36', locale:'en-US', viewport:{width:1500,height:1100} });
const fh = fs.createWriteStream(outPath,{flags:'a'});
for (const d of domains){
  const p = await ctx.newPage();
  let rec={domain:d, ads:null, plus:false, advertisers:null, ts:new Date().toISOString()};
  try{
    await p.goto(`https://adstransparency.google.com/?region=US&domain=${encodeURIComponent(d)}`,{waitUntil:'domcontentloaded',timeout:60000});
    await p.waitForTimeout(9000);
    const t = await p.evaluate(()=>document.body.innerText||'');
    const m = t.match(/([\d,]+)(\+?)\s*ads?\b/i);
    if(m){ rec.ads = parseInt(m[1].replace(/,/g,''),10); rec.plus = m[2]==='+'; }
    else if(/No ads found/i.test(t)) rec.ads=0;
    else rec.ads=null;
    rec.multi = /multiple advertiser accounts/i.test(t);
  }catch(e){ rec.err=e.message.split('\n')[0].slice(0,120); }
  fh.write(JSON.stringify(rec)+'\n');
  console.log(JSON.stringify(rec));
  await p.close();
}
fh.end(); await b.close();
