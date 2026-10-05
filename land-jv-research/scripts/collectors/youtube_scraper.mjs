import { chromium } from '/opt/node22/lib/node_modules/playwright/index.mjs';
import fs from 'fs';
const names = JSON.parse(fs.readFileSync(process.argv[2],'utf8'));
const b = await chromium.launch({ executablePath:'/opt/pw-browsers/chromium', args:['--no-sandbox','--disable-dev-shm-usage'] });
const ctx = await b.newContext({ userAgent:'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36', locale:'en-US', viewport:{width:1400,height:1000} });
const fh = fs.createWriteStream(process.argv[3],{flags:'a'});
for (const [label, handle] of names){
  const p = await ctx.newPage();
  let rec = {label, handle, subs:null, videos:null, views:null, joined:null, err:null};
  try{
    await p.goto(`https://www.youtube.com/${handle}/about`,{waitUntil:'domcontentloaded',timeout:50000});
    await p.waitForTimeout(5000);
    const t = await p.evaluate(()=>document.body.innerText||'');
    const s = t.match(/([\d.,]+[KMB]?)\s*subscribers?/i); if(s) rec.subs=s[1];
    const v = t.match(/([\d.,]+[KMB]?)\s*videos?/i); if(v) rec.videos=v[1];
    const w = t.match(/([\d.,]+[KMB]?)\s*views?/i); if(w) rec.views=w[1];
    const j = t.match(/Joined\s+([A-Za-z]+\s+\d{1,2},\s*\d{4})/); if(j) rec.joined=j[1];
    if(/404|not available|doesn't exist/i.test(t.slice(0,300))) rec.err='notfound';
  }catch(e){ rec.err=e.message.split('\n')[0].slice(0,80); }
  fh.write(JSON.stringify(rec)+'\n'); console.log(JSON.stringify(rec));
  await p.close();
}
fh.end(); await b.close();
