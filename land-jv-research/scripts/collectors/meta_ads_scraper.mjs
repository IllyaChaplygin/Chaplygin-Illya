import { chromium } from '/opt/node22/lib/node_modules/playwright/index.mjs';
import fs from 'fs';
const names = JSON.parse(fs.readFileSync(process.argv[2],'utf8'));
const out = process.argv[3];
const b = await chromium.launch({ executablePath:'/opt/pw-browsers/chromium',
  args:['--no-sandbox','--disable-dev-shm-usage','--disable-blink-features=AutomationControlled'] });
const ctx = await b.newContext({ userAgent:'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36', locale:'en-US', viewport:{width:1500,height:1100} });
const fh = fs.createWriteStream(out,{flags:'a'});
async function probe(q){
  for(let a=0;a<2;a++){
    const p = await ctx.newPage();
    const url=`https://www.facebook.com/ads/library/?active_status=active&ad_type=all&country=US&q=${encodeURIComponent(q)}&search_type=keyword_unordered&media_type=all`;
    try{
      await p.goto(url,{waitUntil:'domcontentloaded',timeout:60000});
      let res=null;
      for(let i=0;i<13;i++){
        await p.waitForTimeout(3000);
        const t=await p.evaluate(()=>document.body.innerText||'');
        const m=t.match(/~?([\d,]+)\s*results?/i);
        if(m){res={n:parseInt(m[1].replace(/,/g,'')),state:'count'};break;}
        if(/No ads match your search criteria/i.test(t)&&i>=5){res={n:0,state:'empty'};break;}
      }
      await p.close();
      if(res&&(res.n>0||a===1)) return res;
    }catch(e){ await p.close().catch(()=>{}); }
  }
  return {n:null,state:'fail'};
}
for(const q of names){
  const r=await probe(q);
  fh.write(JSON.stringify({q,meta_results:r.n,state:r.state})+'\n');
  console.log(`${q.padEnd(28)} -> ${r.n===null?'FAIL':r.n} (${r.state})`);
}
fh.end(); await b.close();
