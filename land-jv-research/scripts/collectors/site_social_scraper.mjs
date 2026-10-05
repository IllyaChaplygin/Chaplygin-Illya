import { chromium } from '/opt/node22/lib/node_modules/playwright/index.mjs';
import fs from 'fs';
const sites = JSON.parse(fs.readFileSync(process.argv[2],'utf8'));
const b = await chromium.launch({ executablePath:'/opt/pw-browsers/chromium', args:['--no-sandbox','--disable-dev-shm-usage'] });
const ctx = await b.newContext({ userAgent:'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36', locale:'en-US', viewport:{width:1400,height:1000} });
const fh = fs.createWriteStream(process.argv[3],{flags:'a'});
for (const [name, url] of sites){
  const p = await ctx.newPage();
  let rec = {name, url, social:{}, hreflang:[], langHints:[], err:null};
  try{
    await p.goto(url,{waitUntil:'domcontentloaded',timeout:45000});
    await p.waitForTimeout(5000);
    rec = await p.evaluate((rec)=>{
      const hrefs=[...document.querySelectorAll('a[href]')].map(a=>a.href);
      const pat={facebook:/facebook\.com\//i, linkedin:/linkedin\.com\//i, x:/(twitter\.com|x\.com)\//i,
                 youtube:/youtube\.com\//i, instagram:/instagram\.com\//i, tiktok:/tiktok\.com\//i};
      for(const k in pat){ const h=hrefs.find(u=>pat[k].test(u)); if(h) rec.social[k]=h.split('?')[0]; }
      rec.hreflang=[...document.querySelectorAll('link[hreflang]')].map(l=>l.getAttribute('hreflang'));
      const txt=document.body.innerText||'';
      const langs=['繁體中文','简体中文','日本語','DEUTSCH','Deutsch','Español','Français','한국어','Português','Italiano','Русский','العربية'];
      rec.langHints=langs.filter(L=>txt.includes(L));
      rec.htmlLang=document.documentElement.lang||'';
      return rec;
    }, rec);
  }catch(e){ rec.err=e.message.split('\n')[0].slice(0,70); }
  fh.write(JSON.stringify(rec)+'\n');
  const s=Object.keys(rec.social).join(',')||'—';
  console.log(`${name.slice(0,26).padEnd(27)} social=[${s}] langs=[${rec.langHints.join(' ')||'-'}] hreflang=${rec.hreflang.length} ${rec.err||''}`);
  await p.close();
}
fh.end(); await b.close();
