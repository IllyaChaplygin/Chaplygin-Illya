import { chromium } from '/opt/node22/lib/node_modules/playwright/index.mjs';
import fs from 'fs';
const targets = JSON.parse(fs.readFileSync(process.argv[2],'utf8'));
const b = await chromium.launch({ executablePath:'/opt/pw-browsers/chromium', args:['--no-sandbox','--disable-dev-shm-usage'] });
const ctx = await b.newContext({ userAgent:'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36', locale:'en-US', viewport:{width:1400,height:1100} });
const fh = fs.createWriteStream(process.argv[3],{flags:'a'});
for (const [name,url] of targets){
  const p = await ctx.newPage();
  let rec={name,url};
  try{
    await p.goto(url,{waitUntil:'domcontentloaded',timeout:45000});
    await p.waitForTimeout(6000);
    rec = await p.evaluate((rec)=>{
      const txt = n => (n&&n.innerText||'').trim().replace(/\s+/g,' ');
      rec.h1 = txt(document.querySelector('h1')).slice(0,220);
      rec.h2 = [...document.querySelectorAll('h2')].slice(0,6).map(txt).filter(Boolean).map(s=>s.slice(0,120));
      // CTA кнопки и ссылки
      const ctas=[...document.querySelectorAll('a,button')].map(txt)
        .filter(t=>t && t.length<60 && /invest|get started|book|schedule|call|webinar|download|register|sign up|join|learn more|request|apply|access|guide|brochure|portfolio|offering/i.test(t));
      rec.ctas=[...new Set(ctas)].slice(0,14);
      // поля форм = что просят у лида
      rec.fields=[...new Set([...document.querySelectorAll('input,select,textarea')]
        .map(i=>(i.name||i.placeholder||i.type||'').toLowerCase()).filter(Boolean)
        .filter(f=>!/hidden|csrf|token|utm|^submit$/.test(f)))].slice(0,18);
      // трекеры = как мерят трафик
      const html=document.documentElement.innerHTML;
      const tr=[];
      if(/gtag\(|googletagmanager/i.test(html)) tr.push('Google Analytics/GTM');
      if(/fbq\(|connect\.facebook\.net/i.test(html)) tr.push('Meta Pixel');
      if(/snaptr\(/i.test(html)) tr.push('Snap Pixel');
      if(/ttq\.|analytics\.tiktok/i.test(html)) tr.push('TikTok Pixel');
      if(/licdn\.com|_linkedin_partner_id/i.test(html)) tr.push('LinkedIn Insight');
      if(/hubspot|hs-scripts/i.test(html)) tr.push('HubSpot');
      if(/marketo|munchkin/i.test(html)) tr.push('Marketo');
      if(/pardot/i.test(html)) tr.push('Pardot');
      if(/calendly/i.test(html)) tr.push('Calendly');
      if(/hotjar/i.test(html)) tr.push('Hotjar');
      if(/clarity\.ms/i.test(html)) tr.push('MS Clarity');
      if(/doubleclick|googleadservices/i.test(html)) tr.push('Google Ads remarketing');
      if(/bat\.bing|uetq/i.test(html)) tr.push('Microsoft/Bing UET');
      if(/twq\(|static\.ads-twitter/i.test(html)) tr.push('X Pixel');
      rec.trackers=tr;
      rec.minMention=(document.body.innerText.match(/\$[\d,]+(?:\s*(?:minimum|min\.?))/i)||[])[0]||null;
      return rec;
    }, rec);
  }catch(e){ rec.err=e.message.split('\n')[0].slice(0,70); }
  fh.write(JSON.stringify(rec)+'\n');
  console.log(`\n### ${name}`);
  console.log('  H1:', (rec.h1||'').slice(0,140));
  console.log('  CTA:', (rec.ctas||[]).slice(0,8).join(' | '));
  console.log('  FORM:', (rec.fields||[]).join(', ')||'—');
  console.log('  TRACK:', (rec.trackers||[]).join(', ')||'—', rec.err||'');
  await p.close();
}
fh.end(); await b.close();
