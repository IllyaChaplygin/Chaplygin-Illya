#!/usr/bin/env python3
"""Static pages for search engines and link previews.

The site itself is a single page (hash routes are invisible to Google), so for every lot this writes
docs/lot/<ID>/index.html with a title, description, Open Graph card and the lot facts as plain text;
visitors with JavaScript are sent straight to the real page (#/p/<ID>). Also writes docs/sitemap.xml,
docs/robots.txt and refreshes the <!--seo--> block of docs/index.html.
Run after build.py:  python3 tools/build_pages.py   (base URL is in tools/site.json; change it when the domain is bought)."""
import json, subprocess, html, datetime, re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DOCS = ROOT / 'docs'
BASE = json.loads((ROOT / 'tools' / 'site.json').read_text())['base'].rstrip('/') + '/'
lots = json.loads(subprocess.check_output(['node', '-e', "const w={};new Function('w',require('fs').readFileSync('docs/data.js','utf8').replace('const LOTS','w.LOTS')+';')(w);console.log(JSON.stringify(w.LOTS))"], cwd=ROOT))
REG = {'Budva': 'Будва', 'Kotor': 'Котор', 'Kolasin': 'Колашин'}
PUR = {'joint': 'Партнёрская застройка', 'sale': 'Продажа'}
VIEW = {'море': 'у моря', 'гора': 'в горах'}
e = html.escape
today = datetime.date.today().isoformat()

def og_image(l):
    if l['photos']:
        return f"{BASE}photos/{l['photos'][0]}.jpg"
    if l.get('tourCover'):
        return f"{BASE}photos/{l['tourCover']}"
    return f"{BASE}img/hero.jpg"

def fm(n):
    return f'{n:,.0f}'.replace(',', ' ')

PAGE = '''<!doctype html>
<html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<meta name="description" content="{desc}">
<link rel="canonical" href="{url}">
<meta property="og:type" content="website"><meta property="og:site_name" content="ALLIANZ KAPITAL">
<meta property="og:title" content="{title}"><meta property="og:description" content="{desc}">
<meta property="og:url" content="{url}"><meta property="og:image" content="{img}">
<meta name="twitter:card" content="summary_large_image"><meta name="twitter:image" content="{img}">
<link rel="icon" href="{base}logo.png">
<style>body{{margin:0;font:16px/1.55 system-ui,Arial,sans-serif;color:#0a2733;background:#f6f9f8}}main{{max-width:760px;margin:0 auto;padding:40px 20px}}a{{color:#0b7285}}img{{max-width:100%;border-radius:12px}}</style>
</head><body><main>
<p><a href="{base}">ALLIANZ KAPITAL</a></p>
<h1>{h1}</h1>
<p>{lead}</p>
<ul>{facts}</ul>
<p><img src="{img}" alt="{h1}" width="760"></p>
<p><a href="{app}">Открыть лот на сайте</a></p>
<p>ALLIANZ KAPITAL d.o.o., Poluostrvo Zavala, Zgrada Harmonija, Budva, Crna Gora.</p>
</main>
<script>location.replace("{app_rel}")</script>
</body></html>
'''

urls = [BASE]
for l in lots:
    reg = REG.get(l['region'], l['region'])
    area = f"{fm(l['area'])} м²" if l.get('area') else ''
    title = f"Участок {l['id']}, {l['municipality']}: {area}, {VIEW.get(l['view'], '')}".replace(', : ', ': ').strip(' ,:') + ' | ALLIANZ KAPITAL'
    priced = l.get('total') and l.get('purpose') != 'joint'
    bits = [f"Земельный участок {l['id']} ({area}) в районе {l['municipality']}, {reg}, Черногория", PUR.get(l.get('purpose'), '')]
    if priced:
        bits.append(f"Цена € {fm(l['total'])}")
    desc = '. '.join(b for b in bits if b) + '. Фото, видео и карта на сайте ALLIANZ KAPITAL.'
    facts = [('Лот', l['id']), ('Регион', reg), ('Населённый пункт', l['municipality']), ('Расположение', VIEW.get(l['view'], '').capitalize()),
             ('Площадь', area), ('Кадастровых участков', str(l['n'])), ('Назначение', PUR.get(l.get('purpose'), ''))]
    if priced:
        facts.append(('Цена', f"€ {fm(l['total'])}"))
    d = DOCS / 'lot' / l['id']
    d.mkdir(parents=True, exist_ok=True)
    url = f"{BASE}lot/{l['id']}/"
    (d / 'index.html').write_text(PAGE.format(
        title=e(title), desc=e(desc), url=url, img=og_image(l), base=BASE, h1=e(f"Участок {l['id']}, {l['municipality']}"),
        lead=e(desc), facts=''.join(f'<li><b>{e(k)}:</b> {e(v)}</li>' for k, v in facts if v),
        app=f"{BASE}#/p/{l['id']}", app_rel=f"../../#/p/{l['id']}"), encoding='utf-8')
    urls.append(url)

(DOCS / 'sitemap.xml').write_text('<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n' +
    ''.join(f'  <url><loc>{u}</loc><lastmod>{today}</lastmod></url>\n' for u in urls) + '</urlset>\n', encoding='utf-8')
(DOCS / 'robots.txt').write_text(f'User-agent: *\nAllow: /\n\nSitemap: {BASE}sitemap.xml\n', encoding='utf-8')

seo = f'''<!--seo-->
<link rel="canonical" href="{BASE}">
<meta property="og:type" content="website"><meta property="og:site_name" content="ALLIANZ KAPITAL">
<meta property="og:title" content="ALLIANZ KAPITAL | Земля в Будве, Которе и Колашине">
<meta property="og:description" content="Земельный банк Allianz Kapital: участки у моря и в горах для покупки и партнёрской застройки.">
<meta property="og:url" content="{BASE}"><meta property="og:image" content="{BASE}img/hero.jpg">
<meta name="twitter:card" content="summary_large_image"><meta name="twitter:image" content="{BASE}img/hero.jpg">
<!--/seo-->'''
idx = DOCS / 'index.html'
t = idx.read_text(encoding='utf-8')
if '<!--seo-->' in t:
    t = re.sub(r'<!--seo-->.*?<!--/seo-->', lambda m: seo, t, flags=re.S)
else:
    t = t.replace('<link rel="icon" href="logo.png">', '<link rel="icon" href="logo.png">\n' + seo, 1)
idx.write_text(t, encoding='utf-8')
print(len(lots), 'lot pages;', len(urls), 'urls in sitemap')
