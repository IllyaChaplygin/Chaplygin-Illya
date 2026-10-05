#!/usr/bin/env python3
"""Генерация HTML-дашборда из собранных данных."""
import json, os, io

HERE = os.path.dirname(os.path.abspath(__file__))
D = json.load(open(os.path.join(HERE, 'data', 'dashboard_data.json'), encoding='utf-8'))

HTML = """<title>Земельный капитал США</title>
<style>
/* Макет: один столбец разделов; внутри — сетка карточек, схлопывается в одну колонку на телефоне */
:root{
  --bg:#f7f6f3; --surface:#fffefc; --surface-2:#f1efe9; --line:#e0ddd4;
  --ink:#16150f; --ink-2:#56544a; --ink-3:#8a8578;
  --accent:#8a5a2b; --accent-soft:#f0e4d6;
  --neg:#b4472f; --pos:#2f6b4f;
  --s1:#2a78d6; --s2:#eb6834; --s3:#1baf7a; --s4:#eda100;
  --grid:#e6e3da;
  --display:"Fraunces",Georgia,"Times New Roman",serif;
  --body:"Inter Tight",-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;
  --mono:"IBM Plex Mono",ui-monospace,SFMono-Regular,Menlo,monospace;
}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){
  --bg:#13120f; --surface:#1c1b17; --surface-2:#242320; --line:#34322c;
  --ink:#f5f3ec; --ink-2:#b5b2a6; --ink-3:#817d71;
  --accent:#d19a62; --accent-soft:#2e2619;
  --neg:#e0775c; --pos:#5fab86;
  --s1:#3987e5; --s2:#d95926; --s3:#199e70; --s4:#c98500;
  --grid:#2c2a25; color-scheme:dark;
}}
:root[data-theme="dark"]{
  --bg:#13120f; --surface:#1c1b17; --surface-2:#242320; --line:#34322c;
  --ink:#f5f3ec; --ink-2:#b5b2a6; --ink-3:#817d71;
  --accent:#d19a62; --accent-soft:#2e2619;
  --neg:#e0775c; --pos:#5fab86;
  --s1:#3987e5; --s2:#d95926; --s3:#199e70; --s4:#c98500;
  --grid:#2c2a25; color-scheme:dark;
}
*{box-sizing:border-box}
body{background:var(--bg);color:var(--ink);font-family:var(--body);margin:0;
  font-size:15px;line-height:1.55;-webkit-font-smoothing:antialiased}
.wrap{max-width:1180px;margin:0 auto;padding-inline:20px;padding-block:40px 72px}
h1,h2,h3{font-family:var(--display);font-weight:600;text-wrap:balance;margin:0}
h1{font-size:clamp(30px,5.2vw,50px);line-height:1.05;letter-spacing:-.02em}
h2{font-size:clamp(20px,2.9vw,27px);letter-spacing:-.01em;margin-bottom:6px}
h3{font-size:16px;font-family:var(--body);font-weight:650}
.eyebrow{font-family:var(--mono);font-size:11px;letter-spacing:.14em;text-transform:uppercase;
  color:var(--accent);margin-bottom:12px}
.lede{font-size:17px;color:var(--ink-2);max-width:68ch;margin-top:16px}
section{margin-top:52px}
.sub{color:var(--ink-2);font-size:14px;max-width:76ch;margin:0 0 20px}
.card{background:var(--surface);border:1px solid var(--line);border-radius:10px;padding:20px}
.grid{display:grid;gap:14px}
.tiles{grid-template-columns:repeat(auto-fit,minmax(215px,1fr));margin-top:28px}
.tile .n{font-family:var(--display);font-size:38px;line-height:1;letter-spacing:-.02em;
  font-variant-numeric:tabular-nums;display:block}
.tile .k{font-family:var(--mono);font-size:10.5px;letter-spacing:.11em;text-transform:uppercase;
  color:var(--ink-3);margin-bottom:9px}
.tile .d{font-size:12.5px;color:var(--ink-2);margin-top:9px;line-height:1.45}
.tile.flag .n{color:var(--neg)}
.verdict{border-left:3px solid var(--accent);background:var(--accent-soft);
  padding:18px 22px;border-radius:0 8px 8px 0;margin-top:26px}
.verdict p{margin:0;font-size:15.5px}
.verdict p+p{margin-top:10px}
.tablewrap{overflow-x:auto;border:1px solid var(--line);border-radius:10px;background:var(--surface)}
table{border-collapse:collapse;width:100%;font-size:13px;min-width:860px}
th,td{padding:9px 11px;text-align:left;border-bottom:1px solid var(--line);vertical-align:top}
thead th{background:var(--surface-2);font-family:var(--mono);font-size:10px;letter-spacing:.07em;
  text-transform:uppercase;color:var(--ink-2);position:sticky;top:0;cursor:pointer;
  white-space:nowrap;font-weight:500}
thead th:hover{color:var(--accent)}
tbody tr:hover{background:var(--surface-2)}
td.num{text-align:right;font-variant-numeric:tabular-nums;font-family:var(--mono);font-size:12px;
  white-space:nowrap}
td.nm{font-weight:600;min-width:188px}
.rk{font-family:var(--display);font-size:16px;color:var(--ink-3);width:30px;text-align:right}
.pill{display:inline-block;padding:1px 7px;border-radius:99px;font-size:10.5px;
  font-family:var(--mono);letter-spacing:.03em;white-space:nowrap}
.p-zero{background:color-mix(in srgb,var(--neg) 16%,transparent);color:var(--neg)}
.p-some{background:color-mix(in srgb,var(--pos) 18%,transparent);color:var(--pos)}
.p-t1{background:var(--surface-2);color:var(--ink-2);border:1px solid var(--line)}
.bar-s{display:flex;height:9px;border-radius:2px;overflow:hidden;gap:2px;min-width:118px}
.bar-s i{display:block}
.legend{display:flex;flex-wrap:wrap;gap:14px;margin:0 0 16px;padding:0;list-style:none;
  font-size:12.5px;color:var(--ink-2)}
.legend li{display:flex;align-items:center;gap:6px}
.sw{width:11px;height:11px;border-radius:2px;flex:none}
figure{margin:0}
figcaption{font-size:12.5px;color:var(--ink-3);margin-top:10px;line-height:1.5}
.chartbox{width:100%;overflow-x:auto}
svg{display:block;max-width:100%}
.hbar{display:grid;grid-template-columns:minmax(135px,auto) 1fr auto;gap:10px;align-items:center;
  margin-bottom:9px;font-size:13px}
.hbar .t{color:var(--ink-2)}
.hbar .track{background:var(--surface-2);border-radius:3px;height:20px;position:relative;min-width:0}
.hbar .fill{background:var(--accent);height:100%;border-radius:3px;min-width:2px}
.hbar .v{font-family:var(--mono);font-size:12px;font-variant-numeric:tabular-nums;
  color:var(--ink);white-space:nowrap}
.two{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.note{font-size:12.5px;color:var(--ink-3);line-height:1.55}
.status{font-family:var(--mono);font-size:10px;padding:1px 5px;border-radius:3px;
  background:var(--surface-2);color:var(--ink-2);border:1px solid var(--line)}
ul.clean{margin:10px 0 0;padding-left:19px;color:var(--ink-2);font-size:14px}
ul.clean li{margin-bottom:7px}
.foot{margin-top:60px;padding-top:22px;border-top:1px solid var(--line);
  font-size:12.5px;color:var(--ink-3)}
@media (max-width:760px){ .two{grid-template-columns:1fr} }
</style>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600&family=Inter+Tight:wght@400;500;600;650&family=IBM+Plex+Mono:wght@400;500&display=swap">

<div class="wrap">

<header>
  <div class="eyebrow">Срез рынка США · 5 октября 2026</div>
  <h1>Кто владеет землёй,<br>тот не покупает клики</h1>
  <p class="lede">20 американских компаний, которые собирают инвесторов в совместные проекты
  на земле — отранжированы по измеримым показателям. Главный результат исследования опровергает
  исходную гипотезу: масштаб земельного банка и рекламная активность в США связаны
  <strong>обратно</strong>.</p>

  <div class="grid tiles">
    <div class="card tile flag">
      <div class="k">Объявлений Google</div>
      <span class="n">0</span>
      <div class="d">У всех 13 крупнейших землевладельцев выборки. Walton, Millrose, Domain,
      Crow, Five Point, St. Joe, Tejon, BTI — ни одного активного объявления.</div>
    </div>
    <div class="card tile">
      <div class="k">Медианный чек Crow Holdings</div>
      <span class="n">$5 млн</span>
      <div class="d">Ровно искомый профиль: земля, совместное участие, крупные чеки.
      И 54 из 54 предложений по 506(b) — рекламировать запрещено законом.</div>
    </div>
    <div class="card tile flag">
      <div class="k">Доля 506(c) у Walton</div>
      <span class="n">0%</span>
      <div class="d">52 из 52 предложений — 506(b). Нулевая реклама у лидера рынка
      не выбор маркетолога, а режим регистрации выпуска.</div>
    </div>
    <div class="card tile">
      <div class="k">В платные каналы</div>
      <span class="n">15%</span>
      <div class="d">Оценка доли совокупного бюджета привлечения, идущей в платный поиск
      и соцсети. Остальное — люди, события и контент.</div>
    </div>
  </div>

  <div class="verdict">
    <p><strong>Что это значит практически.</strong> Копировать «как они льют трафик» у Walton
    или Domain нечего — они его не льют. Это две разные машины, и копировать нужно обе по
    отдельности: машину земельного сорсинга у первой группы и машину привлечения
    розничного капитала у второй.</p>
    <p><strong>Первое ограничение — юридическое, а не бюджетное.</strong> Rule 506(b) запрещает
    публичную рекламу предложения. Пока структура выпуска не 506(c) или Reg A+, никакой
    медиаплан не имеет силы. Это первая развилка, а не последняя.</p>
  </div>
</header>

<section>
  <h2>Рейтинг: 20 компаний</h2>
  <p class="sub">100 баллов, четыре блока. Нажмите на заголовок столбца для сортировки.
  Столбец «данных %» показывает долю заполненных ключевых полей: низкий балл при низкой
  полноте означает непрозрачность компании, а не её малый масштаб.</p>

  <ul class="legend">
    <li><span class="sw" style="background:var(--s1)"></span>A — масштаб земли (30)</li>
    <li><span class="sw" style="background:var(--s2)"></span>B — капитальная машина (30)</li>
    <li><span class="sw" style="background:var(--s3)"></span>C — маркетинг (30)</li>
    <li><span class="sw" style="background:var(--s4)"></span>D — конверсия (10)</li>
  </ul>

  <div class="tablewrap">
    <table id="t">
      <thead><tr>
        <th data-k="rank" data-t="n">#</th>
        <th data-k="name">Компания</th>
        <th data-k="hq">Город</th>
        <th data-k="TOTAL" data-t="n">Балл</th>
        <th>Структура</th>
        <th data-k="acres" data-t="n">Акры /<br>homesites</th>
        <th data-k="regd_entities" data-t="n">Reg D<br>структур</th>
        <th data-k="pct_506c" data-t="n">% 506(c)</th>
        <th data-k="google_ads" data-t="n">Google<br>объявл.</th>
        <th data-k="min_check_usd" data-t="n">Мин.<br>чек $</th>
        <th data-k="investors" data-t="n">Инвест.</th>
        <th data-k="DATA_COMPLETENESS_PCT" data-t="n">данных<br>%</th>
        <th>Как ищут инвесторов</th>
      </tr></thead>
      <tbody></tbody>
    </table>
  </div>
  <figcaption>Акры, Reg D-структуры, % 506(c), объявления Google, минимальный чек и число
  инвесторов — прямые измерения (Google Ads Transparency Center и SEC EDGAR Form D, 05.10.2026).
  Балл — расчёт по модели. Прочерк означает отсутствие данных, а не нулевое значение.</figcaption>
</section>

<section>
  <h2>Право рекламировать против фактической рекламы</h2>
  <p class="sub">Каждая точка — компания. По горизонтали — доля предложений по Rule 506(c),
  то есть юридическое право на публичную рекламу. По вертикали — сколько объявлений реально
  работает. Левый нижний угол — те, кому нельзя и кто не рекламируется.</p>
  <div class="card">
    <div class="chartbox"><div id="scatter"></div></div>
    <figcaption>Источники: SEC EDGAR Form D (поле federalExemptionsExclusions) и Google Ads
    Transparency Center. Обе оси — измеренные значения.</figcaption>
  </div>
</section>

<section>
  <h2>Куда на самом деле уходят деньги</h2>
  <p class="sub">Совокупный годовой бюджет привлечения капитала по 20 компаниям —
  <strong>около $82 млн</strong>. Это <em>модельная оценка</em>: бюджет привязан к объёму
  привлечения и ставке затрат на привлечение, поскольку сами расходы нигде не раскрываются.
  Структура при этом устойчива — почти 40% уходит людям, а не в медиа.</p>
  <div class="two">
    <div class="card">
      <h3 style="margin-bottom:16px">Распределение бюджета по каналам</h3>
      <div id="channels"></div>
      <figcaption>ОЦЕНКА. Точность — порядок величины. Доли выведены из наблюдаемых
      сигналов: наличия рекламы, объёма контента, событий, структуры дистрибуции.</figcaption>
    </div>
    <div class="card">
      <h3 style="margin-bottom:12px">Почему так, а не иначе</h3>
      <ul class="clean">
        <li><strong>Юридически.</strong> 506(b) запрещает general solicitation. У Walton и
        Crow Holdings — 100% предложений именно такие.</li>
        <li><strong>Экономически.</strong> Лид аккредитованного инвестора стоит $50–500,
        а один профинансировавший инвестор — $3 500–4 500. Под чек $5–20 млн дешевле
        содержать команду capital markets и платить дистрибьютору.</li>
        <li><strong>Организационно.</strong> Капитал берут там, где он уже лежит:
        институциональные LP, страховые балансы, сети wealth-менеджеров. У Walton —
        300+ одобренных каналов дистрибуции и инвесторы из 91 страны.</li>
      </ul>
      <p class="note" style="margin-top:16px"><strong>Аномалия, которую стоит изучить
      отдельно:</strong> у AcreTrader 101 из 101 предложения — 506(c), то есть полное право
      рекламировать, и при этом всего 3 объявления. Контент-воронка вместо платного трафика
      принесла им 7 035 инвесторов.</p>
    </div>
  </div>
</section>

<section>
  <h2>У кого копировать рекламную механику</h2>
  <p class="sub">Это не земельные игроки — они не прошли фильтр по активу. Но именно они
  умеют привлекать капитал платным трафиком, и их воронку имеет смысл разбирать.</p>
  <div class="tablewrap">
    <table id="tb">
      <thead><tr>
        <th>Компания</th><th>Домен</th><th>Google объявл.</th><th>YouTube подп.</th>
        <th>Reg D структур</th><th>Чему учиться</th>
      </tr></thead>
      <tbody></tbody>
    </table>
  </div>
  <figcaption>Внимание на Brookfield Residential: 700 объявлений — это продажа домов конечным
  покупателям, а не привлечение инвесторов. Наивное чтение рекламных счётчиков даёт здесь
  прямо противоположный вывод.</figcaption>
</section>

<section>
  <h2>Что считать фактом, а что оценкой</h2>
  <div class="two">
    <div class="card">
      <h3>Факт — измерено</h3>
      <ul class="clean">
        <li>Количество активных объявлений Google — Ads Transparency Center, прямое снятие</li>
        <li>Число Reg D-структур, форм D, доля 506(c), суммы продаж, медианный минимальный
        чек, число инвесторов — SEC EDGAR, разбор primary_doc.xml</li>
        <li>Подписчики и видео на YouTube</li>
        <li>Акры, AUM, число проектов — раскрыто компаниями и в отчётности</li>
      </ul>
    </div>
    <div class="card">
      <h3>Оценка — модель</h3>
      <ul class="clean">
        <li>Все долларовые бюджеты и распределение по каналам</li>
        <li><strong>Почему:</strong> рекламные расходы частных компаний в США не раскрываются
        нигде. Ни Meta Ad Library, ни Google ATC не показывают суммы для неполитической
        рекламы — только наличие, количество и креативы.</li>
        <li><strong>Точность:</strong> порядок величины, ±50–70%. Годится как ориентир
        структуры, но не как бюджетный план.</li>
        <li><strong>Пробел:</strong> Meta закрыла поиск по ключевым словам без авторизации,
        поэтому данные по Meta неполные.</li>
      </ul>
    </div>
  </div>
</section>

<div class="foot">
  Выборка: 64 проверенные компании, 20 прошли все пять фильтров включения.
  228 записей в журнале доказательств, из них 173 — прямые измерения.
  Источники: Google Ads Transparency Center, SEC EDGAR Form D, отчётность компаний,
  отраслевые пресс-релизы, YouTube. Данные на 5 октября 2026.
</div>
</div>

<script id="data" type="application/json">__DATA__</script>
<script>
const D = JSON.parse(document.getElementById('data').textContent);
const num = v => (v===''||v==null) ? null : Number(v);
const fmt = v => v==null ? '<span style="color:var(--ink-3)">—</span>'
  : v.toLocaleString('ru-RU');

/* ---------- таблица рейтинга ---------- */
const tb = document.querySelector('#t tbody');
function rowHTML(c){
  const t=+c.TOTAL, A=+c.BLOCK_A, B=+c.BLOCK_B, C=+c.BLOCK_C, Dd=+c.BLOCK_D;
  const seg=(v,col)=> v>0?`<i style="width:${v/100*100}%;background:${col}"></i>`:'';
  const ads=num(c.google_ads)||0;
  const pct=num(c.pct_506c);
  return `<tr>
   <td class="rk">${c.rank}</td>
   <td class="nm">${c.name}<div style="font-weight:400;color:var(--ink-3);font-size:11.5px">${c.domain}</div></td>
   <td style="color:var(--ink-2);font-size:12px;white-space:nowrap">${c.hq}</td>
   <td class="num" style="font-weight:700;font-size:14px">${t}</td>
   <td><div class="bar-s" title="A ${A} / B ${B} / C ${C} / D ${Dd}">
     ${seg(A,'var(--s1)')}${seg(B,'var(--s2)')}${seg(C,'var(--s3)')}${seg(Dd,'var(--s4)')}</div></td>
   <td class="num">${fmt(num(c.acres))}</td>
   <td class="num">${fmt(num(c.regd_entities))}</td>
   <td class="num">${pct==null?'<span style="color:var(--ink-3)">—</span>':pct.toFixed(1)}</td>
   <td class="num"><span class="pill ${ads===0?'p-zero':'p-some'}">${ads}</span></td>
   <td class="num">${fmt(num(c.min_check_usd))}</td>
   <td class="num">${fmt(num(c.investors))}</td>
   <td class="num" style="color:${+c.DATA_COMPLETENESS_PCT<50?'var(--neg)':'var(--ink-2)'}">${c.DATA_COMPLETENESS_PCT}</td>
   <td style="font-size:12px;color:var(--ink-2);min-width:230px">${c.channel_model}</td>
  </tr>`;
}
let rows=[...D.companies];
function draw(){ tb.innerHTML = rows.map(rowHTML).join(''); }
draw();
let lastKey=null, asc=true;
document.querySelectorAll('#t thead th[data-k]').forEach(th=>{
  th.addEventListener('click',()=>{
    const k=th.dataset.k, isNum=th.dataset.t==='n';
    asc = (lastKey===k) ? !asc : true; lastKey=k;
    rows.sort((a,b)=>{
      let x=isNum?num(a[k]):a[k], y=isNum?num(b[k]):b[k];
      if(x==null&&y==null) return 0; if(x==null) return 1; if(y==null) return -1;
      return (x>y?1:x<y?-1:0)*(asc?1:-1);
    });
    draw();
  });
});

/* ---------- бенчмарк ---------- */
document.querySelector('#tb tbody').innerHTML = D.benchmark.map(b=>`<tr>
  <td class="nm">${b.name}</td>
  <td style="font-size:12px;color:var(--ink-3)">${b.domain}</td>
  <td class="num"><span class="pill ${+b.google_ads===0?'p-zero':'p-some'}">${b.google_ads}</span></td>
  <td class="num">${fmt(num(b.yt_subs))}</td>
  <td class="num">${fmt(num(b.regd_entities))}</td>
  <td style="font-size:12px;color:var(--ink-2);min-width:300px">${b.note}</td></tr>`).join('');

/* ---------- каналы: горизонтальные столбцы ---------- */
(function(){
  const tot=D.channels.reduce((s,c)=>s+c.usd_m,0), mx=Math.max(...D.channels.map(c=>c.usd_m));
  document.getElementById('channels').innerHTML = D.channels.map(c=>`
    <div class="hbar">
      <span class="t">${c.channel}</span>
      <span class="track"><span class="fill" style="width:${c.usd_m/mx*100}%"></span></span>
      <span class="v">$${c.usd_m.toFixed(1)} млн · ${(c.usd_m/tot*100).toFixed(1)}%</span>
    </div>`).join('');
})();

/* ---------- диаграмма рассеяния: 506(c) против рекламы ---------- */
(function(){
  const pts = D.companies.filter(c=>c.pct_506c!=='' && c.pct_506c!=null)
    .map(c=>({x:+c.pct_506c, y:+(c.google_ads||0), n:c.name}));
  const W=780,H=420,M={t:22,r:96,b:56,l:62};
  const iw=W-M.l-M.r, ih=H-M.t-M.b;
  const maxY=Math.ceil(Math.max(25,...pts.map(p=>p.y))/5)*5;
  const sx=v=>M.l+v/100*iw, sy=v=>M.t+ih-(v/maxY)*ih;
  let g='';
  for(let i=0;i<=5;i++){const v=i*20,X=sx(v);
    g+=`<line x1="${X}" y1="${M.t}" x2="${X}" y2="${M.t+ih}" stroke="var(--grid)" stroke-width="1"/>
        <text x="${X}" y="${M.t+ih+21}" text-anchor="middle" font-size="11" fill="var(--ink-3)" font-family="var(--mono)">${v}%</text>`;}
  const steps=maxY/5;
  for(let v=0;v<=maxY;v+=steps){const Y=sy(v);
    g+=`<line x1="${M.l}" y1="${Y}" x2="${M.l+iw}" y2="${Y}" stroke="var(--grid)" stroke-width="1"/>
        <text x="${M.l-10}" y="${Y+4}" text-anchor="end" font-size="11" fill="var(--ink-3)" font-family="var(--mono)">${v}</text>`;}
  // Walton Global и Crow Holdings лежат в одной точке (0%, 0) — подписываем кластер один раз
  const CLUSTER=['Walton Global','Crow Holdings'];
  const label={'AcreTrader':[-14,-14],'FarmTogether':[14,5],'DLP Capital':[-14,-12],
    'Caliber (NASDAQ: CWD)':[-14,-12],'MLG Capital':[-14,-14],'GSP REI':[16,-14],
    'Urban Catalyst':[14,-10]};
  const marks=pts.map(p=>{
    const X=sx(p.x),Y=sy(p.y), zero=p.y===0;
    const off=label[p.n];
    const short=p.n.replace(/ \\(.*\\)/,'');
    const lab= off? `<text x="${X+off[0]}" y="${Y+off[1]}" font-size="11.5" fill="var(--ink-2)"
      text-anchor="${off[0]<0?'end':'start'}" font-family="var(--body)">${short}</text>`:'';
    return `<circle cx="${X}" cy="${Y}" r="6.5" fill="${zero?'var(--neg)':'var(--s1)'}"
      fill-opacity=".85" stroke="var(--surface)" stroke-width="2"><title>${p.n}: ${p.x}% 506(c), ${p.y} объявлений</title></circle>${lab}`;
  }).join('');
  const cx0=sx(0), cy0=sy(0);
  const clusterNote=`<g>
    <line x1="${cx0+9}" y1="${cy0-6}" x2="${cx0+54}" y2="${cy0-40}" stroke="var(--ink-3)" stroke-width="1"/>
    <text x="${cx0+58}" y="${cy0-44}" font-size="11.5" fill="var(--neg)" font-family="var(--body)" font-weight="600">Walton Global · Crow Holdings</text>
    <text x="${cx0+58}" y="${cy0-30}" font-size="10.5" fill="var(--ink-3)" font-family="var(--mono)">0% 506(c) · 0 объявлений</text>
  </g>`;
  document.getElementById('scatter').innerHTML=`
  <svg viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" role="img"
    aria-label="Диаграмма рассеяния: доля 506(c) против числа объявлений Google">
    ${g}
    <line x1="${M.l}" y1="${M.t+ih}" x2="${M.l+iw}" y2="${M.t+ih}" stroke="var(--ink-3)" stroke-width="1"/>
    <line x1="${M.l}" y1="${M.t}" x2="${M.l}" y2="${M.t+ih}" stroke="var(--ink-3)" stroke-width="1"/>
    ${marks}
    ${clusterNote}
    <text x="${M.l+iw/2}" y="${H-10}" text-anchor="middle" font-size="12" fill="var(--ink-2)">доля предложений 506(c) — право рекламировать</text>
    <text x="16" y="${M.t+ih/2}" transform="rotate(-90 16 ${M.t+ih/2})" text-anchor="middle" font-size="12" fill="var(--ink-2)">активных объявлений Google</text>
  </svg>`;
})();
</script>
"""

out = HTML.replace('__DATA__', json.dumps(D, ensure_ascii=False))
path = os.path.join(HERE, 'dashboard.html')
io.open(path, 'w', encoding='utf-8').write(out)
print('written', path, len(out), 'bytes')
