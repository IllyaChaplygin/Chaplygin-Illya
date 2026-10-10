'use strict';
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const ls = {
  get(k, d) { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} }
};
const guess = () => { const n = (navigator.language || 'en').toLowerCase(); return n.startsWith('ru') ? 'ru' : /^(sr|hr|bs|cnr|me)/.test(n) ? 'me' : 'en'; };
let lang = ls.get('lang', null); if (!I18N[lang]) lang = guess();
const D = () => I18N[lang];
const t = (k, ...a) => { const v = D()[k]; return typeof v === 'function' ? v(...a) : (v ?? k); };
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const safeUrl = u => (/^https:\/\//.test(u || '') ? esc(u) : '#');
const fmt = (v, d = 0) => v == null ? '-' : v.toLocaleString(D().loc, { maximumFractionDigits: d });
const eur = v => '€ ' + fmt(v);
const fav = new Set(ls.get('fav', []));
const CO = { name: 'ALLIANZ KAPITAL d.o.o.', addr: 'Poluostrvo Zavala, Zgrada Harmonija, Budva, Crna Gora', reg: new Date(2007, 2, 13), registry: 'https://www.companywall.me/firma/allianz-kapital/MMxF2yCD' };
const regDate = () => CO.reg.toLocaleDateString(D().loc, { day: 'numeric', month: 'long', year: 'numeric' });
const REGION_XY = { Kolasin: [42.8228, 19.5206], Budva: [42.2864, 18.84], Kotor: [42.4247, 18.7712] };
const REGION_EN = { Kolasin: 'Kolasin', Budva: 'Budva', Kotor: 'Kotor' };
const ESRI = 'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}';
const DEF = () => ({ view: '', region: '', muni: '', pmin: '', pmax: '', amin: '', amax: '', photo: false, video: false, tour: false, saved: false, sort: 'rel', mode: 'grid' });
let S = DEF();
let homeMap = null, lotMap = null;

/* ---------- helpers ---------- */
const lotById = id => LOTS.find(l => l.id === id);
const photoSm = i => `photos/${i}_s.jpg`;
const photoBig = i => `photos/${i}.jpg`;
const has3d = l => l.tours.length > 0;
const hasMedia = l => l.photos.length || l.tourCover;
const regionName = r => D().reg[r] || r;
const viewName = v => D().view[v] || v;
const unit = l => l.view === 'гора' ? 'mountains' : 'waves';
const contactSet = () => { const c = (typeof CONTACT !== 'undefined' && CONTACT) || {}; return (c.email || c.whatsapp || c.phone || c.telegram) ? c : null; };

function applyFilters() {
  const n = x => x === '' ? null : Number(x);
  const pmin = n(S.pmin), pmax = n(S.pmax), amin = n(S.amin), amax = n(S.amax);
  const list = LOTS.filter(l =>
    (!S.view || l.view === S.view) && (!S.region || l.region === S.region) && (!S.muni || l.municipality === S.muni) &&
    (pmin == null || (l.total != null && l.total >= pmin)) && (pmax == null || (l.total != null && l.total <= pmax)) &&
    (amin == null || (l.area != null && l.area >= amin)) && (amax == null || (l.area != null && l.area <= amax)) &&
    (!S.photo || l.photos.length) && (!S.video || l.videos.length) && (!S.tour || has3d(l)) && (!S.saved || fav.has(l.id)));
  const rank = l => (l.photos.length ? 0 : l.tourCover ? 1 : 2);
  const key = { pa: l => l.total ?? Infinity, pd: l => -(l.total ?? -Infinity), ad: l => -(l.area ?? 0), aa: l => l.area ?? Infinity, pm: l => l.ppm ?? Infinity,
    rel: l => rank(l) }[S.sort];
  return list.sort((a, b) => key(a) - key(b) || (a.id < b.id ? -1 : 1));
}

/* ---------- cards ---------- */
function card(l) {
  let img;
  if (l.photos.length) img = `<img src="${photoSm(l.photos[0])}" alt="" loading="lazy" decoding="async">`;
  else if (l.tourCover) img = `<img src="photos/${esc(l.tourCover)}" alt="" loading="lazy" decoding="async">`;
  else img = `<div class="noimg"><i class="ph ph-${unit(l)}"></i><b>${l.area != null ? fmt(l.area) + ' ' + esc(t('m2')) : ''}</b><span>${esc(t('noPhoto'))}</span></div>`;
  const badge = (l.photos.length || l.videos.length || has3d(l)) ? `<span class="badge">${l.photos.length ? `<span><i class="ph ph-camera"></i> ${l.photos.length}</span>` : ''}${l.videos.length ? '<i class="ph ph-video-camera"></i>' : ''}${has3d(l) ? '<i class="ph ph-cube"></i>' : ''}</span>` : '';
  const on = fav.has(l.id);
  const price = l.total != null ? `${eur(l.total)}${l.ppm ? `<small>${eur(l.ppm)} ${esc(t('perM2'))}</small>` : ''}` : esc(t('onReq'));
  return `<article class="card" data-id="${esc(l.id)}">
    <a class="im" href="#/p/${esc(l.id)}" aria-label="${esc(t('lot'))} ${esc(l.id)}">${img}<span class="code">${esc(l.id)}</span>${badge}</a>
    <button class="fav ${on ? 'on' : ''}" data-fav="${esc(l.id)}" aria-label="${esc(t(on ? 'unsave' : 'save'))}" aria-pressed="${on}"><i class="ph${on ? '-fill' : ''} ph-heart"></i></button>
    <a class="bd" href="#/p/${esc(l.id)}">
      <div class="price">${price}</div>
      <div class="loc">${l.area != null ? `${fmt(l.area)} ${esc(t('m2'))}, ` : ''}${esc(viewName(l.view))}</div>
      <div class="meta">${esc(l.municipality)}, ${esc(regionName(l.region))}</div>
      <div class="tags">${l.n > 1 ? `<span class="tag">${esc(t('parcelsN', l.n))}</span>` : ''}${l.partial ? `<span class="tag">${esc(t('partialOf', l.partial, l.n))}</span>` : ''}</div>
    </a></article>`;
}

/* ---------- home ---------- */
function opts(arr, cur, first, label = x => x) {
  return `<option value="">${esc(first)}</option>` + arr.map(v => `<option value="${esc(v)}"${v === cur ? ' selected' : ''}>${esc(label(v))}</option>`).join('');
}
const regions = () => [...new Set(LOTS.map(l => l.region))].sort();
const munis = () => [...new Set(LOTS.filter(l => !S.region || l.region === S.region).map(l => l.municipality))].sort((a, b) => a.localeCompare(b, 'sr'));

function heroSub() {
  const ar = LOTS.map(l => l.area).filter(Boolean), pr = LOTS.map(l => l.total).filter(Boolean);
  return t('heroS', fmt(Math.min(...ar)), fmt(Math.max(...ar) / 10000, 1), fmt(Math.min(...pr)));
}

function renderHome() {
  document.title = t('title');
  $('#app').innerHTML = `
  <section class="hero" style="background-image:url('img/hero.jpg')"><div class="wrap">
    <h1>${esc(t('heroT', LOTS.length))}</h1><p>${esc(heroSub())}</p>
    <form class="search" id="hs">
      <div class="seg" id="hseg">${[['', 'any'], ['море', 'sea'], ['гора', 'mountain']].map(([v, k]) => `<button type="button" data-v="${v}" class="${S.view === v ? 'on' : ''}">${esc(t(k))}</button>`).join('')}</div>
      <select id="hreg" aria-label="${esc(t('region'))}">${opts(regions(), S.region, t('allRegions'), regionName)}</select>
      <input id="hmax" type="number" min="0" step="1000" inputmode="numeric" placeholder="${esc(t('maxPrice'))}" aria-label="${esc(t('maxPrice'))}" value="${esc(S.pmax)}">
      <button class="btn" type="submit"><i class="ph ph-magnifying-glass"></i>${esc(t('show'))}</button>
    </form>
  </div></section>
  <div class="wrap" id="catalog"><div class="cat">
    <aside class="filters" id="flt" aria-label="${esc(t('filters'))}"></aside>
    <section>
      <div class="bar">
        <h2 id="rc"></h2>
        <button class="hbtn fbtn" id="fopen"><i class="ph ph-sliders-horizontal"></i>${esc(t('filters'))}</button>
        <select id="sort" aria-label="${esc(t('sort'))}">${[['rel', 'sRel'], ['pa', 'sPA'], ['pd', 'sPD'], ['ad', 'sAD'], ['aa', 'sAA'], ['pm', 'sPm']].map(([v, k]) => `<option value="${v}"${S.sort === v ? ' selected' : ''}>${esc(t(k))}</option>`).join('')}</select>
        <div class="views" id="views">${[['grid', 'squares-four'], ['list', 'rows'], ['map', 'map-trifold']].map(([m, ic]) => `<button data-m="${m}" class="${S.mode === m ? 'on' : ''}" aria-label="${esc(t(m))}"><i class="ph ph-${ic}"></i><span class="t">${esc(t(m))}</span></button>`).join('')}</div>
      </div>
      <div id="res"></div>
    </section>
  </div></div>`;
  renderFilters();
  updateResults();
  bindHome();
}

function renderFilters() {
  const num = (id, k, ph, lab) => `<input id="${id}" type="number" min="0" inputmode="numeric" placeholder="${esc(t(ph))}" aria-label="${esc(t(lab))} ${esc(t(ph))}" value="${esc(S[k])}">`;
  $('#flt').innerHTML = `
    <h3>${esc(t('filters'))} <button class="linkbtn" id="freset">${esc(t('reset'))}</button></h3>
    <div class="fld"><label for="fview">${esc(t('type'))}</label><select id="fview">${[['', 'any'], ['море', 'sea'], ['гора', 'mountain']].map(([v, k]) => `<option value="${v}"${S.view === v ? ' selected' : ''}>${esc(t(k))}</option>`).join('')}</select></div>
    <div class="fld"><label for="freg">${esc(t('region'))}</label><select id="freg">${opts(regions(), S.region, t('allRegions'), regionName)}</select></div>
    <div class="fld"><label for="fmuni">${esc(t('muni'))}</label><select id="fmuni">${opts(munis(), S.muni, t('allMuni'))}</select></div>
    <div class="fld"><span>${esc(t('price'))}</span><div class="two">${num('pmin', 'pmin', 'from', 'price')}${num('pmax', 'pmax', 'to', 'price')}</div></div>
    <div class="fld"><span>${esc(t('area'))}</span><div class="two">${num('amin', 'amin', 'from', 'area')}${num('amax', 'amax', 'to', 'area')}</div></div>
    <div class="fld">${[['photo', 'withPhoto'], ['video', 'withVideo'], ['tour', 'with3d'], ['saved', 'onlySaved']].map(([k, l]) => `<label class="chk"><input type="checkbox" id="c-${k}"${S[k] ? ' checked' : ''}> ${esc(t(l))}</label>`).join('')}</div>
    <button class="btn apply fbtn" id="fclose">${esc(t('apply'))}</button>`;
}

function updateResults() {
  const list = applyFilters();
  $('#rc').textContent = t('results', list.length);
  $$('#views button').forEach(b => b.classList.toggle('on', b.dataset.m === S.mode));
  const res = $('#res');
  if (homeMap) { homeMap.remove(); homeMap = null; }
  if (!list.length) {
    res.innerHTML = `<div class="empty"><i class="ph ph-magnifying-glass"></i><h3>${esc(t('noRes'))}</h3><p>${esc(t('noResS'))}</p><button class="btn ghost" id="rst2">${esc(t('reset'))}</button></div>`;
    return;
  }
  if (S.mode === 'map') { res.innerHTML = `<div id="map" role="application"></div><p class="maplab">${esc(t('mapNote'))}</p>`; drawHomeMap(list); return; }
  res.innerHTML = `<div class="${S.mode === 'list' ? 'list' : 'grid'}">${list.map(card).join('')}</div>`;
}

function drawHomeMap(list) {
  const counts = {};
  list.forEach(l => counts[l.region] = (counts[l.region] || 0) + 1);
  homeMap = L.map('map', { scrollWheelZoom: false }).setView([42.55, 19.1], 8);
  L.tileLayer('https://tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 18, attribution: '&copy; OpenStreetMap' }).addTo(homeMap);
  const pts = [];
  Object.entries(counts).forEach(([r, n]) => {
    const xy = REGION_XY[r]; if (!xy) return; pts.push(xy);
    const sz = 44 + Math.min(30, n * 2);
    L.marker(xy, { icon: L.divIcon({ className: '', html: `<div class="pin" style="width:${sz}px;height:${sz}px">${n}</div>`, iconSize: [sz, sz] }) }).addTo(homeMap)
      .bindPopup(`<b>${esc(regionName(r))}</b><br>${esc(t('results', n))}<br><a href="#/" data-pick="${esc(r)}">${esc(t('mapShow'))}</a>`);
  });
  list.filter(l => l.geo).forEach(l => {
    L.circle([l.geo.lat, l.geo.lng], { radius: Math.max(l.geo.r, 120), color: '#c04f15', weight: 2, fillColor: '#c04f15', fillOpacity: .3 }).addTo(homeMap)
      .bindPopup(`<b>${esc(t('lot'))} ${esc(l.id)}</b><br>${esc(l.municipality)}, ${l.area != null ? fmt(l.area) + ' ' + esc(t('m2')) : ''}<br><a href="#/p/${esc(l.id)}">${esc(t('sGeneral'))}</a>`);
  });
  if (pts.length) homeMap.fitBounds(pts, { padding: [60, 60], maxZoom: 10 });
}

function bindHome() {
  const upd = () => updateResults();
  $('#hseg').onclick = e => { const b = e.target.closest('button'); if (!b) return; S.view = b.dataset.v; $$('#hseg button').forEach(x => x.classList.toggle('on', x === b)); $('#fview').value = S.view; upd(); };
  $('#hreg').onchange = e => { S.region = e.target.value; S.muni = ''; renderFilters(); upd(); };
  $('#hmax').oninput = e => { S.pmax = e.target.value; $('#pmax').value = S.pmax; upd(); };
  $('#hs').onsubmit = e => { e.preventDefault(); $('#catalog').scrollIntoView({ behavior: 'smooth' }); };
  $('#flt').addEventListener('change', e => {
    const id = e.target.id;
    if (id === 'fview') { S.view = e.target.value; $$('#hseg button').forEach(x => x.classList.toggle('on', x.dataset.v === S.view)); }
    else if (id === 'freg') { S.region = e.target.value; S.muni = ''; $('#hreg').value = S.region; renderFilters(); }
    else if (id === 'fmuni') S.muni = e.target.value;
    else if (id.startsWith('c-')) S[id.slice(2)] = e.target.checked;
    upd();
  });
  $('#flt').addEventListener('input', e => {
    const k = { pmin: 'pmin', pmax: 'pmax', amin: 'amin', amax: 'amax' }[e.target.id]; if (!k) return;
    S[k] = e.target.value; if (k === 'pmax') $('#hmax').value = S.pmax; upd();
  });
  $('#sort').onchange = e => { S.sort = e.target.value; upd(); };
  $('#views').onclick = e => { const b = e.target.closest('button'); if (!b) return; S.mode = b.dataset.m; upd(); };
  $('#fopen').onclick = () => $('#flt').classList.add('open');
  $('#flt').addEventListener('click', e => {
    if (e.target.closest('#fclose')) $('#flt').classList.remove('open');
    if (e.target.closest('#freset')) { const m = S.mode; S = DEF(); S.mode = m; renderHome(); }
  });
  $('#res').addEventListener('click', e => { if (e.target.closest('#rst2')) { const m = S.mode; S = DEF(); S.mode = m; renderHome(); } });
}

/* ---------- favorites, toast ---------- */
function toast(msg) {
  $$('.toast').forEach(x => x.remove());
  const d = document.createElement('div'); d.className = 'toast'; d.setAttribute('role', 'status'); d.textContent = msg; document.body.append(d); setTimeout(() => d.remove(), 2200);
}
function updateSavedBadge() {
  $('#savedT').textContent = t('saved');
  const n = $('#savedN'); n.hidden = !fav.size; n.textContent = fav.size;
}
document.addEventListener('click', e => {
  const f = e.target.closest('[data-fav]');
  if (f) {
    e.preventDefault(); const id = f.dataset.fav;
    fav.has(id) ? fav.delete(id) : fav.add(id);
    ls.set('fav', [...fav]); updateSavedBadge();
    if (S.saved && $('#res')) { updateResults(); return; }
    const on = fav.has(id);
    $$(`[data-fav="${CSS.escape(id)}"]`).forEach(b => { b.classList.toggle('on', on); b.setAttribute('aria-pressed', on); const i = b.querySelector('i'); if (i) i.className = `ph${on ? '-fill' : ''} ph-heart`; const s = b.querySelector('span'); if (s) s.textContent = t(on ? 'unsave' : 'save'); });
    return;
  }
  const pk = e.target.closest('[data-pick]');
  if (pk) { e.preventDefault(); S = Object.assign(DEF(), { region: pk.dataset.pick }); if (location.hash === '#/' || location.hash === '') renderRoute(); else location.hash = '#/'; }
});

/* ---------- lot page ---------- */
function mediaBlock(src, bg, icon, label) {
  return `<div class="media" data-src="${esc(src)}"><button class="go" ${bg ? `style="background-image:url('${esc(bg)}')"` : ''}><span><i class="ph ${icon}"></i>${esc(label)}</span></button></div>`;
}
const nfmt = v => v == null ? '-' : fmt(v, v % 1 ? 1 : 0);

function geoLinks(l) {
  const g = l.geo;
  const q = g ? `${g.lat},${g.lng}` : encodeURIComponent(`${l.municipality}, ${REGION_EN[l.region] || l.region}, Montenegro`);
  return { open: `https://www.google.com/maps/search/?api=1&query=${q}`, route: g ? `https://www.google.com/maps/dir/?api=1&destination=${q}` : '' };
}

function locationPanel(l) {
  const g = l.geo, links = geoLinks(l);
  if (!g) return `<div class="card-p"><div class="lbl"><i class="ph ph-map-pin"></i>${esc(t('sLocation'))}</div><div class="nogeo"><span>${esc(l.municipality)}, ${esc(regionName(l.region))}. ${esc(t('noGeoT'))}</span><a class="btn ghost" href="${links.open}" target="_blank" rel="noopener"><i class="ph ph-map-pin"></i>${esc(t('findG'))}</a></div></div>`;
  return `<div class="card-p"><div class="lbl"><i class="ph ph-map-pin"></i>${esc(t('sLocation'))}</div>
    <div class="tabs" id="mtabs"><button data-k="sat" class="on">${esc(t('tabSat'))}</button><button data-k="map">${esc(t('tabMap'))}</button><button data-k="g">${esc(t('tabG'))}</button></div>
    <div class="mapbox"><div id="mapcanvas"></div><iframe id="gframe" hidden title="Google Maps" loading="lazy" allowfullscreen referrerpolicy="no-referrer-when-downgrade"></iframe></div>
    <div class="maplinks"><span class="sp">${esc(t('circleNote'))}</span>
      <a class="hbtn" href="${links.open}" target="_blank" rel="noopener"><i class="ph ph-map-pin"></i>${esc(t('openG'))}</a>
      <a class="hbtn" href="${links.route}" target="_blank" rel="noopener"><i class="ph ph-compass"></i>${esc(t('routeG'))}</a></div></div>`;
}

function mountLotMap(l) {
  const g = l.geo; if (!g || !$('#mapcanvas')) return;
  const sat = L.tileLayer(ESRI, { maxZoom: 19, attribution: 'Tiles &copy; Esri, Maxar, Earthstar Geographics' });
  const osm = L.tileLayer('https://tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 19, attribution: '&copy; OpenStreetMap' });
  lotMap = L.map('mapcanvas', { scrollWheelZoom: false }).setView([g.lat, g.lng], 17);
  sat.addTo(lotMap);
  L.circle([g.lat, g.lng], { radius: g.r, color: '#c04f15', weight: 4, fillColor: '#c04f15', fillOpacity: .18 }).addTo(lotMap)
    .bindTooltip(`${t('lot')} ${l.id}`, { permanent: true, direction: 'top', className: 'lotlbl', offset: [0, -4] });
  $('#mtabs').onclick = e => {
    const b = e.target.closest('button'); if (!b) return;
    $$('#mtabs button').forEach(x => x.classList.toggle('on', x === b));
    const k = b.dataset.k, gf = $('#gframe');
    if (k === 'g') { if (!gf.src) gf.src = `https://www.google.com/maps?q=${g.lat},${g.lng}&z=18&t=k&output=embed`; gf.hidden = false; $('#mapcanvas').hidden = true; return; }
    gf.hidden = true; $('#mapcanvas').hidden = false;
    if (k === 'map') { lotMap.removeLayer(sat); osm.addTo(lotMap); } else { lotMap.removeLayer(osm); sat.addTo(lotMap); }
    lotMap.invalidateSize();
  };
}

const COMM_ICON = { electricity: 'lightning', water: 'drop', gas: 'flame', sewage: 'pipe' };
const TR_ICON = { beach: 'umbrella-simple', portBudva: 'anchor', airportTivat: 'airplane-tilt' };

function renderLot(id) {
  const l = lotById(id);
  if (!l) { location.hash = '#/'; return; }
  const info = l.info || {}, ph = l.photos;
  document.title = `${t('lot')} ${l.id}, ${l.municipality} | ALLIANZ KAPITAL`;
  let top;
  if (ph.length) {
    const shown = ph.slice(0, 5);
    top = `<div class="mosaic ${shown.length < 3 ? 'one' : ''}" id="mos">${shown.map((i, k) => `<button data-i="${k}" aria-label="${esc(t('photos'))} ${k + 1}"><img src="${photoBig(i)}" alt="" ${k ? 'loading="lazy"' : 'fetchpriority="high"'}></button>`).join('')}<button class="allph" data-i="0"><i class="ph ph-images"></i>${esc(t('allPhotos', ph.length))}</button></div>`;
  } else if (l.tourCover) {
    top = `<div class="cover3d"><img src="photos/${esc(l.tourCover)}" alt=""><a class="btn" href="#m3d" id="to3d"><i class="ph ph-cube"></i>${esc(t('open3d'))}</a></div>`;
  } else top = `<div class="cover3d"><div class="noimg"><i class="ph ph-${unit(l)}"></i><b>${l.area != null ? fmt(l.area) + ' ' + esc(t('m2')) : ''}</b><span>${esc(t('noPhoto'))}</span></div></div>`;

  const restr = info.restrictions === 'none' ? D().v.none : !l.restricted ? t('restrNo') : l.restricted === l.n ? t('restrY') : t('restrSome', l.restricted, l.n);
  const facts = [
    [t('totalArea'), l.area != null ? `${fmt(l.area)} ${t('m2')}${l.area >= 10000 ? ` (${fmt(l.area / 10000, 2)} ${t('ha')})` : ''}` : '-', 'big'],
    [t('cadastreN'), l.n], [t('type'), viewName(l.view)], [t('region'), regionName(l.region)], [t('muni'), l.municipality],
    info.category && [t('category'), D().v[info.category]], info.right && [t('right'), D().v[info.right]],
    [t('restr'), restr], info.cooperation && [t('coop'), D().v[info.cooperation]],
    [t('avail'), l.partial ? t('partialOf', l.partial, l.n) : t('statusN')],
    info.since && [t('since', info.since).split(' ')[0] === t('since', info.since).split(' ')[0] ? '' : '', ''] ].filter(Boolean).filter(f => f[0] !== '' || f[1] !== '');
  const comms = info.comms ? `<div class="comms" aria-label="${esc(t('commsT'))}">${info.comms.map(c => `<span><i class="ph ph-${COMM_ICON[c]}"></i>${esc(D().comms[c])}</span>`).join('')}</div>` : '';
  const sinceNote = info.since ? `<p class="meta" style="margin:14px 0 0"><i class="ph ph-seal-check"></i> ${esc(t('since', info.since))}</p>` : '';
  const vol = info.volume && info.volume.gross ? `<div class="card-p"><div class="lbl"><i class="ph ph-buildings"></i>${esc(t('sVolume'))}</div><div class="in"><div class="vol">${Object.entries(info.volume).map(([k, v]) => `<div class="${k === 'gross' ? 'main' : ''}"><b>${fmt(v)} ${esc(t('m2'))}</b><span>${esc(D().vol[k])}</span></div>`).join('')}</div></div></div>` : '';
  const tr = info.transport ? `<div class="card-p"><div class="lbl"><i class="ph ph-road-horizon"></i>${esc(t('sTransport'))}</div><div class="in"><div class="tr">${info.transport.map(x => `<a href="${safeUrl(x.link)}" target="_blank" rel="noopener"><i class="ph ph-${TR_ICON[x.k]}"></i><b>${esc(x.how === 'walk' ? t('walk', x.min, fmt(x.m)) : t('car', x.min, nfmt(x.km)))}</b><span>${esc(D().tr[x.k])}, ${esc(t('route'))}</span></a>`).join('')}</div></div></div>` : '';
  const comp = l.n > 1 ? `<div class="card-p"><div class="lbl"><i class="ph ph-stack"></i>${esc(t('sComposition'))} (${l.n})</div><div class="tbl-w"><table class="tbl"><thead><tr><th>${esc(t('thParcel'))}</th><th class="n">${esc(t('thArea'))}</th><th class="n">${esc(t('thPrice'))}</th><th>${esc(t('thNote'))}</th></tr></thead><tbody>${l.parcels.map(p => `<tr><td>${esc(p.cadastre)}</td><td class="n">${fmt(p.area)}</td><td class="n">${fmt(p.total)}</td><td>${[p.partial ? t('partial') : '', p.restricted ? t('encumb') : ''].filter(Boolean).join(', ')}</td></tr>`).join('')}</tbody></table></div></div>` : '';
  const plan = [[t('occ'), l.occupancy], [t('far'), l.far], [t('floors'), l.floors], [t('plan'), l.plan]].filter(([, v]) => v);
  const planP = plan.length ? `<div class="card-p"><div class="lbl"><i class="ph ph-ruler"></i>${esc(t('sPlanning'))}</div><div class="in"><dl class="facts">${plan.map(([k, v]) => `<div><dt>${esc(k)}</dt><dd>${esc(v)}</dd></div>`).join('')}</dl>${l.floors ? `<p class="meta" style="margin:14px 0 0">${esc(t('floorsHint'))}</p>` : ''}</div></div>` : '';
  const imgs = info.img ? [...info.img.cad.map(x => [x, t('cad')]), info.img.concept && [info.img.concept, t('concept')]].filter(Boolean) : [];
  const schemes = imgs.length ? `<div class="card-p"><div class="lbl"><i class="ph ph-map-trifold"></i>${esc(t('sSchemes'))}</div><div class="in"><div class="schemes">${imgs.map(([f, a]) => `<img src="img/${esc(f)}" alt="${esc(a)}" loading="lazy">`).join('')}</div></div></div>` : '';
  const vids = l.videos.length ? `<div class="card-p"><div class="lbl"><i class="ph ph-video-camera"></i>${esc(t('sVideo'))}</div><div class="in">${l.videos.slice(0, 3).map(v => mediaBlock(`https://drive.google.com/file/d/${encodeURIComponent(v)}/preview`, ph.length ? photoBig(ph[0]) : '', 'ph-play', t('watch'))).join('')}</div></div>` : '';
  const tours = has3d(l) ? `<div class="card-p" id="m3d"><div class="lbl"><i class="ph ph-cube"></i>${esc(t('s3d'))}</div><div class="in">${l.tours.slice(0, 3).map((u, k) => mediaBlock(u, l.tourCover && !k ? `photos/${l.tourCover}` : '', 'ph-cube', t('open3d')) + `<p class="meta"><a href="${safeUrl(u)}" target="_blank" rel="noopener">${esc(t('loadFail'))}</a></p>`).join('')}</div></div>` : '';
  const docs = [['idea', 'docIdea'], ['utu', 'docUtu'], ['deck', 'docDeck']].filter(([k]) => l.docs[k]);
  const docsP = docs.length ? `<div class="card-p"><div class="lbl"><i class="ph ph-file-text"></i>${esc(t('sDocs'))}</div><div class="in"><div class="doclinks">${docs.map(([k, n]) => `<a class="hbtn" href="${safeUrl(l.docs[k])}" target="_blank" rel="noopener"><i class="ph ph-file-text"></i>${esc(t(n))}</a>`).join('')}</div></div></div>` : '';
  const c = contactSet();
  const price = l.total != null ? eur(l.total) : t('onReq');
  const sim = LOTS.filter(x => x.id !== l.id && x.region === l.region && hasMedia(x)).slice(0, 3);
  const on = fav.has(l.id);

  $('#app').innerHTML = `<div class="wrap">
    <nav class="crumbs" aria-label="breadcrumb"><a href="#/"><i class="ph ph-arrow-left"></i> ${esc(t('back'))}</a><span>/</span><span>${esc(regionName(l.region))}</span><span>/</span><span>${esc(l.municipality)}</span></nav>
    ${top}
    <div class="dgrid"><div>
      <h1 class="title">${l.area != null ? `${fmt(l.area)} ${esc(t('m2'))}, ` : ''}${esc(l.municipality)}</h1>
      <p class="meta" style="margin:0 0 16px">${esc(t('lot'))} ${esc(l.id)}. ${esc(regionName(l.region))}, ${esc(viewName(l.view))}</p>
      <div class="card-p"><div class="lbl"><i class="ph ph-info"></i>${esc(t('sGeneral'))}</div><div class="in"><dl class="facts">${facts.map(([k, v, cl]) => `<div><dt>${esc(k)}</dt><dd class="${cl || ''}">${esc(v)}</dd></div>`).join('')}</dl>${comms}${sinceNote}</div></div>
      ${vol}${locationPanel(l)}${tr}${comp}${planP}${schemes}${vids}${tours}${docsP}
    </div>
    <aside class="side"><div class="card-p"><div class="in">
      <div class="price">${price}</div>
      ${l.ppm && l.total != null ? `<div class="meta">${eur(l.ppm)} ${esc(t('perM2'))}</div>` : ''}
      <div class="stats"><div><b>${l.area != null ? fmt(l.area) + ' ' + esc(t('m2')) : '-'}</b><span>${esc(t('totalArea'))}</span></div><div><b>${l.n}</b><span>${esc(t('cadastreN'))}</span></div></div>
      ${c ? `<div class="form"><label>${esc(t('yourMsg'))}<textarea id="msg" rows="3">${esc(t('msgDef', l.id, l.municipality))}</textarea></label>
        <div style="display:grid;gap:8px">
        ${c.whatsapp ? `<a class="btn block" data-ct="wa" href="#"><i class="ph ph-whatsapp-logo"></i>${esc(t('viaWa'))}</a>` : ''}
        ${c.telegram ? `<a class="btn ghost block" data-ct="tg" href="#"><i class="ph ph-telegram-logo"></i>${esc(t('viaTg'))}</a>` : ''}
        ${c.email ? `<a class="btn ghost block" data-ct="mail" href="#"><i class="ph ph-envelope-simple"></i>${esc(t('viaMail'))}</a>` : ''}
        ${c.phone ? `<a class="btn ghost block" href="tel:+${esc(String(c.phone).replace(/\D/g, ''))}"><i class="ph ph-phone"></i>${esc(t('call'))}</a>` : ''}</div></div>` : ''}
      <div class="row"><button class="sbtn ${on ? 'on' : ''}" data-fav="${esc(l.id)}" aria-pressed="${on}"><i class="ph${on ? '-fill' : ''} ph-heart"></i><span>${esc(t(on ? 'unsave' : 'save'))}</span></button>
      <button class="sbtn" id="share"><i class="ph ph-share-network"></i>${esc(t('share'))}</button></div>
    </div></div></aside></div>
    ${sim.length ? `<section style="padding-bottom:56px"><h2>${esc(t('similar'))}</h2><div class="sim">${sim.map(card).join('')}</div></section>` : ''}
  </div>`;

  mountLotMap(l);
  $('#share').onclick = async () => {
    const url = location.href;
    try { if (navigator.share) { await navigator.share({ url, title: document.title }); return; } await navigator.clipboard.writeText(url); toast(t('copied')); } catch (e) { }
  };
  $$('.media').forEach(m => $('.go', m).onclick = () => { m.innerHTML = `<iframe src="${m.dataset.src}" allow="fullscreen; xr-spatial-tracking; gyroscope; accelerometer; autoplay" allowfullscreen loading="lazy" title="media"></iframe>`; });
  const to3d = $('#to3d'); if (to3d) to3d.onclick = e => { e.preventDefault(); $('#m3d').scrollIntoView({ behavior: 'smooth' }); };
  $$('#mos [data-i]').forEach(b => b.onclick = () => openLb(ph, +b.dataset.i));
  $$('[data-ct]').forEach(a => a.addEventListener('click', () => {
    const m = encodeURIComponent($('#msg').value);
    a.href = a.dataset.ct === 'wa' ? `https://wa.me/${String(c.whatsapp).replace(/\D/g, '')}?text=${m}`
      : a.dataset.ct === 'tg' ? `https://t.me/${encodeURIComponent(c.telegram)}`
      : `mailto:${c.email}?subject=${encodeURIComponent(t('lot') + ' ' + l.id)}&body=${m}`;
    a.target = '_blank'; a.rel = 'noopener';
  }));
}

/* ---------- about ---------- */
function renderAbout() {
  document.title = `${t('abTitle')} | ALLIANZ KAPITAL`;
  const ha = LOTS.reduce((s, l) => s + (l.area || 0), 0) / 10000;
  const regs = regions();
  const cnt = r => LOTS.filter(l => l.region === r).length;
  const bg = { Budva: 'img/coast.jpg', Kotor: 'img/hero.jpg', Kolasin: 'img/kolasin.jpg' };
  const regOrder = ['Budva', 'Kotor', 'Kolasin'].filter(r => regs.includes(r));
  const flags = ['M01', 'M03', 'M07'].map(lotById).filter(Boolean);
  const c = contactSet();
  $('#app').innerHTML = `
  <section class="abt"><div class="wrap"><h1>${esc(t('abTitle'))}</h1>${D().ab.map(p => `<p>${esc(p)}</p>`).join('')}</div></section>
  <div class="wrap">
    <div class="nums reveal"><div><b>${fmt(LOTS.reduce((n, l) => n + l.n, 0))}</b><span>${esc(t('abParcels'))}</span></div><div><b>${LOTS.length}</b><span>${esc(t('abLots'))}</span></div><div><b>${fmt(ha, 1)}</b><span>${esc(t('abHa'))}</span></div><div><b>${regs.length}</b><span>${esc(t('abRegions'))}</span></div></div>
    <section class="sec reveal"><h2>${esc(t('abDoT'))}</h2><div class="do">
      <div class="big-t"><i class="ph ph-tree"></i><p>${esc(t('abDo1'))}</p></div>
      <div class="small-t"><i class="ph ph-buildings"></i><p>${esc(t('abDo2'))}</p></div></div></section>
    <section class="sec reveal"><h2>${esc(t('abRegT'))}</h2><div class="regs">${regOrder.map(r => `<a class="reg" href="#/" data-pick="${r}" style="background-image:url('${bg[r]}')"><div><b>${esc(regionName(r))}</b><span>${esc(t('abRegN', cnt(r)))}</span></div></a>`).join('')}</div></section>
    <section class="sec reveal"><h2>${esc(t('abFlagT'))}</h2>${flags.map(l => {
      const pic = l.photos.length ? photoBig(l.photos[Math.min(1, l.photos.length - 1)]) : `img/${l.info && l.info.img ? l.info.img.concept : 'coast.jpg'}`;
      return `<article class="flag"><div class="pic" style="background-image:url('${pic}')"></div><div class="tx"><h3>${esc(t('lot'))} ${esc(l.id)}, ${esc(l.municipality)}</h3>
        <dl><div><dt>${esc(t('totalArea'))}</dt><dd>${fmt(l.area)} ${esc(t('m2'))}</dd></div><div><dt>${esc(t('cadastreN'))}</dt><dd>${l.n}</dd></div>
        <div><dt>${esc(t('region'))}</dt><dd>${esc(regionName(l.region))}</dd></div><div><dt>${esc(t('price'))}</dt><dd>${l.total != null ? eur(l.total) : esc(t('onReq'))}</dd></div></dl>
        <div><a class="btn" href="#/p/${esc(l.id)}">${esc(t('abOpen'))}</a></div></div></article>`; }).join('')}</section>
    <section class="sec reveal"><div class="card-p"><div class="lbl"><i class="ph ph-identification-card"></i>${esc(t('coT'))}</div><div class="in">
      <dl class="facts"><div><dt>${esc(t('coName'))}</dt><dd>${esc(CO.name)}</dd></div><div><dt>${esc(t('coAddr'))}</dt><dd>${esc(CO.addr)}</dd></div><div><dt>${esc(t('coReg'))}</dt><dd>${esc(regDate())}</dd></div></dl>
      <div class="doclinks" style="margin-top:16px"><a class="hbtn" href="https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(CO.addr)}" target="_blank" rel="noopener"><i class="ph ph-map-pin"></i>${esc(t('coMap'))}</a>
      <a class="hbtn" href="${esc(CO.registry)}" target="_blank" rel="noopener"><i class="ph ph-arrow-up-right"></i>${esc(t('coReg3'))}</a></div></div></div></section>
    <section class="cta reveal"><div><h2>${esc(t('abCtaT'))}</h2><p>${esc(t('abCtaS'))}</p></div>
      <div style="display:flex;gap:10px;flex-wrap:wrap"><a class="btn light" href="#/">${esc(t('abCat'))}</a>
      ${c && c.whatsapp ? `<a class="btn" href="https://wa.me/${esc(String(c.whatsapp).replace(/\D/g, ''))}" target="_blank" rel="noopener"><i class="ph ph-whatsapp-logo"></i>${esc(t('viaWa'))}</a>` : ''}
      ${c && c.email ? `<a class="btn" href="mailto:${esc(c.email)}"><i class="ph ph-envelope-simple"></i>${esc(t('viaMail'))}</a>` : ''}</div></section>
  </div>`;
  const io = 'IntersectionObserver' in window ? new IntersectionObserver(es => es.forEach(e => { if (e.isIntersecting) { e.target.classList.add('in'); io.unobserve(e.target); } }), { threshold: .12 }) : null;
  $$('.reveal').forEach(el => io ? io.observe(el) : el.classList.add('in'));
}

/* ---------- lightbox ---------- */
let lbList = [], lbI = 0;
function showLb() { $('#lbImg').src = photoBig(lbList[lbI]); $('#lbCt').textContent = `${lbI + 1} / ${lbList.length}`; }
function openLb(list, i) { lbList = list; lbI = i; showLb(); $('#lb').showModal(); }
$('#lbX').onclick = () => $('#lb').close();
$('#lbP').onclick = () => { lbI = (lbI - 1 + lbList.length) % lbList.length; showLb(); };
$('#lbN').onclick = () => { lbI = (lbI + 1) % lbList.length; showLb(); };
$('#lb').addEventListener('keydown', e => { if (e.key === 'ArrowLeft') $('#lbP').click(); if (e.key === 'ArrowRight') $('#lbN').click(); });
$('#lb').addEventListener('click', e => { if (e.target === $('#lb') || e.target.classList.contains('st')) $('#lb').close(); });

/* ---------- router ---------- */
function renderRoute() {
  if (lotMap) { lotMap.remove(); lotMap = null; }
  if (homeMap) { homeMap.remove(); homeMap = null; }
  const h = location.hash || '#/';
  const m = h.match(/^#\/p\/([\w-]+)/);
  updateSavedBadge();
  $('#disc').textContent = t('disc');
  $('#co').textContent = `${CO.name}, ${CO.addr}. ${t('coReg2')} ${regDate()}${regDate().endsWith('.') ? '' : '.'}`;
  $('#navCat').textContent = t('navCat'); $('#navAbout').textContent = t('navAbout');
  $('#navCat').classList.toggle('on', !h.startsWith('#/about')); $('#navAbout').classList.toggle('on', h.startsWith('#/about'));
  $$('#lang button').forEach(b => b.classList.toggle('on', b.dataset.l === lang));
  document.documentElement.lang = D().html;
  if (m) { renderLot(m[1]); window.scrollTo(0, 0); return; }
  if (h.startsWith('#/about')) { renderAbout(); window.scrollTo(0, 0); return; }
  if (h === '#/saved') S = Object.assign(DEF(), { saved: true });
  else if (S.saved && h === '#/') S.saved = false;
  renderHome();
  if (h === '#/saved') $('#catalog').scrollIntoView(); else window.scrollTo(0, 0);
}
window.addEventListener('hashchange', renderRoute);
$('#lang').onclick = e => { const b = e.target.closest('button'); if (!b) return; lang = b.dataset.l; ls.set('lang', lang); renderRoute(); };
renderRoute();
