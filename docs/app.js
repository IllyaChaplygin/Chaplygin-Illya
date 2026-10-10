'use strict';
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const ls = {
  get(k, d) { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} }
};
const guess = () => { const n = (navigator.language || 'en').toLowerCase(); return n.startsWith('ru') ? 'ru' : /^(sr|hr|bs|cnr|me)/.test(n) ? 'me' : 'en'; };
let lang = ls.get('lang', null); if (!I18N[lang]) lang = guess();
const L = () => I18N[lang];
const t = (k, ...a) => { const v = L()[k]; return typeof v === 'function' ? v(...a) : (v ?? k); };
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const safeUrl = u => (/^https:\/\//.test(u || '') ? esc(u) : '#');
const fmt = (v, d = 0) => v == null ? '-' : v.toLocaleString(L().loc, { maximumFractionDigits: d });
const eur = v => '€ ' + fmt(v);
const fav = new Set(ls.get('fav', []));
const REGION_XY = { Kolasin: [42.8228, 19.5206], Budva: [42.2864, 18.84], Kotor: [42.4247, 18.7712] };
const DEF = () => ({ view: '', region: '', muni: '', pmin: '', pmax: '', amin: '', amax: '', photo: false, video: false, tour: false, saved: false, sort: 'rel', mode: 'grid' });
let S = DEF();
let map = null;

/* ---------- helpers ---------- */
const hasMedia = p => p.photos.length || p.tourCover;
const photoSm = i => `photos/${i}_s.jpg`;
const photoBig = i => `photos/${i}.jpg`;
const coverOf = p => p.photos.length ? photoSm(p.photos[p.id % Math.min(4, p.photos.length)]) : (p.tourCover ? `photos/${p.tourCover}` : '');
const has3d = p => !!(p.tour || p.tourGroup);
const parcelById = id => PARCELS.find(p => p.id === id);
const regionName = r => L().reg[r] || r;
const viewName = v => L().view[v] || v;

function applyFilters() {
  const n = x => x === '' ? null : Number(x);
  const pmin = n(S.pmin), pmax = n(S.pmax), amin = n(S.amin), amax = n(S.amax);
  let list = PARCELS.filter(p =>
    (!S.view || p.view === S.view) && (!S.region || p.region === S.region) && (!S.muni || p.municipality === S.muni) &&
    (pmin == null || (p.total != null && p.total >= pmin)) && (pmax == null || (p.total != null && p.total <= pmax)) &&
    (amin == null || (p.area != null && p.area >= amin)) && (amax == null || (p.area != null && p.area <= amax)) &&
    (!S.photo || p.photos.length) && (!S.video || p.video) && (!S.tour || has3d(p)) && (!S.saved || fav.has(p.id)));
  const key = { pa: p => p.total ?? Infinity, pd: p => -(p.total ?? -Infinity), ad: p => -(p.area ?? 0), aa: p => p.area ?? Infinity, pm: p => p.ppm ?? Infinity,
    rel: p => (p.photos.length ? 0 : p.tourCover ? 1 : 2) * 1e6 + p.id }[S.sort];
  return list.sort((a, b) => key(a) - key(b) || a.id - b.id);
}

/* ---------- cards ---------- */
function card(p) {
  const src = coverOf(p);
  const img = src ? `<img src="${esc(src)}" alt="" loading="lazy" decoding="async">`
    : `<div class="noimg"><i class="ph ph-image"></i>${esc(t('noPhoto'))}</div>`;
  const badge = (p.photos.length || p.video || has3d(p)) ? `<span class="badge">${p.photos.length ? `<span><i class="ph ph-camera"></i> ${p.photos.length}</span>` : ''}${p.video ? '<i class="ph ph-video-camera"></i>' : ''}${has3d(p) ? '<i class="ph ph-cube"></i>' : ''}</span>` : '';
  const on = fav.has(p.id);
  const price = p.total != null ? `${eur(p.total)}${p.ppm ? `<small>${eur(p.ppm)} ${esc(t('perM2'))}</small>` : ''}` : esc(t('onReq'));
  return `<article class="card" data-id="${p.id}">
    <a class="im" href="#/p/${p.id}" aria-label="${esc(p.cadastre)}">${img}${badge}</a>
    <button class="fav ${on ? 'on' : ''}" data-fav="${p.id}" aria-label="${esc(t(on ? 'unsave' : 'save'))}" aria-pressed="${on}"><i class="ph${on ? '-fill' : ''} ph-heart"></i></button>
    <a class="bd" href="#/p/${p.id}">
      <div class="price">${price}</div>
      <div class="loc">${p.area != null ? `${fmt(p.area)} ${esc(t('m2'))}` : ''}${p.area != null ? ', ' : ''}${esc(viewName(p.view))}</div>
      <div class="meta">${esc(p.municipality)}, ${esc(regionName(p.region))}</div>
      <div class="meta">${esc(t('cadastre'))} ${esc(p.cadastre)}</div>
      ${p.status === 'частично' ? `<div class="tags"><span class="tag w">${esc(t('statusP'))}</span></div>` : ''}
    </a></article>`;
}

/* ---------- home ---------- */
function opts(arr, cur, first, label = x => x) {
  return `<option value="">${esc(first)}</option>` + arr.map(v => `<option value="${esc(v)}"${v === cur ? ' selected' : ''}>${esc(label(v))}</option>`).join('');
}
const regions = () => [...new Set(PARCELS.map(p => p.region))].sort();
const munis = () => [...new Set(PARCELS.filter(p => !S.region || p.region === S.region).map(p => p.municipality))].sort((a, b) => a.localeCompare(b, 'sr'));

function renderHome() {
  const heroP = parcelById(105) || PARCELS.find(p => p.view === 'море' && p.photos.length);
  const heroImg = heroP && heroP.photos.length ? photoBig(heroP.photos[Math.min(2, heroP.photos.length - 1)]) : '';
  document.title = t('title');
  $('#app').innerHTML = `
  <section class="hero" ${heroImg ? `style="background-image:url('${heroImg}')"` : ''}><div class="wrap">
    <h1>${esc(t('heroT'))}</h1><p>${esc(t('heroS'))}</p>
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
  $('#flt').innerHTML = `
    <h3>${esc(t('filters'))} <button class="linkbtn" id="freset">${esc(t('reset'))}</button></h3>
    <div class="fld"><label for="fview">${esc(t('type'))}</label><select id="fview">${[['', 'any'], ['море', 'sea'], ['гора', 'mountain']].map(([v, k]) => `<option value="${v}"${S.view === v ? ' selected' : ''}>${esc(t(k))}</option>`).join('')}</select></div>
    <div class="fld"><label for="freg">${esc(t('region'))}</label><select id="freg">${opts(regions(), S.region, t('allRegions'), regionName)}</select></div>
    <div class="fld"><label for="fmuni">${esc(t('muni'))}</label><select id="fmuni">${opts(munis(), S.muni, t('allMuni'))}</select></div>
    <div class="fld"><span>${esc(t('price'))}</span><div class="two"><input id="pmin" type="number" min="0" inputmode="numeric" placeholder="${esc(t('from'))}" aria-label="${esc(t('price'))} ${esc(t('from'))}" value="${esc(S.pmin)}"><input id="pmax" type="number" min="0" inputmode="numeric" placeholder="${esc(t('to'))}" aria-label="${esc(t('price'))} ${esc(t('to'))}" value="${esc(S.pmax)}"></div></div>
    <div class="fld"><span>${esc(t('area'))}</span><div class="two"><input id="amin" type="number" min="0" inputmode="numeric" placeholder="${esc(t('from'))}" aria-label="${esc(t('area'))} ${esc(t('from'))}" value="${esc(S.amin)}"><input id="amax" type="number" min="0" inputmode="numeric" placeholder="${esc(t('to'))}" aria-label="${esc(t('area'))} ${esc(t('to'))}" value="${esc(S.amax)}"></div></div>
    <div class="fld">${[['photo', 'withPhoto'], ['video', 'withVideo'], ['tour', 'with3d'], ['saved', 'onlySaved']].map(([k, l]) => `<label class="chk"><input type="checkbox" id="c-${k}"${S[k] ? ' checked' : ''}> ${esc(t(l))}</label>`).join('')}</div>
    <button class="btn apply fbtn" id="fclose">${esc(t('apply'))}</button>`;
}

function updateResults() {
  const list = applyFilters();
  $('#rc').textContent = t('results', list.length);
  $$('#views button').forEach(b => b.classList.toggle('on', b.dataset.m === S.mode));
  const res = $('#res');
  if (map) { map.remove(); map = null; }
  if (!list.length) {
    res.innerHTML = `<div class="empty"><i class="ph ph-magnifying-glass"></i><h3>${esc(t('noRes'))}</h3><p>${esc(t('noResS'))}</p><button class="btn ghost" id="rst2">${esc(t('reset'))}</button></div>`;
    return;
  }
  if (S.mode === 'map') { res.innerHTML = `<div id="map" role="application"></div><p class="maplab">${esc(t('mapNote'))}</p>`; drawMap(list); return; }
  res.innerHTML = `<div class="${S.mode === 'list' ? 'list' : 'grid'}">${list.map(card).join('')}</div>`;
}

function drawMap(list) {
  const counts = {};
  list.forEach(p => counts[p.region] = (counts[p.region] || 0) + 1);
  map = L_.map('map', { scrollWheelZoom: false }).setView([42.55, 19.1], 8);
  L_.tileLayer('https://tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 14, attribution: '&copy; OpenStreetMap' }).addTo(map);
  const pts = [];
  Object.entries(counts).forEach(([r, n]) => {
    const xy = REGION_XY[r]; if (!xy) return; pts.push(xy);
    const sz = 44 + Math.min(30, n / 2);
    const m = L_.marker(xy, { icon: L_.divIcon({ className: '', html: `<div class="pin" style="width:${sz}px;height:${sz}px">${n}</div>`, iconSize: [sz, sz] }) }).addTo(map);
    m.bindPopup(`<b>${esc(regionName(r))}</b><br>${esc(t('results', n))}<br><a href="#/" data-pick="${esc(r)}">${esc(t('mapShow'))}</a>`);
  });
  if (pts.length) map.fitBounds(pts, { padding: [60, 60], maxZoom: 10 });
}

function bindHome() {
  const upd = () => { updateResults(); };
  const bindNum = (id, key) => $(id).addEventListener('input', e => { S[key] = e.target.value; upd(); });
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
  [['#pmin', 'pmin'], ['#pmax', 'pmax'], ['#amin', 'amin'], ['#amax', 'amax']].forEach(([i, k]) => $('#flt').addEventListener('input', e => { if (e.target.matches(i)) { S[k] = e.target.value; if (k === 'pmax') $('#hmax').value = S.pmax; upd(); } }));
  $('#sort').onchange = e => { S.sort = e.target.value; upd(); };
  $('#views').onclick = e => { const b = e.target.closest('button'); if (!b) return; S.mode = b.dataset.m; upd(); };
  $('#fopen').onclick = () => $('#flt').classList.add('open');
  $('#flt').addEventListener('click', e => {
    if (e.target.closest('#fclose')) $('#flt').classList.remove('open');
    if (e.target.closest('#freset')) { const m = S.mode, sv = S.saved && false; S = DEF(); S.mode = m; renderHome(); }
  });
  $('#res').addEventListener('click', e => {
    if (e.target.closest('#rst2')) { const m = S.mode; S = DEF(); S.mode = m; renderHome(); }
  });
}

/* ---------- favorites, toast ---------- */
function toast(msg) {
  $$('.toast').forEach(x => x.remove());
  const d = document.createElement('div'); d.className = 'toast'; d.setAttribute('role', 'status'); d.textContent = msg; document.body.append(d); setTimeout(() => d.remove(), 2200);
}
function toggleFav(id) {
  fav.has(id) ? fav.delete(id) : fav.add(id);
  ls.set('fav', [...fav]); updateSavedBadge();
}
function updateSavedBadge() {
  $('#savedT').textContent = t('saved');
  const n = $('#savedN'); n.hidden = !fav.size; n.textContent = fav.size;
}
document.addEventListener('click', e => {
  const f = e.target.closest('[data-fav]');
  if (f) {
    e.preventDefault(); const id = +f.dataset.fav; toggleFav(id);
    if (S.saved && $('#res') && !location.hash.startsWith('#/p/')) { updateResults(); return; }
    const on = fav.has(id);
    $$(`[data-fav="${id}"]`).forEach(b => { b.classList.toggle('on', on); b.setAttribute('aria-pressed', on); const i = b.querySelector('i'); if (i) i.className = `ph${on ? '-fill' : ''} ph-heart`; const s = b.querySelector('span'); if (s) s.textContent = t(on ? 'unsave' : 'save'); });
    return;
  }
  const pk = e.target.closest('[data-pick]');
  if (pk) { S = Object.assign(DEF(), { region: pk.dataset.pick }); location.hash = '#/'; renderRoute(); }
});

/* ---------- detail ---------- */
function mediaBlock(kind, p) {
  if (kind === 'video') {
    const src = `https://drive.google.com/file/d/${encodeURIComponent(p.video)}/preview`;
    const bg = p.photos.length ? photoBig(p.photos[0]) : '';
    return `<div class="media" data-src="${esc(src)}"><button class="go" ${bg ? `style="background-image:url('${bg}')"` : ''}><span><i class="ph-fill ph-play"></i>${esc(t('watch'))}</span></button></div>`;
  }
  const u = p.tour || p.tourGroup;
  const bg = p.tourCover ? `photos/${p.tourCover}` : '';
  return `<div class="media" data-src="${safeUrl(u)}"><button class="go" ${bg ? `style="background-image:url('${bg}')"` : ''}><span><i class="ph ph-cube"></i>${esc(t('open3d'))}</span></button></div>
    <p class="meta"><a href="${safeUrl(u)}" target="_blank" rel="noopener">${esc(t('loadFail'))}</a></p>`;
}

function renderDetail(id) {
  const p = parcelById(id);
  if (!p) { location.hash = '#/'; return; }
  document.title = `${t('cadastre')} ${p.cadastre}, ${p.municipality} | ALLIANZ KAPITAL`;
  const ph = p.photos;
  let top;
  if (ph.length) {
    const shown = ph.slice(0, 5);
    top = `${''}<div class="mosaic ${shown.length < 3 ? 'one' : ''}" id="mos">${shown.map((i, k) => `<button data-i="${k}" aria-label="${esc(t('photos'))} ${k + 1}"><img src="${photoBig(i)}" alt="" ${k ? 'loading="lazy"' : 'fetchpriority="high"'}></button>`).join('')}<button class="allph" data-i="0"><i class="ph ph-images"></i>${esc(t('allPhotos', ph.length))}</button></div>`;
  } else if (p.tourCover) {
    top = `<div class="cover3d"><img src="photos/${esc(p.tourCover)}" alt=""><a class="btn" href="#m3d" id="to3d"><i class="ph ph-cube"></i>${esc(t('open3d'))}</a></div>`;
  } else top = `<div class="cover3d"><div class="noimg" style="height:100%"><i class="ph ph-image"></i>${esc(t('noPhoto'))}</div></div>`;

  const facts = [
    [t('cadastre'), p.cadastre], [t('type'), viewName(p.view)], [t('region'), regionName(p.region)], [t('muni'), p.municipality],
    [t('area'), p.area != null ? `${fmt(p.area)} ${t('m2')}${p.area >= 10000 ? ` (${fmt(p.area / 10000, 2)} ${t('ha')})` : ''}` : '-'],
    [t('status'), p.status === 'частично' ? t('statusP') : t('statusN')],
    [t('restr'), p.restricted ? t('restrY') : t('restrN')]];
  const plan = [[t('occ'), p.occupancy], [t('far'), p.far], [t('floors'), p.floors], [t('plan'), p.plan]].filter(([, v]) => v);
  const docs = [['idea', 'docIdea'], ['utu', 'docUtu'], ['deck', 'docDeck']].filter(([k]) => p.docs[k]);
  const c = CONTACT || {};
  const hasContact = c.email || c.whatsapp || c.phone || c.telegram;
  const price = p.total != null ? eur(p.total) : t('onReq');
  const sim = PARCELS.filter(x => x.id !== p.id && x.municipality === p.municipality && hasMedia(x)).slice(0, 4);
  const sim2 = sim.length ? sim : PARCELS.filter(x => x.id !== p.id && x.region === p.region && hasMedia(x)).slice(0, 4);
  const on = fav.has(p.id);

  $('#app').innerHTML = `<div class="wrap">
    <nav class="crumbs" aria-label="breadcrumb"><a href="#/"><i class="ph ph-arrow-left"></i> ${esc(t('back'))}</a><span>/</span><span>${esc(regionName(p.region))}</span><span>/</span><span>${esc(p.municipality)}</span></nav>
    ${top}
    ${p.shared > 1 && ph.length ? `<p class="meta" style="margin:10px 2px 0"><i class="ph ph-info"></i> ${esc(t('sharedPh', p.shared))}</p>` : ''}
    <div class="dgrid"><div>
      <h1 class="title">${p.area != null ? `${fmt(p.area)} ${esc(t('m2'))}, ` : ''}${esc(p.municipality)}</h1>
      <p class="meta">${esc(regionName(p.region))}, ${esc(viewName(p.view))}. ${esc(t('cadastre'))} ${esc(p.cadastre)}</p>
      <div class="panel" style="margin-top:16px"><h2>${esc(t('details'))}</h2><dl class="facts">${facts.map(([k, v]) => `<div><dt>${esc(k)}</dt><dd>${esc(v)}</dd></div>`).join('')}</dl></div>
      ${p.restricted ? `<div class="panel"><div class="warnbox"><i class="ph ph-warning"></i><span>${esc(t('restr'))}: ${esc(t('restrY'))}</span></div></div>` : ''}
      ${plan.length ? `<div class="panel"><h2>${esc(t('planning'))}</h2><dl class="facts">${plan.map(([k, v]) => `<div><dt>${esc(k)}</dt><dd>${esc(v)}</dd></div>`).join('')}</dl>${p.floors ? `<p class="meta" style="margin:12px 0 0">${esc(t('floorsHint'))}</p>` : ''}</div>` : ''}
      ${p.video ? `<div class="panel"><h2>${esc(t('video'))}</h2>${mediaBlock('video', p)}</div>` : ''}
      ${has3d(p) ? `<div class="panel" id="m3d"><h2>${esc(t('tour3d'))}</h2>${mediaBlock('3d', p)}</div>` : ''}
      ${docs.length ? `<div class="panel"><h2>${esc(t('docs'))}</h2><div class="doclinks">${docs.map(([k, l]) => `<a class="hbtn" href="${safeUrl(p.docs[k])}" target="_blank" rel="noopener"><i class="ph ph-file-text"></i>${esc(t(l))}</a>`).join('')}</div></div>` : ''}
    </div>
    <aside class="side"><div class="panel">
      <div class="price">${price}</div>
      ${p.ppm && p.total != null ? `<div class="meta">${eur(p.ppm)} ${esc(t('perM2'))}</div>` : ''}
      ${hasContact ? `<div class="form"><label>${esc(t('yourMsg'))}<textarea id="msg" rows="3">${esc(t('msgDef', p.cadastre))}</textarea></label>
        <div style="display:grid;gap:8px">
        ${c.whatsapp ? `<a class="btn block" data-ct="wa" href="#"><i class="ph ph-whatsapp-logo"></i>${esc(t('viaWa'))}</a>` : ''}
        ${c.telegram ? `<a class="btn ghost block" data-ct="tg" href="#"><i class="ph ph-telegram-logo"></i>${esc(t('viaTg'))}</a>` : ''}
        ${c.email ? `<a class="btn ghost block" data-ct="mail" href="#"><i class="ph ph-envelope-simple"></i>${esc(t('viaMail'))}</a>` : ''}
        ${c.phone ? `<a class="btn ghost block" href="tel:+${esc(String(c.phone).replace(/\D/g, ''))}"><i class="ph ph-phone"></i>${esc(t('call'))}</a>` : ''}</div></div>` : ''}
      <div class="row"><button class="sbtn ${on ? 'on' : ''}" data-fav="${p.id}" aria-pressed="${on}"><i class="ph${on ? '-fill' : ''} ph-heart"></i><span>${esc(t(on ? 'unsave' : 'save'))}</span></button>
      <button class="sbtn" id="share"><i class="ph ph-share-network"></i>${esc(t('share'))}</button></div>
    </div></aside></div>
    ${sim2.length ? `<section style="padding-bottom:56px"><h2>${esc(t('similar'))}</h2><div class="sim">${sim2.map(card).join('')}</div></section>` : ''}
  </div>`;

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
      : `mailto:${c.email}?subject=${encodeURIComponent(t('cadastre') + ' ' + p.cadastre)}&body=${m}`;
    a.target = '_blank'; a.rel = 'noopener';
  }));
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
const L_ = window.L;
function renderRoute() {
  const h = location.hash || '#/';
  const m = h.match(/^#\/p\/(\d+)/);
  updateSavedBadge();
  $('#disc').textContent = t('disc');
  $$('#lang button').forEach(b => b.classList.toggle('on', b.dataset.l === lang));
  document.documentElement.lang = L().html;
  if (m) { renderDetail(+m[1]); window.scrollTo(0, 0); return; }
  if (h === '#/saved') { S = Object.assign(DEF(), { saved: true }); }
  else if (S.saved && h === '#/') S.saved = false;
  renderHome();
  if (h === '#/saved') $('#catalog').scrollIntoView();
  else if (!(m)) window.scrollTo(0, 0);
}
window.addEventListener('hashchange', renderRoute);
$('#lang').onclick = e => { const b = e.target.closest('button'); if (!b) return; lang = b.dataset.l; ls.set('lang', lang); renderRoute(); };
renderRoute();
