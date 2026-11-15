'use strict';
/* Ma Cave — gestion de cave à vin, données dans OneDrive (Microsoft Graph). */

const CFG = window.CAVE_CONFIG || {};
const LOCAL_MODE = !CFG.clientId || new URLSearchParams(location.search).has('local');
const GRAPH = 'https://graph.microsoft.com/v1.0';
const SCOPES = ['Files.ReadWrite'];
const LS = { data: 'cave.data', etag: 'cave.etag', dirty: 'cave.dirty', ui: 'cave.ui' };
const SHEETJS = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';

const COULEURS = ['Rouge', 'Blanc', 'Rosé', 'Jaune', 'Effervescent', 'Liquoreux'];
const FORMATS = ['Bouteille', 'Demi-bouteille', 'Magnum', 'Jéroboam'];
const REGIONS = ['Alsace', 'Beaujolais', 'Bordeaux', 'Bourgogne', 'Champagne', 'Jura', 'Languedoc', 'Loire', 'Provence', 'Rhône', 'Savoie', 'Sud-Ouest', 'Étranger'];

let data = null;
let etag = null;
let dirty = false;
let rev = 0;
let ui = { tab: 'cave', cave: 'all', q: '', couleur: '', epuises: false, verif: false, incomplet: false, jtype: '' };
let msalApp = null;
let account = null;
let remoteChecked = false;
let syncState = 'idle';
let syncMsg = '';
let STOCK = new Map();

/* ---------- Utilitaires ---------- */
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const uid = p => p + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const nowIso = () => new Date().toISOString();
const YEAR = () => new Date().getFullYear();
function today() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}
function fmtDate(iso) {
  if (!iso) return '';
  const [y, m, d] = iso.slice(0, 10).split('-');
  return d && m ? `${d}/${m}/${y}` : y;
}
const EUR = new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 });
const EUR2 = new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR' });
const num = v => (v === '' || v == null || isNaN(Number(v))) ? null : Number(v);
const int = v => { const n = num(v); return n == null ? null : Math.round(n); };
const plural = (n, s, p) => `${n} ${n > 1 ? (p || s + 's') : s}`;
const couleurClass = c => 'c-' + (c || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
const norm = s => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');

function toast(msg) {
  const t = $('#toast');
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(toast._t);
  toast._t = setTimeout(() => t.classList.remove('show'), 2600);
}

/* ---------- Données ---------- */
function emptyData() {
  return { schema: 1, maj: nowIso(), caves: [{ id: 'vieux-conde', nom: 'Vieux-Condé', maj: nowIso() }, { id: 'ucel', nom: 'Ucel', maj: nowIso() }], vins: [], mouvements: [] };
}
function normalize(d) {
  if (!d || !Array.isArray(d.vins) || !Array.isArray(d.mouvements)) throw new Error('Fichier de cave invalide');
  d.caves = Array.isArray(d.caves) && d.caves.length ? d.caves : emptyData().caves;
  for (const v of d.vins) { v.emplacements = v.emplacements && typeof v.emplacements === 'object' ? v.emplacements : {}; }
  return d;
}
function loadLocal() {
  try {
    const raw = localStorage.getItem(LS.data);
    data = raw ? normalize(JSON.parse(raw)) : null;
    etag = localStorage.getItem(LS.etag) || null;
    dirty = localStorage.getItem(LS.dirty) === '1';
    const u = JSON.parse(localStorage.getItem(LS.ui) || 'null');
    if (u) ui = { ...ui, cave: u.cave || 'all', epuises: !!u.epuises, tab: u.tab || 'cave' };
  } catch (e) { data = null; }
}
function saveLocal() {
  try {
    if (data) localStorage.setItem(LS.data, JSON.stringify(data)); else localStorage.removeItem(LS.data);
    localStorage.setItem(LS.etag, etag || '');
    localStorage.setItem(LS.dirty, dirty ? '1' : '0');
  } catch (e) { toast('Impossible d’enregistrer sur cet appareil'); }
}
function saveUi() {
  try { localStorage.setItem(LS.ui, JSON.stringify({ cave: ui.cave, epuises: ui.epuises, tab: ui.tab })); } catch (e) { /* sans importance */ }
}
function commit(fn, msg) {
  fn();
  data.maj = nowIso();
  rev++;
  dirty = true;
  saveLocal();
  render();
  if (msg) toast(msg);
  schedulePush();
}

const vinsActifs = () => data.vins.filter(v => !v.supprime);
const mvtsActifs = () => data.mouvements.filter(m => !m.supprime);
const vinById = id => data.vins.find(v => v.id === id);
const caveNom = id => (data.caves.find(c => c.id === id) || {}).nom || id;

function computeStock() {
  const s = new Map();
  const add = (vin, cave, q) => {
    let o = s.get(vin);
    if (!o) { o = {}; s.set(vin, o); }
    o[cave] = (o[cave] || 0) + q;
  };
  for (const m of data.mouvements) {
    if (m.supprime) continue;
    const q = Number(m.qte) || 0;
    if (m.type === 'entree') add(m.vinId, m.cave, q);
    else if (m.type === 'sortie') add(m.vinId, m.cave, -q);
    else if (m.type === 'transfert') { add(m.vinId, m.cave, -q); add(m.vinId, m.caveDest, q); }
    else if (m.type === 'ajustement') add(m.vinId, m.cave, q);
  }
  STOCK = s;
}
function qte(v, cave = 'all') {
  const o = STOCK.get(v.id) || {};
  if (cave === 'all') return Object.values(o).reduce((a, b) => a + b, 0);
  return o[cave] || 0;
}
function prixMoyen(v) {
  let q = 0, t = 0;
  for (const m of data.mouvements) {
    if (m.supprime || m.vinId !== v.id || m.type !== 'entree' || m.prix == null) continue;
    q += m.qte; t += m.qte * m.prix;
  }
  return q ? t / q : null;
}
const valeurUnit = v => v.prixEstime ?? prixMoyen(v);

function statut(v) {
  const y = YEAR();
  if (v.gardeDebut && y < v.gardeDebut) return { k: 'jeune', label: `Attendre ${v.gardeDebut}` };
  if (v.gardeFin && y > v.gardeFin) return { k: 'depasse', label: `Limite ${v.gardeFin} dépassée` };
  if (v.gardeFin && y >= v.gardeFin - 1) return { k: 'urgent', label: `À boire avant fin ${v.gardeFin}` };
  if (v.gardeFin) return { k: 'pret', label: `Prêt · jusqu'à ${v.gardeFin}` };
  if (v.gardeDebut) return { k: 'pret', label: `Prêt depuis ${v.gardeDebut}` };
  return { k: 'inconnu', label: 'Garde à renseigner' };
}
// Informations utiles encore vides sur une fiche
function manquants(v) {
  const m = [];
  if (!v.appellation) m.push('appellation');
  if (!v.region) m.push('région');
  if (!v.cepages) m.push('cépages');
  if (!v.gardeDebut && !v.gardeFin) m.push('période de garde');
  if (!v.photo) m.push('photo');
  return m;
}
function titreVin(v) { return v.cuvee || v.appellation || v.domaine || 'Vin sans nom'; }
function sousTitre(v) {
  const parts = [v.domaine, v.appellation && v.appellation !== v.cuvee ? v.appellation : ''].filter(Boolean);
  return parts.join(' · ');
}
function nomComplet(v) { return [titreVin(v), v.millesime, v.format !== 'Bouteille' ? v.format : ''].filter(Boolean).join(' '); }
function emplacementTexte(v) {
  const e = v.emplacements || {};
  if (ui.cave !== 'all') return e[ui.cave] || '';
  return Object.entries(e).filter(([c, x]) => x && qte(v, c) > 0).map(([c, x]) => data.caves.length > 1 ? `${caveNom(c)} : ${x}` : x).join(' · ');
}
function uniques(champ) {
  return [...new Set(vinsActifs().map(v => v[champ]).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'fr'));
}

/* ---------- Authentification & OneDrive ---------- */
function redirectUri() {
  return CFG.redirectUri || (location.origin + location.pathname.replace(/index\.html$/, ''));
}
async function initAuth() {
  if (LOCAL_MODE) return;
  if (!window.msal) { setSync('offline', 'Connexion Microsoft indisponible (hors ligne ?)'); return; }
  msalApp = new window.msal.PublicClientApplication({
    auth: {
      clientId: CFG.clientId,
      authority: 'https://login.microsoftonline.com/' + (CFG.tenantId || 'organizations'),
      redirectUri: redirectUri()
    },
    cache: { cacheLocation: 'localStorage' }
  });
  await msalApp.initialize();
  try {
    const res = await msalApp.handleRedirectPromise();
    account = (res && res.account) || msalApp.getAllAccounts()[0] || null;
  } catch (e) {
    account = msalApp.getAllAccounts()[0] || null;
    setSync('error', 'Connexion : ' + e.message);
  }
  if (account) msalApp.setActiveAccount(account);
}
function login() { msalApp && msalApp.loginRedirect({ scopes: SCOPES }); }
function logout() {
  if (!confirm('Se déconnecter ? Les données restent dans OneDrive.')) return;
  localStorage.removeItem(LS.data); localStorage.removeItem(LS.etag); localStorage.removeItem(LS.dirty);
  msalApp.logoutRedirect({ account });
}
async function getToken() {
  try {
    const r = await msalApp.acquireTokenSilent({ scopes: SCOPES, account });
    return r.accessToken;
  } catch (e) {
    if (e instanceof window.msal.InteractionRequiredAuthError) {
      await msalApp.acquireTokenRedirect({ scopes: SCOPES, account });
    }
    throw e;
  }
}
async function gfetch(url, opts = {}) {
  const token = await getToken();
  return fetch(url.startsWith('http') ? url : GRAPH + url, {
    ...opts, headers: { ...(opts.headers || {}), Authorization: 'Bearer ' + token }
  });
}
const cheminEnc = p => p.split('/').map(encodeURIComponent).join('/');
const filePath = () => cheminEnc(CFG.cheminFichier || 'Cave/cave.json');
const dossier = () => (CFG.cheminFichier || 'Cave/cave.json').split('/').slice(0, -1).join('/');

async function fetchRemote() {
  const r = await gfetch(`/me/drive/root:/${filePath()}`);
  if (r.status === 404) return null;
  if (!r.ok) throw new Error('Lecture OneDrive (' + r.status + ')');
  const meta = await r.json();
  const c = await fetch(meta['@microsoft.graph.downloadUrl'], { cache: 'no-store' });
  if (!c.ok) throw new Error('Téléchargement OneDrive (' + c.status + ')');
  return { data: normalize(await c.json()), etag: meta.eTag };
}
async function ensureFolder(path) {
  if (!path) return;
  let parent = '';
  for (const name of path.split('/')) {
    const url = parent ? `/me/drive/root:/${cheminEnc(parent)}:/children` : '/me/drive/root/children';
    await gfetch(url, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' })
    });
    parent = parent ? parent + '/' + name : name;
  }
}
async function putFile(pathEnc, body, contentType, ifMatch) {
  const h = { 'Content-Type': contentType };
  if (ifMatch) h['If-Match'] = ifMatch;
  let r = await gfetch(`/me/drive/root:/${pathEnc}:/content`, { method: 'PUT', headers: h, body });
  if (r.status === 404 || r.status === 409) {
    await ensureFolder(decodeURIComponent(pathEnc).split('/').slice(0, -1).join('/'));
    r = await gfetch(`/me/drive/root:/${pathEnc}:/content`, { method: 'PUT', headers: h, body });
  }
  if (r.status === 412) return { conflict: true };
  if (!r.ok) throw new Error('Écriture OneDrive (' + r.status + ')');
  const j = await r.json();
  return { etag: j.eTag };
}

/* ---------- Photos des bouteilles ----------
   v.photo peut valoir :
   - 'od:<chemin>'  photo prise avec l'appli, stockée dans OneDrive (dossier Cave/photos)
   - 'https://…'    image trouvée sur le web
   - 'data:…'       photo réduite gardée dans le fichier (mode local ou envoi impossible)
   v.photoMini : même chose en petit format pour la liste. */
const PHOTO_CACHE = 'cave-photos';
const photoUrls = new Map();
const photoAttente = new Map();
const dossierPhotos = () => (dossier() ? dossier() + '/' : '') + 'photos';
const cleCache = chemin => new Request(location.origin + location.pathname.replace(/[^/]*$/, '') + '__photo/' + encodeURIComponent(chemin));

async function urlPhoto(ref) {
  if (!ref) return '';
  if (!ref.startsWith('od:')) return ref;
  if (photoUrls.has(ref)) return photoUrls.get(ref);
  if (photoAttente.has(ref)) return photoAttente.get(ref);
  const p = (async () => {
    const chemin = ref.slice(3);
    let blob = null;
    try {
      const c = await caches.open(PHOTO_CACHE);
      const r = await c.match(cleCache(chemin));
      if (r) blob = await r.blob();
    } catch (e) { /* cache indisponible */ }
    if (!blob) {
      if (LOCAL_MODE || !account || !navigator.onLine) return '';
      const r = await gfetch(`/me/drive/root:/${cheminEnc(chemin)}`);
      if (!r.ok) return '';
      const meta = await r.json();
      const f = await fetch(meta['@microsoft.graph.downloadUrl']);
      if (!f.ok) return '';
      blob = await f.blob();
      try { (await caches.open(PHOTO_CACHE)).put(cleCache(chemin), new Response(blob, { headers: { 'Content-Type': blob.type || 'image/jpeg' } })); } catch (e) { /* sans importance */ }
    }
    const u = URL.createObjectURL(blob);
    photoUrls.set(ref, u);
    return u;
  })().catch(() => '').finally(() => photoAttente.delete(ref));
  photoAttente.set(ref, p);
  return p;
}
// Charge les images (balises <img data-photo>) quand elles arrivent à l'écran
const photoObs = 'IntersectionObserver' in window ? new IntersectionObserver(entries => {
  for (const e of entries) if (e.isIntersecting) { photoObs.unobserve(e.target); chargerImg(e.target); }
}, { rootMargin: '300px' }) : null;
async function chargerImg(img) {
  const u = await urlPhoto(img.dataset.photo);
  if (u) { img.src = u; img.closest('.thumb, .photo-wrap')?.classList.add('ok'); }
  else img.closest('.thumb, .photo-wrap')?.classList.add('ko');
}
function hydraterPhotos(root = document) {
  $$('img[data-photo]:not([data-vu])', root).forEach(img => {
    img.dataset.vu = '1';
    if (photoObs) photoObs.observe(img); else chargerImg(img);
  });
}
function imageDepuisFichier(file) {
  return new Promise((res, rej) => {
    const img = new Image();
    const u = URL.createObjectURL(file);
    img.onload = () => { URL.revokeObjectURL(u); res(img); };
    img.onerror = () => { URL.revokeObjectURL(u); rej(new Error('Image illisible')); };
    img.src = u;
  });
}
function reduire(img, maxCote, qualite) {
  const k = Math.min(1, maxCote / Math.max(img.naturalWidth, img.naturalHeight));
  const c = document.createElement('canvas');
  c.width = Math.round(img.naturalWidth * k); c.height = Math.round(img.naturalHeight * k);
  c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
  return new Promise(res => c.toBlob(b => res({ blob: b, dataUrl: () => c.toDataURL('image/jpeg', qualite) }), 'image/jpeg', qualite));
}
async function enregistrerPhoto(v, file) {
  const img = await imageDepuisFichier(file);
  const grand = await reduire(img, 1400, 0.82);
  const mini = await reduire(img, 220, 0.7);
  if (!LOCAL_MODE && account && navigator.onLine) {
    try {
      const base = `${dossierPhotos()}/${v.id}-${Date.now().toString(36)}`;
      await putFile(cheminEnc(base + '.jpg'), grand.blob, 'image/jpeg', null);
      await putFile(cheminEnc(base + '-mini.jpg'), mini.blob, 'image/jpeg', null);
      try {
        const c = await caches.open(PHOTO_CACHE);
        await c.put(cleCache(base + '.jpg'), new Response(grand.blob, { headers: { 'Content-Type': 'image/jpeg' } }));
        await c.put(cleCache(base + '-mini.jpg'), new Response(mini.blob, { headers: { 'Content-Type': 'image/jpeg' } }));
      } catch (e) { /* sans importance */ }
      return { photo: 'od:' + base + '.jpg', photoMini: 'od:' + base + '-mini.jpg' };
    } catch (e) {
      console.error(e);
      toast('Envoi OneDrive impossible : photo gardée en petit format');
    }
  }
  // Mode local ou hors ligne : version réduite stockée dans le fichier de la cave
  const moyen = await reduire(img, 640, 0.72);
  return { photo: moyen.dataUrl(), photoMini: mini.dataUrl() };
}
function choisirPhoto(id) {
  const v = vinById(id);
  const inp = document.createElement('input');
  inp.type = 'file'; inp.accept = 'image/*';
  inp.onchange = async () => {
    const f = inp.files[0];
    if (!f) return;
    toast('Enregistrement de la photo…');
    try {
      const r = await enregistrerPhoto(v, f);
      commit(() => { Object.assign(v, r); v.maj = nowIso(); }, 'Photo enregistrée');
      ouvrirVin(v.id);
    } catch (e) { toast('Photo impossible : ' + e.message); }
  };
  inp.click();
}
function lienPhotoWeb(id) {
  const v = vinById(id);
  const actuel = v.photo && /^https?:/.test(v.photo) ? v.photo : '';
  const u = prompt('Collez le lien d’une image de la bouteille (clic long sur l’image → « Copier l’adresse de l’image »)', actuel);
  if (u == null) return;
  const x = u.trim();
  if (x && !/^https?:\/\/\S+$/.test(x)) { toast('Lien invalide'); return; }
  commit(() => { v.photo = x; v.photoMini = x; v.maj = nowIso(); }, x ? 'Image enregistrée' : 'Image retirée');
  ouvrirVin(v.id);
}
const rechercheImage = v => 'https://www.google.com/search?tbm=isch&q=' + encodeURIComponent([v.domaine, v.cuvee, v.appellation, v.millesime, 'bouteille'].filter(Boolean).join(' '));

function mergeById(a = [], b = []) {
  const m = new Map();
  for (const x of [...b, ...a]) {
    const p = m.get(x.id);
    if (!p || (x.maj || '') >= (p.maj || '')) m.set(x.id, x);
  }
  return [...m.values()];
}
function merge(local, remote) {
  return {
    ...remote, ...local,
    caves: mergeById(local.caves, remote.caves),
    vins: mergeById(local.vins, remote.vins),
    mouvements: mergeById(local.mouvements, remote.mouvements),
    maj: nowIso()
  };
}

let syncing = false, again = false, pushTimer = null;
let remplacer = false; // import : écraser le fichier OneDrive au lieu de fusionner
function schedulePush() {
  if (LOCAL_MODE) return;
  clearTimeout(pushTimer);
  pushTimer = setTimeout(sync, 800);
}
function setSync(state, msg = '') {
  syncState = state; syncMsg = msg;
  const b = $('#syncBtn');
  if (!b) return;
  const labels = { ok: 'Synchronisé', syncing: 'Synchro…', offline: 'Hors ligne', error: 'Erreur', pending: 'À envoyer', local: 'Mode local', idle: '' };
  b.dataset.state = state;
  $('#syncLabel').textContent = labels[state] || '';
  b.title = msg || labels[state] || '';
}
async function sync() {
  if (LOCAL_MODE) { setSync('local', 'Données enregistrées uniquement dans ce navigateur'); return; }
  if (!account) { setSync('idle'); return; }
  if (syncing) { again = true; return; }
  syncing = true;
  setSync('syncing');
  const rev0 = rev;
  try {
    if (!navigator.onLine) throw Object.assign(new Error('Pas de réseau : les changements seront envoyés plus tard'), { offline: true });
    let remote = await fetchRemote();
    remoteChecked = true;
    if (!remote) {
      if (data) {
        const r = await putFile(filePath(), JSON.stringify(data), 'application/json', null);
        etag = r.etag;
        if (rev === rev0) dirty = false; else again = true;
      }
    } else if (!dirty) {
      if (remote.etag !== etag || !data) { data = remote.data; etag = remote.etag; }
    } else {
      for (let i = 0; i < 3; i++) {
        const merged = (remplacer || remote.etag === etag) ? data : merge(data, remote.data);
        const r = await putFile(filePath(), JSON.stringify(merged), 'application/json', remote.etag);
        if (!r.conflict) {
          etag = r.etag;
          remplacer = false;
          if (rev === rev0) { data = merged; dirty = false; } else { data = merge(data, merged); again = true; }
          break;
        }
        remote = await fetchRemote();
        if (!remote) break;
      }
    }
    saveLocal();
    setSync(dirty ? 'pending' : 'ok', dirty ? 'Changements en attente d’envoi' : 'Données à jour dans OneDrive');
  } catch (e) {
    console.error(e);
    setSync(e.offline || !navigator.onLine ? (dirty ? 'pending' : 'offline') : 'error', e.message);
  } finally {
    syncing = false;
    render();
    if (again) { again = false; schedulePush(); }
  }
}

/* ---------- Rendu général ---------- */
function render() {
  const content = $('#content');
  renderCaveFilter();
  $$('.tabs button').forEach(b => b.setAttribute('aria-current', b.dataset.tab === ui.tab ? 'page' : 'false'));
  $('#fab').hidden = !(data && ui.tab === 'cave');
  $('#fabBoire').hidden = !(data && (ui.tab === 'cave' || ui.tab === 'boire'));
  if (!data) { content.innerHTML = renderWelcome(); return; }
  computeStock();
  if (ui.tab === 'cave') renderCave(content);
  else if (ui.tab === 'boire') content.innerHTML = renderBoire();
  else if (ui.tab === 'journal') content.innerHTML = renderJournal();
  else content.innerHTML = renderBilan();
  hydraterPhotos(content);
}
function renderCaveFilter() {
  const el = $('#caveFilter');
  if (!data) { el.innerHTML = ''; return; }
  computeStock();
  const total = c => vinsActifs().reduce((a, v) => a + qte(v, c), 0);
  const opts = [{ id: 'all', nom: 'Toutes' }, ...data.caves];
  el.innerHTML = opts.map(c => `<button type="button" role="tab" data-cave="${esc(c.id)}" aria-selected="${ui.cave === c.id}">${esc(c.nom)}<small>${total(c.id)}</small></button>`).join('');
}
function renderWelcome() {
  if (!LOCAL_MODE && !account) {
    return `<div class="welcome"><div style="font-size:44px">🍷</div><h2>Bienvenue dans votre cave</h2>
      <p>Connectez-vous avec votre compte Microsoft pour retrouver vos bouteilles enregistrées dans OneDrive.</p>
      <div class="row"><button class="btn primary" data-act="login">Se connecter avec Microsoft</button></div></div>`;
  }
  if (!LOCAL_MODE && !remoteChecked) {
    return `<div class="welcome"><p>${syncState === 'syncing' ? 'Chargement depuis OneDrive…' : esc(syncMsg || 'Impossible de joindre OneDrive pour le moment.')}</p>
      <div class="row"><button class="btn" data-act="sync">Réessayer</button></div></div>`;
  }
  return `<div class="welcome"><div style="font-size:44px">🍷</div><h2>Aucune cave pour l'instant</h2>
    <p>${LOCAL_MODE ? 'Mode local : les données restent dans ce navigateur.' : 'Aucun fichier <b>' + esc(CFG.cheminFichier || 'Cave/cave.json') + '</b> dans votre OneDrive.'}
    Importez le fichier <b>cave.json</b> préparé à partir de votre livre de cave, ou partez de zéro.</p>
    <div class="row"><button class="btn primary" data-act="import">Importer cave.json</button>
    <button class="btn" data-act="vide">Commencer une cave vide</button></div></div>`;
}

/* ---------- Onglet Cave ---------- */
function renderCave(content) {
  if (!$('#caveToolbar', content)) {
    content.innerHTML = `<div id="caveToolbar"><div class="search"><input id="q" type="search" placeholder="Rechercher : domaine, cuvée, millésime, casier…" autocomplete="off"></div>
      <div class="chips" id="chips"></div></div><div id="list"></div>`;
    $('#q').value = ui.q;
  }
  const verif = vinsActifs().filter(v => v.aVerifier).length;
  const incompl = vinsActifs().filter(v => qte(v, ui.cave) > 0 && manquants(v).length).length;
  $('#chips').innerHTML = ['', ...COULEURS].filter(c => !c || vinsActifs().some(v => v.couleur === c)).map(c =>
    `<button type="button" class="chip" data-couleur="${esc(c)}" aria-pressed="${ui.couleur === c}">${c || 'Toutes couleurs'}</button>`).join('') +
    `<button type="button" class="chip" data-toggle="epuises" aria-pressed="${ui.epuises}">Épuisés</button>` +
    (verif ? `<button type="button" class="chip" data-toggle="verif" aria-pressed="${ui.verif}">À vérifier (${verif})</button>` : '') +
    (incompl ? `<button type="button" class="chip" data-toggle="incomplet" aria-pressed="${ui.incomplet}">À compléter (${incompl})</button>` : '');
  renderList();
}
function filtreVins() {
  const q = norm(ui.q).trim();
  const words = q ? q.split(/\s+/) : [];
  return vinsActifs().filter(v => {
    const n = qte(v, ui.cave);
    if (!ui.epuises && n <= 0) return false;
    if (ui.couleur && v.couleur !== ui.couleur) return false;
    if (ui.verif && !v.aVerifier) return false;
    if (ui.incomplet && !manquants(v).length) return false;
    if (words.length) {
      const hay = norm([v.domaine, v.cuvee, v.appellation, v.region, v.millesime, v.cepages, v.couleur, v.format, ...Object.values(v.emplacements || {})].join(' '));
      if (!words.every(w => hay.includes(w))) return false;
    }
    return true;
  });
}
function trierVins(list) {
  return list.sort((a, b) =>
    (a.region || '').localeCompare(b.region || '', 'fr') ||
    (a.domaine || '').localeCompare(b.domaine || '', 'fr') ||
    titreVin(a).localeCompare(titreVin(b), 'fr') ||
    (a.millesime || 0) - (b.millesime || 0));
}
function renderList() {
  const list = trierVins(filtreVins());
  const total = list.reduce((a, v) => a + qte(v, ui.cave), 0);
  let html = `<div class="summary">${plural(total, 'bouteille')} · ${plural(list.length, 'vin')}${ui.cave !== 'all' ? ' · ' + esc(caveNom(ui.cave)) : ''}</div>`;
  if (!list.length) {
    html += `<div class="empty-state">${ui.q || ui.couleur || ui.verif || ui.incomplet ? 'Aucun vin ne correspond.' : ui.cave !== 'all' ? 'Aucune bouteille dans cette cave pour le moment.' : 'Aucun vin. Ajoutez-en avec le bouton +.'}</div>`;
  } else {
    let region = null;
    for (const v of list) {
      if (v.region !== region) {
        region = v.region;
        const n = list.filter(x => x.region === region).reduce((a, x) => a + qte(x, ui.cave), 0);
        html += `<div class="group-title"><span>${esc(region || 'Sans région')}</span><span>${n}</span></div>`;
      }
      html += carteVin(v);
    }
  }
  $('#list').innerHTML = html;
  hydraterPhotos($('#list'));
}
function carteVin(v) {
  const n = qte(v, ui.cave);
  const st = statut(v);
  const loc = emplacementTexte(v);
  const mini = v.photoMini || v.photo;
  return `<div class="vin-wrap"><button type="button" class="vin${n <= 0 ? ' empty' : ''}${n > 0 ? ' has-quick' : ''}" data-vin="${esc(v.id)}">
    ${mini ? `<span class="thumb ${couleurClass(v.couleur)}"><img data-photo="${esc(mini)}" alt="" decoding="async"></span>` : `<span class="dot ${couleurClass(v.couleur)}" title="${esc(v.couleur)}"></span>`}
    <span class="vin-main"><span class="vin-title">${esc(titreVin(v))}</span>
      <span class="vin-sub" style="display:block">${esc(sousTitre(v))}${v.format && v.format !== 'Bouteille' ? ' · ' + esc(v.format) : ''}</span>
      <span class="vin-meta"><span class="pill s-${st.k}">${esc(st.label)}</span>${v.aVerifier ? '<span class="pill verif">À vérifier</span>' : ''}${loc ? `<span class="loc">📍 ${esc(loc)}</span>` : ''}</span></span>
    <span class="vin-right"><span class="vin-year" style="display:block">${esc(v.millesime || 'NM')}</span><span class="vin-qty">${n}<small>btl</small></span></span>
  </button>${n > 0 ? `<button type="button" class="quick-boire" data-act="boire" data-id="${esc(v.id)}" aria-label="Boire une bouteille de ${esc(nomComplet(v))}">🍷 Bu</button>` : ''}</div>`;
}

/* ---------- Onglet À boire ---------- */
function renderBoire() {
  const enStock = vinsActifs().filter(v => qte(v, ui.cave) > 0);
  const byFin = (a, b) => (a.gardeFin || 0) - (b.gardeFin || 0) || titreVin(a).localeCompare(titreVin(b), 'fr');
  const depasse = enStock.filter(v => statut(v).k === 'depasse').sort(byFin);
  const urgent = enStock.filter(v => statut(v).k === 'urgent').sort(byFin);
  const inconnu = enStock.filter(v => statut(v).k === 'inconnu');
  const somme = l => l.reduce((a, v) => a + qte(v, ui.cave), 0);
  let html = `<div class="kpis"><div class="kpi"><b style="color:var(--bad)">${somme(depasse)}</b><span>bouteilles ont dépassé leur limite</span></div>
    <div class="kpi"><b style="color:var(--warn)">${somme(urgent)}</b><span>à boire d'ici fin ${YEAR() + 1}</span></div></div>`;
  if (depasse.length) html += `<div class="group-title"><span>Limite dépassée</span><span>${somme(depasse)}</span></div>` + depasse.map(carteVin).join('');
  if (urgent.length) html += `<div class="group-title"><span>À boire bientôt</span><span>${somme(urgent)}</span></div>` + urgent.map(carteVin).join('');
  if (!depasse.length && !urgent.length) html += `<div class="empty-state">Rien d'urgent : tous les vins renseignés ont encore le temps.</div>`;
  if (inconnu.length) {
    html += `<div class="card" style="margin-top:16px"><h2>Garde non renseignée</h2><p class="muted small" style="margin:0 0 8px">${plural(inconnu.length, 'vin')} (${plural(somme(inconnu), 'bouteille')}) n'ont pas de dates de garde : ils n'apparaissent pas ici.</p>
      <button class="btn link" data-act="voir-inconnus">Voir ces vins</button></div>`;
  }
  return html;
}

/* ---------- Onglet Journal ---------- */
const TYPES = { entree: ['+', 'Entrée'], sortie: ['−', 'Sortie'], transfert: ['⇄', 'Déplacement'], ajustement: ['±', 'Correction'], commentaire: ['✎', 'Commentaire'] };
function renderJournal() {
  let list = mvtsActifs().filter(m => !ui.jtype || m.type === ui.jtype);
  if (ui.cave !== 'all') list = list.filter(m => m.cave === ui.cave || m.caveDest === ui.cave);
  list.sort((a, b) => (b.date || '0').localeCompare(a.date || '0') || (b.maj || '').localeCompare(a.maj || ''));
  let html = `<div class="chips" style="padding-top:0">${[['', 'Tout'], ['sortie', 'Sorties'], ['commentaire', 'Commentaires'], ['entree', 'Entrées'], ['transfert', 'Déplacements'], ['ajustement', 'Corrections']]
    .map(([k, l]) => `<button type="button" class="chip" data-jtype="${k}" aria-pressed="${ui.jtype === k}">${l}</button>`).join('')}</div>`;
  if (!list.length) return html + '<div class="empty-state">Aucun mouvement.</div>';
  let mois = null;
  for (const m of list.slice(0, 400)) {
    const k = m.date ? new Date(m.date.slice(0, 7) + '-01T12:00:00').toLocaleDateString('fr-FR', { month: 'long', year: 'numeric' }) : 'Date inconnue';
    if (k !== mois) { mois = k; html += `<div class="group-title"><span>${esc(k)}</span></div>`; }
    const v = vinById(m.vinId) || {};
    const [ico] = TYPES[m.type] || ['?'];
    if (m.type === 'commentaire') {
      html += `<button type="button" class="mvt" data-mvt="${esc(m.id)}"><span class="ico commentaire">✎</span>
        <span class="mvt-main"><span class="t" style="display:block">${esc(nomComplet(v))}</span><span class="s" style="display:block">${esc([m.noteSur20 != null ? m.noteSur20 + '/20' : '', m.occasion, m.commentaire].filter(Boolean).join(' · '))}</span></span>
        <span class="d">${esc(fmtDate(m.date))}</span></button>`;
      continue;
    }
    const signe = m.type === 'sortie' ? '−' : m.type === 'entree' ? '+' : m.type === 'ajustement' && m.qte > 0 ? '+' : '';
    const lieu = m.type === 'transfert' ? `${caveNom(m.cave)} → ${caveNom(m.caveDest)}` : caveNom(m.cave);
    const details = [lieu, m.occasion, m.noteSur20 != null ? m.noteSur20 + '/20' : '', m.prix != null ? EUR2.format(m.prix) + '/btl' : '', m.fournisseur, m.note && m.note !== 'Reprise du livre de cave Excel' ? m.note : ''].filter(Boolean).join(' · ');
    html += `<button type="button" class="mvt" data-mvt="${esc(m.id)}"><span class="ico ${m.type}">${ico}</span>
      <span class="mvt-main"><span class="t" style="display:block">${signe}${Math.abs(m.qte)} · ${esc(nomComplet(v))}</span><span class="s" style="display:block">${esc(details)}</span></span>
      <span class="d">${esc(fmtDate(m.date))}</span></button>`;
  }
  return html;
}

/* ---------- Onglet Bilan ---------- */
function bars(rows) {
  const max = Math.max(1, ...rows.map(r => r[1]));
  return `<div class="bars">${rows.map(([l, n]) => `<div class="bar-row"><span>${esc(l)}</span><span class="track"><span class="fill" style="display:block;width:${(n / max * 100).toFixed(1)}%"></span></span><span class="n">${n}</span></div>`).join('')}</div>`;
}
function renderBilan() {
  const vins = vinsActifs();
  const c = ui.cave;
  const total = vins.reduce((a, v) => a + qte(v, c), 0);
  const refs = vins.filter(v => qte(v, c) > 0).length;
  let valeur = 0, qValo = 0;
  for (const v of vins) { const n = qte(v, c); const p = valeurUnit(v); if (n > 0 && p != null) { valeur += n * p; qValo += n; } }
  const y = String(YEAR());
  const bus = mvtsActifs().filter(m => m.type === 'sortie' && (m.date || '').startsWith(y) && (c === 'all' || m.cave === c)).reduce((a, m) => a + m.qte, 0);
  const grp = champ => {
    const m = new Map();
    for (const v of vins) { const n = qte(v, c); if (n > 0) m.set(v[champ] || '—', (m.get(v[champ] || '—') || 0) + n); }
    return [...m.entries()].sort((a, b) => b[1] - a[1]);
  };
  const verif = vins.filter(v => v.aVerifier).length;
  let html = `<div class="kpis">
    <div class="kpi"><b>${total}</b><span>bouteilles · ${plural(refs, 'vin')}</span></div>
    <div class="kpi"><b>${qValo ? EUR.format(valeur) : '—'}</b><span>${qValo ? `valeur de ${qValo} btl dont le prix est connu` : 'valeur : aucun prix renseigné'}</span></div>
    <div class="kpi"><b>${bus}</b><span>bouteilles bues en ${y}</span></div>
    <div class="kpi"><b>${vins.filter(v => qte(v, c) > 0 && ['depasse', 'urgent'].includes(statut(v).k)).reduce((a, v) => a + qte(v, c), 0)}</b><span>à boire en priorité</span></div></div>`;
  if (c === 'all' && data.caves.length > 1) html += `<div class="card"><h2>Par cave</h2>${bars(data.caves.map(cv => [cv.nom, vins.reduce((a, v) => a + qte(v, cv.id), 0)]))}</div>`;
  html += `<div class="card"><h2>Par région</h2>${bars(grp('region'))}</div>`;
  html += `<div class="card"><h2>Par couleur</h2>${bars(grp('couleur'))}</div>`;
  if (verif) html += `<div class="card"><h2>Fiches à compléter</h2><p class="muted small" style="margin:0 0 6px">${plural(verif, 'fiche')} signalée${verif > 1 ? 's' : ''} lors de la reprise du fichier Excel (millésime manquant, appellation corrigée…).</p><button class="btn link" data-act="voir-verif">Les passer en revue</button></div>`;
  html += `<div class="card"><h2>Caves</h2><ul class="list-plain">${data.caves.map(cv => `<li><span>${esc(cv.nom)}</span><button class="btn link" data-act="renommer-cave" data-id="${esc(cv.id)}">Renommer</button></li>`).join('')}</ul>
    <button class="btn" data-act="ajouter-cave" style="margin-top:8px">Ajouter une cave</button></div>`;
  const th = themeActuel();
  html += `<div class="card"><h2>Apparence</h2><div class="theme-pick">${[['auto', 'Automatique'], ['etiquette', 'Étiquette'], ['cave', 'Cave']].map(([k, l]) =>
    `<button type="button" class="chip" data-act="theme" data-id="${k}" aria-pressed="${th === k}">${l}</button>`).join('')}</div>
    <p class="muted small" style="margin:8px 0 0">Automatique : « Étiquette » le jour, « Cave » quand le téléphone est en mode sombre.</p></div>`;
  html += `<div class="card"><h2>Exports et sauvegardes</h2><div class="row">
    <button class="btn" data-act="export-xlsx">Télécharger en Excel</button>
    ${!LOCAL_MODE && account ? '<button class="btn" data-act="export-xlsx-od">Excel dans OneDrive</button>' : ''}
    <button class="btn" data-act="export-json">Sauvegarde JSON</button>
    <button class="btn" data-act="import">Importer un fichier</button></div></div>`;
  html += `<div class="card"><h2>Compte</h2>`;
  if (LOCAL_MODE) html += `<p class="muted small" style="margin:0">Mode local : les données sont enregistrées uniquement dans ce navigateur. Renseignez <code>config.js</code> pour utiliser OneDrive.</p>`;
  else html += `<p class="small" style="margin:0 0 8px">${account ? 'Connecté : ' + esc(account.username) : 'Non connecté'}<br><span class="muted">Fichier : OneDrive / ${esc(CFG.cheminFichier || 'Cave/cave.json')}</span><br><span class="muted">${esc(syncMsg)}</span></p>
      <div class="row">${account ? '<button class="btn" data-act="sync">Synchroniser</button><button class="btn danger" data-act="logout">Se déconnecter</button>' : '<button class="btn primary" data-act="login">Se connecter</button>'}</div>`;
  html += `</div>`;
  return html;
}

/* ---------- Thème ---------- */
function themeActuel() { return document.documentElement.dataset.theme || 'auto'; }
function appliquerTheme(t) {
  if (t === 'auto') delete document.documentElement.dataset.theme; else document.documentElement.dataset.theme = t;
  try { if (t === 'auto') localStorage.removeItem('cave.theme'); else localStorage.setItem('cave.theme', t); } catch (e) { /* sans importance */ }
  majCouleurBarre();
}
function majCouleurBarre() {
  const m = document.querySelector('meta[name="theme-color"]');
  if (m) m.content = getComputedStyle(document.documentElement).getPropertyValue('--bg').trim() || '#f5efe4';
}
matchMedia('(prefers-color-scheme: dark)').addEventListener?.('change', majCouleurBarre);

/* ---------- Feuilles (dialogues) ---------- */
function openSheet(html, onSubmit, vinId = '') {
  const dlg = $('#sheet');
  dlg.dataset.vin = vinId;
  $('#sheetBody').innerHTML = html;
  dlg._onSubmit = onSubmit || null;
  if (!dlg.open) dlg.showModal();
  const f = $('#sheetBody form');
  if (f) {
    f.addEventListener('submit', e => {
      e.preventDefault();
      const fd = Object.fromEntries(new FormData(f).entries());
      const ok = dlg._onSubmit && dlg._onSubmit(fd, f);
      if (ok !== false) closeSheet();
    });
  }
  $('#sheetBody').scrollTop = 0;
}
function closeSheet() { const d = $('#sheet'); if (d.open) d.close(); }
const head = (titre, sub = '') => `<div class="sheet-head"><div><h2>${esc(titre)}</h2>${sub ? `<div class="sub">${esc(sub)}</div>` : ''}</div><button type="button" class="close" data-act="close" aria-label="Fermer">×</button></div>`;
const options = (list, sel) => list.map(o => { const [val, lab] = Array.isArray(o) ? o : [o, o]; return `<option value="${esc(val)}"${String(val) === String(sel ?? '') ? ' selected' : ''}>${esc(lab)}</option>`; }).join('');
const datalist = (id, list) => `<datalist id="${id}">${list.map(x => `<option value="${esc(x)}">`).join('')}</datalist>`;
const stepper = (name, val, min, max) => `<div class="stepper"><button type="button" data-step="-1" aria-label="Moins">−</button><input name="${name}" type="number" inputmode="numeric" value="${val}" min="${min}" ${max != null ? `max="${max}"` : ''} required><button type="button" data-step="1" aria-label="Plus">+</button></div>`;

function ouvrirVin(id) {
  const v = vinById(id);
  if (!v) return;
  computeStock();
  const st = statut(v);
  const total = qte(v);
  const pm = prixMoyen(v);
  const tous = mvtsActifs().filter(m => m.vinId === v.id).sort((a, b) => (b.date || '0').localeCompare(a.date || '0') || (b.maj || '').localeCompare(a.maj || ''));
  const mv = tous.filter(m => m.type !== 'commentaire');
  const degust = tous.filter(m => m.type === 'commentaire' || (m.type === 'sortie' && (m.commentaire || m.noteSur20 != null || m.occasion)));
  const nbBues = tous.filter(m => m.type === 'sortie').reduce((a, m) => a + m.qte, 0);
  const manque = manquants(v);
  const lignes = data.caves.filter(c => qte(v, c.id) !== 0 || (v.emplacements || {})[c.id]).map(c =>
    `<tr><td>${esc(c.nom)}<div class="muted small">${v.emplacements[c.id] ? '📍 ' + esc(v.emplacements[c.id]) : 'Emplacement non renseigné'}</div></td><td class="q">${qte(v, c.id)}</td></tr>`).join('');
  openSheet(`${head(nomComplet(v), [v.domaine, v.appellation, v.region].filter(Boolean).join(' · '))}
    ${v.photo ? `<div class="photo-wrap ${couleurClass(v.couleur)}"><img data-photo="${esc(v.photo)}" alt="Bouteille ${esc(nomComplet(v))}"></div>` : ''}
    <div class="row photo-actions">
      <button class="btn small-btn" data-act="photo" data-id="${v.id}">📷 ${v.photo ? 'Changer la photo' : 'Prendre / choisir une photo'}</button>
      <button class="btn small-btn" data-act="photo-web" data-id="${v.id}">🔗 Image du web</button>
      ${v.photo ? `<button class="btn small-btn danger" data-act="photo-suppr" data-id="${v.id}">Retirer</button>` : ''}
    </div>
    ${!v.photo || (v.photo && /^https?:/.test(v.photo)) ? `<p class="muted small" style="margin:-4px 0 10px">Pour une image du web : <a href="${esc(rechercheImage(v))}" target="_blank" rel="noopener">chercher la bouteille</a>, copier l'adresse de l'image puis « Image du web ».</p>` : ''}
    ${v.aVerifier ? `<div class="banner">⚠️ ${esc(v.aVerifier)} <button class="btn link" data-act="verifie" data-id="${v.id}">C'est vérifié</button></div>` : ''}
    <div class="vin-meta" style="margin-bottom:12px"><span class="pill s-${st.k}">${esc(st.label)}</span><span class="pill">${esc(v.couleur || 'Couleur ?')}</span>${v.format !== 'Bouteille' ? `<span class="pill">${esc(v.format)}</span>` : ''}</div>
    <div class="row">
      <button class="btn primary" data-act="boire" data-id="${v.id}" ${total <= 0 ? 'disabled' : ''}>🍷 Boire</button>
      <button class="btn" data-act="entree" data-id="${v.id}">+ Ajouter</button>
      ${data.caves.length > 1 ? `<button class="btn" data-act="transfert" data-id="${v.id}" ${total <= 0 ? 'disabled' : ''}>⇄ Déplacer</button>` : ''}
    </div>
    <div class="section-title">Dégustations & commentaires${nbBues ? ` · ${plural(nbBues, 'bouteille bue', 'bouteilles bues')}` : ''}</div>
    ${degust.length ? degust.map(carteDegust).join('') : '<p class="muted small">Pas encore de commentaire. Ils s’ajoutent quand vous buvez une bouteille, ou avec le bouton ci-dessous.</p>'}
    <div class="row" style="margin-top:6px"><button class="btn" data-act="commentaire" data-id="${v.id}">✎ Ajouter un commentaire</button></div>
    <div class="section-title">Stock · ${plural(total, 'bouteille')}</div>
    ${lignes ? `<table class="stock-table">${lignes}</table>` : '<p class="muted small">Aucune bouteille.</p>'}
    <div class="row" style="margin-top:4px"><button class="btn link" data-act="emplacements" data-id="${v.id}">Modifier les emplacements</button><button class="btn link" data-act="corriger" data-id="${v.id}">Corriger le stock</button></div>
    <div class="section-title">Fiche</div>
    ${manque.filter(x => x !== 'photo').length ? `<div class="banner warn">À compléter : ${esc(manque.filter(x => x !== 'photo').join(', '))} <button class="btn link" data-act="modifier" data-id="${v.id}">Compléter</button></div>` : ''}
    <dl class="info">
      <dt>Appellation</dt><dd>${esc(v.appellation || '—')}</dd>
      <dt>Région</dt><dd>${esc(v.region || '—')}</dd>
      <dt>Garde</dt><dd>${v.gardeDebut || v.gardeFin ? `${v.gardeDebut || '?'} → ${v.gardeFin || '?'}` : '—'}</dd>
      <dt>Cépages</dt><dd>${esc(v.cepages || '—')}</dd>
      <dt>Prix d'achat moyen</dt><dd>${pm != null ? EUR2.format(pm) : '—'}</dd>
      <dt>Valeur estimée</dt><dd>${v.prixEstime != null ? EUR2.format(v.prixEstime) + ' / btl' : '—'}</dd>
      ${v.notes ? `<dt>Notes</dt><dd>${esc(v.notes)}</dd>` : ''}
    </dl>
    <div class="row" style="margin-top:10px"><button class="btn" data-act="modifier" data-id="${v.id}">Modifier la fiche</button><button class="btn" data-act="millesime" data-id="${v.id}">Autre millésime</button></div>
    <div class="section-title">Mouvements de stock</div>
    ${mv.length ? mv.map(m => {
      const signe = m.type === 'sortie' ? '−' : m.type === 'entree' ? '+' : '';
      const txt = m.type === 'transfert' ? `${caveNom(m.cave)} → ${caveNom(m.caveDest)}` : caveNom(m.cave);
      const det = [m.occasion, m.noteSur20 != null ? m.noteSur20 + '/20' : '', m.prix != null ? EUR2.format(m.prix) : '', m.fournisseur, m.commentaire].filter(Boolean).join(' · ');
      return `<div class="list-plain" style="padding:6px 0;border-top:1px solid var(--line)"><div style="display:flex;justify-content:space-between;gap:8px"><span><b>${signe}${Math.abs(m.qte)}</b> ${esc((TYPES[m.type] || [])[1] || '')} · ${esc(txt)}</span><span class="muted small">${esc(fmtDate(m.date) || 'date ?')}</span></div>${det ? `<div class="muted small">${esc(det)}</div>` : ''}</div>`;
    }).join('') : '<p class="muted small">Aucun mouvement.</p>'}
    <div style="margin-top:18px"><button class="btn danger" data-act="supprimer-vin" data-id="${v.id}">Supprimer ce vin</button></div>`, null, v.id);
  hydraterPhotos($('#sheetBody'));
}
function dateLongue(iso) {
  if (!iso) return 'Date inconnue';
  const d = new Date(iso.slice(0, 10) + 'T12:00:00');
  return isNaN(d) ? iso : d.toLocaleDateString('fr-FR', { day: 'numeric', month: 'long', year: 'numeric' });
}
function carteDegust(m) {
  const tag = m.type === 'sortie' ? `${plural(m.qte, 'bouteille bue', 'bouteilles bues')}${data.caves.length > 1 ? ' · ' + esc(caveNom(m.cave)) : ''}` : 'Commentaire';
  return `<div class="degust"><div class="degust-head"><span class="degust-date">${esc(dateLongue(m.date))}</span>${m.noteSur20 != null ? `<span class="note">${m.noteSur20}/20</span>` : ''}</div>
    <div class="muted small">${tag}${m.occasion ? ' · ' + esc(m.occasion) : ''}</div>
    ${m.commentaire ? `<p class="degust-txt">${esc(m.commentaire)}</p>` : ''}
    <button class="btn link" data-act="edit-commentaire" data-id="${esc(m.id)}">Modifier</button></div>`;
}
// Ajout ou modification d'un commentaire daté (sans changer le stock)
function formCommentaire(vinId, mvtId) {
  const m = mvtId ? data.mouvements.find(x => x.id === mvtId) : null;
  const v = vinById(m ? m.vinId : vinId);
  const estSortie = m && m.type === 'sortie';
  openSheet(`${head(m ? 'Modifier le commentaire' : 'Ajouter un commentaire', nomComplet(v))}<form class="f">
    ${estSortie ? `<p class="muted small" style="margin:0">Sortie de ${plural(m.qte, 'bouteille')} · le stock n'est pas modifié ici.</p>` : '<p class="muted small" style="margin:0">Un commentaire seul ne retire aucune bouteille du stock (bouteille bue chez des amis, notes d’un salon…).</p>'}
    <div class="two"><label>Date de consommation<input type="date" name="date" value="${esc((m && m.date) || today())}" required></label>
    <label>Note /20<input type="number" name="noteSur20" min="0" max="20" step="0.5" inputmode="decimal" value="${esc(m && m.noteSur20 != null ? m.noteSur20 : '')}"></label></div>
    <label>Occasion<input name="occasion" value="${esc((m && m.occasion) || '')}" placeholder="Dîner, anniversaire, avec…"></label>
    <label>Commentaire<textarea name="commentaire" rows="4" placeholder="Nez, bouche, accord, à racheter ?">${esc((m && m.commentaire) || '')}</textarea></label>
    <button class="btn primary">Enregistrer</button>
    ${m && !estSortie ? `<button type="button" class="btn danger" data-act="suppr-commentaire" data-id="${esc(m.id)}">Supprimer ce commentaire</button>` : ''}</form>`, fd => {
    const champs = { date: fd.date, noteSur20: num(fd.noteSur20), occasion: fd.occasion.trim(), commentaire: fd.commentaire.trim(), maj: nowIso() };
    if (!m && !champs.commentaire && champs.noteSur20 == null) { toast('Écrivez un commentaire ou une note'); return false; }
    commit(() => {
      if (m) Object.assign(m, champs);
      else data.mouvements.push({ id: uid('m'), vinId: v.id, type: 'commentaire', qte: 0, cave: ui.cave !== 'all' ? ui.cave : data.caves[0].id, ...champs });
    }, 'Commentaire enregistré');
    setTimeout(() => ouvrirVin(v.id), 0);
  }, v.id);
}
// Saisie rapide « J'ai bu » : recherche puis sortie d'une bouteille
function choisirVinBu() {
  openSheet(`${head('J’ai bu une bouteille', 'Cherchez le vin, puis touchez-le')}
    <div class="search"><input id="qb" type="search" placeholder="Domaine, cuvée, millésime…" autocomplete="off"></div>
    <div id="qbList" style="margin-top:10px"></div>`);
  listeVinBu('');
  setTimeout(() => $('#qb')?.focus(), 50);
}
function listeVinBu(q) {
  computeStock();
  const words = norm(q).trim().split(/\s+/).filter(Boolean);
  const l = trierVins(vinsActifs().filter(v => qte(v, ui.cave) > 0 && words.every(w => norm([v.domaine, v.cuvee, v.appellation, v.region, v.millesime, v.couleur].join(' ')).includes(w))));
  $('#qbList').innerHTML = l.length ? l.slice(0, 60).map(v => `<button type="button" class="pick" data-act="boire" data-id="${esc(v.id)}">
    ${v.photoMini || v.photo ? `<span class="thumb sm ${couleurClass(v.couleur)}"><img data-photo="${esc(v.photoMini || v.photo)}" alt=""></span>` : `<span class="dot ${couleurClass(v.couleur)}"></span>`}
    <span class="vin-main"><span class="vin-title">${esc(titreVin(v))} ${esc(v.millesime || '')}</span><span class="vin-sub" style="display:block">${esc(sousTitre(v))}</span></span>
    <span class="muted small">${qte(v, ui.cave)} btl</span></button>`).join('') : '<p class="empty-state">Aucun vin en stock ne correspond.</p>';
  hydraterPhotos($('#qbList'));
}

function cavesAvecStock(v) { return data.caves.filter(c => qte(v, c.id) > 0); }
function caveParDefaut(list) { return list.find(c => c.id === ui.cave) || list[0]; }

function formBoire(id) {
  const v = vinById(id);
  const d = $('#sheet');
  const depuisFiche = d.open && d.dataset.vin === id && !!$('#sheetBody [data-act="commentaire"]');
  const caves = cavesAvecStock(v);
  const def = caveParDefaut(caves);
  openSheet(`${head('Boire', nomComplet(v))}<form class="f">
    ${caves.length > 1 ? `<label>Cave<select name="cave">${options(caves.map(c => [c.id, `${c.nom} (${qte(v, c.id)})`]), def.id)}</select></label>` : `<input type="hidden" name="cave" value="${esc(def.id)}"><p class="muted small" style="margin:0">${esc(def.nom)} · ${qte(v, def.id)} en stock</p>`}
    <label>Nombre de bouteilles${stepper('qte', 1, 1, qte(v, def.id))}</label>
    <div class="two"><label>Date de consommation<input type="date" name="date" value="${today()}" required></label><label>Note /20<input type="number" name="noteSur20" min="0" max="20" step="0.5" inputmode="decimal"></label></div>
    <label>Commentaire de dégustation<textarea name="commentaire" rows="3" placeholder="Facultatif : vous pourrez l'ajouter plus tard depuis la fiche"></textarea></label>
    <label>Occasion<input name="occasion" placeholder="Dîner, anniversaire, offert à…"></label>
    <button class="btn primary">🍷 Enregistrer</button></form>`, fd => {
    const q = int(fd.qte);
    const dispo = qte(v, fd.cave);
    if (!q || q < 1 || q > dispo) { toast(`Quantité invalide (${dispo} en stock)`); return false; }
    commit(() => data.mouvements.push({ id: uid('m'), vinId: v.id, type: 'sortie', date: fd.date, qte: q, cave: fd.cave, occasion: fd.occasion.trim(), noteSur20: num(fd.noteSur20), commentaire: fd.commentaire.trim(), maj: nowIso() }),
      `${plural(q, 'bouteille')} sortie${q > 1 ? 's' : ''} · santé !`);
    if (depuisFiche) setTimeout(() => ouvrirVin(v.id), 0);
  }, v.id);
}
function formEntree(id) {
  const v = vinById(id);
  const def = ui.cave !== 'all' ? ui.cave : data.caves[0].id;
  openSheet(`${head('Ajouter des bouteilles', nomComplet(v))}<form class="f">
    <label>Cave<select name="cave">${options(data.caves.map(c => [c.id, c.nom]), def)}</select></label>
    <label>Nombre de bouteilles${stepper('qte', 6, 1)}</label>
    <div class="two"><label>Prix unitaire (€)<input type="number" name="prix" min="0" step="0.01" inputmode="decimal"></label><label>Date d'achat<input type="date" name="date" value="${today()}"></label></div>
    <label>Fournisseur<input name="fournisseur" list="dl-fourn"></label>${datalist('dl-fourn', fournisseurs())}
    <label>Emplacement dans la cave<input name="emplacement" placeholder="Casier, rangée…" value="${esc((v.emplacements || {})[def] || '')}"></label>
    <button class="btn primary">Ajouter au stock</button></form>`, fd => {
    const q = int(fd.qte);
    if (!q || q < 1) { toast('Quantité invalide'); return false; }
    commit(() => {
      data.mouvements.push({ id: uid('m'), vinId: v.id, type: 'entree', date: fd.date || null, qte: q, cave: fd.cave, prix: num(fd.prix), fournisseur: fd.fournisseur.trim(), maj: nowIso() });
      if (fd.emplacement.trim()) { v.emplacements[fd.cave] = fd.emplacement.trim(); v.maj = nowIso(); }
    }, `${plural(q, 'bouteille')} ajoutée${q > 1 ? 's' : ''}`);
  });
}
function formTransfert(id) {
  const v = vinById(id);
  const src = cavesAvecStock(v);
  const def = caveParDefaut(src);
  const dest = data.caves.find(c => c.id !== def.id);
  openSheet(`${head('Déplacer des bouteilles', nomComplet(v))}<form class="f">
    <div class="two"><label>De<select name="cave">${options(src.map(c => [c.id, `${c.nom} (${qte(v, c.id)})`]), def.id)}</select></label>
    <label>Vers<select name="caveDest">${options(data.caves.map(c => [c.id, c.nom]), dest.id)}</select></label></div>
    <label>Nombre de bouteilles${stepper('qte', 1, 1, qte(v, def.id))}</label>
    <label>Emplacement à l'arrivée<input name="emplacement" placeholder="Casier, rangée…"></label>
    <label>Date<input type="date" name="date" value="${today()}"></label>
    <button class="btn primary">Déplacer</button></form>`, fd => {
    const q = int(fd.qte);
    if (fd.cave === fd.caveDest) { toast('Choisissez deux caves différentes'); return false; }
    if (!q || q < 1 || q > qte(v, fd.cave)) { toast(`Quantité invalide (${qte(v, fd.cave)} disponibles)`); return false; }
    commit(() => {
      data.mouvements.push({ id: uid('m'), vinId: v.id, type: 'transfert', date: fd.date || today(), qte: q, cave: fd.cave, caveDest: fd.caveDest, maj: nowIso() });
      if (fd.emplacement.trim()) { v.emplacements[fd.caveDest] = fd.emplacement.trim(); v.maj = nowIso(); }
    }, 'Bouteilles déplacées');
  }, v.id);
}
function formCorriger(id) {
  const v = vinById(id);
  openSheet(`${head('Corriger le stock', nomComplet(v))}<form class="f"><p class="muted small" style="margin:0">Indiquez le nombre réel de bouteilles après un inventaire. La différence est notée dans le journal.</p>
    ${data.caves.map(c => `<label>${esc(c.nom)}<input type="number" name="c_${esc(c.id)}" min="0" value="${qte(v, c.id)}" inputmode="numeric"></label>`).join('')}
    <button class="btn primary">Enregistrer</button></form>`, fd => {
    const deltas = data.caves.map(c => [c.id, (int(fd['c_' + c.id]) ?? qte(v, c.id)) - qte(v, c.id)]).filter(([, d]) => d !== 0);
    if (!deltas.length) return;
    commit(() => deltas.forEach(([cave, d]) => data.mouvements.push({ id: uid('m'), vinId: v.id, type: 'ajustement', date: today(), qte: d, cave, note: 'Inventaire', maj: nowIso() })), 'Stock corrigé');
  });
}
function formEmplacements(id) {
  const v = vinById(id);
  openSheet(`${head('Emplacements', nomComplet(v))}<form class="f">
    ${data.caves.map(c => `<label>${esc(c.nom)}<input name="e_${esc(c.id)}" value="${esc(v.emplacements[c.id] || '')}" placeholder="Casier, rangée…"></label>`).join('')}
    <button class="btn primary">Enregistrer</button></form>`, fd => {
    commit(() => {
      for (const c of data.caves) { const x = (fd['e_' + c.id] || '').trim(); if (x) v.emplacements[c.id] = x; else delete v.emplacements[c.id]; }
      v.maj = nowIso();
    }, 'Emplacements enregistrés');
  });
}
function fournisseurs() { return [...new Set(data.mouvements.map(m => m.fournisseur).filter(Boolean))].sort(); }

function formVin(id, mode) {
  const src = id ? vinById(id) : null;
  const edit = mode === 'edit';
  const v = src ? { ...src } : { region: '', domaine: '', appellation: '', couleur: 'Rouge', cuvee: '', millesime: null, format: 'Bouteille', cepages: '', gardeDebut: null, gardeFin: null, prixEstime: null, notes: '' };
  if (mode === 'millesime') { v.millesime = null; v.gardeDebut = null; v.gardeFin = null; v.prixEstime = null; }
  const def = ui.cave !== 'all' ? ui.cave : data.caves[0].id;
  const titre = edit ? 'Modifier la fiche' : mode === 'millesime' ? 'Autre millésime' : 'Nouveau vin';
  openSheet(`${head(titre, mode === 'millesime' ? titreVin(src) : '')}<form class="f">
    <label>Domaine<input name="domaine" list="dl-dom" value="${esc(v.domaine)}" required></label>${datalist('dl-dom', uniques('domaine'))}
    <label>Cuvée<input name="cuvee" list="dl-cuv" value="${esc(v.cuvee)}"></label>${datalist('dl-cuv', uniques('cuvee'))}
    <div class="two"><label>Appellation<input name="appellation" list="dl-app" value="${esc(v.appellation)}"></label>${datalist('dl-app', uniques('appellation'))}
    <label>Région<input name="region" list="dl-reg" value="${esc(v.region)}" required></label>${datalist('dl-reg', [...new Set([...REGIONS, ...uniques('region')])].sort())}</div>
    <div class="two"><label>Couleur<select name="couleur">${options(COULEURS, v.couleur)}</select></label>
    <label>Millésime<input type="number" name="millesime" min="1900" max="2100" inputmode="numeric" value="${esc(v.millesime ?? '')}" placeholder="Non millésimé"></label></div>
    <div class="two"><label>Format<select name="format">${options(FORMATS, v.format)}</select></label>
    <label>Cépages<input name="cepages" value="${esc(v.cepages)}"></label></div>
    <div class="two"><label>À boire à partir de<input type="number" name="gardeDebut" min="1900" max="2150" inputmode="numeric" value="${esc(v.gardeDebut ?? '')}"></label>
    <label>À boire avant fin<input type="number" name="gardeFin" min="1900" max="2150" inputmode="numeric" value="${esc(v.gardeFin ?? '')}"></label></div>
    <label>Valeur estimée actuelle (€ / btl)<input type="number" name="prixEstime" min="0" step="0.01" inputmode="decimal" value="${esc(v.prixEstime ?? '')}"></label>
    <label>Notes<textarea name="notes">${esc(v.notes)}</textarea></label>
    ${edit ? '' : `<fieldset><legend>Bouteilles à ajouter</legend>
      <div class="two"><label>Cave<select name="cave">${options(data.caves.map(c => [c.id, c.nom]), def)}</select></label><label>Quantité<input type="number" name="qte" min="0" value="6" inputmode="numeric"></label></div>
      <div class="two"><label>Prix unitaire (€)<input type="number" name="prix" min="0" step="0.01" inputmode="decimal"></label><label>Date d'achat<input type="date" name="date" value="${today()}"></label></div>
      <label>Fournisseur<input name="fournisseur" list="dl-fourn"></label>${datalist('dl-fourn', fournisseurs())}
      <label>Emplacement<input name="emplacement" placeholder="Casier, rangée…"></label></fieldset>`}
    <button class="btn primary">${edit ? 'Enregistrer' : 'Ajouter ce vin'}</button></form>`, fd => {
    const fiche = {
      domaine: fd.domaine.trim(), cuvee: fd.cuvee.trim(), appellation: fd.appellation.trim(), region: fd.region.trim(),
      couleur: fd.couleur, millesime: int(fd.millesime), format: fd.format, cepages: fd.cepages.trim(),
      gardeDebut: int(fd.gardeDebut), gardeFin: int(fd.gardeFin), prixEstime: num(fd.prixEstime), notes: fd.notes.trim(), maj: nowIso()
    };
    if (fiche.gardeDebut && fiche.gardeFin && fiche.gardeFin < fiche.gardeDebut) { toast('La fin de garde doit être après le début'); return false; }
    if (edit) {
      commit(() => Object.assign(src, fiche), 'Fiche enregistrée');
      setTimeout(() => ouvrirVin(src.id), 0);
      return;
    }
    const doublon = vinsActifs().find(x => norm(x.domaine) === norm(fiche.domaine) && norm(x.cuvee) === norm(fiche.cuvee) && x.millesime === fiche.millesime && x.format === fiche.format && norm(x.appellation) === norm(fiche.appellation));
    if (doublon && !confirm(`« ${nomComplet(doublon)} » existe déjà. Créer quand même une nouvelle fiche ?\n(Annuler pour ajouter les bouteilles à la fiche existante.)`)) {
      const q = int(fd.qte);
      if (q > 0) commit(() => data.mouvements.push({ id: uid('m'), vinId: doublon.id, type: 'entree', date: fd.date || null, qte: q, cave: fd.cave, prix: num(fd.prix), fournisseur: fd.fournisseur.trim(), maj: nowIso() }), `${plural(q, 'bouteille')} ajoutée${q > 1 ? 's' : ''}`);
      return;
    }
    const nv = { id: uid('v'), ...fiche, emplacements: {}, aVerifier: '' };
    if (fd.emplacement.trim()) nv.emplacements[fd.cave] = fd.emplacement.trim();
    const q = int(fd.qte);
    commit(() => {
      data.vins.push(nv);
      if (q > 0) data.mouvements.push({ id: uid('m'), vinId: nv.id, type: 'entree', date: fd.date || null, qte: q, cave: fd.cave, prix: num(fd.prix), fournisseur: fd.fournisseur.trim(), maj: nowIso() });
    }, 'Vin ajouté');
  });
}
function ouvrirMvt(id) {
  const m = data.mouvements.find(x => x.id === id);
  if (m.type === 'commentaire') { formCommentaire(null, m.id); return; }
  const v = vinById(m.vinId) || {};
  const [, lab] = TYPES[m.type] || ['', ''];
  openSheet(`${head(`${lab} · ${Math.abs(m.qte)} btl`, nomComplet(v))}
    <dl class="info"><dt>Date</dt><dd>${esc(fmtDate(m.date) || 'inconnue')}</dd>
    <dt>Cave</dt><dd>${esc(m.type === 'transfert' ? caveNom(m.cave) + ' → ' + caveNom(m.caveDest) : caveNom(m.cave))}</dd>
    ${m.prix != null ? `<dt>Prix</dt><dd>${EUR2.format(m.prix)} / btl</dd>` : ''}${m.fournisseur ? `<dt>Fournisseur</dt><dd>${esc(m.fournisseur)}</dd>` : ''}
    ${m.occasion ? `<dt>Occasion</dt><dd>${esc(m.occasion)}</dd>` : ''}${m.noteSur20 != null ? `<dt>Note</dt><dd>${m.noteSur20}/20</dd>` : ''}
    ${m.commentaire ? `<dt>Commentaire</dt><dd>${esc(m.commentaire)}</dd>` : ''}${m.note ? `<dt>Remarque</dt><dd>${esc(m.note)}</dd>` : ''}</dl>
    <div class="row" style="margin-top:16px">${m.type === 'sortie' ? `<button class="btn" data-act="edit-commentaire" data-id="${m.id}">✎ Commentaire</button>` : ''}<button class="btn" data-act="voir-vin" data-id="${esc(v.id || '')}">Voir le vin</button><button class="btn danger" data-act="annuler-mvt" data-id="${m.id}">Annuler ce mouvement</button></div>`);
}

/* ---------- Import / export ---------- */
function choisirFichier() {
  const inp = document.createElement('input');
  inp.type = 'file'; inp.accept = '.json,application/json';
  inp.onchange = async () => {
    const f = inp.files[0];
    if (!f) return;
    try { importerDonnees(JSON.parse(await f.text())); } catch (e) { toast('Fichier illisible : ' + e.message); }
  };
  inp.click();
}
// Fichier de compléments : { "complements": [ { "id": "…", "cepages": "…", "gardeFin": 2030, … } ] }
// Seuls les champs de la fiche sont modifiés ; le stock et l'historique ne bougent pas.
const CHAMPS_COMPLEMENT = ['domaine', 'cuvee', 'appellation', 'region', 'couleur', 'millesime', 'cepages', 'gardeDebut', 'gardeFin', 'prixEstime', 'notes', 'photo', 'photoMini'];
function importerComplements(liste) {
  if (!data) { toast('Chargez d’abord votre cave'); return; }
  let n = 0;
  const ok = liste.filter(c => c && vinById(c.id));
  if (!ok.length) { toast('Aucune fiche correspondante dans ce fichier'); return; }
  if (!confirm(`Compléter ${plural(ok.length, 'fiche')} avec ce fichier ? Les champs déjà remplis à la main sont conservés.`)) return;
  commit(() => {
    for (const c of ok) {
      const v = vinById(c.id);
      let change = false;
      for (const k of CHAMPS_COMPLEMENT) {
        if (!(k in c) || c[k] === '' || c[k] == null) continue;
        if (v[k] !== '' && v[k] != null && !c.forcer) continue;
        v[k] = c[k]; change = true;
      }
      if (change) { v.maj = nowIso(); n++; }
    }
  }, '');
  toast(`${plural(n, 'fiche complétée', 'fiches complétées')}`);
}
function importerDonnees(obj) {
  if (obj && Array.isArray(obj.complements)) { importerComplements(obj.complements); return; }
  const d = normalize(obj);
  if (data && vinsActifs().length && !confirm(`Remplacer la cave actuelle (${vinsActifs().length} vins) par le fichier importé (${d.vins.length} vins) ?`)) return;
  const now = nowIso();
  for (const x of [...d.vins, ...d.mouvements, ...d.caves]) x.maj = x.maj || now;
  data = d;
  remplacer = true;
  commit(() => {}, `${plural(d.vins.length, 'vin')} importé${d.vins.length > 1 ? 's' : ''}`);
}
function telecharger(blob, nom) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = nom;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}
function loadScript(src) {
  return new Promise((res, rej) => {
    if ([...document.scripts].some(s => s.src === src)) return res();
    const s = document.createElement('script');
    s.src = src; s.onload = res; s.onerror = () => rej(new Error('Chargement impossible (hors ligne ?)'));
    document.head.appendChild(s);
  });
}
async function classeurExcel() {
  await loadScript(SHEETJS);
  computeStock();
  const XLSX = window.XLSX;
  const vins = trierVins(vinsActifs());
  const stock = vins.filter(v => qte(v) > 0).map(v => {
    const r = { 'Région': v.region, 'Domaine': v.domaine, 'Appellation': v.appellation, 'Couleur': v.couleur, 'Cuvée': v.cuvee, 'Millésime': v.millesime, 'Format': v.format, 'Cépages': v.cepages, 'Boire à partir de': v.gardeDebut, 'Boire avant fin': v.gardeFin, 'Statut': statut(v).label };
    for (const c of data.caves) { r['Qté ' + c.nom] = qte(v, c.id); r['Emplacement ' + c.nom] = v.emplacements[c.id] || ''; }
    const pm = prixMoyen(v), vu = valeurUnit(v);
    Object.assign(r, { 'Total': qte(v), "Prix d'achat moyen": pm != null ? Math.round(pm * 100) / 100 : null, 'Valeur estimée / btl': v.prixEstime, 'Valeur totale': vu != null ? Math.round(vu * qte(v) * 100) / 100 : null, 'À vérifier': v.aVerifier, 'Notes': v.notes });
    return r;
  });
  const mvts = mvtsActifs().sort((a, b) => (a.date || '').localeCompare(b.date || '')).map(m => {
    const v = vinById(m.vinId) || {};
    return { 'Date': m.date, 'Type': (TYPES[m.type] || [])[1], 'Quantité': m.qte, 'Cave': caveNom(m.cave), 'Vers': m.caveDest ? caveNom(m.caveDest) : '', 'Domaine': v.domaine, 'Cuvée': v.cuvee, 'Millésime': v.millesime, 'Prix unitaire': m.prix, 'Fournisseur': m.fournisseur, 'Occasion': m.occasion, 'Note /20': m.noteSur20, 'Commentaire': m.commentaire || m.note };
  });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(stock), 'Stock');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(mvts), 'Mouvements');
  return { XLSX, wb };
}
async function exportXlsx(versOneDrive) {
  try {
    const { XLSX, wb } = await classeurExcel();
    const nom = `Cave ${today()}.xlsx`;
    if (!versOneDrive) { XLSX.writeFile(wb, nom); return; }
    toast('Envoi vers OneDrive…');
    const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    const chemin = (dossier() ? dossier() + '/' : '') + nom;
    await putFile(cheminEnc(chemin), buf, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', null);
    toast('Enregistré : OneDrive / ' + chemin);
  } catch (e) { toast('Export impossible : ' + e.message); }
}

/* ---------- Événements ---------- */
document.addEventListener('click', e => {
  const t = e.target.closest('button');
  if (!t) return;
  if (t.dataset.step) {
    const inp = t.parentElement.querySelector('input');
    const v = (int(inp.value) || 0) + Number(t.dataset.step);
    const min = inp.min !== '' ? Number(inp.min) : -Infinity, max = inp.max !== '' ? Number(inp.max) : Infinity;
    inp.value = Math.min(max, Math.max(min, v));
    return;
  }
  if (t.dataset.tab) { ui.tab = t.dataset.tab; saveUi(); $('#content').innerHTML = ''; render(); window.scrollTo(0, 0); return; }
  if (t.dataset.cave) { ui.cave = t.dataset.cave; saveUi(); render(); return; }
  if (t.dataset.couleur !== undefined) { ui.couleur = t.dataset.couleur; render(); return; }
  if (t.dataset.toggle) { ui[t.dataset.toggle] = !ui[t.dataset.toggle]; saveUi(); render(); return; }
  if (t.dataset.jtype !== undefined) { ui.jtype = t.dataset.jtype; render(); return; }
  if (t.dataset.vin) { ouvrirVin(t.dataset.vin); return; }
  if (t.dataset.mvt) { ouvrirMvt(t.dataset.mvt); return; }
  if (t.id === 'fab') { formVin(null, 'new'); return; }
  if (t.id === 'fabBoire') { choisirVinBu(); return; }
  if (t.id === 'syncBtn') { if (syncMsg) toast(syncMsg); if (account) sync(); return; }
  const id = t.dataset.id;
  switch (t.dataset.act) {
    case 'close': closeSheet(); break;
    case 'login': login(); break;
    case 'logout': logout(); break;
    case 'sync': sync(); break;
    case 'import': choisirFichier(); break;
    case 'vide': data = emptyData(); commit(() => {}, 'Cave créée'); break;
    case 'boire': formBoire(id); break;
    case 'boire-rapide': choisirVinBu(); break;
    case 'commentaire': formCommentaire(id); break;
    case 'edit-commentaire': formCommentaire(null, id); break;
    case 'suppr-commentaire': {
      const m = data.mouvements.find(x => x.id === id);
      if (!confirm('Supprimer ce commentaire ?')) break;
      commit(() => { m.supprime = true; m.maj = nowIso(); }, 'Commentaire supprimé');
      ouvrirVin(m.vinId);
      break;
    }
    case 'photo': choisirPhoto(id); break;
    case 'theme': appliquerTheme(id); render(); break;
    case 'photo-web': lienPhotoWeb(id); break;
    case 'photo-suppr': {
      const v = vinById(id);
      if (!confirm('Retirer la photo de cette bouteille ?')) break;
      commit(() => { v.photo = ''; v.photoMini = ''; v.maj = nowIso(); }, 'Photo retirée');
      ouvrirVin(id);
      break;
    }
    case 'entree': formEntree(id); break;
    case 'transfert': formTransfert(id); break;
    case 'corriger': formCorriger(id); break;
    case 'emplacements': formEmplacements(id); break;
    case 'modifier': formVin(id, 'edit'); break;
    case 'millesime': formVin(id, 'millesime'); break;
    case 'voir-vin': if (id) ouvrirVin(id); break;
    case 'verifie': { const v = vinById(id); commit(() => { v.aVerifier = ''; v.maj = nowIso(); }, 'Fiche validée'); ouvrirVin(id); break; }
    case 'supprimer-vin': {
      const v = vinById(id);
      if (!confirm(`Supprimer « ${nomComplet(v)} » et tout son historique ?`)) break;
      closeSheet();
      commit(() => {
        v.supprime = true; v.maj = nowIso();
        data.mouvements.filter(m => m.vinId === id).forEach(m => { m.supprime = true; m.maj = nowIso(); });
      }, 'Vin supprimé');
      break;
    }
    case 'annuler-mvt': {
      const m = data.mouvements.find(x => x.id === id);
      if (!confirm('Annuler ce mouvement ? Le stock sera recalculé.')) break;
      closeSheet();
      commit(() => { m.supprime = true; m.maj = nowIso(); }, 'Mouvement annulé');
      break;
    }
    case 'voir-verif': ui.tab = 'cave'; ui.verif = true; ui.epuises = true; ui.q = ''; $('#content').innerHTML = ''; render(); break;
    case 'voir-inconnus': ui.tab = 'cave'; ui.q = ''; $('#content').innerHTML = ''; render(); toast('Astuce : ouvrez un vin puis « Modifier la fiche » pour saisir sa garde'); break;
    case 'renommer-cave': {
      const c = data.caves.find(x => x.id === id);
      const nom = prompt('Nouveau nom de la cave', c.nom);
      if (nom && nom.trim()) commit(() => { c.nom = nom.trim(); c.maj = nowIso(); }, 'Cave renommée');
      break;
    }
    case 'ajouter-cave': {
      const nom = prompt('Nom de la nouvelle cave');
      if (nom && nom.trim()) commit(() => data.caves.push({ id: uid('c'), nom: nom.trim(), maj: nowIso() }), 'Cave ajoutée');
      break;
    }
    case 'export-xlsx': exportXlsx(false); break;
    case 'export-xlsx-od': exportXlsx(true); break;
    case 'export-json': telecharger(new Blob([JSON.stringify(data, null, 1)], { type: 'application/json' }), `cave-sauvegarde-${today()}.json`); break;
  }
});
document.addEventListener('input', e => {
  if (e.target.id === 'q') { ui.q = e.target.value; renderList(); }
  if (e.target.id === 'qb') { listeVinBu(e.target.value); return; }
  // Quantité maximale = stock de la cave choisie (sortie, déplacement)
  const vinId = $('#sheet').dataset.vin;
  if (e.target.name === 'cave' && e.target.form && vinId) {
    const qInp = e.target.form.querySelector('input[name="qte"][max]');
    const v = vinById(vinId);
    if (qInp && v) { computeStock(); qInp.max = qte(v, e.target.value); if (int(qInp.value) > Number(qInp.max)) qInp.value = qInp.max; }
  }
});
$('#sheet').addEventListener('click', e => { if (e.target.id === 'sheet') closeSheet(); });
window.addEventListener('online', () => { if (account) sync(); });
window.addEventListener('offline', () => setSync(dirty ? 'pending' : 'offline', 'Pas de réseau'));
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && account && !dirty) sync(); });

/* ---------- Démarrage ---------- */
(async function start() {
  loadLocal();
  majCouleurBarre();
  if (LOCAL_MODE) remoteChecked = true;
  render();
  if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});
  try { await initAuth(); } catch (e) { setSync('error', e.message); }
  if (LOCAL_MODE) setSync('local', 'Données enregistrées uniquement dans ce navigateur');
  render();
  if (account) await sync();
})();

window.cave = { importerDonnees, get data() { return data; } };
