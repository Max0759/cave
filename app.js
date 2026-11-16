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
let ui = { tab: 'cave', cave: 'all', q: '', couleur: '', epuises: false, verif: false, incomplet: false, alertes: false, jtype: '', periode: '12m' };
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
  ALERTES = new Map();
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
// Alerte de stock : seuil choisi sur la fiche (0 = jamais), sinon automatique (25 % des bouteilles achetées)
let ALERTES = new Map();
function alerteStock(v) {
  if (!ALERTES.has(v.id)) ALERTES.set(v.id, calculerAlerte(v));
  return ALERTES.get(v.id);
}
function calculerAlerte(v) {
  if (v.seuilAlerte === 0) return null;
  let achetees = 0, dernSortie = '', bues90 = 0;
  const il90 = new Date(Date.now() - 90 * 864e5).toISOString().slice(0, 10);
  for (const m of data.mouvements) {
    if (m.supprime || m.vinId !== v.id) continue;
    if (m.type === 'entree') achetees += m.qte;
    if (m.type === 'sortie') { if ((m.date || '') > dernSortie) dernSortie = m.date || ''; if ((m.date || '') >= il90) bues90 += m.qte; }
  }
  const n = qte(v);
  const seuil = v.seuilAlerte ?? (achetees >= 3 ? Math.max(1, Math.round(achetees * 0.25)) : null);
  if (n <= 0) {
    const il60 = new Date(Date.now() - 60 * 864e5).toISOString().slice(0, 10);
    return dernSortie >= il60 && achetees > 0 ? { k: 'epuise', label: 'Épuisé', detail: `Dernière bouteille bue le ${fmtDate(dernSortie)}` } : null;
  }
  const rapide = bues90 >= 2 && bues90 >= n;
  if (seuil != null && n <= seuil) {
    const semaines = bues90 ? Math.max(1, Math.round(n / (bues90 / 13))) : null;
    return { k: 'bas', label: `Stock bas · ${n}`, detail: `Plus que ${plural(n, 'bouteille')} sur ${achetees}${semaines && rapide ? ` · épuisé dans ~${plural(semaines, 'semaine')} au rythme actuel` : ''}` };
  }
  if (rapide) return { k: 'rapide', label: 'Part vite', detail: `${bues90} bues en 3 mois, il en reste ${n}` };
  return null;
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
function choisirPhoto(id, apres) {
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
      if (apres) apres(); else ouvrirVin(v.id);
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
/* ---------- Lecture d'étiquette par Claude ----------
   La photo est lue par la page « Lecteur d'étiquettes » sur claude.ai (avec l'abonnement Claude,
   sans API). Elle renvoie ici le résultat : lien « #scan=… » ou code « MACAVE:… » à coller. */
const LECTEUR = () => CFG.lecteurUrl || 'https://claude.ai/artifact/L4YjkwBxeqGYQZyWwAPhWv';
let scan = null; // { file, apercu, lu, resultats }
const MOTS_VIDES = new Set(['domaine', 'dom', 'chateau', 'ch', 'clos', 'de', 'du', 'des', 'la', 'le', 'les', 'et', 'en', 'cuvee', 'vin', 'vins', 'aoc', 'aop', 'igp', 'appellation', 'controlee', 'protegee', 'mis', 'bouteille', 'au', 'a', 'l', 'd']);
const motsCles = t => norm(t).replace(/[^a-z0-9]+/g, ' ').split(' ').filter(w => w.length > 1 && !MOTS_VIDES.has(w));
function recouvrement(a, b) {
  const A = motsCles(a), B = new Set(motsCles(b));
  if (!A.length || !B.size) return 0;
  return A.filter(w => B.has(w) || [...B].some(x => x.length > 3 && w.length > 3 && (x.startsWith(w) || w.startsWith(x)))).length / Math.max(A.length, B.size);
}
function scoreVin(v, lu) {
  const texteVin = [v.domaine, v.cuvee, v.appellation].join(' ');
  const texteLu = [lu.domaine, lu.cuvee, lu.appellation, lu.texteLu].join(' ');
  let s = 3 * recouvrement(lu.domaine, v.domaine) + 2 * recouvrement(lu.cuvee || lu.appellation, v.cuvee || v.appellation) + 1 * recouvrement(lu.appellation, v.appellation) + 2 * recouvrement(texteVin, texteLu);
  if (lu.millesime && v.millesime) s += lu.millesime === v.millesime ? 2 : -1.5;
  if (lu.couleur && v.couleur && lu.couleur !== v.couleur) s -= 1;
  if (lu.format && v.format && lu.format !== v.format) s -= 0.5;
  return s;
}
function chercherCorrespondances(lu) {
  computeStock();
  return vinsActifs().map(v => ({ v, s: scoreVin(v, lu) })).filter(x => x.s >= 2).sort((a, b) => b.s - a.s || qte(b.v) - qte(a.v)).slice(0, 4);
}
function scannerEtiquette() {
  openSheet(`${head('Scanner une bouteille', 'Lecture par Claude, incluse dans votre abonnement')}
    <ol class="etapes"><li><a class="btn primary" href="${esc(LECTEUR())}" target="_blank" rel="noopener">📷 Ouvrir le lecteur d'étiquettes</a><span class="muted small">Photographiez l'étiquette dans Claude.</span></li>
    <li>Touchez « Ouvrir dans Ma Cave », ou copiez le code et collez-le ici :</li></ol>
    <form class="f" id="formCode"><textarea name="code" id="codeScan" rows="3" placeholder="MACAVE:…" autocomplete="off"></textarea>
    <button class="btn primary">Retrouver la bouteille</button></form>
    <div class="row"><button class="btn link" data-act="boire-rapide">Chercher le vin à la main</button></div>`, fd => {
    if (!recevoirScan(fd.code)) { toast('Code non reconnu : copiez-le en entier depuis le lecteur'); return false; }
    return false; // la feuille de résultat remplace celle-ci
  });
}
function decoderScan(code) {
  const brut = String(code || '').trim().replace(/^.*?(MACAVE:|#scan=)/s, '').replace(/\s+/g, '');
  if (!brut) return null;
  try {
    const b64 = brut.replace(/-/g, '+').replace(/_/g, '/');
    const bin = atob(b64 + '='.repeat((4 - b64.length % 4) % 4));
    const lu = JSON.parse(new TextDecoder().decode(Uint8Array.from(bin, c => c.charCodeAt(0))));
    return lu && typeof lu === 'object' ? lu : null;
  } catch (e) { return null; }
}
function recevoirScan(code) {
  const lu = decoderScan(code);
  if (!lu) return false;
  if (!data) { scanEnAttente = lu; toast('Bouteille lue : connectez-vous pour la retrouver'); return true; }
  scan = { file: null, apercu: '', lu, resultats: chercherCorrespondances(lu) };
  afficherScan();
  return true;
}
let scanEnAttente = null;
function scanDepuisAdresse() {
  const h = location.hash;
  if (!h.startsWith('#scan=')) return;
  history.replaceState(null, '', location.pathname + location.search);
  recevoirScan(h);
}
function nomLu(lu) { return [lu.domaine, lu.cuvee && lu.cuvee !== lu.domaine ? lu.cuvee : '', lu.appellation && lu.appellation !== lu.cuvee ? lu.appellation : ''].filter(Boolean).join(' · ') || 'Vin non identifié'; }
function complementsPossibles(v, lu) {
  const c = {};
  for (const k of ['appellation', 'region', 'cepages', 'gardeDebut', 'gardeFin']) if ((v[k] === '' || v[k] == null) && lu[k]) c[k] = lu[k];
  if (v.prixEstime == null && lu.prixIndicatif) c.prixEstime = lu.prixIndicatif;
  return c;
}
const LIB_CHAMPS = { appellation: 'appellation', region: 'région', cepages: 'cépages', gardeDebut: 'garde dès', gardeFin: 'garde jusqu’à', prixEstime: 'prix indicatif' };
function afficherScan() {
  const { lu, resultats } = scan;
  computeStock();
  const top = resultats[0];
  const sur = top && top.s >= 6 && (!lu.millesime || top.v.millesime === lu.millesime);
  const cartes = resultats.map(({ v }, i) => {
    const n = qte(v);
    const compl = complementsPossibles(v, lu);
    const nbC = Object.keys(compl).length;
    const memeMill = !lu.millesime || !v.millesime || v.millesime === lu.millesime;
    return `<div class="scan-match${i === 0 && sur ? ' best' : ''}">
      <button type="button" class="pick" data-vin="${esc(v.id)}">
        ${v.photoMini || v.photo ? `<span class="thumb sm ${couleurClass(v.couleur)}"><img data-photo="${esc(v.photoMini || v.photo)}" alt=""></span>` : `<span class="dot ${couleurClass(v.couleur)}"></span>`}
        <span class="vin-main"><span class="vin-title">${esc(titreVin(v))} ${esc(v.millesime || 'NM')}</span><span class="vin-sub" style="display:block">${esc(sousTitre(v))}${memeMill ? '' : ' · <b>autre millésime</b>'}</span></span>
        <span class="muted small">${n} btl</span></button>
      <div class="row">
        <button class="btn primary" data-act="scan-boire" data-id="${esc(v.id)}" ${n <= 0 ? 'disabled' : ''}>🍷 Bu (−1)</button>
        <button class="btn" data-act="scan-ajouter" data-id="${esc(v.id)}">+ Ajouter</button>
      </div>
      ${nbC ? `<button class="btn link" data-act="scan-completer" data-id="${esc(v.id)}">Compléter la fiche : ${esc(Object.keys(compl).map(k => LIB_CHAMPS[k]).join(', '))}</button>` : ''}
    </div>`;
  }).join('');
  openSheet(`${head(sur ? 'Bouteille reconnue' : resultats.length ? 'Est-ce l’un de ces vins ?' : 'Vin absent de la cave')}
    <div class="scan-lu">${scan.apercu ? `<img src="${scan.apercu}" alt="">` : ''}<div><div class="vin-title">${esc(nomLu(lu))}</div>
      <div class="muted small">${esc([lu.millesime || 'Millésime non lu', lu.couleur, lu.region, lu.format !== 'Bouteille' ? lu.format : ''].filter(Boolean).join(' · '))}</div>
      ${lu.cepages ? `<div class="muted small">${esc(lu.cepages)}${lu.gardeFin ? ` · garde ${lu.gardeDebut || '?'}–${lu.gardeFin}` : ''}${lu.prixIndicatif ? ' · ~' + EUR.format(lu.prixIndicatif) : ''}</div>` : ''}
      ${lu.remarque ? `<div class="muted small"><i>${esc(lu.remarque)}</i></div>` : ''}</div></div>
    ${cartes}
    ${scan.file ? `<label class="check"><input type="checkbox" id="scanPhoto" checked> Utiliser cette photo pour la fiche (si elle n'en a pas)</label>` : ''}
    <div class="row" style="margin-top:6px"><button class="btn" data-act="scan-nouveau">➕ Nouveau vin à partir de l'étiquette</button></div>
    <div class="row"><button class="btn link" data-act="boire-rapide">Pas dans la liste ? Chercher à la main</button><button class="btn link" data-act="scanner">Autre bouteille</button></div>`);
  hydraterPhotos($('#sheetBody'));
}
async function photoScanSurFiche(v) {
  const coche = $('#scanPhoto');
  if (!scan || !scan.file || v.photo || (coche && !coche.checked)) return;
  const file = scan.file; scan.file = null;
  try { const r = await enregistrerPhoto(v, file); commit(() => { Object.assign(v, r); v.maj = nowIso(); }); } catch (e) { /* la photo est facultative */ }
}

// Assistant photos : les vins en stock sans photo, avec appareil photo ou image du web pour chacun
function assistantPhotos() {
  computeStock();
  const sans = trierVins(vinsActifs().filter(v => qte(v) > 0 && !v.photo));
  const total = vinsActifs().filter(v => qte(v) > 0).length;
  openSheet(`${head('Photos des bouteilles', `${total - sans.length} / ${total} vins en stock ont une photo`)}
    <p class="muted small" style="margin:0 0 10px">📷 photographie la bouteille. 🔎 ouvre la recherche d'images : appui long sur une image, « Copier l'adresse de l'image », revenez ici et collez-la dans le champ.</p>
    ${sans.length ? `<div class="assist">${sans.map(v => `<div class="assist-ligne" data-assist="${esc(v.id)}">
      <div class="inv-nom"><b>${esc(titreVin(v))} ${esc(v.millesime || '')}</b><span class="muted small">${esc(v.domaine || v.appellation || '')} · ${qte(v)} btl</span></div>
      <div class="assist-actions"><button type="button" class="btn small-btn" data-act="assist-photo" data-id="${esc(v.id)}" aria-label="Photographier">📷</button>
      <a class="btn small-btn" href="${esc(rechercheImage(v))}" target="_blank" rel="noopener" aria-label="Chercher une image">🔎</a></div>
      <input class="assist-url" data-id="${esc(v.id)}" type="url" inputmode="url" placeholder="Coller l'adresse de l'image ici" autocomplete="off"></div>`).join('')}</div>`
      : '<p>Tous les vins en stock ont une photo. 🎉</p>'}`);
}
function enregistrerUrlPhoto(id, u) {
  const x = String(u || '').trim();
  if (!/^https?:\/\/\S+$/.test(x)) { toast('Ce n’est pas une adresse d’image'); return; }
  const v = vinById(id);
  commit(() => { v.photo = x; v.photoMini = x; v.maj = nowIso(); }, 'Image enregistrée');
  assistantPhotos();
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
  $('#fabScan').hidden = !(data && (ui.tab === 'cave' || ui.tab === 'boire'));
  if (!data) { content.innerHTML = renderWelcome(); return; }
  computeStock();
  if (ui.tab === 'cave') renderCave(content);
  else if (ui.tab === 'boire') content.innerHTML = renderBoire();
  else if (ui.tab === 'journal') content.innerHTML = renderJournal();
  else if (ui.tab === 'analyse') content.innerHTML = renderAnalyse();
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
  const nbAlertes = vinsActifs().filter(alerteStock).length;
  $('#chips').innerHTML = ['', ...COULEURS].filter(c => !c || vinsActifs().some(v => v.couleur === c)).map(c =>
    `<button type="button" class="chip" data-couleur="${esc(c)}" aria-pressed="${ui.couleur === c}">${c || 'Toutes couleurs'}</button>`).join('') +
    (nbAlertes ? `<button type="button" class="chip chip-alerte" data-toggle="alertes" aria-pressed="${ui.alertes}">🔔 Stock (${nbAlertes})</button>` : '') +
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
    if (!ui.epuises && n <= 0 && !(ui.alertes && alerteStock(v))) return false;
    if (ui.couleur && v.couleur !== ui.couleur) return false;
    if (ui.verif && !v.aVerifier) return false;
    if (ui.incomplet && !manquants(v).length) return false;
    if (ui.alertes && !alerteStock(v)) return false;
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
  const alertes = vinsActifs().filter(alerteStock);
  let html = alertes.length && !ui.alertes && !ui.q ? `<button type="button" class="alerte-banner" data-toggle="alertes">🔔 <span><b>${plural(alertes.length, 'vin')} à surveiller</b> : ${esc(alertes.slice(0, 3).map(v => nomComplet(v)).join(', '))}${alertes.length > 3 ? '…' : ''}</span><span class="voir">Voir</span></button>` : '';
  html += `<div class="summary">${plural(total, 'bouteille')} · ${plural(list.length, 'vin')}${ui.cave !== 'all' ? ' · ' + esc(caveNom(ui.cave)) : ''}</div>`;
  if (!list.length) {
    html += `<div class="empty-state">${ui.q || ui.couleur || ui.verif || ui.incomplet || ui.alertes ? 'Aucun vin ne correspond.' : ui.cave !== 'all' ? 'Aucune bouteille dans cette cave pour le moment.' : 'Aucun vin. Ajoutez-en avec le bouton +.'}</div>`;
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
  const al = alerteStock(v);
  return `<div class="vin-wrap"><button type="button" class="vin${n <= 0 ? ' empty' : ''}${n > 0 ? ' has-quick' : ''}" data-vin="${esc(v.id)}">
    ${mini ? `<span class="thumb ${couleurClass(v.couleur)}"><img data-photo="${esc(mini)}" alt="" decoding="async"></span>` : `<span class="dot ${couleurClass(v.couleur)}" title="${esc(v.couleur)}"></span>`}
    <span class="vin-main"><span class="vin-title">${esc(titreVin(v))}</span>
      <span class="vin-sub" style="display:block">${esc(sousTitre(v))}${v.format && v.format !== 'Bouteille' ? ' · ' + esc(v.format) : ''}</span>
      <span class="vin-meta"><span class="pill s-${st.k}">${esc(st.label)}</span>${al ? `<span class="pill alerte a-${al.k}">🔔 ${esc(al.label)}</span>` : ''}${v.aVerifier ? '<span class="pill verif">À vérifier</span>' : ''}${loc ? `<span class="loc">📍 ${esc(loc)}</span>` : ''}</span></span>
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
  const alertes = vinsActifs().filter(alerteStock);
  if (alertes.length) html += `<div class="group-title"><span>🔔 Stock à surveiller</span><span>${alertes.length}</span></div>` + alertes.map(v => `<div class="alerte-ligne"><button type="button" class="btn link" data-vin="${esc(v.id)}">${esc(nomComplet(v))}</button><span class="muted small">${esc(alerteStock(v).detail)}</span></div>`).join('');
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

/* ---------- Onglet Analyse : consommation et tarifs ---------- */
const TRANCHES = [[0, 10, 'Moins de 10 €'], [10, 20, '10 – 20 €'], [20, 40, '20 – 40 €'], [40, 80, '40 – 80 €'], [80, Infinity, '80 € et plus']];
const trancheDe = p => p == null ? -1 : TRANCHES.findIndex(([a, b]) => p >= a && p < b);
const prixRef = v => prixMoyen(v) ?? v.prixEstime ?? null; // prix d'achat, sinon valeur estimée
function debutPeriode() {
  const d = new Date();
  if (ui.periode === '12m') { d.setMonth(d.getMonth() - 11, 1); return d.toISOString().slice(0, 7) + '-01'; }
  if (ui.periode === 'annee') return d.getFullYear() + '-01-01';
  return '0000';
}
const EUR0 = n => EUR.format(Math.round(n));
function renderAnalyse() {
  const c = ui.cave;
  const deb = debutPeriode();
  const sorties = mvtsActifs().filter(m => m.type === 'sortie' && (c === 'all' || m.cave === c) && (m.date ? m.date >= deb : ui.periode === 'tout'));
  const nbBues = sorties.reduce((a, m) => a + m.qte, 0);
  let html = `<div class="chips" style="padding-top:0">${[['12m', '12 derniers mois'], ['annee', 'Année ' + YEAR()], ['tout', 'Depuis le début']]
    .map(([k, l]) => `<button type="button" class="chip" data-periode="${k}" aria-pressed="${ui.periode === k}">${l}</button>`).join('')}</div>`;
  if (!nbBues) return html + '<div class="empty-state">Aucune bouteille bue sur cette période. Les sorties enregistrées avec « Bu » apparaîtront ici.</div>';

  // Valeurs
  let valeurBue = 0, qPrixBu = 0;
  for (const m of sorties) { const p = prixRef(vinById(m.vinId) || {}); if (p != null) { valeurBue += p * m.qte; qPrixBu += m.qte; } }
  const enStock = vinsActifs().filter(v => qte(v, c) > 0);
  let valeurStock = 0, qPrixStock = 0, nbStock = 0;
  for (const v of enStock) { const n = qte(v, c); nbStock += n; const p = prixRef(v); if (p != null) { valeurStock += p * n; qPrixStock += n; } }
  const moisPeriode = ui.periode === '12m' ? 12 : ui.periode === 'annee' ? new Date().getMonth() + 1 : Math.max(1, moisEntre(mvtsActifs().filter(m => m.type === 'entree' || m.type === 'sortie').reduce((a, m) => (m.date && m.date < a ? m.date : a), today()), today()));
  const rythme = nbBues / moisPeriode;
  const pmBu = qPrixBu ? valeurBue / qPrixBu : null, pmStock = qPrixStock ? valeurStock / qPrixStock : null;
  html += `<div class="kpis">
    <div class="kpi"><b>${nbBues}</b><span>bouteilles bues · ${rythme.toFixed(1).replace('.', ',')} / mois</span></div>
    <div class="kpi"><b>${qPrixBu ? EUR0(valeurBue) : '—'}</b><span>valeur bue${qPrixBu < nbBues ? ` (${qPrixBu} btl au prix connu)` : ''}</span></div>
    <div class="kpi"><b>${pmBu != null ? EUR2.format(pmBu) : '—'}</b><span>prix moyen d'une bouteille bue</span></div>
    <div class="kpi"><b>${pmStock != null ? EUR2.format(pmStock) : '—'}</b><span>prix moyen d'une bouteille en stock</span></div>
    <div class="kpi"><b>${rythme > 0 ? Math.round(nbStock / rythme) : '—'}</b><span>mois de stock au rythme actuel</span></div>
    <div class="kpi"><b>${qPrixStock ? EUR0(valeurStock) : '—'}</b><span>valeur du stock (${nbStock} btl)</span></div></div>`;
  if (pmBu != null && pmStock != null) {
    const ecart = (pmStock - pmBu) / pmBu;
    html += `<div class="banner">${Math.abs(ecart) < 0.15 ? 'Vous buvez des bouteilles d’un prix proche de celui de votre stock.'
      : ecart > 0 ? `Votre stock vaut en moyenne <b>${Math.round(ecart * 100)} % de plus</b> par bouteille que ce que vous buvez : les belles bouteilles attendent.`
      : `Vous buvez en moyenne des bouteilles <b>${Math.round(-ecart * 100)} % plus chères</b> que celles qui restent en stock.`}</div>`;
  }

  // Graphique 1 : consommation par mois (une série)
  const mois = new Map();
  for (const m of sorties) { if (!m.date) continue; const k = m.date.slice(0, 7); const o = mois.get(k) || { n: 0, val: 0 }; o.n += m.qte; const p = prixRef(vinById(m.vinId) || {}); if (p != null) o.val += p * m.qte; mois.set(k, o); }
  const cles = listeMois(ui.periode === 'tout' ? [...mois.keys()].sort()[0] || today().slice(0, 7) : deb.slice(0, 7), today().slice(0, 7));
  const maxM = Math.max(1, ...cles.map(k => (mois.get(k) || {}).n || 0));
  const etiquetteMois = k => new Date(k + '-15T12:00').toLocaleDateString('fr-FR', { month: 'short' }).replace('.', '');
  const pas = cles.length > 18 ? Math.ceil(cles.length / 12) : 1;
  const sansDate = sorties.filter(m => !m.date).reduce((a, m) => a + m.qte, 0);
  if (sansDate < nbBues) html += `<div class="card"><h2>Bouteilles bues par mois</h2>
    <div class="chart-mois" style="--n:${cles.length}">${cles.map((k, i) => { const o = mois.get(k) || { n: 0, val: 0 };
      const lib = new Date(k + '-15T12:00').toLocaleDateString('fr-FR', { month: 'long', year: 'numeric' });
      return `<div class="col" tabindex="0" data-tip="${esc(`<b>${lib}</b><br>${plural(o.n, 'bouteille')}${o.val ? ' · ' + EUR0(o.val) : ''}`)}">
        <span class="bar" style="height:${(o.n / maxM * 100).toFixed(1)}%"></span><span class="x">${i % pas === 0 ? esc(etiquetteMois(k)) : ''}</span></div>`; }).join('')}</div>
    <div class="axe-note muted small">max ${maxM} btl / mois</div></div>`;
  if (sansDate) html += `<p class="muted small" style="margin:-4px 2px 12px">${plural(sansDate, 'bouteille bue', 'bouteilles bues')} sans date (reprises de l'Excel) ne figurent pas dans le graphique par mois.</p>`;

  // Graphique 2 : tranches de prix, bues vs en stock (parts en %)
  const bu = TRANCHES.map(() => 0), st = TRANCHES.map(() => 0);
  let buInc = 0, stInc = 0;
  for (const m of sorties) { const t = trancheDe(prixRef(vinById(m.vinId) || {})); if (t < 0) buInc += m.qte; else bu[t] += m.qte; }
  for (const v of enStock) { const t = trancheDe(prixRef(v)); if (t < 0) stInc += qte(v, c); else st[t] += qte(v, c); }
  const totB = bu.reduce((a, b) => a + b, 0), totS = st.reduce((a, b) => a + b, 0);
  if (totB || totS) {
    const maxP = Math.max(1, ...bu.map(x => totB ? x / totB : 0), ...st.map(x => totS ? x / totS : 0));
    html += `<div class="card"><h2>Par tranche de prix</h2>
      <div class="legende"><span><i class="sw bu"></i>Bues (${totB})</span><span><i class="sw st"></i>En stock (${totS})</span></div>
      <div class="chart-tranches">${TRANCHES.map(([, , l], i) => {
        const pb = totB ? bu[i] / totB : 0, ps = totS ? st[i] / totS : 0;
        return `<div class="tr-row"><span class="tr-lib">${l}</span><span class="tr-bars">
          <span class="hb bu" style="width:${(pb / maxP * 100).toFixed(1)}%" data-tip="${esc(`<b>${l}</b><br>Bues : ${bu[i]} btl (${Math.round(pb * 100)} %)`)}"></span>
          <span class="hb st" style="width:${(ps / maxP * 100).toFixed(1)}%" data-tip="${esc(`<b>${l}</b><br>En stock : ${st[i]} btl (${Math.round(ps * 100)} %)`)}"></span></span>
          <span class="tr-val">${Math.round(pb * 100)} %<br>${Math.round(ps * 100)} %</span></div>`; }).join('')}</div>
      ${buInc || stInc ? `<p class="muted small" style="margin:8px 0 0">Sans prix connu : ${buInc} bues, ${stInc} en stock (non comptées).</p>` : ''}
      <details class="tableau"><summary>Voir le tableau</summary><table><tr><th>Tranche</th><th>Bues</th><th>En stock</th></tr>
        ${TRANCHES.map(([, , l], i) => `<tr><td>${l}</td><td>${bu[i]}</td><td>${st[i]}</td></tr>`).join('')}</table></details></div>`;
  }

  // Vins les plus bus et rapport note / prix
  const parVin = new Map();
  for (const m of sorties) { const o = parVin.get(m.vinId) || { n: 0, notes: [] }; o.n += m.qte; if (m.noteSur20 != null) o.notes.push(m.noteSur20); parVin.set(m.vinId, o); }
  for (const m of mvtsActifs()) if (m.type === 'commentaire' && m.noteSur20 != null && (m.date || '') >= deb && parVin.has(m.vinId)) parVin.get(m.vinId).notes.push(m.noteSur20);
  const lignes = [...parVin.entries()].map(([id, o]) => { const v = vinById(id) || {}; const p = prixRef(v); const note = o.notes.length ? o.notes.reduce((a, b) => a + b, 0) / o.notes.length : null; return { v, n: o.n, p, note }; });
  const top = lignes.slice().sort((a, b) => b.n - a.n).slice(0, 8);
  html += `<div class="card"><h2>Les plus bus</h2><table class="t-analyse"><tr><th>Vin</th><th>Bues</th><th>Prix</th><th>Note</th></tr>
    ${top.map(l => `<tr><td><button class="btn link" data-vin="${esc(l.v.id || '')}">${esc(nomComplet(l.v))}</button></td><td>${l.n}</td><td>${l.p != null ? EUR0(l.p) : '—'}</td><td>${l.note != null ? l.note.toFixed(1).replace('.', ',') : '—'}</td></tr>`).join('')}</table></div>`;
  const notes = lignes.filter(l => l.note != null && l.p);
  if (notes.length >= 2) {
    html += renderNuageNotes(notes);
    const rq = notes.map(l => ({ ...l, r: l.note / Math.log2(l.p + 2) })).sort((a, b) => b.r - a.r).slice(0, 5);
    html += `<div class="card"><h2>Meilleurs rapports qualité / prix</h2><ol class="rq">${rq.map(l => `<li><button class="btn link" data-vin="${esc(l.v.id)}">${esc(nomComplet(l.v))}</button> <span class="muted small">${l.note.toFixed(1).replace('.', ',')}/20 pour ${EUR0(l.p)}</span></li>`).join('')}</ol>
      <p class="muted small" style="margin:6px 0 0">Classement : note obtenue au regard du prix payé (les prix élevés sont pondérés).</p></div>`;
  } else {
    html += `<p class="muted small">Notez vos dégustations (/20) pour voir le rapport qualité / prix de vos vins.</p>`;
  }
  return html;
}
function renderNuageNotes(pts) {
  const W = 320, H = 200, g = { l: 34, r: 10, t: 10, b: 28 };
  const maxP = Math.max(...pts.map(p => p.p)) * 1.1, minN = Math.min(10, ...pts.map(p => p.note)), maxN = 20;
  const x = p => g.l + (p / maxP) * (W - g.l - g.r), y = n => H - g.b - ((n - minN) / (maxN - minN)) * (H - g.t - g.b);
  const ticksX = [0, maxP / 3, 2 * maxP / 3].map(v => Math.round(v / 5) * 5);
  const ticksY = [minN, (minN + maxN) / 2, maxN].map(Math.round);
  return `<div class="card"><h2>Note selon le prix</h2><svg class="nuage" viewBox="0 0 ${W} ${H}" role="img" aria-label="Note sur 20 selon le prix par bouteille">
    ${ticksY.map(n => `<line class="grille" x1="${g.l}" x2="${W - g.r}" y1="${y(n)}" y2="${y(n)}"/><text class="ax" x="${g.l - 6}" y="${y(n) + 4}" text-anchor="end">${n}</text>`).join('')}
    ${ticksX.map(p => `<text class="ax" x="${x(p)}" y="${H - 8}" text-anchor="middle">${p} €</text>`).join('')}
    ${pts.map(p => `<circle class="pt" cx="${x(p.p).toFixed(1)}" cy="${y(p.note).toFixed(1)}" r="5" data-vin="${esc(p.v.id)}" data-tip="${esc(`<b>${esc(nomComplet(p.v))}</b><br>${p.note.toFixed(1).replace('.', ',')}/20 · ${EUR0(p.p)}`)}"/>`).join('')}
  </svg><p class="muted small" style="margin:4px 0 0">Chaque point est un vin bu et noté. En haut à gauche : bonnes notes pour peu cher.</p></div>`;
}
function moisEntre(a, b) { const [ya, ma] = a.split('-').map(Number), [yb, mb] = b.split('-').map(Number); return (yb - ya) * 12 + (mb - ma) + 1; }
function listeMois(deb, fin) {
  const r = []; let [y, m] = deb.split('-').map(Number); const [yf, mf] = fin.split('-').map(Number);
  while (y < yf || (y === yf && m <= mf)) { r.push(`${y}-${String(m).padStart(2, '0')}`); if (++m > 12) { m = 1; y++; } if (r.length > 240) break; }
  return r;
}
// Infobulle partagée pour les graphiques ([data-tip])
function infobulle() {
  let tip = $('#tip');
  if (!tip) { tip = document.createElement('div'); tip.id = 'tip'; tip.className = 'tip'; document.body.appendChild(tip); }
  return tip;
}
function montrerTip(el, x, y) {
  const tip = infobulle();
  tip.innerHTML = el.dataset.tip;
  tip.classList.add('show');
  const r = tip.getBoundingClientRect();
  tip.style.left = Math.min(window.innerWidth - r.width - 8, Math.max(8, x - r.width / 2)) + 'px';
  tip.style.top = Math.max(8, y - r.height - 14) + 'px';
}
document.addEventListener('pointerover', e => { const el = e.target.closest?.('[data-tip]'); if (el) { const b = el.getBoundingClientRect(); montrerTip(el, e.clientX || b.left + b.width / 2, b.top); } });
document.addEventListener('pointerout', e => { if (e.target.closest?.('[data-tip]')) infobulle().classList.remove('show'); });
document.addEventListener('focusin', e => { const el = e.target.closest?.('[data-tip]'); if (el) { const b = el.getBoundingClientRect(); montrerTip(el, b.left + b.width / 2, b.top); } });
document.addEventListener('scroll', () => $('#tip')?.classList.remove('show'), true);

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
  const dbl = chercherDoublons();
  const inv = lireInv();
  html += `<div class="card"><h2>Contrôle de la cave</h2>
    <p class="small" style="margin:0 0 8px">${dbl.entrees.length || dbl.fiches.length ? `⚠️ ${dbl.entrees.length ? plural(dbl.entrees.length, 'entrée en double', 'entrées en double') + ` (${plural(dbl.btlEnTrop, 'bouteille')} peut-être en trop)` : ''}${dbl.entrees.length && dbl.fiches.length ? ' · ' : ''}${dbl.fiches.length ? plural(dbl.fiches.length, 'fiche en double', 'fiches en double') : ''}` : '✅ Aucun doublon détecté.'}
    ${inv ? `<br>Inventaire en cours : ${esc(caveNom(inv.cave))}, ${plural(Object.keys(inv.comptes).length, 'vin compté', 'vins comptés')}.` : ''}</p>
    <div class="row"><button class="btn" data-act="doublons">Chercher les doublons</button><button class="btn" data-act="inventaire">${inv ? 'Reprendre l’inventaire' : 'Faire l’inventaire'}</button></div></div>`;
  html += `<div class="card"><h2>Caves</h2><ul class="list-plain">${data.caves.map(cv => `<li><span>${esc(cv.nom)}</span><button class="btn link" data-act="renommer-cave" data-id="${esc(cv.id)}">Renommer</button></li>`).join('')}</ul>
    <button class="btn" data-act="ajouter-cave" style="margin-top:8px">Ajouter une cave</button></div>`;
  html += `<div class="card"><h2>Lecture des étiquettes</h2><p class="small" style="margin:0 0 8px">Le bouton 📷 Scanner ouvre le lecteur d'étiquettes sur claude.ai : la lecture est incluse dans votre abonnement Claude.</p>
    <a class="btn" href="${esc(LECTEUR())}" target="_blank" rel="noopener">Ouvrir le lecteur</a></div>`;
  const enStockB = vinsActifs().filter(v => qte(v) > 0), avecPhoto = enStockB.filter(v => v.photo).length;
  html += `<div class="card"><h2>Photos des bouteilles</h2><p class="small" style="margin:0 0 8px">${avecPhoto} / ${enStockB.length} vins en stock ont une photo.</p>
    <button class="btn" data-act="photos">Ajouter les photos manquantes</button></div>`;
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
      <dt>Alerte stock</dt><dd>${v.seuilAlerte === 0 ? 'Désactivée' : v.seuilAlerte != null ? `Sous ${plural(v.seuilAlerte, 'bouteille')}` : 'Automatique (25 % des achats)'}${alerteStock(v) ? ` · <b>🔔 ${esc(alerteStock(v).detail)}</b>` : ''}</dd>
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
    <div class="row" style="margin-bottom:10px"><button class="btn" data-act="scanner">📷 Photographier l'étiquette</button></div>
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
    const al = alerteStock(v);
    if (al) setTimeout(() => toast(`🔔 ${nomComplet(v)} : ${al.k === 'epuise' ? 'dernière bouteille bue' : al.detail}`), 2700);
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
/* ---------- Contrôle de la cave : doublons et inventaire ---------- */
// Entrées identiques (même vin, date, quantité, prix, cave) et fiches identiques (même vin, millésime, format)
function chercherDoublons() {
  computeStock();
  const ent = new Map();
  for (const m of mvtsActifs()) {
    if (m.type !== 'entree') continue;
    const k = [m.vinId, m.date || '', m.qte, m.prix ?? '', m.cave, norm(m.fournisseur)].join('|');
    if (!ent.has(k)) ent.set(k, []);
    ent.get(k).push(m);
  }
  const entrees = [...ent.values()].filter(l => l.length > 1 && vinById(l[0].vinId) && !vinById(l[0].vinId).supprime);
  const fic = new Map();
  for (const v of vinsActifs()) { const k = cleVin(v); if (!fic.has(k)) fic.set(k, []); fic.get(k).push(v); }
  const fiches = [...fic.values()].filter(l => l.length > 1);
  return { entrees, fiches, btlEnTrop: entrees.reduce((a, l) => a + l.slice(1).reduce((b, m) => b + m.qte, 0), 0) };
}
function ouvrirDoublons() {
  const { entrees, fiches, btlEnTrop } = chercherDoublons();
  if (!entrees.length && !fiches.length) { openSheet(`${head('Contrôle des doublons')}<p>Aucun doublon trouvé : pas d'entrée enregistrée deux fois, pas de fiche en double.</p><p class="muted small">Si le stock paraît encore faux, faites l'inventaire d'une cave pour recompter.</p><div class="row"><button class="btn primary" data-act="inventaire">Faire l'inventaire</button></div>`); return; }
  openSheet(`${head('Contrôle des doublons', entrees.length ? `${plural(btlEnTrop, 'bouteille')} peut-être en trop` : '')}
    ${entrees.length ? `<div class="section-title">Entrées enregistrées plusieurs fois</div>
    <p class="muted small" style="margin:0 0 8px">Même vin, même date, même quantité et même prix. C'est souvent un import ou une synchronisation en double. Gardez une seule entrée si c'est une erreur.</p>
    ${entrees.map((l, i) => { const v = vinById(l[0].vinId); const m = l[0];
      return `<div class="doublon"><div><b>${esc(nomComplet(v))}</b><div class="muted small">${l.length} × +${m.qte} btl · ${esc(fmtDate(m.date) || 'date inconnue')} · ${esc(caveNom(m.cave))}${m.prix != null ? ' · ' + EUR2.format(m.prix) : ''} · stock actuel ${qte(v)}</div></div>
        <button class="btn small-btn" data-act="garder-une" data-id="${i}">Garder une seule (−${l.slice(1).reduce((a, x) => a + x.qte, 0)})</button></div>`; }).join('')}
    <div class="row" style="margin-top:8px"><button class="btn" data-act="garder-une-tout">Corriger toutes les entrées en double</button></div>` : ''}
    ${fiches.length ? `<div class="section-title">Fiches en double</div>
    <p class="muted small" style="margin:0 0 8px">Le même vin a plusieurs fiches. Fusionner regroupe le stock et l'historique sur une seule fiche.</p>
    ${fiches.map((l, i) => `<div class="doublon"><div><b>${esc(nomComplet(l[0]))}</b><div class="muted small">${l.length} fiches · stocks ${l.map(v => qte(v)).join(' + ')}</div></div>
      <button class="btn small-btn" data-act="fusionner" data-id="${i}">Fusionner</button></div>`).join('')}` : ''}`);
  ouvrirDoublons.etat = { entrees, fiches };
}
function garderUne(groupes) {
  let n = 0;
  commit(() => { for (const l of groupes) for (const m of l.slice(1)) { m.supprime = true; m.maj = nowIso(); n += m.qte; } }, '');
  toast(`${plural(n, 'bouteille retirée', 'bouteilles retirées')} du stock`);
}
function fusionnerFiches(l) {
  const [garde, ...autres] = l;
  commit(() => {
    for (const v of autres) {
      for (const m of data.mouvements) if (m.vinId === v.id) { m.vinId = garde.id; m.maj = nowIso(); }
      for (const [c, e] of Object.entries(v.emplacements || {})) if (!garde.emplacements[c]) garde.emplacements[c] = e;
      for (const k of ['appellation', 'region', 'cepages', 'gardeDebut', 'gardeFin', 'prixEstime', 'photo', 'photoMini', 'notes']) if ((garde[k] === '' || garde[k] == null) && v[k]) garde[k] = v[k];
      v.supprime = true; v.maj = nowIso();
    }
    garde.maj = nowIso();
  }, 'Fiches fusionnées');
}

// Inventaire : recompter une cave, emplacement par emplacement. Le brouillon reste sur l'appareil jusqu'à validation.
const CLE_INV = 'cave.inventaire';
function lireInv() { try { return JSON.parse(localStorage.getItem(CLE_INV) || 'null'); } catch (e) { return null; } }
function ecrireInv(x) { try { if (x) localStorage.setItem(CLE_INV, JSON.stringify(x)); else localStorage.removeItem(CLE_INV); } catch (e) { /* sans importance */ } }
function ouvrirInventaire(caveId) {
  const brouillon = lireInv();
  if (!caveId && brouillon) caveId = brouillon.cave;
  if (!caveId) {
    openSheet(`${head('Inventaire', 'Recompter les bouteilles, casier par casier')}
      <p class="muted small">Choisissez la cave à recompter. Vous pouvez vous arrêter en cours de route : le comptage est gardé sur cet appareil jusqu'à ce que vous l'enregistriez.</p>
      <div class="row">${data.caves.map(c => `<button class="btn" data-act="inventaire-cave" data-id="${esc(c.id)}">${esc(c.nom)} (${vinsActifs().reduce((a, v) => a + qte(v, c.id), 0)})</button>`).join('')}</div>`);
    return;
  }
  computeStock();
  const inv = brouillon && brouillon.cave === caveId ? brouillon : { cave: caveId, debut: today(), comptes: {} };
  ecrireInv(inv);
  const vins = vinsActifs().filter(v => qte(v, caveId) > 0 || inv.comptes[v.id] != null);
  const groupes = new Map();
  // Groupes : l'emplacement s'il est connu, sinon le domaine (plus facile à retrouver dans la cave)
  for (const v of vins) { const e = (v.emplacements || {})[caveId] || (v.domaine ? 'Domaine : ' + v.domaine : 'Sans emplacement'); if (!groupes.has(e)) groupes.set(e, []); groupes.get(e).push(v); }
  const cles = [...groupes.keys()].sort((a, b) => a === 'Sans emplacement' ? 1 : b === 'Sans emplacement' ? -1 : a.localeCompare(b, 'fr', { numeric: true }));
  const vus = vins.filter(v => inv.comptes[v.id] != null).length;
  openSheet(`${head('Inventaire · ' + caveNom(caveId), `${vus} / ${vins.length} vins comptés`)}
    <p class="muted small" style="margin:0 0 10px">Pour chaque vin, touchez ✓ si le nombre est juste, ou corrigez-le. Les vins non comptés ne changent pas.</p>
    <div class="inv">${cles.map(e => `<div class="inv-groupe"><div class="inv-titre"><span>📍 ${esc(e)}</span><button class="btn link" data-act="inv-groupe-ok" data-id="${esc(e)}">Tout est juste</button></div>
      ${trierVins(groupes.get(e)).map(v => { const att = qte(v, caveId); const c = inv.comptes[v.id];
        return `<div class="inv-ligne${c != null ? (c === att ? ' ok' : ' diff') : ''}" data-vin-inv="${esc(v.id)}" data-emp="${esc(e)}">
          <div class="inv-nom"><b>${esc(titreVin(v))} ${esc(v.millesime || '')}</b><span class="muted small">${esc(v.domaine || '')} · prévu ${att}</span></div>
          <div class="inv-saisie"><button type="button" class="btn small-btn inv-ok" data-act="inv-ok" data-id="${esc(v.id)}" aria-label="Le nombre est juste">✓</button>
          <input type="number" min="0" inputmode="numeric" class="inv-n" data-id="${esc(v.id)}" value="${c ?? att}" aria-label="Bouteilles comptées"></div></div>`; }).join('')}</div>`).join('')}</div>
    <div class="row" style="margin-top:12px"><button class="btn primary" data-act="inv-valider">Voir les écarts et enregistrer</button><button class="btn" data-act="inv-abandon">Abandonner</button></div>`);
}
function noterCompte(id, n) {
  const inv = lireInv(); if (!inv) return;
  inv.comptes[id] = Math.max(0, int(n) ?? 0); ecrireInv(inv);
  const l = $(`[data-vin-inv="${CSS.escape(id)}"]`);
  if (l) { const att = qte(vinById(id), inv.cave); l.classList.toggle('ok', inv.comptes[id] === att); l.classList.toggle('diff', inv.comptes[id] !== att); }
  const t = $('#sheetBody .sheet-head .sub'); if (t) t.textContent = `${Object.keys(inv.comptes).length} / ${$$('.inv-ligne').length} vins comptés`;
}
function validerInventaire(confirme) {
  const inv = lireInv(); if (!inv) return;
  computeStock();
  const ecarts = Object.entries(inv.comptes).map(([id, n]) => ({ v: vinById(id), n, att: qte(vinById(id), inv.cave) })).filter(x => x.v && x.n !== x.att);
  const nouveaux = inv.nouveaux || [], emps = Object.entries(inv.emplacements || {}).filter(([id]) => vinById(id));
  if (!confirme) {
    const moins = ecarts.reduce((a, x) => a + Math.max(0, x.att - x.n), 0), plus = ecarts.reduce((a, x) => a + Math.max(0, x.n - x.att), 0);
    openSheet(`${head('Écarts d’inventaire', caveNom(inv.cave))}
      ${ecarts.length ? `<p>${plural(ecarts.length, 'vin')} à corriger : <b>−${moins}</b> / <b>+${plural(plus, 'bouteille')}</b>.${(inv.bues || []).length ? ` Les bouteilles manquantes des vins indiqués comme bus sont enregistrées comme bues.` : ''}</p>
      <table class="stock-table">${ecarts.map(x => `<tr><td>${esc(nomComplet(x.v))}</td><td class="q">${x.att} → ${x.n}</td></tr>`).join('')}</table>
      <p class="muted small">Chaque écart est noté « Inventaire » dans le journal ; on peut l'annuler depuis le journal.</p>` : '<p>Aucun écart : tout ce que vous avez compté correspond au stock.</p>'}
      ${nouveaux.length ? `<div class="section-title">Vins trouvés en cave, sans fiche</div><table class="stock-table">${nouveaux.map(x => `<tr><td>${esc(nomComplet(x.vin))}<div class="muted small">${esc([x.vin.domaine, x.vin.appellation].filter(Boolean).join(' · '))} · nouvelle fiche</div></td><td class="q">+${x.n}</td></tr>`).join('')}</table>` : ''}
      ${emps.length ? `<p class="small">📍 Emplacement noté pour ${plural(emps.length, 'vin')} : ${esc([...new Set(emps.map(([, e]) => e))].join(', '))}.</p>` : ''}
      <div class="row" style="margin-top:12px"><button class="btn primary" data-act="inv-enregistrer">${ecarts.length || nouveaux.length ? 'Enregistrer les corrections' : 'Terminer l’inventaire'}</button><button class="btn" data-act="inventaire">Continuer à compter</button></div>`);
    return;
  }
  commit(() => {
    const bues = new Set(inv.bues || []);
    ecarts.forEach(x => data.mouvements.push(bues.has(x.v.id) && x.n < x.att
      ? { id: uid('m'), vinId: x.v.id, type: 'sortie', date: null, qte: x.att - x.n, cave: inv.cave, occasion: 'Bue (constaté à l’inventaire du ' + fmtDate(today()) + ')', maj: nowIso() }
      : { id: uid('m'), vinId: x.v.id, type: 'ajustement', date: today(), qte: x.n - x.att, cave: inv.cave, note: 'Inventaire', maj: nowIso() }));
    for (const x of nouveaux) {
      data.vins.push({ emplacements: {}, aVerifier: '', format: 'Bouteille', ...x.vin, maj: nowIso() });
      if (x.n > 0) data.mouvements.push({ id: uid('m'), vinId: x.vin.id, type: 'ajustement', date: today(), qte: x.n, cave: inv.cave, note: 'Inventaire : trouvé en cave', maj: nowIso() });
    }
    for (const [id, e] of emps) { const v = vinById(id); v.emplacements[inv.cave] = e; v.maj = nowIso(); }
  }, ecarts.length || nouveaux.length ? `Inventaire enregistré : ${plural(ecarts.length + nouveaux.length, 'vin corrigé', 'vins corrigés')}` : 'Inventaire terminé');
  ecrireInv(null);
  closeSheet();
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

function formVin(id, mode, pre) {
  const src = id ? vinById(id) : null;
  const edit = mode === 'edit';
  const v = src ? { ...src } : { region: '', domaine: '', appellation: '', couleur: 'Rouge', cuvee: '', millesime: null, format: 'Bouteille', cepages: '', gardeDebut: null, gardeFin: null, prixEstime: null, notes: '' };
  if (pre) for (const k of Object.keys(v)) if (pre[k] !== undefined && pre[k] !== null && pre[k] !== '') v[k] = pre[k];
  const photoScan = pre && pre.photoFile;
  if (mode === 'millesime') { v.millesime = null; v.gardeDebut = null; v.gardeFin = null; v.prixEstime = null; }
  const def = ui.cave !== 'all' ? ui.cave : data.caves[0].id;
  const titre = edit ? 'Modifier la fiche' : mode === 'millesime' ? 'Autre millésime' : pre ? 'Nouveau vin (étiquette lue)' : 'Nouveau vin';
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
    <div class="two"><label>Valeur estimée actuelle (€ / btl)<input type="number" name="prixEstime" min="0" step="0.01" inputmode="decimal" value="${esc(v.prixEstime ?? '')}"></label>
    <label>M'alerter sous (btl)<input type="number" name="seuilAlerte" min="0" inputmode="numeric" value="${esc(v.seuilAlerte ?? '')}" placeholder="Auto · 0 = jamais"></label></div>
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
      gardeDebut: int(fd.gardeDebut), gardeFin: int(fd.gardeFin), prixEstime: num(fd.prixEstime), seuilAlerte: int(fd.seuilAlerte), notes: fd.notes.trim(), maj: nowIso()
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
    if (photoScan) enregistrerPhoto(nv, photoScan).then(r => commit(() => { Object.assign(nv, r); nv.maj = nowIso(); })).catch(() => {});
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
// Fichier d'achats : { "achats": [ { domaine, cuvee, appellation, region, couleur, millesime, format, cepages,
//   gardeDebut, gardeFin, qte, prix, date, cave, fournisseur, emplacement } ] }
// Chaque achat rejoint la fiche existante (même vin, même millésime, même format) ou crée une fiche.
function cleVin(x) { return [norm(x.domaine), norm(x.cuvee), norm(x.appellation), x.millesime || '', x.format || 'Bouteille'].join('|'); }
function trouverCave(nom) {
  if (!nom) return ui.cave !== 'all' ? ui.cave : data.caves[0].id;
  return (data.caves.find(c => c.id === nom || norm(c.nom) === norm(nom)) || {}).id || null;
}
function importerAchats(achats) {
  if (!data) { toast('Chargez d’abord votre cave'); return; }
  computeStock();
  const lignes = achats.filter(a => a && a.domaine && int(a.qte) > 0).map(a => {
    const fiche = vinsActifs().find(v => cleVin(v) === cleVin({ ...a, format: a.format || 'Bouteille' }));
    return { a, fiche, cave: trouverCave(a.cave) };
  });
  const inconnues = lignes.filter(l => !l.cave);
  if (!lignes.length) { toast('Aucun achat lisible dans ce fichier'); return; }
  if (inconnues.length) { toast(`Cave inconnue : ${inconnues[0].a.cave}`); return; }
  const total = lignes.reduce((s, l) => s + int(l.a.qte), 0);
  openSheet(`${head('Ajouter les achats', `${plural(total, 'bouteille')} · ${plural(lignes.length, 'vin')}`)}
    <div class="achats">${lignes.map(({ a, fiche, cave }) => `<div class="achat"><div><b>${esc(nomComplet({ ...a, format: a.format || 'Bouteille' }))}</b>
      <div class="muted small">${esc([a.domaine !== (a.cuvee || a.appellation) ? a.domaine : '', caveNom(cave), a.emplacement ? '📍 ' + a.emplacement : '', a.prix != null ? EUR2.format(a.prix) + '/btl' : 'prix ?'].filter(Boolean).join(' · '))}</div>
      <div class="small ${fiche ? '' : 'nouveau'}">${fiche ? `Ajouté à la fiche existante (${qte(fiche)} en stock)` : 'Nouvelle fiche'}</div></div>
      <span class="achat-q">+${int(a.qte)}</span></div>`).join('')}</div>
    <div class="row" style="margin-top:12px"><button class="btn primary" data-act="valider-achats">Ajouter au stock</button><button class="btn" data-act="close">Annuler</button></div>`);
  importerAchats.enAttente = lignes;
}
function validerAchats() {
  const lignes = importerAchats.enAttente || [];
  importerAchats.enAttente = null;
  let nouveaux = 0, btl = 0;
  commit(() => {
    for (const { a, fiche, cave } of lignes) {
      let v = fiche;
      if (!v) {
        v = { id: uid('v'), domaine: a.domaine.trim(), cuvee: (a.cuvee || '').trim(), appellation: (a.appellation || '').trim(), region: a.region || '', couleur: COULEURS.includes(a.couleur) ? a.couleur : 'Rouge',
          millesime: int(a.millesime), format: FORMATS.includes(a.format) ? a.format : 'Bouteille', cepages: a.cepages || '', gardeDebut: int(a.gardeDebut), gardeFin: int(a.gardeFin),
          prixEstime: num(a.prixEstime), notes: a.notes || '', emplacements: {}, aVerifier: a.aVerifier || '', maj: nowIso() };
        data.vins.push(v); nouveaux++;
      }
      if (a.emplacement) { v.emplacements[cave] = a.emplacement; v.maj = nowIso(); }
      data.mouvements.push({ id: uid('m'), vinId: v.id, type: 'entree', date: a.date || today(), qte: int(a.qte), cave, prix: num(a.prix), fournisseur: (a.fournisseur || '').trim(), maj: nowIso() });
      btl += int(a.qte);
    }
  }, `${plural(btl, 'bouteille ajoutée', 'bouteilles ajoutées')}${nouveaux ? ` · ${plural(nouveaux, 'nouvelle fiche', 'nouvelles fiches')}` : ''}`);
  closeSheet();
}
// Historique : { "historique": { "vins": [...], "mouvements": [...] } } ajoute fiches et mouvements absents (par id), sans rien écraser.
function importerHistorique(h) {
  if (!data) { toast('Chargez d’abord votre cave'); return; }
  const vins = (h.vins || []).filter(v => v && v.id && !data.vins.some(x => x.id === v.id));
  const mvts = (h.mouvements || []).filter(m => m && m.id && m.vinId && !data.mouvements.some(x => x.id === m.id));
  if (!vins.length && !mvts.length) { toast('Cet historique est déjà dans la cave'); return; }
  const bues = mvts.filter(m => m.type === 'sortie').reduce((a, m) => a + m.qte, 0);
  const ent = mvts.filter(m => m.type === 'entree').reduce((a, m) => a + m.qte, 0);
  openSheet(`${head('Reprendre l’historique', h.titre || '')}
    <p>${plural(bues, 'bouteille bue', 'bouteilles bues')} vont être ajoutées à l'historique${vins.length ? `, dont celles de ${plural(vins.length, 'vin entièrement bu', 'vins entièrement bus')} (nouvelles fiches, épuisées)` : ''}.</p>
    <p class="muted small">Pour chaque vin, l'achat complet est rétabli (+${ent}) et les bouteilles bues sont sorties (−${bues}) : le stock actuel ne change pas${ent - bues ? ` (écart ${ent - bues})` : ''}. L'analyse et les alertes tiennent compte de ces bouteilles.</p>
    <div class="row" style="margin-top:12px"><button class="btn primary" data-act="valider-historique">Ajouter l'historique</button><button class="btn" data-act="close">Annuler</button></div>`);
  importerHistorique.enAttente = { vins, mvts, bues };
}
function validerHistorique() {
  const x = importerHistorique.enAttente; importerHistorique.enAttente = null;
  if (!x) return;
  const now = nowIso();
  commit(() => {
    for (const v of x.vins) data.vins.push({ emplacements: {}, aVerifier: '', ...v, maj: now });
    for (const m of x.mvts) data.mouvements.push({ ...m, maj: now });
  }, `Historique ajouté : ${plural(x.bues, 'bouteille bue', 'bouteilles bues')}`);
  closeSheet();
}
// Vins encore en cave alors que l'inventaire les a passés à « bus » : { "restaurer": { "<idVin>": nombre réel, ... } }
function importerRestauration(r) {
  if (!data) { toast('Chargez d’abord votre cave'); return; }
  computeStock();
  const lignes = Object.entries(r).map(([id, n]) => ({ v: vinById(id), n: Math.max(0, int(n) ?? 0),
    sorties: mvtsActifs().filter(m => m.vinId === id && m.type === 'sortie' && String(m.occasion || '').startsWith('Bue (constaté')) })).filter(x => x.v);
  if (!lignes.length) { toast('Aucun vin correspondant'); return; }
  openSheet(`${head('Remettre en stock', 'Vins retirés à tort par l’inventaire')}
    <table class="stock-table">${lignes.map(x => `<tr><td>${esc(nomComplet(x.v))}<div class="muted small">${x.sorties.length ? 'la sortie « bue à l’inventaire » est annulée' : 'correction du stock'}</div></td><td class="q">${qte(x.v)} → ${x.n}</td></tr>`).join('')}</table>
    <div class="row" style="margin-top:12px"><button class="btn primary" data-act="valider-restauration">Remettre en stock</button><button class="btn" data-act="close">Annuler</button></div>`);
  importerRestauration.enAttente = lignes;
}
function validerRestauration() {
  const lignes = importerRestauration.enAttente || []; importerRestauration.enAttente = null;
  commit(() => {
    for (const x of lignes) x.sorties.forEach(m => { m.supprime = true; m.maj = nowIso(); });
    computeStock();
    for (const x of lignes) { const d = x.n - qte(x.v); if (d) data.mouvements.push({ id: uid('m'), vinId: x.v.id, type: 'ajustement', date: today(), qte: d, cave: data.caves[0].id, note: 'Inventaire (correction)', maj: nowIso() }); }
  }, `${plural(lignes.length, 'vin remis', 'vins remis')} en stock`);
  closeSheet();
}
function importerDonnees(obj) {
  if (obj && obj.restaurer) { importerRestauration(obj.restaurer); return; }
  // Comptages dictés : { "inventaire": { "cave": "vieux-conde", "comptes": { "<idVin>": 4, ... } } } → écran des écarts à valider
  if (obj && obj.inventaire && obj.inventaire.comptes) {
    if (!data) { toast('Chargez d’abord votre cave'); return; }
    const inv = lireInv() && lireInv().cave === obj.inventaire.cave ? lireInv() : { cave: obj.inventaire.cave || data.caves[0].id, debut: today(), comptes: {} };
    for (const [id, n] of Object.entries(obj.inventaire.comptes)) if (vinById(id)) inv.comptes[id] = Math.max(0, int(n) ?? 0);
    inv.nouveaux = (obj.inventaire.nouveaux || []).filter(x => x && x.vin && x.vin.id && !vinById(x.vin.id));
    inv.emplacements = obj.inventaire.emplacements || {};
    inv.bues = obj.inventaire.bues || []; // vins dont les bouteilles manquantes ont été bues (sorties plutôt que corrections)
    ecrireInv(inv); validerInventaire(false); return;
  }
  if (obj && obj.historique) { importerHistorique(obj.historique); return; }
  if (obj && Array.isArray(obj.achats)) { importerAchats(obj.achats); return; }
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
  const pt = e.target.closest('circle[data-vin]');
  if (pt) { ouvrirVin(pt.dataset.vin); return; }
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
  if (t.dataset.periode) { ui.periode = t.dataset.periode; render(); return; }
  if (t.dataset.vin) { ouvrirVin(t.dataset.vin); return; }
  if (t.dataset.mvt) { ouvrirMvt(t.dataset.mvt); return; }
  if (t.id === 'fab') { formVin(null, 'new'); return; }
  if (t.id === 'fabScan') { scannerEtiquette(); return; }
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
    case 'scanner': scannerEtiquette(); break;
    case 'scan-boire': case 'scan-ajouter': {
      const v = vinById(id);
      photoScanSurFiche(v);
      if (t.dataset.act === 'scan-boire') formBoire(id); else formEntree(id);
      break;
    }
    case 'scan-completer': {
      const v = vinById(id);
      const c = complementsPossibles(v, scan.lu);
      commit(() => { Object.assign(v, c); v.maj = nowIso(); }, 'Fiche complétée');
      afficherScan();
      break;
    }
    case 'scan-nouveau': {
      const lu = scan.lu;
      const coche = $('#scanPhoto');
      formVin(null, 'new', { domaine: lu.domaine, cuvee: lu.cuvee, appellation: lu.appellation, region: lu.region, couleur: lu.couleur, millesime: lu.millesime, format: lu.format, cepages: lu.cepages, gardeDebut: lu.gardeDebut, gardeFin: lu.gardeFin, prixEstime: lu.prixIndicatif, photoFile: coche && !coche.checked ? null : scan.file });
      break;
    }
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
    case 'valider-achats': validerAchats(); break;
    case 'valider-historique': validerHistorique(); break;
    case 'valider-restauration': validerRestauration(); break;
    case 'photos': assistantPhotos(); break;
    case 'assist-photo': choisirPhoto(id, assistantPhotos); break;
    case 'doublons': ouvrirDoublons(); break;
    case 'garder-une': garderUne([ouvrirDoublons.etat.entrees[+id]]); ouvrirDoublons(); break;
    case 'garder-une-tout': garderUne(ouvrirDoublons.etat.entrees); ouvrirDoublons(); break;
    case 'fusionner': fusionnerFiches(ouvrirDoublons.etat.fiches[+id]); ouvrirDoublons(); break;
    case 'inventaire': ouvrirInventaire(); break;
    case 'inventaire-cave': ouvrirInventaire(id); break;
    case 'inv-ok': { const inp = $(`.inv-n[data-id="${CSS.escape(id)}"]`); noterCompte(id, inp ? inp.value : 0); break; }
    case 'inv-groupe-ok': $$(`.inv-ligne[data-emp="${CSS.escape(id)}"] .inv-n`).forEach(i => noterCompte(i.dataset.id, i.value)); break;
    case 'inv-valider': validerInventaire(false); break;
    case 'inv-enregistrer': validerInventaire(true); break;
    case 'inv-abandon': ecrireInv(null); closeSheet(); toast('Inventaire abandonné'); render(); break;
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
  if (e.target.classList.contains('assist-url') && /^https?:\/\/\S+\.\S+/.test(e.target.value.trim())) { enregistrerUrlPhoto(e.target.dataset.id, e.target.value); return; }
  if (e.target.classList.contains('inv-n') && e.target.value !== '') { noterCompte(e.target.dataset.id, e.target.value); return; }
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
  scanDepuisAdresse();
  if (account) await sync();
  if (scanEnAttente && data) { const lu = scanEnAttente; scanEnAttente = null; scan = { file: null, apercu: '', lu, resultats: chercherCorrespondances(lu) }; afficherScan(); }
})();
window.addEventListener('hashchange', scanDepuisAdresse);
// Coller un code « MACAVE:… » n'importe où dans l'appli
document.addEventListener('paste', e => {
  if (e.target.closest && e.target.closest('input, textarea') && e.target.id !== 'codeScan') return;
  const t = (e.clipboardData || window.clipboardData)?.getData('text') || '';
  if (/MACAVE:|#scan=/.test(t) && recevoirScan(t)) e.preventDefault();
});

window.cave = { importerDonnees, get data() { return data; } };
