/* Wildschaden Feld – Stufe 1
   Erfassung am Feld, Vorschau-Berechnung, Export in die Excel-Vorlage. Alles offline. */
'use strict';
const APP_VERSION = '1.2.1';

/* ---------- Auswahllisten (aus Blatt „Listen“ der Vorlage) ---------- */
const L = {
  wildart: ['Schwarzwild', 'Rotwild', 'Damwild', 'Rehwild', 'Muffelwild', 'Sikawild', 'Elchwild', 'Wisent', 'Gamswild', 'Steinwild', 'Wildkaninchen', 'Fasan'],
  verfahren: ['Gütliche Einigung', 'Vorverfahren Amt/Gemeinde'],
  methode: ['Probedrusch', 'Betriebsdurchschnitt', 'Schätzung am Bestand'],
  quelle: ['KTBL/MR-Satz', 'Eigene Kosten', 'Lohnunternehmer'],
  saatmethode: ['kg/ha × Preis', 'Pauschal €/ha', 'Laut Rechnung'],
  ustGesch: ['Netto (regelbesteuert)', 'Pauschalierend (Durchschnittssatz)'],
  termin: ['Vorbesichtigung', 'Endschätzung', 'Nachschätzung', 'Sonstiger Termin'],
  taetigkeit: ['Vorbereitung', 'Ortstermin / Schätzung', 'Nachbereitung / Protokoll', 'Sonstiges'],
  grad: ['leicht', 'mittel', 'schwer'],
  status: ['offen', 'geschätzt', 'abgerechnet'],
  empfaenger: ['Amt/Gemeinde', 'Geschädigter', 'Ersatzpflichtiger', 'Sonstiger'],
  ustRechnung: ['Kleinunternehmer (§ 19 UStG)', 'Regelbesteuert'],
  sigRolle: ['Geschädigter', 'Ersatzpflichtiger', 'Zeuge', 'Vertreter Jagdvorstand', 'Sonstige Person']
};
const SIG_TEXTE = ['Anwesenheit beim Ortstermin bestätigt.', 'Protokoll zur Kenntnis genommen.', 'Mit dem Ergebnis der Schätzung einverstanden (gütliche Einigung).'];
const GRAD_MASSNAHME = { leicht: 'Schleppen / Walzen', mittel: 'Schleppen + Walzen + Nachsaat', schwer: 'Pflügen / Fräsen + Neuansaat' };

/* Werte aus deiner Vorlage (Stand der hochgeladenen Datei). Beim Laden der Vorlage werden sie automatisch ersetzt. */
const DEFAULT_SETTINGS = {
  stamm: { name: 'Christian Zaunig', strasse: 'Hainstraße 8', telefon: '0160 6498407', email: 'zaunigchristian@gmail.com',
    steuernr: '', iban: 'DE78 1806 2678 0000 2995 96', bank: '', bestellt: '', ustModus: 'Kleinunternehmer (§ 19 UStG)', ort: '03238 Lieskau',
    stundensatz: 40, tagesmax: 200, kmgeld: 0.3, ust: 0.19, durchschnitt: 0.078, frist: 7 },
  preise: [
    ['Winterweizen (Brotweizen)', 'Getreide', 't Korn', 200, 80, 150], ['Winterweizen (Futterweizen)', 'Getreide', 't Korn', 185, 80, 150],
    ['Winterroggen', 'Getreide', 't Korn', 160, 80, 150], ['Wintergerste', 'Getreide', 't Korn', 170, 80, 150],
    ['Sommergerste / Braugerste', 'Getreide', 't Korn', 220, 80, 150], ['Triticale', 'Getreide', 't Korn', 165, 80, 150],
    ['Hafer', 'Getreide', 't Korn', 170, 80, 150], ['Silomais', 'Mais', 't FM', 40, null, 200], ['Körnermais', 'Mais', 't Korn', 190, null, 250],
    ['Grünland Wiese/Mähweide', 'Grünland', 't FM', 30, null, 100], ['Grünland Weide', 'Grünland', 't FM', 25, null, 40],
    ['Ackergras / Feldgras', 'Grünland', 't FM', 30, null, 100], ['Winterraps', 'Sonstige', 't Korn', 450, null, 150], ['Sonstige Kultur', 'Sonstige', 't', null, null, null]
  ].map(([kultur, gruppe, einheit, preis, stroh, erspart]) => ({ kultur, gruppe, einheit, preis, stroh, erspart })),
  ktbl: [['Schleppen', 25], ['Walzen', 25], ['Schleppen + Walzen + Nachsaat (Kombination)', 80], ['Fräsen', 120], ['Grubbern', 45],
    ['Nachsaat (Aussaat)', 40], ['Pflügen', 90], ['Saatbettbereitung', 50], ['Neuansaat (Aussaat)', 45], ['Sonstiger Arbeitsgang', null]]
    .map(([gang, satz]) => ({ gang, satz })),
  saatgut: [['Grünland-Nachsaatmischung', 25, 6], ['Grünland-Neuansaatmischung', 40, 5], ['Getreide-Saatgut', 160, 0.8], ['Mais-Saatgut (kg)', 25, 12], ['Sonstiges Saatgut', null, null]]
    .map(([sorte, kgha, preis]) => ({ sorte, kgha, preis })),
  vorlageName: null, vorlageDatum: null, quelle: 'Werte aus der Vorlage vom 13.10.2026 (eingebaut)'
};

/* ---------- Speicher (IndexedDB) ---------- */
const DB = {
  db: null,
  open() {
    return new Promise((res, rej) => {
      const r = indexedDB.open('wildschaden-feld', 1);
      r.onupgradeneeded = () => {
        const d = r.result;
        d.createObjectStore('faelle', { keyPath: 'id' });
        d.createObjectStore('kontakte', { keyPath: 'id' });
        d.createObjectStore('fotos', { keyPath: 'id' }).createIndex('fallId', 'fallId');
        d.createObjectStore('kv');
      };
      r.onsuccess = () => { DB.db = r.result; res(); };
      r.onerror = () => rej(r.error);
    });
  },
  tx(store, mode, fn) {
    return new Promise((res, rej) => {
      const t = DB.db.transaction(store, mode); const s = t.objectStore(store); let out;
      const r = fn(s); if (r) r.onsuccess = () => { out = r.result; };
      t.oncomplete = () => res(out); t.onerror = () => rej(t.error); t.onabort = () => rej(t.error);
    });
  },
  all: s => DB.tx(s, 'readonly', st => st.getAll()),
  get: (s, k) => DB.tx(s, 'readonly', st => st.get(k)),
  put: (s, v, k) => DB.tx(s, 'readwrite', st => k === undefined ? st.put(v) : st.put(v, k)),
  del: (s, k) => DB.tx(s, 'readwrite', st => st.delete(k)),
  byIndex: (s, i, v) => DB.tx(s, 'readonly', st => st.index(i).getAll(v))
};

/* ---------- Helfer ---------- */
const $ = s => document.querySelector(s);
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const today = () => { const d = new Date(); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); };
function num(v) { if (v === '' || v == null) return null; const x = Number(String(v).replace(/\s/g, '').replace(',', '.')); return isNaN(x) ? null : x; }
const fmt = (v, d = 2) => v == null || isNaN(v) ? '–' : v.toLocaleString('de-DE', { minimumFractionDigits: d, maximumFractionDigits: d });
const eur = v => fmt(v, 2) + ' €';
const dDE = iso => iso ? iso.split('-').reverse().join('.') : '';
function getPath(o, p) { return p.split('.').reduce((a, k) => a == null ? a : a[k], o); }
function setPath(o, p, v) { const ks = p.split('.'); const last = ks.pop(); const t = ks.reduce((a, k) => a[k], o); t[last] = v; }
function toast(msg, ms = 2600) {
  const t = document.createElement('div'); t.className = 'toast'; t.textContent = msg; document.body.appendChild(t);
  setTimeout(() => t.remove(), ms);
}
function daysBetween(a, b) { return Math.round((Date.parse(b) - Date.parse(a)) / 86400000); }
function fileSafe(s) { return String(s || '').normalize('NFKD').replace(/[̀-ͯ]/g, '').replace(/ß/g, 'ss').replace(/[^\w\-]+/g, '_').replace(/_+/g, '_').replace(/^_|_$/g, '').slice(0, 40); }

/* ---------- Zustand ---------- */
let S = null;          // Einstellungen
let faelle = [];       // alle Fälle
let kontakte = [];
let cur = null;        // geöffneter Fall
let view = { name: 'list', tab: 'fall' };
let lastExport = null; // { blob, name }
let saveTimer = null;

function newFall() {
  const year = new Date().getFullYear();
  const nr = faelle.filter(f => (f.az || '').includes(String(year))).length + 1;
  return {
    id: uid(), created: Date.now(), updated: Date.now(), status: 'offen',
    az: 'WS-' + year + '-' + String(nr).padStart(3, '0'), verfahren: '', stelle: '', stelleAnschrift: '',
    gesch: { name: '', anschrift: '' }, ersatz: { name: '', anschrift: '' }, jagdbezirk: '', zeuge1: '', zeuge2: '', vertreter: '',
    gemarkung: '', flur: '', flurstueck: '', schlag: '', schlagHa: '', kultur: '', wildart: '',
    kenntnis: '', anmeldung: '', sbProz: '', sbEuro: '', deckel: '', pauschale: '', ustGesch: '', bemerkung: '',
    termine: [], stellen: [newStelle()],
    ertrag: { methode: '', proben: [{}, {}, {}, {}, {}], betrieb: '', schaetzung: '', stroh: '' },
    wieder: [], saat: [], zeiten: [], fotoSeq: 0,
    prot: { ort: '', datum: '', fotos: true, karte: true }, sigs: [], rech: { empf: '', sonstName: '', sonstAnschrift: '', nr: '', datum: '', ust: '' }
  };
}
/* ältere Fälle (Stufe 1) um neue Felder ergänzen */
function normFall(f) {
  f.prot = Object.assign({ ort: '', datum: '', fotos: true, karte: true }, f.prot || {});
  f.sigs = f.sigs || [];
  f.rech = Object.assign({ empf: '', sonstName: '', sonstAnschrift: '', nr: '', datum: '', ust: '' }, f.rech || {});
  return f;
}
function newStelle() { return { id: uid(), modus: 'lb', laenge: '', breite: '', poly: [], verlust: '', grad: '', lat: '', lon: '', acc: '', fotos: [], termin: '', bem: '' }; }

function save() {
  if (!cur) return;
  cur.updated = Date.now();
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => DB.put('faelle', cur).catch(e => toast('Speichern fehlgeschlagen: ' + e.message)), 300);
}
async function saveSettings() { await DB.put('kv', S, 'settings'); }

/* ---------- Berechnung (Vorschau – maßgeblich bleibt Excel) ---------- */
function polyArea(pts) {
  if (!pts || pts.length < 3) return 0;
  const lat0 = pts.reduce((a, p) => a + p[0], 0) / pts.length, lon0 = pts.reduce((a, p) => a + p[1], 0) / pts.length;
  const kx = 111320 * Math.cos(lat0 * Math.PI / 180), ky = 110574;
  const xy = pts.map(p => [(p[1] - lon0) * kx, (p[0] - lat0) * ky]);
  let a = 0; for (let i = 0; i < xy.length; i++) { const [x1, y1] = xy[i], [x2, y2] = xy[(i + 1) % xy.length]; a += x1 * y2 - x2 * y1; }
  return Math.abs(a) / 2;
}
function stelleFlaeche(s) { return s.modus === 'gps' ? polyArea(s.poly) : ((num(s.laenge) || 0) * (num(s.breite) || 0)); }
function preisRow(kultur) { return S.preise.find(p => p.kultur === kultur) || null; }
function gruppe(f) { const p = preisRow(f.kultur); return p ? p.gruppe : ''; }

function calc(f) {
  const r = {};
  const st = f.stellen.filter(s => stelleFlaeche(s) > 0);
  r.anzahl = st.length;
  r.flaeche = st.reduce((a, s) => a + stelleFlaeche(s), 0);
  r.eff = st.reduce((a, s) => { const v = num(s.verlust); return a + stelleFlaeche(s) * (v == null ? 1 : v / 100); }, 0);
  r.ha = r.eff / 10000;
  r.grad = { leicht: 0, mittel: 0, schwer: 0 };
  st.forEach(s => { if (r.grad[s.grad] != null) r.grad[s.grad] += stelleFlaeche(s); });
  const e = f.ertrag;
  const proben = e.proben.map(p => (num(p.kg) && num(p.m2)) ? num(p.kg) / num(p.m2) * 10 : null).filter(x => x != null);
  r.probenMittel = proben.length ? proben.reduce((a, b) => a + b, 0) / proben.length : null;
  r.ertrag = e.methode === 'Probedrusch' ? r.probenMittel : e.methode === 'Betriebsdurchschnitt' ? num(e.betrieb) : e.methode === 'Schätzung am Bestand' ? num(e.schaetzung) : null;
  const p = preisRow(f.kultur); r.gruppe = p ? p.gruppe : '';
  r.preis = p ? p.preis : null; r.strohpreis = p ? p.stroh : null; r.erspartHa = p ? p.erspart : null;
  r.wert = (r.ertrag || 0) * (r.preis || 0) * r.ha;
  r.stroh = r.gruppe === 'Getreide' ? (num(e.stroh) || 0) * (r.strohpreis || 0) * r.ha : 0;
  r.erspart = (r.erspartHa || 0) * r.ha;
  r.ertragsschaden = r.wert + r.stroh - r.erspart;
  r.wieder = f.wieder.map(w => {
    const ha = num(w.ha) || 0, q = w.quelle || 'KTBL/MR-Satz';
    if (q === 'Lohnunternehmer') return num(w.lu) || 0;
    if (q === 'Eigene Kosten') return (num(w.eigen) || 0) * ha;
    const k = S.ktbl.find(x => x.gang === w.gang); return ((k && k.satz) || 0) * ha;
  });
  r.wiederSumme = r.wieder.reduce((a, b) => a + b, 0);
  r.saat = f.saat.map(s => {
    const ha = num(s.ha) || 0, m = s.methode || 'kg/ha × Preis';
    if (m === 'Laut Rechnung') return num(s.rechnung) || 0;
    if (m === 'Pauschal €/ha') return (num(s.pauschal) || 0) * ha;
    const g = S.saatgut.find(x => x.sorte === s.sorte); return g ? (g.kgha || 0) * (g.preis || 0) * ha : 0;
  });
  r.saatSumme = r.saat.reduce((a, b) => a + b, 0);
  r.summe = r.ertragsschaden + r.wiederSumme + r.saatSumme;
  r.sb = Math.min(Math.max(r.summe, 0), r.summe * (num(f.sbProz) || 0) / 100 + (num(f.sbEuro) || 0));
  const nachSb = r.summe - r.sb;
  const deckel = num(f.deckel);
  r.kuerzung = deckel && nachSb > deckel ? nachSb - deckel : 0;
  r.netto = nachSb - r.kuerzung;
  r.zuschlag = (f.ustGesch || '').startsWith('Pauschalierend') ? r.netto * (S.stamm.durchschnitt || 0) : 0;
  r.gesamt = r.netto + r.zuschlag;
  // Honorar
  const tage = {}; let km = 0, std = 0;
  f.zeiten.filter(z => !z.start).forEach(z => { const h = num(z.stunden) || 0; std += h; tage[z.datum || '?'] = (tage[z.datum || '?'] || 0) + h; km += num(z.km) || 0; });
  r.std = std; r.km = km;
  r.honorarRoh = std * S.stamm.stundensatz;
  r.honorar = Object.values(tage).reduce((a, h) => a + Math.min(h * S.stamm.stundensatz, S.stamm.tagesmax), 0);
  r.fahrt = km * S.stamm.kmgeld;
  return r;
}
function fristInfo(f) {
  const frist = S.stamm.frist || 7;
  if (f.kenntnis && f.anmeldung) {
    const d = daysBetween(f.kenntnis, f.anmeldung);
    if (d < 0) return ['bad', 'Anmeldung liegt vor der Kenntnis – Daten prüfen.'];
    if (d > frist) return ['bad', 'Anmeldung ' + d + ' Tage nach Kenntnis – Frist von ' + frist + ' Tagen überschritten (§ 34 BJagdG).'];
    return ['ok', 'Frist eingehalten: Anmeldung ' + d + ' Tag(e) nach Kenntnis.'];
  }
  if (f.kenntnis && !f.anmeldung) {
    const d = daysBetween(f.kenntnis, today());
    return d > frist ? ['warn', 'Noch kein Anmeldedatum – seit Kenntnis sind ' + d + ' Tage vergangen.'] : ['info', 'Anmeldedatum fehlt noch.'];
  }
  return ['info', 'Daten eintragen, dann wird die Frist geprüft.'];
}

/* ---------- Bausteine ---------- */
function F(label, k, o = {}) {
  const v = getPath(cur, k) ?? '';
  const t = o.type || 'text';
  const im = o.num ? ' inputmode="decimal"' : '';
  if (t === 'textarea') return `<label><span>${esc(label)}</span><textarea data-k="${k}" rows="3">${esc(v)}</textarea></label>`;
  return `<label><span>${esc(label)}</span><input type="${t}" data-k="${k}" value="${esc(v)}"${im}${o.ph ? ` placeholder="${esc(o.ph)}"` : ''}></label>`;
}
function Sel(label, k, opts, o = {}) {
  const v = getPath(cur, k) ?? '';
  const list = opts.includes(v) || !v ? opts : [v, ...opts];
  return `<label><span>${esc(label)}</span><select data-k="${k}"${o.re ? ' data-re="1"' : ''}><option value="">– wählen –</option>${list.map(x => `<option${x === v ? ' selected' : ''}>${esc(x)}</option>`).join('')}</select></label>`;
}

/* ---------- Ansichten ---------- */
const TABS = [['fall', 'Fall'], ['termine', 'Termine'], ['stellen', 'Schadstellen'], ['ertrag', 'Ertrag'], ['kosten', 'Wiederherstellung'], ['zeiten', 'Zeiten'], ['export', 'Ergebnis & Export']];

function render() {
  const main = $('#main'); const tabs = $('#tabs');
  $('#back').classList.toggle('hidden', view.name === 'list');
  $('#settingsBtn').classList.toggle('hidden', view.name === 'settings');
  if (view.name === 'list') { $('#title').textContent = 'Wildschadensfälle'; tabs.classList.add('hidden'); main.innerHTML = vList(); }
  else if (view.name === 'settings') { $('#title').textContent = 'Einstellungen'; tabs.classList.add('hidden'); main.innerHTML = vSettings(); }
  else {
    $('#title').textContent = cur.az || 'Fall';
    tabs.classList.remove('hidden');
    tabs.innerHTML = TABS.map(([k, t]) => `<button data-tab="${k}" class="${view.tab === k ? 'on' : ''}">${t}</button>`).join('');
    main.innerHTML = ({ fall: vFall, termine: vTermine, stellen: vStellen, ertrag: vErtrag, kosten: vKosten, zeiten: vZeiten, export: vExport })[view.tab]();
    const on = tabs.querySelector('.on'); if (on) on.scrollIntoView({ inline: 'center', block: 'nearest' });
    if (view.tab === 'stellen') loadThumbs();
  }
  live();
}

function vList() {
  const q = (view.q || '').toLowerCase();
  const list = faelle.slice().sort((a, b) => b.updated - a.updated).filter(f => !q ||
    [f.az, f.gesch.name, f.ersatz.name, f.gemarkung, f.flur, f.flurstueck, f.schlag, f.kultur].join(' ').toLowerCase().includes(q))
    .filter(f => !view.st || f.status === view.st);
  const cnt = s => faelle.filter(f => f.status === s).length;
  return `
  <label><input type="search" id="search" placeholder="Suchen: Name, Flurstück, Az …" value="${esc(view.q || '')}"></label>
  <div class="seg">${[['', 'Alle (' + faelle.length + ')'], ...L.status.map(s => [s, s[0].toUpperCase() + s.slice(1) + ' (' + cnt(s) + ')'])].map(([k, t]) => `<button data-act="filter" data-v="${k}" class="${(view.st || '') === k ? 'on' : ''}">${t}</button>`).join('')}</div>
  ${faelle.length && (!S.lastBackupTs || Date.now() - S.lastBackupTs > 7 * 86400000) ? `<div class="note warn">${S.lastBackupTs ? 'Letzte Datensicherung ist über eine Woche her.' : 'Noch keine Datensicherung gemacht.'} <button class="btn sm" data-act="goBackup">💾 Jetzt sichern</button></div>` : ''}
  <div class="btns" style="margin:0 0 10px"><label class="btn sm" style="margin:0">📥 Excel-Datei einlesen<input type="file" id="impIn" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" class="hidden"></label></div>
  ${S.vorlageName ? '' : `<div class="note warn">Noch keine Excel-Vorlage geladen. Erfassen geht trotzdem – für den Export einmal unter ⚙ Einstellungen die Vorlage laden.</div>`}
  <div class="list">${list.map(f => {
    const r = calc(f); const fr = fristInfo(f)[0] === 'bad';
    return `<button class="item" data-act="open" data-id="${f.id}">
      <div style="display:flex;justify-content:space-between;gap:8px"><span class="t">${esc(f.az || 'ohne Az.')}</span><span class="pill ${f.status === 'abgerechnet' ? 'done' : f.status === 'offen' ? 'warn' : ''}">${esc(f.status)}</span></div>
      <div>${esc(f.gesch.name || '—')}${f.kultur ? ' · ' + esc(f.kultur) : ''}</div>
      <div class="s">${esc([f.gemarkung, f.flur && 'Flur ' + f.flur, f.flurstueck && 'Fl.st. ' + f.flurstueck].filter(Boolean).join(', ') || 'Fläche noch nicht erfasst')}</div>
      <div class="s">${r.anzahl} Schadstelle(n) · ${fmt(r.eff, 0)} m² wirksam · ${eur(r.gesamt)}${fr ? ' · <b style="color:var(--bad)">Frist!</b>' : ''}</div>
    </button>`;
  }).join('') || `<p class="muted">${faelle.length ? 'Nichts gefunden.' : 'Noch keine Fälle. Unten rechts einen neuen Fall anlegen.'}</p>`}</div>
  <button class="btn pri fab" data-act="new">＋ Neuer Fall</button>`;
}

function contactBtn(prefix) { return `<div class="btns"><button class="btn sm" data-act="pick" data-p="${prefix}">📇 Aus Adressbuch</button>${'contacts' in navigator && 'ContactsManager' in window ? `<button class="btn sm" data-act="phone" data-p="${prefix}">📱 Handy-Kontakt</button>` : ''}<button class="btn sm" data-act="savec" data-p="${prefix}">＋ Ins Adressbuch</button></div>`; }

function vFall() {
  const fr = fristInfo(cur);
  return `
  <div class="card"><h2>Status</h2>
    <div class="seg">${L.status.map(s => `<button data-act="status" data-v="${s}" class="${cur.status === s ? 'on' : ''}">${s}</button>`).join('')}</div></div>
  <div class="card"><h2>Verfahren</h2>
    ${F('Aktenzeichen', 'az')}${Sel('Verfahrensart', 'verfahren', L.verfahren)}
    ${F('Verfahrensführende Stelle (Amt/Gemeinde)', 'stelle')}${F('Ansprechpartner / Anschrift Stelle', 'stelleAnschrift')}
    ${contactBtn('stelle')}</div>
  <div class="card"><h2>Beteiligte</h2>
    ${F('Geschädigter (Name)', 'gesch.name')}${F('Anschrift Geschädigter', 'gesch.anschrift')}${contactBtn('gesch')}
    <hr style="border:0;border-top:1px solid var(--line);margin:14px 0">
    ${F('Ersatzpflichtiger (Jagdausübungsberechtigter / Jagdgenossenschaft)', 'ersatz.name')}${F('Anschrift Ersatzpflichtiger', 'ersatz.anschrift')}${contactBtn('ersatz')}
    <hr style="border:0;border-top:1px solid var(--line);margin:14px 0">
    ${F('Jagdbezirk', 'jagdbezirk')}
    <div class="row">${F('Zeuge 1', 'zeuge1')}${F('Zeuge 2', 'zeuge2')}</div>
    ${F('Vertreter Jagdvorstand', 'vertreter')}</div>
  <div class="card"><h2>Fläche und Schaden</h2>
    <div class="row">${F('Gemarkung', 'gemarkung')}${F('Flur', 'flur')}</div>
    <div class="row">${F('Flurstück(e)', 'flurstueck')}${F('Schlaggröße (ha)', 'schlagHa', { num: 1 })}</div>
    ${F('Schlagname (betrieblich)', 'schlag')}
    ${Sel('Kultur', 'kultur', S.preise.map(p => p.kultur), { re: 1 })}
    ${cur.kultur ? `<p class="muted">Gruppe: <b>${esc(gruppe(cur))}</b> · Preis laut Tabelle: ${preisRow(cur.kultur) && preisRow(cur.kultur).preis != null ? eur(preisRow(cur.kultur).preis) + '/t' : 'nicht hinterlegt'}</p>` : ''}
    ${Sel('Wildart (Verursacher)', 'wildart', L.wildart)}</div>
  <div class="card"><h2>Fristen</h2>
    <div class="row">${F('Kenntnis vom Schaden', 'kenntnis', { type: 'date' })}${F('Anmeldung bei Amt/Gemeinde', 'anmeldung', { type: 'date' })}</div>
    <div class="note ${fr[0]}" data-live="frist">${esc(fr[1])}</div></div>
  <div class="card"><h2>Jagdpachtvertrag</h2><p class="muted" style="margin-top:0">Leer lassen, wenn es keine Regelung gibt.</p>
    <div class="row">${F('Selbstbehalt in %', 'sbProz', { num: 1 })}${F('Selbstbehalt in €', 'sbEuro', { num: 1 })}</div>
    <div class="row">${F('Höchstbetrag je Fall (€)', 'deckel', { num: 1 })}${F('Pauschale (€/Jahr)', 'pauschale', { num: 1 })}</div>
    <div data-live="pauschHinweis"></div></div>
  <div class="card"><h2>Umsatzsteuer beim Geschädigten</h2>${Sel('Besteuerung des Geschädigten', 'ustGesch', L.ustGesch)}</div>
  <div class="card"><h2>Bemerkungen fürs Protokoll</h2>${F('Bemerkungen', 'bemerkung', { type: 'textarea' })}
    <p class="muted">Tipp: Mikrofon auf der Tastatur antippen und diktieren.</p></div>
  <button class="btn bad wide" data-act="delFall">Fall löschen</button>`;
}

function vTermine() {
  return `<p class="muted">Beliebig viele Termine. In die Excel-Vorlage passen 5 – weitere stehen dann nur hier und später im PDF.</p>
  ${cur.termine.map((t, i) => `<div class="card">
    <div class="stelle-head"><h3>Termin ${i + 1}</h3><button class="btn sm bad" data-act="delTermin" data-i="${i}">Entfernen</button></div>
    <div class="row">${F('Datum', 'termine.' + i + '.datum', { type: 'date' })}${Sel('Art des Termins', 'termine.' + i + '.art', L.termin)}</div>
    ${F('Anwesende', 'termine.' + i + '.anwesende')}
    <div class="btns" style="margin:-4px 0 10px"><button class="btn sm" data-act="anwesend" data-i="${i}">Beteiligte übernehmen</button></div>
    ${F('Feststellungen', 'termine.' + i + '.feststellungen', { type: 'textarea' })}</div>`).join('')}
  <button class="btn pri wide" data-act="addTermin">＋ Termin hinzufügen</button>`;
}

function vStellen() {
  const r = calc(cur); const gl = gruppe(cur) === 'Grünland';
  return `<div class="card"><div class="row3">
      <div><div class="muted">Stellen</div><div class="big" data-live="anz">${r.anzahl}</div></div>
      <div><div class="muted">Fläche m²</div><div class="big" data-live="fl">${fmt(r.flaeche, 1)}</div></div>
      <div><div class="muted">wirksam m²</div><div class="big" data-live="eff">${fmt(r.eff, 1)}</div></div></div>
    <p class="muted" style="margin:6px 0 0">Leeres Verlust-Feld = 100 %. Die Excel-Vorlage nimmt bis zu 30 Stellen auf.</p>
    <div class="btns"><button class="btn" data-act="map">🗺 Lageplan anzeigen</button></div></div>
  ${cur.stellen.map((s, i) => {
    const k = 'stellen.' + i;
    return `<div class="card" id="st-${s.id}">
    <div class="stelle-head"><h3>Schadstelle ${i + 1}</h3><span class="pill" data-live="sf-${i}">${fmt(stelleFlaeche(s), 1)} m²</span></div>
    <div class="seg"><button data-act="modus" data-i="${i}" data-v="lb" class="${s.modus === 'lb' ? 'on' : ''}">Länge × Breite</button><button data-act="modus" data-i="${i}" data-v="gps" class="${s.modus === 'gps' ? 'on' : ''}">GPS-Umgang</button></div>
    ${s.modus === 'lb' ? `<div class="row">${F('Länge (m)', k + '.laenge', { num: 1 })}${F('Breite (m)', k + '.breite', { num: 1 })}</div>` :
      `<div class="note info">Eckpunkte der Schadstelle der Reihe nach ablaufen und an jeder Ecke „Punkt setzen“ tippen. GPS ist auf etwa 3–5 m genau – bei kleinen Stellen lieber Länge × Breite nehmen.</div>
       <p><b>${s.poly.length}</b> Punkt(e)${s.poly.length ? ', letzte Genauigkeit ±' + Math.round(s.poly[s.poly.length - 1][2]) + ' m' : ''}${s.poly.length >= 3 ? ' · Fläche <b>' + fmt(polyArea(s.poly), 1) + ' m²</b>' : ''}</p>
       <div class="btns"><button class="btn pri" data-act="polyAdd" data-i="${i}">📍 Punkt setzen</button><button class="btn" data-act="polyUndo" data-i="${i}" ${s.poly.length ? '' : 'disabled'}>Letzten löschen</button></div>`}
    <div class="row">${F('Verlust (%)', k + '.verlust', { num: 1, ph: '100' })}${gl ? Sel('Grünland-Schadgrad', k + '.grad', L.grad) : Sel('Ortstermin', k + '.termin', L.termin)}</div>
    ${gl ? Sel('Ortstermin', k + '.termin', L.termin) : ''}
    <div class="row"><label><span>GPS Breite (N)</span><input data-k="${k}.lat" value="${esc(s.lat)}" inputmode="decimal"></label><label><span>GPS Länge (E)</span><input data-k="${k}.lon" value="${esc(s.lon)}" inputmode="decimal"></label></div>
    <div class="btns" style="margin-top:-4px"><button class="btn" data-act="pos" data-i="${i}">📍 Position übernehmen</button>${s.acc ? `<span class="muted" style="align-self:center">±${Math.round(s.acc)} m</span>` : ''}</div>
    <div class="btns"><label class="btn" style="margin:0">📷 Foto aufnehmen<input type="file" accept="image/*" capture="environment" data-photo="${i}" class="hidden"></label>
      <label class="btn" style="margin:0">🖼 Aus Galerie<input type="file" accept="image/*" multiple data-photo="${i}" class="hidden"></label></div>
    <div class="thumbs" data-thumbs="${s.id}"></div>
    ${F('Bemerkung', k + '.bem')}
    <button class="btn sm bad" data-act="delStelle" data-i="${i}">Schadstelle entfernen</button></div>`;
  }).join('')}
  <button class="btn pri wide" data-act="addStelle">＋ Schadstelle hinzufügen</button>`;
}

function vErtrag() {
  const r = calc(cur); const e = cur.ertrag; const g = gruppe(cur);
  return `<div class="card"><h2>Ertragsermittlung</h2>
    ${cur.kultur ? `<p class="muted" style="margin-top:0">${esc(cur.kultur)} · Gruppe ${esc(g)} · Mengen in t${g === 'Grünland' || cur.kultur === 'Silomais' ? ' Frischmasse' : ''}</p>` : '<div class="note warn">Erst unter „Fall“ die Kultur wählen.</div>'}
    ${Sel('Methode', 'ertrag.methode', L.methode, { re: 1 })}
    ${e.methode === 'Probedrusch' ? `<table class="mini"><tr><th>Probe</th><th>Erntemenge kg</th><th>Probefläche m²</th><th class="num">t/ha</th></tr>
      ${e.proben.map((p, i) => `<tr><td>${i + 1}</td><td><input data-k="ertrag.proben.${i}.kg" value="${esc(p.kg || '')}" inputmode="decimal"></td><td><input data-k="ertrag.proben.${i}.m2" value="${esc(p.m2 || '')}" inputmode="decimal"></td><td class="num" data-live="probe-${i}">${num(p.kg) && num(p.m2) ? fmt(num(p.kg) / num(p.m2) * 10) : '–'}</td></tr>`).join('')}
      <tr><td colspan="3"><b>Mittelwert</b></td><td class="num"><b data-live="pmittel">${fmt(r.probenMittel)}</b></td></tr></table>` : ''}
    ${e.methode === 'Betriebsdurchschnitt' ? F('Betriebsdurchschnitt (t/ha)', 'ertrag.betrieb', { num: 1 }) : ''}
    ${e.methode === 'Schätzung am Bestand' ? F(g === 'Grünland' ? 'Restaufwuchs des Jahres (t FM/ha)' : 'Schätzung am Bestand (t/ha)', 'ertrag.schaetzung', { num: 1 }) : ''}
    ${g === 'Getreide' ? F('Strohertrag (t/ha)', 'ertrag.stroh', { num: 1 }) : ''}
  </div>
  <div class="card"><h2>Vorschau Ertragsschaden</h2><div data-live="ertragBox"></div></div>`;
}
function ertragBox(r) {
  return `<div class="sum"><span>Angesetzter Ertrag</span><span class="num">${fmt(r.ertrag)} t/ha</span></div>
  <div class="sum"><span>Preis</span><span class="num">${r.preis == null ? '–' : eur(r.preis) + '/t'}</span></div>
  <div class="sum"><span>Schadfläche wirksam</span><span class="num">${fmt(r.ha, 4)} ha</span></div>
  <div class="sum"><span>Wert Korn / Aufwuchs</span><span class="num">${eur(r.wert)}</span></div>
  ${r.gruppe === 'Getreide' ? `<div class="sum"><span>Wert Stroh</span><span class="num">${eur(r.stroh)}</span></div>` : ''}
  <div class="sum"><span>abzüglich ersparte Kosten</span><span class="num">−${eur(r.erspart)}</span></div>
  <div class="sum total"><span>Ertragsschaden netto</span><span class="num">${eur(r.ertragsschaden)}</span></div>`;
}

function vKosten() {
  const r = calc(cur); const gl = gruppe(cur) === 'Grünland';
  const gradHint = gl ? `<div class="note info">${L.grad.filter(g => r.grad[g] > 0).map(g => `<b>${g}</b>: ${fmt(r.grad[g], 1)} m² (${fmt(r.grad[g] / 10000, 4)} ha) → ${GRAD_MASSNAHME[g]}`).join('<br>') || 'Noch kein Schadgrad bei den Schadstellen gewählt.'}</div>` : '';
  return `<div class="card"><h2>Wiederherstellung</h2>${gradHint}
    <p class="muted" style="margin-top:0">Ohne Kostenquelle gilt der KTBL/MR-Satz. Die Vorlage hat 8 Zeilen.</p>
    ${cur.wieder.map((w, i) => {
      const k = 'wieder.' + i; const q = w.quelle || 'KTBL/MR-Satz'; const kt = S.ktbl.find(x => x.gang === w.gang);
      return `<div style="border-top:1px solid var(--line);padding-top:10px;margin-top:10px">
        ${Sel('Arbeitsgang', k + '.gang', S.ktbl.map(x => x.gang), { re: 1 })}
        <div class="row">${Sel('Kostenquelle', k + '.quelle', L.quelle, { re: 1 })}${F('Fläche (ha)', k + '.ha', { num: 1 })}</div>
        ${q === 'KTBL/MR-Satz' ? `<p class="muted">KTBL/MR-Satz: ${kt && kt.satz != null ? eur(kt.satz) + '/ha' : 'nicht hinterlegt'}</p>` : ''}
        ${q === 'Eigene Kosten' ? F('Eigener Satz (€/ha)', k + '.eigen', { num: 1 }) : ''}
        ${q === 'Lohnunternehmer' ? F('Lohnunternehmer laut Rechnung (€)', k + '.lu', { num: 1 }) : ''}
        <div class="sum"><span>Betrag</span><span class="num" data-live="w-${i}">${eur(r.wieder[i])}</span></div>
        <div class="btns"><button class="btn sm" data-act="haFromStellen" data-t="wieder" data-i="${i}">Fläche aus Schadstellen</button><button class="btn sm bad" data-act="delRow" data-t="wieder" data-i="${i}">Entfernen</button></div></div>`;
    }).join('')}
    <div class="btns"><button class="btn pri" data-act="addRow" data-t="wieder">＋ Arbeitsgang</button></div></div>
  <div class="card"><h2>Saatgut und Material</h2><p class="muted" style="margin-top:0">Die Vorlage hat 3 Zeilen.</p>
    ${cur.saat.map((s, i) => {
      const k = 'saat.' + i; const m = s.methode || 'kg/ha × Preis'; const g = S.saatgut.find(x => x.sorte === s.sorte);
      return `<div style="border-top:1px solid var(--line);padding-top:10px;margin-top:10px">
        ${Sel('Saatgut', k + '.sorte', S.saatgut.map(x => x.sorte), { re: 1 })}
        <div class="row">${Sel('Methode', k + '.methode', L.saatmethode, { re: 1 })}${F('Fläche (ha)', k + '.ha', { num: 1 })}</div>
        ${m === 'kg/ha × Preis' ? `<p class="muted">Tabelle: ${g && g.kgha != null ? fmt(g.kgha, 0) + ' kg/ha × ' + eur(g.preis) + '/kg' : 'nicht hinterlegt'}</p>` : ''}
        ${m === 'Pauschal €/ha' ? F('Pauschal (€/ha)', k + '.pauschal', { num: 1 }) : ''}
        ${m === 'Laut Rechnung' ? F('Rechnungsbetrag (€)', k + '.rechnung', { num: 1 }) : ''}
        <div class="sum"><span>Betrag</span><span class="num" data-live="s-${i}">${eur(r.saat[i])}</span></div>
        <div class="btns"><button class="btn sm" data-act="haFromStellen" data-t="saat" data-i="${i}">Fläche aus Schadstellen</button><button class="btn sm bad" data-act="delRow" data-t="saat" data-i="${i}">Entfernen</button></div></div>`;
    }).join('')}
    <div class="btns"><button class="btn pri" data-act="addRow" data-t="saat">＋ Saatgut</button></div></div>`;
}

function vZeiten() {
  const run = cur.zeiten.findIndex(z => z.start);
  const r = calc(cur);
  return `<div class="card"><h2>Stoppuhr</h2>
    ${run >= 0 ? `<div class="timer" data-timer="${cur.zeiten[run].start}">0:00:00</div><p class="muted">${esc(cur.zeiten[run].taetigkeit)} seit ${new Date(cur.zeiten[run].start).toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' })}</p>
      <button class="btn pri wide" data-act="stop">■ Stopp</button>` :
      `<div class="btns">${L.taetigkeit.map(t => `<button class="btn" data-act="startT" data-v="${esc(t)}">▶ ${esc(t)}</button>`).join('')}</div>
      <p class="muted">Die Zeit läuft weiter, auch wenn du die App schließt. Beim Stopp wird auf die Viertelstunde gerundet; du kannst danach korrigieren.</p>`}</div>
  ${cur.zeiten.map((z, i) => z.start ? '' : `<div class="card">
    <div class="row">${F('Datum', 'zeiten.' + i + '.datum', { type: 'date' })}${Sel('Tätigkeit', 'zeiten.' + i + '.taetigkeit', L.taetigkeit)}</div>
    ${F('Beschreibung', 'zeiten.' + i + '.beschr')}
    <div class="row">${F('Stunden', 'zeiten.' + i + '.stunden', { num: 1 })}${F('km', 'zeiten.' + i + '.km', { num: 1 })}</div>
    <button class="btn sm bad" data-act="delRow" data-t="zeiten" data-i="${i}">Entfernen</button></div>`).join('')}
  <button class="btn wide" data-act="addZeit">＋ Zeit von–bis eintragen</button>
  <div class="card" style="margin-top:12px"><h2>Vorschau Honorar</h2><div data-live="honorar">${honorarBox(r)}</div></div>`;
}
function honorarBox(r) {
  return `<div class="sum"><span>Stunden</span><span class="num">${fmt(r.std)}</span></div>
  <div class="sum"><span>Honorar ungedeckelt</span><span class="num">${eur(r.honorarRoh)}</span></div>
  <div class="sum"><span>Honorar (max. ${eur(S.stamm.tagesmax)} je Tag)</span><span class="num">${eur(r.honorar)}</span></div>
  <div class="sum"><span>Fahrtkosten ${fmt(r.km, 0)} km × ${eur(S.stamm.kmgeld)}</span><span class="num">${eur(r.fahrt)}</span></div>
  <div class="sum total"><span>Netto</span><span class="num">${eur(r.honorar + r.fahrt)}</span></div>`;
}

function vExport() {
  const r = calc(cur);
  const warn = [];
  if (!cur.kultur) warn.push('Kultur fehlt'); if (!cur.ertrag.methode) warn.push('Ertragsmethode fehlt');
  if (!cur.gesch.name) warn.push('Geschädigter fehlt'); if (cur.stellen.length > 30) warn.push('mehr als 30 Schadstellen');
  if (cur.termine.length > 5) warn.push('mehr als 5 Termine (nur 5 passen in die Vorlage)');
  if (cur.zeiten.filter(z => !z.start).length > 20) warn.push('mehr als 20 Zeiteinträge');
  if (cur.zeiten.some(z => z.start)) warn.push('Stoppuhr läuft noch – laufende Zeit wird nicht exportiert');
  const fr = fristInfo(cur);
  return `<div class="card"><h2>Vorschau Ergebnis</h2>
    <div class="sum"><span>Ertragsschaden netto</span><span class="num">${eur(r.ertragsschaden)}</span></div>
    <div class="sum"><span>Wiederherstellung</span><span class="num">${eur(r.wiederSumme)}</span></div>
    <div class="sum"><span>Saatgut und Material</span><span class="num">${eur(r.saatSumme)}</span></div>
    <div class="sum"><span>Summe Schaden netto</span><span class="num">${eur(r.summe)}</span></div>
    ${r.sb ? `<div class="sum"><span>abzüglich Selbstbehalt</span><span class="num">−${eur(r.sb)}</span></div>` : ''}
    ${r.kuerzung ? `<div class="sum"><span>Kürzung auf Höchstbetrag</span><span class="num">−${eur(r.kuerzung)}</span></div>` : ''}
    ${r.zuschlag ? `<div class="sum"><span>zzgl. Durchschnittssatz ${fmt((S.stamm.durchschnitt || 0) * 100, 1)} %</span><span class="num">${eur(r.zuschlag)}</span></div>` : ''}
    <div class="sum total"><span>Schadensersatz gesamt</span><span class="num">${eur(r.gesamt)}</span></div>
    <p class="muted">Protokoll-PDF und Rechnung rechnen mit genau diesen Werten.</p>
    ${fr[0] === 'bad' ? `<div class="note bad">${esc(fr[1])}</div>` : ''}</div>
${vPdf(r)}
  <div class="card"><h2>Excel-Datei erzeugen</h2>
    ${cur.srcName ? `<p class="muted" style="margin-top:0">Basis: deine eingelesene Datei <b>${esc(cur.srcName)}</b> – was du dort am PC eingetragen hast, bleibt erhalten.</p>` : S.vorlageName ? `<p class="muted" style="margin-top:0">Vorlage: <b>${esc(S.vorlageName)}</b> (geladen ${esc(S.vorlageDatum || '')})</p>` : `<div class="note warn">Noch keine Vorlage geladen. Unter ⚙ Einstellungen einmal deine Excel-Vorlage auswählen.</div>`}
    ${warn.length ? `<div class="note warn">Hinweis: ${warn.map(esc).join(' · ')}</div>` : ''}
    <button class="btn pri wide" data-act="export" ${S.vorlageName || cur.srcName ? '' : 'disabled'}>📄 Excel aus Vorlage erzeugen</button>
    <div id="expOut"></div></div>
  <div class="card"><h2>Fotos</h2><p class="muted" style="margin-top:0">Alle Fotos dieses Falls als ZIP. Dateinamen enthalten Foto-Nr. und Schadstelle, passend zur Spalte „Foto-Nr.“ in Excel.</p>
    <button class="btn wide" data-act="fotozip">🗂 Fotos als ZIP</button><div id="zipOut"></div></div>`;
}

/* ---------- Stufe 2: Protokoll & Rechnung ---------- */
function nextRechNr() {
  const y = new Date().getFullYear();
  const used = faelle.map(f => (f.rech && f.rech.nr) || '').filter(n => n.startsWith(y + '-')).map(n => parseInt(n.split('-')[1], 10) || 0);
  return y + '-' + String((used.length ? Math.max(...used) : 0) + 1).padStart(3, '0');
}
function rechZeiten(f) {
  return f.zeiten.filter(z => !z.start && num(z.stunden)).map(z => ({ datum: z.datum, taetigkeit: z.taetigkeit || '', beschr: z.beschr || '', std: num(z.stunden) || 0 }))
    .sort((a, b) => (a.datum || '').localeCompare(b.datum || ''));
}
function rechBetraege(f) {
  const st = S.stamm, tage = {};
  let km = 0;
  f.zeiten.filter(z => !z.start).forEach(z => { tage[z.datum || '?'] = (tage[z.datum || '?'] || 0) + (num(z.stunden) || 0); km += num(z.km) || 0; });
  const roh = Object.values(tage).reduce((a, h) => a + h * st.stundensatz, 0);
  const ged = Object.values(tage).reduce((a, h) => a + Math.min(h * st.stundensatz, st.tagesmax), 0);
  const fahrt = km * st.kmgeld, netto = ged + fahrt;
  const regel = false, ust = 0; // Ehrenamt: keine Umsatzsteuer
  return { honorarRoh: roh, kuerzung: roh - ged, km, fahrt, netto, regel, ust, brutto: netto + ust };
}
function empfaenger(f) {
  const e = f.rech.empf || 'Amt/Gemeinde';
  if (e === 'Geschädigter') return { name: f.gesch.name, anschrift: f.gesch.anschrift };
  if (e === 'Ersatzpflichtiger') return { name: f.ersatz.name, anschrift: f.ersatz.anschrift };
  if (e === 'Sonstiger') return { name: f.rech.sonstName, anschrift: f.rech.sonstAnschrift };
  return { name: f.stelle, anschrift: f.stelleAnschrift };
}
function vPdf(r) {
  const p = cur.prot, rc = cur.rech, b = rechBetraege(cur), emp = empfaenger(cur);
  const sigs = cur.sigs;
  return `<div class="card"><h2>Schätzprotokoll (PDF)</h2>
    ${S.signatur ? '' : `<div class="note warn">Deine Unterschrift fehlt noch. <button class="btn sm" data-act="sigOwn">✍ Jetzt unterschreiben</button></div>`}
    <div class="row">${F('Ort', 'prot.ort', { ph: 'z. B. Lieskau' })}${F('Datum', 'prot.datum', { type: 'date' })}</div>
    <label style="display:flex;gap:10px;align-items:center"><input type="checkbox" data-chk="prot.fotos" ${p.fotos ? 'checked' : ''} style="width:24px;min-height:24px"> <span style="margin:0;color:var(--ink)">Fotos als Anlage anhängen</span></label>
    <label style="display:flex;gap:10px;align-items:center"><input type="checkbox" data-chk="prot.karte" ${p.karte ? 'checked' : ''} style="width:24px;min-height:24px"> <span style="margin:0;color:var(--ink)">Lageplan der Schadstellen einfügen (Karte braucht Internet)</span></label>
    <h3 style="margin-top:14px">Unterschriften der Beteiligten <span class="muted" style="font-weight:400">(freiwillig)</span></h3>
    ${sigs.map((g, i) => `<div style="border-top:1px solid var(--line);padding-top:10px;margin-top:10px">
      <div class="row">${Sel('Rolle', 'sigs.' + i + '.rolle', L.sigRolle)}${F('Name', 'sigs.' + i + '.name')}</div>
      ${Sel('Erklärung', 'sigs.' + i + '.text', SIG_TEXTE)}
      ${g.png ? `<div style="background:#fff;border:1px solid var(--line);border-radius:10px;padding:6px;text-align:center"><img src="${g.png}" alt="Unterschrift" style="max-width:100%;max-height:90px"></div><p class="muted">unterschrieben ${esc(new Date(g.zeit).toLocaleString('de-DE'))}</p>` : ''}
      <div class="btns"><button class="btn ${g.png ? '' : 'pri'}" data-act="sigParty" data-i="${i}">✍ ${g.png ? 'Neu unterschreiben' : 'Unterschreiben lassen'}</button><button class="btn sm bad" data-act="delSig" data-i="${i}">Entfernen</button></div></div>`).join('')}
    <div class="btns"><button class="btn" data-act="addSig">＋ Beteiligten unterschreiben lassen</button></div>
    <div style="margin-top:14px"><button class="btn pri wide" data-act="protPdf">📄 Schätzprotokoll als PDF</button></div>
    <div id="protOut"></div></div>
  <div class="card"><h2>Schätzerrechnung (PDF)</h2>
    ${Sel('Empfänger', 'rech.empf', L.empfaenger, { re: 1 })}
    ${rc.empf === 'Sonstiger' ? F('Name', 'rech.sonstName') + F('Anschrift', 'rech.sonstAnschrift') : `<p class="muted" style="margin-top:-4px">${esc([emp.name, emp.anschrift].filter(Boolean).join(', ') || 'Name/Anschrift beim Fall eintragen')}</p>`}
    <div class="row">${F('Rechnungs-Nr.', 'rech.nr', { ph: S.naechsteNr || nextRechNr() })}${F('Rechnungsdatum', 'rech.datum', { type: 'date' })}</div>
    <p class="muted">Ehrenamtliche Tätigkeit – die Rechnung weist keine Umsatzsteuer aus.</p>
    <div class="sum"><span>Honorar ${fmt(r.std)} h${b.kuerzung > 0.004 ? ' (gedeckelt)' : ''}</span><span class="num">${eur(b.honorarRoh - b.kuerzung)}</span></div>
    <div class="sum"><span>Fahrtkosten ${fmt(b.km, 0)} km</span><span class="num">${eur(b.fahrt)}</span></div>
    ${b.regel ? `<div class="sum"><span>USt ${fmt(S.stamm.ust * 100, 0)} %</span><span class="num">${eur(b.ust)}</span></div>` : ''}
    <div class="sum total"><span>Rechnungsbetrag</span><span class="num">${eur(b.brutto)}</span></div>
    <div style="margin-top:10px"><button class="btn pri wide" data-act="rechPdf" ${rechZeiten(cur).length ? '' : 'disabled'}>🧾 Rechnung als PDF</button></div>
    ${rechZeiten(cur).length ? '' : '<p class="muted">Erst unter „Zeiten“ Stunden eintragen.</p>'}
    <div id="rechOut"></div></div>`;
}

/* Unterschriften-Feld (Finger) */
function sigPad(title, hint) {
  return new Promise(resolve => {
    openDlg(title, `<p class="muted" style="margin-top:0">${esc(hint || 'Mit dem Finger im weißen Feld unterschreiben.')}</p>
      <canvas id="sigC" style="width:100%;height:42vh;max-height:320px;background:#fff;border:2px dashed var(--line);border-radius:12px;touch-action:none;display:block"></canvas>
      <div class="btns"><button class="btn" id="sigClr">Löschen</button><button class="btn pri" id="sigOk" style="flex:1">Übernehmen</button></div>`);
    const c = $('#sigC'), dpr = window.devicePixelRatio || 1;
    const fit = () => { const r = c.getBoundingClientRect(); c.width = r.width * dpr; c.height = r.height * dpr; };
    fit();
    const g = c.getContext('2d'); let drawing = false, last = null, ink = false;
    const P = e => { const r = c.getBoundingClientRect(); return [(e.clientX - r.left) * dpr, (e.clientY - r.top) * dpr, e.pressure || 0.5]; };
    c.onpointerdown = e => { drawing = true; last = P(e); c.setPointerCapture(e.pointerId); g.beginPath(); g.arc(last[0], last[1], 1.4 * dpr, 0, 7); g.fillStyle = '#16267a'; g.fill(); ink = true; };
    c.onpointermove = e => {
      if (!drawing) return;
      const pts = e.getCoalescedEvents ? e.getCoalescedEvents() : [e];
      for (const ev of pts) {
        const p = P(ev);
        g.strokeStyle = '#16267a'; g.lineCap = 'round'; g.lineJoin = 'round'; g.lineWidth = (2.2 + p[2] * 1.6) * dpr;
        g.beginPath(); g.moveTo(last[0], last[1]); g.lineTo(p[0], p[1]); g.stroke(); last = p;
      }
    };
    c.onpointerup = c.onpointercancel = () => { drawing = false; };
    $('#sigClr').onclick = () => { g.clearRect(0, 0, c.width, c.height); ink = false; };
    let done = false;
    $('#sigOk').onclick = () => {
      if (!ink) { toast('Bitte zuerst unterschreiben'); return; }
      done = true; $('#dlg').close(); resolve(trimCanvas(c));
    };
    $('#dlg').addEventListener('close', () => { if (!done) resolve(null); }, { once: true });
  });
}
function trimCanvas(c) {
  const g = c.getContext('2d'), { width: w, height: h } = c, d = g.getImageData(0, 0, w, h).data;
  let x0 = w, y0 = h, x1 = 0, y1 = 0;
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) if (d[(y * w + x) * 4 + 3] > 10) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
  const pad = 12, sx = Math.max(0, x0 - pad), sy = Math.max(0, y0 - pad), sw = Math.min(w, x1 + pad) - sx, sh = Math.min(h, y1 + pad) - sy;
  const sc = Math.min(1, 900 / sw);
  const o = document.createElement('canvas'); o.width = Math.round(sw * sc); o.height = Math.round(sh * sc);
  o.getContext('2d').drawImage(c, sx, sy, sw, sh, 0, 0, o.width, o.height);
  return o.toDataURL('image/png');
}

/* Daten für die PDFs zusammenstellen */
async function pdfCtx(f, withFotos) {
  const r = calc(f);
  const fotosAll = (await DB.byIndex('fotos', 'fallId', f.id)).sort((a, b) => a.nr - b.nr);
  const idx = Object.fromEntries(f.stellen.map((s, i) => [s.id, i + 1]));
  const stellen = [];
  let nr = 0;
  for (const s of f.stellen) {
    nr++;
    const a = stelleFlaeche(s); if (!(a > 0)) continue;
    const v = num(s.verlust);
    stellen.push({ nr, flaeche: a, eff: a * (v == null ? 1 : v / 100), verlust: v == null ? '100 %' : fmt(v, 0) + ' %', grad: s.grad || '',
      mass: s.modus === 'gps' ? 'GPS-Umgang (' + s.poly.length + ' Pkt.)' : fmt(num(s.laenge), 1) + ' × ' + fmt(num(s.breite), 1) + ' m',
      gps: s.lat && s.lon ? num(s.lat).toFixed(5) + ' N, ' + num(s.lon).toFixed(5) + ' E' : '', fotos: await fotoNrText(s), bem: s.bem || '' });
  }
  const fotos = [];
  if (withFotos === true) for (const p of fotosAll) fotos.push({ nr: p.nr, stelle: idx[p.stelleId] || '?', bytes: new Uint8Array(await p.blob.arrayBuffer()),
    zeit: new Date(p.time).toLocaleString('de-DE', { dateStyle: 'medium', timeStyle: 'short' }), gps: p.lat ? p.lat.toFixed(5) + ' N, ' + p.lon.toFixed(5) + ' E' : '' });
  const z = rechZeiten(f);
  const ds = z.map(x => x.datum).filter(Boolean);
  let map = null;
  if (withFotos !== false && f.prot.karte && window.WSMAP) {
    try {
      const m = await WSMAP.render(mapStellen(f), 1100, 680);
      if (m) map = { bytes: dataUrlToBytes(m.canvas.toDataURL('image/jpeg', 0.85)), online: m.online, count: m.count };
    } catch (e) { console.warn('Lageplan', e); }
  }
  return {
    r, stellen, fotos, map, num, frist: fristInfo, grad: GRAD_MASSNAHME, zeiten: z, betraege: rechBetraege(f), empfaenger,
    zeitraum: ds.length ? (ds[0] === ds[ds.length - 1] ? dDE(ds[0]) : dDE(ds[0]) + ' – ' + dDE(ds[ds.length - 1])) : '',
    wSatz: w => { const q = w.quelle || 'KTBL/MR-Satz'; if (q === 'Lohnunternehmer') return 'laut Rechnung'; if (q === 'Eigene Kosten') return fmt(num(w.eigen)) + ' €/ha'; const k = S.ktbl.find(x => x.gang === w.gang); return k && k.satz != null ? fmt(k.satz) + ' €/ha' : '–'; },
    sSatz: s => { const m = s.methode || 'kg/ha × Preis'; if (m === 'Laut Rechnung') return 'laut Rechnung'; if (m === 'Pauschal €/ha') return fmt(num(s.pauschal)) + ' €/ha'; const g = S.saatgut.find(x => x.sorte === s.sorte); return g ? fmt(g.kgha, 0) + ' kg/ha × ' + fmt(g.preis) + ' €/kg' : '–'; }
  };
}
async function makePdf(kind) {
  const out = $(kind === 'prot' ? '#protOut' : '#rechOut');
  out.innerHTML = '<p class="muted">Erzeuge PDF …</p>';
  try {
    if (kind === 'prot') {
      if (!cur.prot.datum) cur.prot.datum = today();
      const ctx = await pdfCtx(cur, cur.prot.fotos ? true : 'ohneFotos');
      const bytes = await PDFX.protokoll(cur, S, ctx.r, ctx);
      lastExport = { blob: new Blob([bytes], { type: 'application/pdf' }), name: 'Schaetzprotokoll_' + fileSafe(cur.az) + '.pdf' };
      if (cur.status === 'offen') cur.status = 'geschätzt';
    } else {
      if (!cur.rech.nr) { cur.rech.nr = S.naechsteNr || nextRechNr(); S.naechsteNr = incNr(cur.rech.nr); await saveSettings(); }
      if (!cur.rech.datum) cur.rech.datum = today();
      const ctx = await pdfCtx(cur, false);
      const bytes = await PDFX.rechnung(cur, S, ctx);
      lastExport = { blob: new Blob([bytes], { type: 'application/pdf' }), name: 'Rechnung_' + fileSafe(cur.rech.nr) + '_' + fileSafe(cur.az) + '.pdf' };
    }
    save(); render();
    const o = $(kind === 'prot' ? '#protOut' : '#rechOut');
    o.innerHTML = `<div class="note ok">Fertig: <b>${esc(lastExport.name)}</b></div>
      <div class="btns">${canShare(lastExport.blob, lastExport.name) ? '<button class="btn pri" data-act="shareX">📤 Teilen</button>' : ''}<button class="btn" data-act="openX">👁 Ansehen</button><button class="btn" data-act="dlX">⬇ Speichern</button></div>`;
    o.scrollIntoView({ block: 'nearest' });
  } catch (e) { out.innerHTML = `<div class="note bad">Fehler beim PDF: ${esc(e.message)}</div>`; console.error(e); }
}
function mapStellen(f) { return f.stellen.map((s, i) => ({ nr: i + 1, lat: s.lat, lon: s.lon, poly: s.poly || [] })); }
function dataUrlToBytes(u) { const b = atob(u.split(',')[1]); const a = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) a[i] = b.charCodeAt(i); return a; }
let mapZoom = 0;
async function showMap(delta) {
  if (delta === 0) mapZoom = 0; else mapZoom += delta;
  if (!$('#dlg').open || !$('#mapBox')) openDlg('Lageplan', `<div id="mapBox" style="min-height:200px"><p class="muted">Lade Karte …</p></div>
    <div class="btns"><button class="btn" data-act="mapZ" data-v="-1">− weiter weg</button><button class="btn" data-act="mapZ" data-v="1">＋ näher ran</button></div>
    <p class="muted">Rote Nummern = Schadstellen, rote Flächen = GPS-Umgänge. Ohne Internet erscheinen nur die Punkte.</p>`);
  const w = Math.min(1000, Math.round($('#dlgB').clientWidth * (window.devicePixelRatio || 1)));
  const m = await WSMAP.render(mapStellen(cur), w, Math.round(w * 0.75), { zoomDelta: mapZoom, scale: w / Math.max(1, $('#dlgB').clientWidth - 28) });
  const box = $('#mapBox'); if (!box) return;
  if (!m) { box.innerHTML = '<div class="note warn">Noch keine Schadstelle mit GPS-Position. Bei einer Schadstelle „📍 Position übernehmen“ tippen.</div>'; return; }
  m.canvas.style.cssText = 'width:100%;border-radius:10px;display:block';
  box.innerHTML = ''; box.appendChild(m.canvas);
}
function incNr(n) { const m = /^(.*?)(\d+)$/.exec(n); return m ? m[1] + String(+m[2] + 1).padStart(m[2].length, '0') : n; }

function vSettings() {
  const st = S.stamm;
  return `<div class="card"><h2>Excel-Vorlage</h2>
    ${S.vorlageName ? `<div class="note ok">Geladen: <b>${esc(S.vorlageName)}</b><br>am ${esc(S.vorlageDatum)}</div>` : `<div class="note warn">Noch keine Vorlage geladen.</div>`}
    <p class="muted">Wähle deine Datei „…_Vorlage_unterschrieben.xlsx“. Die App speichert sie und füllt für jeden Fall eine Kopie. Preise, Sätze und Stammdaten werden dabei automatisch aus der Vorlage übernommen.</p>
    <label class="btn pri wide" style="margin:0">📂 Vorlage auswählen<input type="file" id="tplIn" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" class="hidden"></label>
    <div id="tplOut"></div></div>
  <div class="card"><h2>Meine Unterschrift</h2>
    ${S.signatur ? `<div style="background:#fff;border:1px solid var(--line);border-radius:10px;padding:8px;text-align:center"><img src="${S.signatur}" alt="Unterschrift" style="max-width:100%;max-height:110px"></div>` : '<div class="note warn">Noch keine Unterschrift hinterlegt. Sie erscheint auf jedem Schätzprotokoll.</div>'}
    <div class="btns"><button class="btn pri" data-act="sigOwn">✍ ${S.signatur ? 'Neu unterschreiben' : 'Jetzt unterschreiben'}</button>${S.signatur ? '<button class="btn bad" data-act="sigOwnDel">Löschen</button>' : ''}</div>
    <p class="muted">Tipp: Handy quer halten, dann hast du mehr Platz.</p></div>
  <div class="card"><h2>Stammdaten für Protokoll und Rechnung</h2>
    <p class="muted" style="margin-top:0">${esc(S.quelle || '')}</p>
    ${[['Name', 'name'], ['Straße, Nr.', 'strasse'], ['PLZ, Ort', 'ort'], ['Telefon', 'telefon'], ['E-Mail', 'email'], ['Bestellt durch (Landkreis/Amt)', 'bestellt'], ['Steuernummer (freiwillig)', 'steuernr'], ['IBAN', 'iban'], ['Bank', 'bank']]
      .map(([l, k]) => `<label><span>${l}</span><input data-s="${k}" value="${esc(st[k] || '')}"></label>`).join('')}
    <p class="muted">Rechnungen ohne Umsatzsteuer (Ehrenamt).</p>
    <label><span>Nächste Rechnungsnummer</span><input data-s="naechsteNr" value="${esc(S.naechsteNr || nextRechNr())}"></label>
    <div class="sum"><span>Stundensatz</span><span class="num">${eur(st.stundensatz)}/h</span></div>
    <div class="sum"><span>Höchstbetrag je Kalendertag</span><span class="num">${eur(st.tagesmax)}</span></div>
    <div class="sum"><span>Kilometergeld</span><span class="num">${eur(st.kmgeld)}/km</span></div>
    <div class="sum"><span>Durchschnittssatz</span><span class="num">${fmt((st.durchschnitt || 0) * 100, 1)} %</span></div>
    <div class="sum"><span>Anmeldefrist</span><span class="num">${st.frist} Tage</span></div></div>
  <div class="card"><h2>Preise (€/t)</h2><p class="muted" style="margin-top:0">Für die Vorschau in der App. In Excel gelten die Preise der Vorlage. Am besten in der Vorlage am PC pflegen und hier neu laden.</p>
    <div class="scrollx"><table class="mini"><tr><th>Kultur</th><th class="num">€/t</th><th class="num">Stroh</th><th class="num">erspart €/ha</th></tr>
    ${S.preise.map(p => `<tr><td>${esc(p.kultur)}</td><td class="num">${p.preis ?? '–'}</td><td class="num">${p.stroh ?? '–'}</td><td class="num">${p.erspart ?? '–'}</td></tr>`).join('')}</table></div></div>
  <div class="card"><h2>Adressbuch</h2>
    ${kontakte.length ? kontakte.slice().sort((a, b) => a.name.localeCompare(b.name)).map(c => `<div class="sum"><span><b>${esc(c.name)}</b><br><span class="muted">${esc(c.anschrift || '')}</span></span><button class="btn sm bad" data-act="delKontakt" data-id="${c.id}">✕</button></div>`).join('') : '<p class="muted">Noch leer. Im Fall bei Beteiligten „＋ Ins Adressbuch“ tippen.</p>'}</div>
  <div class="card"><h2>Datensicherung</h2>
    <p class="muted" style="margin-top:0">Eine Datei mit allen Fällen, Fotos, Adressbuch und Vorlage. Regelmäßig auf den PC kopieren. Wenn du Browserdaten löschst, sind die Fälle sonst weg.</p>
    <div class="btns"><button class="btn pri" data-act="backup">💾 Sicherung erstellen</button>
    <label class="btn" style="margin:0">↩ Sicherung einlesen<input type="file" id="restoreIn" accept=".zip,application/zip" class="hidden"></label></div><div id="bkOut"></div>
    ${S.lastBackup ? `<p class="muted">Letzte Sicherung: ${esc(S.lastBackup)}</p>` : ''}</div>
  <p class="muted" style="text-align:center">Wildschaden Feld · Version ${APP_VERSION} · Stufe 3</p>`;
}

/* ---------- Live-Werte aktualisieren, ohne neu zu zeichnen ---------- */
function live() {
  if (!cur || view.name !== 'fall') return;
  const r = calc(cur);
  const set = (k, html) => { const el = document.querySelector(`[data-live="${k}"]`); if (el) el.innerHTML = html; };
  const fr = fristInfo(cur); const fe = document.querySelector('[data-live="frist"]');
  if (fe) { fe.className = 'note ' + fr[0]; fe.textContent = fr[1]; }
  set('pauschHinweis', num(cur.pauschale) ? `<div class="note warn">Pauschale vereinbart: Prüfen, ob der Schaden damit abgegolten ist (Vertrag).</div>` : '');
  set('anz', r.anzahl); set('fl', fmt(r.flaeche, 1)); set('eff', fmt(r.eff, 1));
  cur.stellen.forEach((s, i) => set('sf-' + i, fmt(stelleFlaeche(s), 1) + ' m²'));
  cur.ertrag.proben.forEach((p, i) => set('probe-' + i, num(p.kg) && num(p.m2) ? fmt(num(p.kg) / num(p.m2) * 10) : '–'));
  set('pmittel', fmt(r.probenMittel)); set('ertragBox', ertragBox(r));
  r.wieder.forEach((v, i) => set('w-' + i, eur(v))); r.saat.forEach((v, i) => set('s-' + i, eur(v)));
  set('honorar', honorarBox(r));
  $('#title').textContent = cur.az || 'Fall';
}
setInterval(() => {
  document.querySelectorAll('[data-timer]').forEach(el => {
    const s = Math.floor((Date.now() - +el.dataset.timer) / 1000);
    el.textContent = Math.floor(s / 3600) + ':' + String(Math.floor(s / 60) % 60).padStart(2, '0') + ':' + String(s % 60).padStart(2, '0');
  });
}, 1000);

/* ---------- GPS ---------- */
function getPos() {
  return new Promise((res, rej) => {
    if (!navigator.geolocation) return rej(new Error('GPS nicht verfügbar'));
    navigator.geolocation.getCurrentPosition(p => res(p.coords), e => rej(new Error(e.code === 1 ? 'Standortzugriff nicht erlaubt' : e.code === 3 ? 'Kein GPS-Signal (Zeitüberschreitung)' : 'Standort nicht ermittelbar')),
      { enableHighAccuracy: true, timeout: 25000, maximumAge: 0 });
  });
}
const r6 = x => Math.round(x * 1e6) / 1e6;

/* ---------- Fotos ---------- */
async function shrink(file) {
  const bmp = await createImageBitmap(file);
  const max = 1600, sc = Math.min(1, max / Math.max(bmp.width, bmp.height));
  const c = document.createElement('canvas'); c.width = Math.round(bmp.width * sc); c.height = Math.round(bmp.height * sc);
  c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
  return new Promise(res => c.toBlob(b => res(b), 'image/jpeg', 0.82));
}
async function addPhotos(i, files) {
  const s = cur.stellen[i];
  let pos = null; try { pos = await Promise.race([getPos(), new Promise((_, j) => setTimeout(() => j(), 8000))]); } catch (e) { /* ohne GPS weiter */ }
  for (const file of files) {
    const blob = await shrink(file);
    cur.fotoSeq = (cur.fotoSeq || 0) + 1;
    const ph = { id: uid(), fallId: cur.id, stelleId: s.id, nr: cur.fotoSeq, blob, time: Date.now(),
      lat: pos ? r6(pos.latitude) : (num(s.lat) || null), lon: pos ? r6(pos.longitude) : (num(s.lon) || null) };
    await DB.put('fotos', ph);
    s.fotos.push(ph.id);
  }
  if (pos && !s.lat) { s.lat = r6(pos.latitude); s.lon = r6(pos.longitude); s.acc = pos.accuracy; }
  save(); render(); toast(files.length + ' Foto(s) gespeichert');
}
const thumbURLs = [];
async function loadThumbs() {
  thumbURLs.splice(0).forEach(u => URL.revokeObjectURL(u));
  const fotos = await DB.byIndex('fotos', 'fallId', cur.id);
  for (const s of cur.stellen) {
    const el = document.querySelector(`[data-thumbs="${s.id}"]`); if (!el) continue;
    el.innerHTML = fotos.filter(f => f.stelleId === s.id).sort((a, b) => a.nr - b.nr).map(f => {
      const u = URL.createObjectURL(f.blob); thumbURLs.push(u);
      return `<figure><img src="${u}" alt="Foto ${f.nr}" data-act="showFoto" data-id="${f.id}"><figcaption>Nr. ${f.nr}</figcaption><button class="x" data-act="delFoto" data-id="${f.id}" aria-label="Foto löschen">✕</button></figure>`;
    }).join('');
  }
}
async function fotoNrText(s) {
  const fotos = (await DB.byIndex('fotos', 'fallId', cur.id)).filter(f => f.stelleId === s.id).map(f => f.nr).sort((a, b) => a - b);
  const parts = []; let a = null, b = null;
  for (const n of fotos) { if (a === null) { a = b = n; } else if (n === b + 1) b = n; else { parts.push(a === b ? '' + a : a + '-' + b); a = b = n; } }
  if (a !== null) parts.push(a === b ? '' + a : a + '-' + b);
  return parts.join(', ');
}

/* ---------- Adressbuch ---------- */
const PREFIX = { gesch: ['gesch.name', 'gesch.anschrift'], ersatz: ['ersatz.name', 'ersatz.anschrift'], stelle: ['stelle', 'stelleAnschrift'] };
function openDlg(title, html) { $('#dlgT').textContent = title; $('#dlgB').innerHTML = html; $('#dlg').showModal(); }
function pickContact(p) {
  const render = q => kontakte.filter(c => !q || (c.name + ' ' + c.anschrift).toLowerCase().includes(q.toLowerCase()))
    .sort((a, b) => a.name.localeCompare(b.name))
    .map(c => `<button class="item" data-act="usec" data-p="${p}" data-id="${c.id}"><div class="t">${esc(c.name)}</div><div class="s">${esc(c.anschrift || '')}</div></button>`).join('') || '<p class="muted">Keine Einträge.</p>';
  openDlg('Adressbuch', `<label><input type="search" id="cq" placeholder="Suchen …"></label><div class="list" id="cl">${render('')}</div>`);
  $('#cq').oninput = e => { $('#cl').innerHTML = render(e.target.value); };
}
async function phoneContact(p) {
  try {
    const [c] = await navigator.contacts.select(['name', 'address', 'tel'], { multiple: false });
    if (!c) return;
    const [kn, ka] = PREFIX[p];
    setPath(cur, kn, (c.name && c.name[0]) || '');
    const a = c.address && c.address[0];
    if (a) setPath(cur, ka, [(a.addressLine || []).join(' '), [a.postalCode, a.city].filter(Boolean).join(' ')].filter(Boolean).join(', '));
    save(); render();
  } catch (e) { toast('Kontakt konnte nicht gelesen werden'); }
}

/* ---------- Export ---------- */
async function stellenRows() {
  const gl = gruppe(cur) === 'Grünland'; const rows = [];
  for (const s of cur.stellen) {
    const a = stelleFlaeche(s); if (!(a > 0)) continue;
    const v = num(s.verlust);
    let l = num(s.laenge), b = num(s.breite), bem = s.bem || '';
    if (s.modus === 'gps') { l = Math.round(a * 10) / 10; b = 1; bem = ('GPS-Umgang ' + s.poly.length + ' Punkte, ' + fmt(a, 1) + ' m²' + (bem ? '; ' + bem : '')); }
    rows.push({ l, b, v: v == null ? '' : v / 100, g: gl ? (s.grad || '') : '', lat: num(s.lat) ?? '', lon: num(s.lon) ?? '',
      foto: await fotoNrText(s), termin: s.termin || '', bem });
  }
  return rows;
}
async function doExport() {
  const out = $('#expOut'); out.innerHTML = '<p class="muted">Erzeuge Datei …</p>';
  try {
    const tpl = (cur.srcName && await DB.get('kv', 'src-' + cur.id)) || await DB.get('kv', 'vorlage');
    if (!tpl) throw new Error('Vorlage nicht gefunden – bitte neu laden.');
    const { blob, rep } = await XL.exportFall(await tpl.arrayBuffer(), cur, { stellenRows: await stellenRows(), stamm: S.stamm });
    const name = 'WS_' + [fileSafe(cur.az), fileSafe(cur.gesch.name), fileSafe(cur.flurstueck && 'Fl' + cur.flurstueck)].filter(Boolean).join('_') + '.xlsx';
    lastExport = { blob, name };
    out.innerHTML = `<div class="note ok">Fertig: <b>${esc(name)}</b> · ${rep.written} Werte übertragen.</div>
      ${rep.warn.map(w => `<div class="note warn">${esc(w)}</div>`).join('')}
      ${rep.missing.length ? `<div class="note warn">In der Vorlage nicht gefunden: ${rep.missing.map(esc).join(', ')}</div>` : ''}
      ${rep.formula.length ? `<div class="note warn">Formelzellen nicht überschrieben: ${rep.formula.map(esc).join(', ')}</div>` : ''}
      <div class="btns">${canShare(blob, name) ? '<button class="btn pri" data-act="shareX">📤 Teilen</button>' : ''}<button class="btn" data-act="dlX">⬇ Speichern</button></div>`;
  } catch (e) { out.innerHTML = `<div class="note bad">Fehler: ${esc(e.message)}</div>`; }
}
function canShare(blob, name) { try { return !!(navigator.canShare && navigator.canShare({ files: [new File([blob], name, { type: blob.type })] })); } catch (e) { return false; } }
async function shareFile(blob, name) {
  try { await navigator.share({ files: [new File([blob], name, { type: blob.type })], title: name }); }
  catch (e) { if (e.name !== 'AbortError') toast('Teilen nicht möglich: ' + e.message); }
}
function download(blob, name) {
  const u = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = u; a.download = name;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(u), 4000);
  toast('Gespeichert unter „Downloads“: ' + name);
}
async function fotoZip() {
  const out = $('#zipOut');
  const fotos = (await DB.byIndex('fotos', 'fallId', cur.id)).sort((a, b) => a.nr - b.nr);
  if (!fotos.length) { out.innerHTML = '<p class="muted">Keine Fotos in diesem Fall.</p>'; return; }
  const z = new JSZip(); const idx = Object.fromEntries(cur.stellen.map((s, i) => [s.id, i + 1]));
  let liste = 'Foto-Nr;Schadstelle;Datum/Uhrzeit;GPS N;GPS E\r\n';
  for (const f of fotos) {
    const st = idx[f.stelleId] || 0;
    z.file('Foto_' + String(f.nr).padStart(2, '0') + '_Stelle' + st + '.jpg', f.blob);
    liste += [f.nr, st, new Date(f.time).toLocaleString('de-DE'), f.lat ?? '', f.lon ?? ''].join(';') + '\r\n';
  }
  z.file('Fotoliste.csv', '﻿' + liste);
  const blob = await z.generateAsync({ type: 'blob', mimeType: 'application/zip' });
  const name = 'Fotos_' + fileSafe(cur.az) + '.zip';
  lastExport = { blob, name };
  out.innerHTML = `<div class="note ok">${fotos.length} Fotos gepackt.</div><div class="btns">${canShare(blob, name) ? '<button class="btn pri" data-act="shareX">📤 Teilen</button>' : ''}<button class="btn" data-act="dlX">⬇ Speichern</button></div>`;
}

/* ---------- Excel-Datei einlesen (am PC begonnen, am Feld weitermachen) ---------- */
async function importExcel(file) {
  try {
    const buf = await file.arrayBuffer();
    const { f: x, info } = await XL.importFall(buf);
    if (x.az && faelle.some(f => f.az === x.az) && !confirm('Es gibt schon einen Fall „' + x.az + '“. Trotzdem als neuen Fall anlegen?')) return;
    const f = newFall();
    for (const k of ['az', 'verfahren', 'stelle', 'stelleAnschrift', 'jagdbezirk', 'zeuge1', 'zeuge2', 'vertreter', 'gemarkung', 'flur', 'flurstueck', 'schlag', 'schlagHa',
      'kultur', 'wildart', 'kenntnis', 'anmeldung', 'sbProz', 'sbEuro', 'deckel', 'pauschale', 'ustGesch', 'bemerkung']) if (x[k]) f[k] = x[k];
    Object.assign(f.gesch, x.gesch || {}); Object.assign(f.ersatz, x.ersatz || {});
    f.termine = x.termine.map(t => ({ id: uid(), ...t }));
    if (x.stellen.length) f.stellen = x.stellen.map(s => Object.assign(newStelle(), { laenge: s.l, breite: s.b, verlust: s.v, grad: s.g,
      lat: s.lat === '' ? '' : s.lat, lon: s.lon === '' ? '' : s.lon, termin: s.termin, bem: [s.bem, s.foto && 'Fotos laut Excel: ' + s.foto].filter(Boolean).join(' · ') }));
    f.ertrag = x.ertrag; f.wieder = x.wieder; f.saat = x.saat;
    f.zeiten = x.zeiten.map(z => ({ id: uid(), ...z }));
    Object.assign(f.rech, x.rech);
    f.srcName = file.name;
    await DB.put('kv', new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'src-' + f.id);
    faelle.push(f); await DB.put('faelle', f);
    cur = f; view = { name: 'fall', tab: 'fall' }; render(); window.scrollTo(0, 0);
    toast('Eingelesen: ' + info.felder + ' Felder, ' + info.stellen + ' Schadstellen, ' + info.termine + ' Termine, ' + info.zeiten + ' Zeiten', 4500);
  } catch (e) { toast('Datei konnte nicht gelesen werden: ' + e.message, 5000); }
}

/* ---------- Vorlage & Sicherung ---------- */
async function loadTemplate(file) {
  const out = $('#tplOut'); out.innerHTML = '<p class="muted">Prüfe Vorlage …</p>';
  try {
    const buf = await file.arrayBuffer();
    const wb = await XL.Workbook.open(buf);
    const need = ['Fall', 'Schadstellen', 'Berechnung', 'Zeiterfassung'].filter(n => !wb.has(n));
    if (need.length) throw new Error('Blätter fehlen: ' + need.join(', ') + '. Ist das die richtige Datei?');
    const st = await XL.readStammdaten(wb);
    // Probelauf: leeren Fall exportieren, um alle Felder zu finden
    const { rep } = await XL.exportFall(buf, newFall(), { stellenRows: [] });
    await DB.put('kv', new Blob([buf], { type: file.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'vorlage');
    if (st.preise && st.preise.length) S.preise = st.preise;
    if (st.ktbl && st.ktbl.length) S.ktbl = st.ktbl;
    if (st.saatgut && st.saatgut.length) S.saatgut = st.saatgut;
    Object.assign(S.stamm, st.stamm || {});
    S.vorlageName = file.name; S.vorlageDatum = new Date().toLocaleString('de-DE');
    S.quelle = 'Übernommen aus „' + file.name + '“ am ' + S.vorlageDatum;
    await saveSettings();
    render();
    $('#tplOut').innerHTML = `<div class="note ok">Vorlage gespeichert. Übernommen: ${st.found.map(esc).join(', ')}.</div>
      ${rep.missing.length || st.missing.length ? `<div class="note warn">Nicht gefunden: ${[...st.missing, ...rep.missing].map(esc).join(', ')}. Diese Felder werden beim Export übersprungen.</div>` : '<div class="note ok">Alle Eingabefelder der Vorlage gefunden.</div>'}
      ${rep.formula.length ? `<div class="note warn">Diese Felder enthalten in der Vorlage Formeln und werden nicht beschrieben: ${rep.formula.map(esc).join(', ')}</div>` : ''}`;
  } catch (e) { out.innerHTML = `<div class="note bad">${esc(e.message)}</div>`; }
}
async function backup() {
  const z = new JSZip();
  const fotos = await DB.all('fotos');
  z.file('daten.json', JSON.stringify({ version: APP_VERSION, erstellt: new Date().toISOString(), faelle, kontakte, settings: S,
    fotos: fotos.map(({ blob, ...m }) => m) }));
  fotos.forEach(f => z.file('fotos/' + f.id + '.jpg', f.blob));
  for (const f of faelle) if (f.srcName) { const b = await DB.get('kv', 'src-' + f.id); if (b) z.file('quellen/' + f.id + '.xlsx', b); }
  const tpl = await DB.get('kv', 'vorlage'); if (tpl) z.file('vorlage.xlsx', tpl);
  const blob = await z.generateAsync({ type: 'blob', mimeType: 'application/zip' });
  const name = 'Wildschaden_Sicherung_' + today() + '.zip';
  lastExport = { blob, name };
  S.lastBackup = new Date().toLocaleString('de-DE'); S.lastBackupTs = Date.now(); await saveSettings();
  $('#bkOut').innerHTML = `<div class="note ok">${faelle.length} Fälle, ${fotos.length} Fotos gesichert.</div><div class="btns">${canShare(blob, name) ? '<button class="btn pri" data-act="shareX">📤 Teilen</button>' : ''}<button class="btn" data-act="dlX">⬇ Speichern</button></div>`;
}
async function restore(file) {
  try {
    const z = await JSZip.loadAsync(file);
    const d = JSON.parse(await z.file('daten.json').async('string'));
    if (!confirm('Sicherung vom ' + new Date(d.erstellt).toLocaleString('de-DE') + ' mit ' + d.faelle.length + ' Fällen einlesen? Fälle mit gleicher Kennung werden überschrieben, andere bleiben erhalten.')) return;
    for (const f of d.faelle) await DB.put('faelle', f);
    for (const c of d.kontakte || []) await DB.put('kontakte', c);
    for (const zf of z.file(/^quellen\/.+\.xlsx$/)) await DB.put('kv', new Blob([await zf.async('arraybuffer')]), 'src-' + zf.name.slice(8, -5));
    for (const m of d.fotos || []) { const zf = z.file('fotos/' + m.id + '.jpg'); if (zf) await DB.put('fotos', { ...m, blob: new Blob([await zf.async('arraybuffer')], { type: 'image/jpeg' }) }); }
    const t = z.file('vorlage.xlsx');
    if (t && d.settings) { await DB.put('kv', new Blob([await t.async('arraybuffer')]), 'vorlage'); S = Object.assign({}, DEFAULT_SETTINGS, d.settings); await saveSettings(); }
    faelle = (await DB.all('faelle')).map(normFall); kontakte = await DB.all('kontakte');
    render(); toast('Sicherung eingelesen');
  } catch (e) { toast('Sicherung konnte nicht gelesen werden: ' + e.message, 4000); }
}

/* ---------- Ereignisse ---------- */
document.addEventListener('input', e => {
  const t = e.target;
  if (t.id === 'search') { view.q = t.value; const pos = t.selectionStart; render(); const s = $('#search'); s.focus(); s.setSelectionRange(pos, pos); return; }
  if (t.dataset.k && cur) { setPath(cur, t.dataset.k, t.value); save(); if (t.tagName === 'SELECT' && t.dataset.re) render(); else live(); }
  if (t.dataset.s) { if (t.dataset.s === 'naechsteNr') S.naechsteNr = t.value; else S.stamm[t.dataset.s] = t.value; clearTimeout(t._st); t._st = setTimeout(saveSettings, 300); }
});
document.addEventListener('change', async e => {
  const t = e.target;
  if (t.dataset.photo !== undefined && t.files.length) { await addPhotos(+t.dataset.photo, [...t.files]); t.value = ''; }
  if (t.id === 'tplIn' && t.files[0]) loadTemplate(t.files[0]);
  if (t.id === 'restoreIn' && t.files[0]) restore(t.files[0]);
  if (t.id === 'impIn' && t.files[0]) { importExcel(t.files[0]); t.value = ''; }
  if (t.dataset.chk && cur) { setPath(cur, t.dataset.chk, t.checked); save(); }
});
document.addEventListener('click', async e => {
  const tab = e.target.closest('[data-tab]'); if (tab) { view.tab = tab.dataset.tab; render(); window.scrollTo(0, 0); return; }
  const b = e.target.closest('[data-act]'); if (!b) return;
  const a = b.dataset.act, i = b.dataset.i != null ? +b.dataset.i : null;
  switch (a) {
    case 'new': cur = newFall(); faelle.push(cur); await DB.put('faelle', cur); view = { name: 'fall', tab: 'fall' }; render(); break;
    case 'open': cur = normFall(faelle.find(f => f.id === b.dataset.id)); view = { name: 'fall', tab: 'fall', q: view.q, st: view.st }; render(); window.scrollTo(0, 0); break;
    case 'filter': view.st = b.dataset.v; render(); break;
    case 'status': cur.status = b.dataset.v; save(); render(); break;
    case 'delFall':
      if (!confirm('Fall „' + (cur.az || '') + '“ mit allen Fotos endgültig löschen?')) break;
      for (const f of await DB.byIndex('fotos', 'fallId', cur.id)) await DB.del('fotos', f.id);
      await DB.del('faelle', cur.id); DB.del('kv', 'src-' + cur.id).catch(() => {}); faelle = faelle.filter(f => f.id !== cur.id); cur = null; view = { name: 'list' }; render(); break;
    case 'addTermin': cur.termine.push({ id: uid(), datum: today(), art: '', anwesende: '', feststellungen: '' }); save(); render(); break;
    case 'delTermin': if (confirm('Termin entfernen?')) { cur.termine.splice(i, 1); save(); render(); } break;
    case 'anwesend': {
      const n = [cur.gesch.name && cur.gesch.name + ' (Geschädigter)', cur.ersatz.name && cur.ersatz.name + ' (Ersatzpflichtiger)', cur.zeuge1, cur.zeuge2, cur.vertreter, S.stamm.name && S.stamm.name + ' (Schätzer)'].filter(Boolean).join(', ');
      cur.termine[i].anwesende = n; save(); render(); break;
    }
    case 'addStelle': cur.stellen.push(newStelle()); save(); render(); setTimeout(() => document.getElementById('st-' + cur.stellen[cur.stellen.length - 1].id).scrollIntoView({ behavior: 'smooth' }), 50); break;
    case 'delStelle':
      if (!confirm('Schadstelle ' + (i + 1) + ' samt Fotos entfernen?')) break;
      for (const id of cur.stellen[i].fotos) await DB.del('fotos', id);
      cur.stellen.splice(i, 1); save(); render(); break;
    case 'modus': cur.stellen[i].modus = b.dataset.v; save(); render(); break;
    case 'pos': case 'polyAdd': {
      b.disabled = true; const old = b.textContent; b.textContent = '⏳ Suche GPS …';
      try {
        const p = await getPos(); const s = cur.stellen[i];
        if (a === 'pos') { s.lat = r6(p.latitude); s.lon = r6(p.longitude); s.acc = p.accuracy; }
        else { s.poly.push([r6(p.latitude), r6(p.longitude), p.accuracy]); if (!s.lat) { s.lat = r6(p.latitude); s.lon = r6(p.longitude); s.acc = p.accuracy; } }
        save(); render(); toast('Position ±' + Math.round(p.accuracy) + ' m');
      } catch (err) { toast(err.message, 4000); b.disabled = false; b.textContent = old; }
      break;
    }
    case 'polyUndo': cur.stellen[i].poly.pop(); save(); render(); break;
    case 'delFoto': {
      if (!confirm('Foto löschen?')) break;
      await DB.del('fotos', b.dataset.id); cur.stellen.forEach(s => { s.fotos = s.fotos.filter(x => x !== b.dataset.id); }); save(); loadThumbs(); break;
    }
    case 'showFoto': { const f = await DB.get('fotos', b.dataset.id); const u = URL.createObjectURL(f.blob);
      openDlg('Foto Nr. ' + f.nr, `<img src="${u}" style="width:100%;border-radius:10px"><p class="muted">${new Date(f.time).toLocaleString('de-DE')}${f.lat ? ' · ' + f.lat + ', ' + f.lon : ''}</p>`); break; }
    case 'addRow': cur[b.dataset.t].push(b.dataset.t === 'wieder' ? { gang: '', quelle: '', ha: '', eigen: '', lu: '' } : { sorte: '', methode: '', ha: '', pauschal: '', rechnung: '' }); save(); render(); break;
    case 'delRow': cur[b.dataset.t].splice(i, 1); save(); render(); break;
    case 'haFromStellen': {
      const r = calc(cur); const row = cur[b.dataset.t][i];
      const ha = r.flaeche / 10000;
      row.ha = String(Math.round(ha * 10000) / 10000).replace('.', ','); save(); render(); toast('Fläche: ' + row.ha + ' ha'); break;
    }
    case 'startT': cur.zeiten.push({ id: uid(), datum: today(), taetigkeit: b.dataset.v, beschr: '', stunden: '', km: '', start: Date.now() }); save(); render(); break;
    case 'stop': {
      const z = cur.zeiten.find(x => x.start);
      const h = Math.max(0.25, Math.round((Date.now() - z.start) / 3600000 * 4) / 4);
      z.stunden = String(h).replace('.', ','); delete z.start; save(); render(); toast(z.stunden + ' h erfasst'); break;
    }
    case 'addZeit': cur.zeiten.push({ id: uid(), datum: today(), taetigkeit: '', beschr: '', stunden: '', km: '' }); save(); render(); break;
    case 'pick': pickContact(b.dataset.p); break;
    case 'phone': phoneContact(b.dataset.p); break;
    case 'usec': { const c = kontakte.find(x => x.id === b.dataset.id); const [kn, ka] = PREFIX[b.dataset.p]; setPath(cur, kn, c.name); setPath(cur, ka, c.anschrift || ''); save(); $('#dlg').close(); render(); break; }
    case 'savec': {
      const [kn, ka] = PREFIX[b.dataset.p]; const name = getPath(cur, kn);
      if (!name) { toast('Erst einen Namen eintragen'); break; }
      let c = kontakte.find(x => x.name === name);
      if (c) c.anschrift = getPath(cur, ka); else { c = { id: uid(), name, anschrift: getPath(cur, ka) }; kontakte.push(c); }
      await DB.put('kontakte', c); toast('Im Adressbuch gespeichert'); break;
    }
    case 'delKontakt': if (confirm('Kontakt löschen?')) { await DB.del('kontakte', b.dataset.id); kontakte = kontakte.filter(c => c.id !== b.dataset.id); render(); } break;
    case 'export': doExport(); break;
    case 'fotozip': fotoZip(); break;
    case 'shareX': shareFile(lastExport.blob, lastExport.name); break;
    case 'dlX': download(lastExport.blob, lastExport.name); break;
    case 'backup': backup(); break;
    case 'sigOwn': { const png = await sigPad('Meine Unterschrift', 'So unterschreibst du später jedes Schätzprotokoll.'); if (png) { S.signatur = png; await saveSettings(); render(); toast('Unterschrift gespeichert'); } break; }
    case 'sigOwnDel': if (confirm('Deine Unterschrift löschen?')) { delete S.signatur; await saveSettings(); render(); } break;
    case 'addSig': cur.sigs.push({ id: uid(), rolle: cur.sigs.some(x => x.rolle === 'Geschädigter') ? 'Ersatzpflichtiger' : 'Geschädigter', name: '', text: SIG_TEXTE[0], png: '', zeit: 0 });
      { const g = cur.sigs[cur.sigs.length - 1]; g.name = g.rolle === 'Geschädigter' ? cur.gesch.name : cur.ersatz.name; } save(); render(); break;
    case 'delSig': if (confirm('Unterschrift entfernen?')) { cur.sigs.splice(i, 1); save(); render(); } break;
    case 'sigParty': {
      const g = cur.sigs[i];
      const png = await sigPad('Unterschrift ' + g.rolle, (g.name ? g.name + ': ' : '') + (g.text || '') + ' Mit dem Finger unterschreiben.');
      if (png) { g.png = png; g.zeit = Date.now(); save(); render(); } break;
    }
    case 'protPdf': makePdf('prot'); break;
    case 'map': showMap(0); break;
    case 'goBackup': view = { name: 'settings', prevTab: view.tab, q: view.q, st: view.st }; render(); setTimeout(() => { const b = document.querySelector('[data-act="backup"]'); if (b) b.scrollIntoView({ block: 'center' }); }, 50); break;
    case 'mapZ': showMap(+b.dataset.v); break;
    case 'rechPdf': makePdf('rech'); break;
    case 'openX': { const u = URL.createObjectURL(lastExport.blob); window.open(u, '_blank'); setTimeout(() => URL.revokeObjectURL(u), 60000); break; }
  }
});
$('#back').onclick = () => {
  if (view.name === 'settings' && cur) { view = { name: 'fall', tab: view.prevTab || 'fall' }; }
  else { view = { name: 'list', q: view.q, st: view.st }; cur = null; faelle = faelle.slice(); }
  render(); window.scrollTo(0, 0);
};
$('#settingsBtn').onclick = () => { view = { name: 'settings', prevTab: view.tab, q: view.q, st: view.st }; render(); window.scrollTo(0, 0); };
$('#dlgX').onclick = () => $('#dlg').close();

/* ---------- Start ---------- */
(async function start() {
  try {
    await DB.open();
    S = Object.assign({}, DEFAULT_SETTINGS, (await DB.get('kv', 'settings')) || {});
    faelle = (await DB.all('faelle')).map(normFall); kontakte = await DB.all('kontakte');
    if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
  } catch (e) {
    $('#main').innerHTML = `<div class="note bad">Speicher nicht verfügbar: ${esc(e.message)}. Bitte die App nicht im Inkognito-Modus öffnen.</div>`; return;
  }
  render();
  if ('serviceWorker' in navigator && window.isSecureContext) {
    try { await navigator.serviceWorker.register('sw.js'); } catch (e) { /* Vorschau ohne Offline-Modus */ }
  }
})();
