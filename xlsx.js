/* Befüllt die Excel-Vorlage direkt im Datei-Inneren (XML), damit Formeln,
   Formate, Auswahllisten und die eingefügte Unterschrift unverändert bleiben.
   Grundregel: Eine Zelle mit Formel wird NIE überschrieben. */
(function (global) {
  'use strict';
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  const norm = s => String(s == null ? '' : s).replace(/\s+/g, ' ').trim().toLowerCase();
  function colToNum(c) { let n = 0; for (const ch of c) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
  function numToCol(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  function parseRef(ref) { const m = /^([A-Z]+)(\d+)$/.exec(ref); return { col: colToNum(m[1]), row: +m[2] }; }
  function dateSerial(iso) {
    if (!iso) return null;
    const [y, m, d] = iso.split('-').map(Number);
    return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
  }

  class Sheet {
    constructor(wb, name, path, doc) {
      this.wb = wb; this.name = name; this.path = path; this.doc = doc; this.dirty = false;
      this.sheetData = doc.getElementsByTagNameNS(NS, 'sheetData')[0];
      this.merges = [...doc.getElementsByTagNameNS(NS, 'mergeCell')].map(m => {
        const [a, b] = m.getAttribute('ref').split(':'); const A = parseRef(a), B = parseRef(b || a);
        return { r1: A.row, c1: A.col, r2: B.row, c2: B.col };
      });
      this.index();
    }
    index() {
      this.map = new Map();
      for (const c of this.sheetData.getElementsByTagNameNS(NS, 'c')) this.map.set(c.getAttribute('r'), c);
    }
    cell(row, col) { return this.map.get(numToCol(col) + row) || null; }
    text(row, col) { return this.wb.cellText(this.cell(row, col)); }
    value(row, col) {
      const c = this.cell(row, col); if (!c) return null;
      const t = c.getAttribute('t'); const txt = this.wb.cellText(c);
      if (txt === '') return null;
      if (!t || t === 'n') { const n = Number(txt); return isNaN(n) ? txt : n; }
      return txt;
    }
    hasFormula(row, col) { const c = this.cell(row, col); return !!(c && c.getElementsByTagNameNS(NS, 'f').length); }
    /* alle Zellen in Zeilen-/Spaltenreihenfolge */
    *iter() {
      for (const r of this.sheetData.getElementsByTagNameNS(NS, 'row')) {
        for (const c of r.getElementsByTagNameNS(NS, 'c')) {
          const p = parseRef(c.getAttribute('r')); yield { row: p.row, col: p.col, el: c };
        }
      }
    }
    find(label, opts = {}) {
      const L = norm(label);
      for (const { row, col, el } of this.iter()) {
        if (el.getElementsByTagNameNS(NS, 'f').length) continue;
        const t = norm(this.wb.cellText(el)); if (!t) continue;
        if (opts.prefix ? t.startsWith(L) : t === L) return { row, col };
      }
      return null;
    }
    colInRow(row, label) {
      const L = norm(label);
      for (const { row: r, col, el } of this.iter()) if (r === row && norm(this.wb.cellText(el)) === L) return col;
      return null;
    }
    mergeAt(row, col) { return this.merges.find(m => row >= m.r1 && row <= m.r2 && col >= m.c1 && col <= m.c2) || null; }
    /* erste Wertzelle rechts neben einer Beschriftung (verbundene Zellen beachten) */
    rightOf(pos) {
      const m = this.mergeAt(pos.row, pos.col); let col = (m ? m.c2 : pos.col) + 1;
      return { row: pos.row, col };
    }
    ensureCell(row, col) {
      const ref = numToCol(col) + row; let c = this.map.get(ref); if (c) return c;
      const rows = [...this.sheetData.getElementsByTagNameNS(NS, 'row')];
      let rEl = rows.find(r => +r.getAttribute('r') === row);
      if (!rEl) {
        rEl = this.doc.createElementNS(NS, 'row'); rEl.setAttribute('r', String(row));
        const after = rows.find(r => +r.getAttribute('r') > row);
        this.sheetData.insertBefore(rEl, after || null);
      }
      c = this.doc.createElementNS(NS, 'c'); c.setAttribute('r', ref);
      // Zeilenformat übernehmen, damit die Zelle nicht „nackt“ aussieht
      const cells = [...rEl.getElementsByTagNameNS(NS, 'c')];
      const after = cells.find(x => parseRef(x.getAttribute('r')).col > col);
      rEl.insertBefore(c, after || null);
      if (rEl.getAttribute('spans')) rEl.removeAttribute('spans');
      this.map.set(ref, c); return c;
    }
    /* Wert schreiben. Rückgabe: true = geschrieben, false = Formelzelle (übersprungen) */
    set(row, col, value) {
      if (this.hasFormula(row, col)) return false;
      if ((value === null || value === undefined || value === '') && !this.cell(row, col)) return true;
      const c = this.ensureCell(row, col);
      while (c.firstChild) c.removeChild(c.firstChild);
      c.removeAttribute('t');
      if (value === null || value === undefined || value === '') { this.dirty = true; return true; }
      if (typeof value === 'number' && isFinite(value)) {
        const v = this.doc.createElementNS(NS, 'v'); v.textContent = String(value); c.appendChild(v);
      } else {
        c.setAttribute('t', 'inlineStr');
        const is = this.doc.createElementNS(NS, 'is'); const t = this.doc.createElementNS(NS, 't');
        t.setAttribute('xml:space', 'preserve'); t.textContent = String(value); is.appendChild(t); c.appendChild(is);
      }
      this.dirty = true; return true;
    }
  }

  class Workbook {
    static async open(buf) {
      const wb = new Workbook(); wb.zip = await JSZip.loadAsync(buf);
      const P = new DOMParser();
      wb.parse = s => P.parseFromString(s, 'application/xml');
      const wbDoc = wb.parse(await wb.zip.file('xl/workbook.xml').async('string'));
      const rels = wb.parse(await wb.zip.file('xl/_rels/workbook.xml.rels').async('string'));
      const relMap = {};
      for (const r of rels.getElementsByTagName('Relationship')) relMap[r.getAttribute('Id')] = r.getAttribute('Target');
      wb.wbDoc = wbDoc; wb.sheetPaths = {};
      for (const s of wbDoc.getElementsByTagNameNS(NS, 'sheet')) {
        let t = relMap[s.getAttributeNS(RNS, 'id') || s.getAttribute('r:id')];
        t = t.startsWith('/') ? t.slice(1) : 'xl/' + t.replace(/^\.\//, '');
        wb.sheetPaths[s.getAttribute('name')] = t;
      }
      wb.shared = [];
      const ss = wb.zip.file('xl/sharedStrings.xml');
      if (ss) {
        const d = wb.parse(await ss.async('string'));
        for (const si of d.getElementsByTagNameNS(NS, 'si')) wb.shared.push([...si.getElementsByTagNameNS(NS, 't')].filter(t => !t.parentNode || t.parentNode.localName !== 'rPh').map(t => t.textContent).join(''));
      }
      wb.sheets = {};
      return wb;
    }
    cellText(c) {
      if (!c) return '';
      const t = c.getAttribute('t');
      if (t === 'inlineStr') return [...c.getElementsByTagNameNS(NS, 't')].map(x => x.textContent).join('');
      const v = c.getElementsByTagNameNS(NS, 'v')[0]; if (!v) return '';
      if (t === 's') return this.shared[+v.textContent] || '';
      return v.textContent;
    }
    has(name) { return !!this.sheetPaths[name]; }
    async sheet(name) {
      if (this.sheets[name]) return this.sheets[name];
      const path = this.sheetPaths[name]; if (!path) return null;
      const doc = this.parse(await this.zip.file(path).async('string'));
      return (this.sheets[name] = new Sheet(this, name, path, doc));
    }
    async save() {
      const S = new XMLSerializer();
      for (const sh of Object.values(this.sheets)) if (sh.dirty) this.zip.file(sh.path, S.serializeToString(sh.doc));
      // Excel soll beim Öffnen alles neu berechnen
      let calc = this.wbDoc.getElementsByTagNameNS(NS, 'calcPr')[0];
      if (!calc) { calc = this.wbDoc.createElementNS(NS, 'calcPr'); this.wbDoc.documentElement.appendChild(calc); }
      calc.setAttribute('fullCalcOnLoad', '1');
      this.zip.file('xl/workbook.xml', S.serializeToString(this.wbDoc));
      return this.zip.generateAsync({ type: 'blob', compression: 'DEFLATE',
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    }
  }

  /* ---------- Vorlage lesen: Stammdaten, Preise, Sätze ---------- */
  async function readStammdaten(wb) {
    const out = { found: [], missing: [] };
    const st = await wb.sheet('Stammdaten');
    if (!st) { out.missing.push('Blatt „Stammdaten“'); return out; }
    const scal = {
      name: ['Name'], strasse: ['Straße, Nr.'], ort: ['PLZ, Ort'], telefon: ['Telefon'], email: ['E-Mail'],
      steuernr: ['Steuernummer'], iban: ['IBAN'], bank: ['Bank'], bestellt: ['Bestellt durch', true],
      stundensatz: ['Stundensatz', true], tagesmax: ['Höchstbetrag Honorar', true], kmgeld: ['Kilometergeld', true],
      ust: ['Umsatzsteuersatz', true], durchschnitt: ['Durchschnittssatz', true], frist: ['Anmeldefrist', true]
    };
    out.stamm = {};
    for (const [k, [lab, pre]] of Object.entries(scal)) {
      const p = st.find(lab, { prefix: !!pre });
      if (p) { const r = st.rightOf(p); const v = st.value(r.row, r.col); if (v !== null) out.stamm[k] = v; }
    }
    // Anker „Preis €/t“ steht nicht in Spalte A → Tabelle über Kopfzeile finden
    const head = st.find('Preis €/t');
    if (head) {
      const cK = st.colInRow(head.row, 'Kultur'), cols = ['Gruppe', 'Einheit', 'Preis €/t', 'Strohpreis €/t', 'Ersparte Kosten €/ha'].map(c => st.colInRow(head.row, c));
      out.preise = [];
      for (let r = head.row + 1; r < head.row + 80; r++) {
        const k = st.text(r, cK); if (!k || norm(k).startsWith('maschinenkosten')) break;
        const [g, e, p, s, x] = cols.map(c => c ? st.value(r, c) : null);
        out.preise.push({ kultur: k, gruppe: g || '', einheit: e || '', preis: num(p), stroh: num(s), erspart: num(x) });
      }
      out.found.push('Preistabelle (' + out.preise.length + ' Kulturen)');
    } else out.missing.push('Preistabelle');
    const kh = st.find('Satz €/ha');
    if (kh) {
      const cA = st.colInRow(kh.row, 'Arbeitsgang');
      out.ktbl = [];
      for (let r = kh.row + 1; r < kh.row + 60; r++) {
        const a = st.text(r, cA); if (!a || norm(a).startsWith('saatgut')) break;
        out.ktbl.push({ gang: a, satz: num(st.value(r, kh.col)) });
      }
      out.found.push('Maschinenkostensätze (' + out.ktbl.length + ')');
    } else out.missing.push('Maschinenkostensätze');
    const sh = st.find('Aussaatmenge kg/ha');
    if (sh) {
      const cS = st.colInRow(sh.row, 'Saatgut'), cP = st.colInRow(sh.row, 'Preis €/kg');
      out.saatgut = [];
      for (let r = sh.row + 1; r < sh.row + 60; r++) {
        const a = st.text(r, cS); if (!a || norm(a).startsWith('grünland:')) break;
        out.saatgut.push({ sorte: a, kgha: num(st.value(r, sh.col)), preis: num(st.value(r, cP)) });
      }
      out.found.push('Saatgut (' + out.saatgut.length + ')');
    } else out.missing.push('Saatgut-Tabelle');
    if (Object.keys(out.stamm).length) out.found.push('Stammdaten (' + Object.keys(out.stamm).length + ' Werte)');
    return out;
  }
  function num(v) { if (v === null || v === undefined || v === '') return null; const n = Number(v); return isNaN(n) ? null : n; }

  /* ---------- Fall in die Vorlage schreiben ---------- */
  const FALL_LABELS = [
    ['Aktenzeichen', f => f.az], ['Verfahrensart', f => f.verfahren],
    ['Verfahrensführende Stelle', f => f.stelle, true], ['Ansprechpartner / Anschrift Stelle', f => f.stelleAnschrift],
    ['Geschädigter (Name)', f => f.gesch.name], ['Anschrift Geschädigter', f => f.gesch.anschrift],
    ['Ersatzpflichtiger (', f => f.ersatz.name, true], ['Anschrift Ersatzpflichtiger', f => f.ersatz.anschrift],
    ['Jagdbezirk', f => f.jagdbezirk], ['Zeuge 1', f => f.zeuge1], ['Zeuge 2', f => f.zeuge2],
    ['Vertreter Jagdvorstand', f => f.vertreter], ['Gemarkung', f => f.gemarkung], ['Flur', f => f.flur],
    ['Flurstück(e)', f => f.flurstueck], ['Schlagname (betrieblich)', f => f.schlag],
    ['Schlaggröße (ha)', f => n(f.schlagHa)], ['Kultur', f => f.kultur], ['Wildart (Verursacher)', f => f.wildart],
    ['Datum Kenntnis vom Schaden', f => dateSerial(f.kenntnis)], ['Datum Anmeldung bei Amt/Gemeinde', f => dateSerial(f.anmeldung)],
    ['Selbstbehalt in % des Schadens', f => pct(f.sbProz)], ['Selbstbehalt in € (fester Betrag)', f => n(f.sbEuro)],
    ['Höchstbetrag / Deckelung je Fall (€)', f => n(f.deckel)], ['Vereinbarte Wildschadenspauschale (€/Jahr)', f => n(f.pauschale)],
    ['Besteuerung des Geschädigten', f => f.ustGesch], ['Bemerkungen fürs Protokoll', f => f.bemerkung]
  ];
  function n(v) { if (v === '' || v === null || v === undefined) return ''; const x = Number(String(v).replace(',', '.')); return isNaN(x) ? '' : x; }
  function pct(v) { const x = n(v); return x === '' ? '' : Math.round(x * 10000) / 1000000; }

  async function fillTable(sh, rep, opt) {
    const a = sh.find(opt.anchor); if (!a) { rep.missing.push(sh.name + ': Tabelle „' + opt.anchor + '“'); return; }
    const cols = {};
    for (const [h, k] of Object.entries(opt.cols)) {
      const c = sh.colInRow(a.row, h); if (c) cols[k] = c; else rep.missing.push(sh.name + ': Spalte „' + h + '“');
    }
    // Datenzeilen bestimmen: bis „Summe“/„Mittelwert“ oder max
    const lastCol = Math.max(...Object.values(cols));
    const rows = [];
    for (let r = a.row + 1; rows.length < opt.max && r < a.row + 200; r++) {
      let stop = false;
      for (let c = 1; c <= lastCol + 1; c++) {
        const t = norm(sh.text(r, c));
        if (t && /^(summe|mittelwert|übertrag|tipp)/.test(t)) { stop = true; break; }
      }
      if (stop) break; rows.push(r);
    }
    // leeren (z. B. Beispielzeile), dann schreiben
    for (const r of rows) for (const c of Object.values(cols)) sh.set(r, c, '');
    const data = opt.data;
    if (data.length > rows.length) rep.warn.push(opt.what + ': ' + data.length + ' Einträge, Vorlage hat nur ' + rows.length + ' Zeilen – Rest nicht übertragen.');
    data.slice(0, rows.length).forEach((d, i) => {
      for (const [k, c] of Object.entries(cols)) {
        const v = d[k]; if (v === undefined) continue;
        if (sh.set(rows[i], c, v)) { if (v !== '' && v !== null) rep.written++; } else rep.formula.push(sh.name + '!' + numToCol(c) + rows[i]);
      }
    });
  }

  async function exportFall(buf, f, opts = {}) {
    const wb = await Workbook.open(buf);
    const rep = { written: 0, missing: [], formula: [], warn: [] };
    const fa = await wb.sheet('Fall');
    if (!fa) { rep.missing.push('Blatt „Fall“'); }
    else {
      for (const [lab, get, pre] of FALL_LABELS) {
        const p = fa.find(lab, { prefix: !!pre });
        if (!p) { rep.missing.push('Fall: „' + lab + '“'); continue; }
        const r = fa.rightOf(p); const v = get(f);
        if (fa.set(r.row, r.col, v === undefined ? '' : v)) { if (v !== '' && v != null) rep.written++; }
        else rep.formula.push('Fall!' + numToCol(r.col) + r.row + ' (' + lab + ')');
      }
      await fillTable(fa, rep, { anchor: 'Art des Termins', max: 5, what: 'Ortstermine',
        cols: { 'Datum': 'datum', 'Art des Termins': 'art', 'Anwesende': 'anw', 'Feststellungen': 'fest' },
        data: f.termine.map(t => ({ datum: dateSerial(t.datum) || '', art: t.art || '', anw: t.anwesende || '', fest: t.feststellungen || '' })) });
    }
    const sc = await wb.sheet('Schadstellen');
    if (!sc) rep.missing.push('Blatt „Schadstellen“');
    else await fillTable(sc, rep, { anchor: 'Länge m', max: 30, what: 'Schadstellen',
      cols: { 'Länge m': 'l', 'Breite m': 'b', 'Verlust %': 'v', 'Grünland-Schadgrad': 'g', 'GPS Breite (N)': 'lat',
        'GPS Länge (E)': 'lon', 'Foto-Nr.': 'foto', 'Ortstermin': 'termin', 'Bemerkung': 'bem' },
      data: (opts.stellenRows || []) });
    const be = await wb.sheet('Berechnung');
    if (!be) rep.missing.push('Blatt „Berechnung“');
    else {
      const e = f.ertrag;
      for (const [lab, v, pre] of [['Ertragsermittlung (Methode)', e.methode || ''], ['Betriebsdurchschnitt (t/ha)', n(e.betrieb)],
        ['Schätzung am Bestand', n(e.schaetzung), true], ['Strohertrag (t/ha)', n(e.stroh), true]]) {
        const p = be.find(lab, { prefix: !!pre }); if (!p) { rep.missing.push('Berechnung: „' + lab + '“'); continue; }
        const r = be.rightOf(p);
        if (be.set(r.row, r.col, v)) { if (v !== '') rep.written++; } else rep.formula.push('Berechnung!' + numToCol(r.col) + r.row);
      }
      await fillTable(be, rep, { anchor: 'Erntemenge kg', max: 5, what: 'Probedrusch',
        cols: { 'Erntemenge kg': 'kg', 'Probefläche m²': 'm2' }, data: e.proben.map(p => ({ kg: n(p.kg), m2: n(p.m2) })) });
      await fillTable(be, rep, { anchor: 'Kostenquelle', max: 8, what: 'Wiederherstellung',
        cols: { 'Arbeitsgang': 'gang', 'Kostenquelle': 'quelle', 'Fläche ha': 'ha', 'Eigener Satz €/ha': 'eigen', 'Lohnunternehmer €': 'lu' },
        data: f.wieder.map(w => ({ gang: w.gang || '', quelle: w.quelle || '', ha: n(w.ha), eigen: n(w.eigen), lu: n(w.lu) })) });
      await fillTable(be, rep, { anchor: 'Pauschal €/ha', max: 3, what: 'Saatgut',
        cols: { 'Saatgut': 'sorte', 'Methode': 'methode', 'Fläche ha': 'ha', 'Pauschal €/ha': 'pauschal', 'Rechnung €': 'rechnung' },
        data: f.saat.map(s => ({ sorte: s.sorte || '', methode: s.methode || '', ha: n(s.ha), pauschal: n(s.pauschal), rechnung: n(s.rechnung) })) });
    }
    const ze = await wb.sheet('Zeiterfassung');
    if (!ze) rep.missing.push('Blatt „Zeiterfassung“');
    else await fillTable(ze, rep, { anchor: 'Tätigkeit', max: 20, what: 'Zeiterfassung',
      cols: { 'Datum': 'datum', 'Tätigkeit': 'taet', 'Beschreibung': 'beschr', 'Stunden': 'std', 'km': 'km' },
      data: f.zeiten.filter(z => !z.start).map(z => ({ datum: dateSerial(z.datum) || '', taet: z.taetigkeit || '', beschr: z.beschr || '', std: n(z.stunden), km: n(z.km) })) });
    const re = await wb.sheet('Rechnung');
    if (re && f.rech) {
      const rc = f.rech;
      const sonst = rc.empf === 'Sonstiger' ? [rc.sonstName, rc.sonstAnschrift].filter(Boolean).join(', ') : '';
      for (const [lab, v, pre] of [['Empfänger:', rc.empf || ''], ['Nur bei „Sonstiger“', sonst, true], ['Rechnungs-Nr.', rc.nr || ''],
        ['Rechnungsdatum', dateSerial(rc.datum) || ''], ['Umsatzsteuer', rc.ust || '']]) {
        const p = re.find(lab, { prefix: !!pre }); if (!p) continue;
        const r = re.rightOf(p);
        if (re.set(r.row, r.col, v)) { if (v !== '') rep.written++; }
      }
    }
    const sd = await wb.sheet('Stammdaten');
    if (sd && opts.stamm) {
      for (const [lab, k, pre] of [['Steuernummer', 'steuernr'], ['Bank', 'bank'], ['Bestellt durch', 'bestellt', true]]) {
        const v = opts.stamm[k]; if (!v) continue;
        const p = sd.find(lab, { prefix: !!pre }); if (!p) continue;
        const r = sd.rightOf(p);
        if (sd.value(r.row, r.col) === null && sd.set(r.row, r.col, v)) rep.written++;
      }
    }
    const blob = await wb.save();
    return { blob, rep };
  }

  global.XL = { Workbook, readStammdaten, exportFall, dateSerial, norm };
})(typeof window !== 'undefined' ? window : globalThis);
