/* Stufe 2: Schätzprotokoll (mit Berechnungsanhang) und Schätzerrechnung als PDF.
   Aufbau folgt den Blättern „Protokoll“, „Anhang“ und „Rechnung“ der Excel-Vorlage. */
(function (global) {
  'use strict';
  const A4 = [595.28, 841.89];
  const M = { l: 56, r: 50, t: 50, b: 60 };
  const INK = () => PDFLib.rgb(0.1, 0.12, 0.11);
  const MUTE = () => PDFLib.rgb(0.38, 0.42, 0.39);
  const LINE = () => PDFLib.rgb(0.78, 0.77, 0.72);
  const ACC = () => PDFLib.rgb(0.18, 0.36, 0.23);
  const FILL = () => PDFLib.rgb(0.95, 0.945, 0.92);

  /* Nur Zeichen, die die Standardschrift (WinAnsi) kann */
  const EXTRA = '€‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ';
  function safe(s) {
    return String(s == null ? '' : s)
      .replace(/[   ]/g, ' ').replace(/[−]/g, '-').replace(/[→]/g, '->')
      .replace(/[\r\t]/g, ' ')
      .split('').map(c => { const n = c.charCodeAt(0); return (n === 10 || (n >= 32 && n <= 126) || (n >= 160 && n <= 255) || EXTRA.includes(c)) ? c : '?'; }).join('');
  }
  const fmt = (v, d = 2) => v == null || isNaN(v) ? '–' : (Math.abs(v) < 0.005 ? 0 : Number(v)).toLocaleString('de-DE', { minimumFractionDigits: d, maximumFractionDigits: d });
  const eur = v => fmt(v) + ' €';
  const dDE = iso => iso ? iso.split('-').reverse().join('.') : '';

  class Doc {
    static async create(meta) {
      const d = new Doc();
      d.pdf = await PDFLib.PDFDocument.create();
      d.pdf.setTitle(safe(meta.title)); d.pdf.setAuthor(safe(meta.author || '')); d.pdf.setCreator('Wildschaden Feld');
      d.f = await d.pdf.embedFont(PDFLib.StandardFonts.Helvetica);
      d.fb = await d.pdf.embedFont(PDFLib.StandardFonts.HelveticaBold);
      d.footer = meta.footer || '';
      d.newPage();
      return d;
    }
    get W() { return A4[0] - M.l - M.r; }
    newPage() { this.page = this.pdf.addPage(A4); this.y = A4[1] - M.t; }
    need(h) { if (this.y - h < M.b) { this.newPage(); return true; } return false; }
    width(t, size, bold) { return (bold ? this.fb : this.f).widthOfTextAtSize(safe(t), size); }
    wrap(t, size, maxW, bold) {
      const out = [];
      for (const para of safe(t).split('\n')) {
        let line = '';
        for (const word of para.split(/ +/)) {
          const test = line ? line + ' ' + word : word;
          if (this.width(test, size, bold) <= maxW) { line = test; continue; }
          if (line) out.push(line);
          // überlange Wörter hart trennen
          let w = word;
          while (this.width(w, size, bold) > maxW) {
            let k = w.length; while (k > 1 && this.width(w.slice(0, k), size, bold) > maxW) k--;
            out.push(w.slice(0, k)); w = w.slice(k);
          }
          line = w;
        }
        out.push(line);
      }
      return out;
    }
    text(t, x, y, o = {}) {
      const size = o.size || 10, bold = !!o.bold; let s = safe(t);
      if (o.align === 'right') x = x - this.width(s, size, bold);
      if (o.align === 'center') x = x - this.width(s, size, bold) / 2;
      this.page.drawText(s, { x, y, size, font: bold ? this.fb : this.f, color: o.color ? o.color() : INK() });
    }
    para(t, o = {}) {
      const size = o.size || 10, lh = size * 1.35, x = M.l + (o.indent || 0);
      for (const ln of this.wrap(t, size, this.W - (o.indent || 0), o.bold)) {
        this.need(lh); this.y -= lh; this.text(ln, x, this.y + 3, o);
      }
      this.y -= (o.after ?? 4);
    }
    gap(h) { this.y -= h; }
    rule(c = LINE, w = 0.6) { this.page.drawLine({ start: { x: M.l, y: this.y }, end: { x: A4[0] - M.r, y: this.y }, thickness: w, color: c() }); }
    h1(t, sub) {
      this.need(50); this.y -= 22; this.text(t, M.l, this.y, { size: 17, bold: true });
      if (sub) { this.y -= 15; this.text(sub, M.l, this.y, { size: 9, color: MUTE }); }
      this.y -= 10;
    }
    h2(t) {
      if (this.need(48)) this.y -= 4;
      this.y -= 20; this.text(t, M.l, this.y, { size: 11.5, bold: true, color: ACC });
      this.y -= 5; this.rule(ACC, 0.8); this.y -= 2;
    }
    /* Zweispaltige Angaben: Beschriftung | Wert */
    kv(rows, o = {}) {
      const lw = o.labelW || 165, size = 10, lh = 13.5;
      for (const [k, v, opt] of rows) {
        const bold = opt && opt.bold;
        const lines = this.wrap(v == null || v === '' ? '–' : v, size, this.W - lw, bold);
        const h = lines.length * lh + 3;
        this.need(h);
        this.text(k, M.l, this.y - lh + 3, { size: 9.5, color: MUTE });
        lines.forEach((ln, i) => this.text(ln, M.l + lw, this.y - lh * (i + 1) + 3, { size, bold }));
        this.y -= h;
      }
    }
    /* Tabelle mit Umbruch je Zelle und Seitenwechsel inkl. Kopfzeile */
    table(cols, rows, o = {}) {
      const size = o.size || 9, lh = size * 1.3, pad = 3;
      const total = cols.reduce((a, c) => a + c.w, 0), k = this.W / total;
      const ws = cols.map(c => c.w * k);
      const xs = ws.reduce((a, w, i) => (a.push(i ? a[i - 1] + ws[i - 1] : M.l), a), []);
      const head = () => {
        const hl = cols.map((c, i) => this.wrap(c.t, size - 0.5, ws[i] - 2 * pad, true));
        const h = Math.max(...hl.map(l => l.length)) * lh + 2 * pad;
        this.need(h + lh * 2);
        this.page.drawRectangle({ x: M.l, y: this.y - h, width: this.W, height: h, color: FILL() });
        hl.forEach((ls, i) => ls.forEach((ln, j) => this.text(ln, cols[i].align === 'right' ? xs[i] + ws[i] - pad : xs[i] + pad, this.y - pad - lh * (j + 1) + 3, { size: size - 0.5, bold: true, align: cols[i].align })));
        this.y -= h;
      };
      head();
      rows.forEach((r, ri) => {
        const bold = r.bold; const cells = r.cells || r;
        const cl = cells.map((c, i) => this.wrap(c == null || c === '' ? '' : c, size, ws[i] - 2 * pad, bold));
        const h = Math.max(1, ...cl.map(l => l.length)) * lh + 2 * pad;
        if (this.need(h)) head();
        if (r.total) this.page.drawLine({ start: { x: M.l, y: this.y }, end: { x: A4[0] - M.r, y: this.y }, thickness: 0.9, color: INK() });
        cl.forEach((ls, i) => ls.forEach((ln, j) => this.text(ln, cols[i].align === 'right' ? xs[i] + ws[i] - pad : xs[i] + pad, this.y - pad - lh * (j + 1) + 3, { size, bold, align: cols[i].align })));
        this.y -= h;
        if (!r.total && ri < rows.length - 1) this.page.drawLine({ start: { x: M.l, y: this.y }, end: { x: A4[0] - M.r, y: this.y }, thickness: 0.3, color: LINE() });
      });
      this.y -= 4;
    }
    sumRows(rows) {
      const size = 10, lh = 15;
      for (const [k, v, o] of rows) {
        const strong = o && o.total;
        this.need(lh + 6);
        if (strong) { this.y -= 3; this.page.drawLine({ start: { x: M.l, y: this.y }, end: { x: A4[0] - M.r, y: this.y }, thickness: 1, color: INK() }); }
        this.y -= lh;
        this.text(k, M.l, this.y + 4, { size: strong ? 11 : size, bold: strong });
        this.text(v, A4[0] - M.r, this.y + 4, { size: strong ? 11 : size, bold: strong, align: 'right' });
      }
      this.y -= 4;
    }
    async image(bytes, type, maxW, maxH, x) {
      const img = type === 'png' ? await this.pdf.embedPng(bytes) : await this.pdf.embedJpg(bytes);
      const sc = Math.min(maxW / img.width, maxH / img.height, 1);
      const w = img.width * sc, h = img.height * sc;
      return { img, w, h, draw: (px, py) => this.page.drawImage(img, { x: px, y: py, width: w, height: h }) };
    }
    finish() {
      const pages = this.pdf.getPages(), n = pages.length;
      pages.forEach((p, i) => {
        p.drawLine({ start: { x: M.l, y: 38 }, end: { x: A4[0] - M.r, y: 38 }, thickness: 0.4, color: LINE() });
        p.drawText(safe(this.footer), { x: M.l, y: 26, size: 8, font: this.f, color: MUTE() });
        const s = 'Seite ' + (i + 1) + ' von ' + n;
        p.drawText(s, { x: A4[0] - M.r - this.f.widthOfTextAtSize(s, 8), y: 26, size: 8, font: this.f, color: MUTE() });
      });
      return this.pdf.save();
    }
  }

  function dataUrlBytes(u) { const b = atob(u.split(',')[1]); const a = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) a[i] = b.charCodeAt(i); return a; }
  function senderLines(st) { return [st.name, st.strasse && st.ort ? st.strasse + ', ' + st.ort : (st.strasse || st.ort || '')].filter(Boolean); }

  /* ---------------- Schätzprotokoll ---------------- */
  async function protokoll(f, S, r, ctx) {
    const st = S.stamm;
    const d = await Doc.create({ title: 'Schätzprotokoll ' + (f.az || ''), author: st.name,
      footer: 'Schätzprotokoll  ·  Az. ' + (f.az || '–') + '  ·  ' + (st.name || '') });
    // Kopf
    const sl = senderLines(st);
    d.text(sl[0] || '', M.l, d.y - 10, { size: 12, bold: true });
    d.text(sl[1] || '', M.l, d.y - 24, { size: 9.5, color: MUTE });
    d.text('Wildschadensschätzer' + (st.bestellt ? ' – bestellt durch ' + st.bestellt : ''), M.l, d.y - 37, { size: 9.5, color: MUTE });
    const contact = [st.telefon && 'Tel. ' + st.telefon, st.email].filter(Boolean);
    contact.forEach((c, i) => d.text(c, A4[0] - M.r, d.y - 24 - i * 13, { size: 9, color: MUTE, align: 'right' }));
    d.y -= 46; d.rule(INK, 1);
    d.h1('Wildschadens-Schätzprotokoll', 'nach §§ 29 ff. Bundesjagdgesetz i. V. m. Jagdgesetz für das Land Brandenburg (BbgJagdG)');

    d.h2('1. Verfahren');
    d.kv([['Aktenzeichen', f.az], ['Verfahrensart', f.verfahren], ['Verfahrensführende Stelle', [f.stelle, f.stelleAnschrift].filter(Boolean).join(', ')]]);
    d.h2('2. Beteiligte');
    d.kv([['Geschädigter', [f.gesch.name, f.gesch.anschrift].filter(Boolean).join(', ')],
      ['Ersatzpflichtiger', [f.ersatz.name, f.ersatz.anschrift].filter(Boolean).join(', ')],
      ['Jagdbezirk', f.jagdbezirk], ['Zeugen', [f.zeuge1, f.zeuge2].filter(Boolean).join(', ')], ['Vertreter Jagdvorstand', f.vertreter]]);
    d.h2('3. Schadensfläche');
    d.kv([['Gemarkung / Flur / Flurstück', [f.gemarkung, f.flur, f.flurstueck].map(x => x || '–').join(' / ')],
      ['Schlag', f.schlag], ['Schlaggröße', f.schlagHa ? fmt(ctx.num(f.schlagHa), 2) + ' ha' : ''],
      ['Kultur', f.kultur ? f.kultur + (r.gruppe ? ' (' + r.gruppe + ')' : '') : ''], ['Wildart', f.wildart]]);
    d.h2('4. Fristen');
    const fr = ctx.frist(f);
    d.kv([['Kenntnis vom Schaden', dDE(f.kenntnis)], ['Anmeldung', dDE(f.anmeldung)], ['Prüfung', fr[1], { bold: fr[0] === 'bad' }]]);
    d.h2('5. Ortstermine');
    if (f.termine.length) d.table([{ t: 'Datum', w: 13 }, { t: 'Art', w: 17 }, { t: 'Anwesende', w: 32 }, { t: 'Feststellungen', w: 38 }],
      f.termine.map(t => [dDE(t.datum), t.art, t.anwesende, t.feststellungen]));
    else d.para('Keine Ortstermine erfasst.', { color: MUTE });
    d.h2('6. Feststellungen zum Schaden');
    const methode = f.ertrag.methode ? f.ertrag.methode + (r.ertrag != null ? ': ' + fmt(r.ertrag) + ' t/ha' : '') : '';
    d.kv([['Anzahl Schadstellen', String(r.anzahl)], ['Schadfläche gesamt', fmt(r.flaeche, 1) + ' m²'],
      ['Schadfläche wirksam', fmt(r.eff, 1) + ' m² (nach Verlustanteil)'], ['Ertragsermittlung', methode], ['Bemerkungen', f.bemerkung]]);
    d.h2('7. Ergebnis der Schätzung');
    const sbk = -(r.sb + r.kuerzung);
    d.sumRows([['Ertragsschaden', eur(r.ertragsschaden)], ['Wiederherstellung', eur(r.wiederSumme)], ['Saatgut und Material', eur(r.saatSumme)],
      ['Selbstbehalt / Höchstbetrag (Jagdpachtvertrag)', eur(sbk)], ['Durchschnittssatz-Zuschlag', eur(r.zuschlag)],
      ['Festgestellter Wildschaden', eur(r.gesamt), { total: true }]]);
    if (ctx.num(f.pauschale)) d.para('Im Jagdpachtvertrag ist eine Wildschadenspauschale von ' + eur(ctx.num(f.pauschale)) + ' je Jahr vereinbart.', { size: 9, color: MUTE });
    d.para('Die Berechnung im Einzelnen ist als Anhang beigefügt.', { size: 9.5 });

    // Unterschrift Schätzer
    d.need(120); d.gap(14);
    const ortDatum = [f.prot.ort, dDE(f.prot.datum)].filter(Boolean).join(', ');
    d.text('Ort, Datum: ' + (ortDatum || ''), M.l, d.y - 12, { size: 10 });
    d.y -= 20;
    if (S.signatur) {
      const im = await d.image(dataUrlBytes(S.signatur), 'png', 200, 60);
      im.draw(M.l, d.y - im.h); d.y -= im.h + 2;
    } else d.y -= 50;
    d.page.drawLine({ start: { x: M.l, y: d.y }, end: { x: M.l + 230, y: d.y }, thickness: 0.6, color: INK() });
    d.y -= 12; d.text('Unterschrift Wildschadensschätzer: ' + (st.name || ''), M.l, d.y, { size: 9, color: MUTE });

    // Unterschriften der Beteiligten (optional)
    const sigs = (f.sigs || []).filter(s => s.png);
    if (sigs.length) {
      d.h2('8. Unterschriften der Beteiligten');
      for (const s of sigs) {
        d.need(110);
        d.para(s.rolle + (s.name ? ': ' + s.name : ''), { bold: true, after: 0 });
        if (s.text) d.para(s.text, { size: 9, after: 2 });
        const im = await d.image(dataUrlBytes(s.png), 'png', 200, 55);
        im.draw(M.l, d.y - im.h); d.y -= im.h + 2;
        d.page.drawLine({ start: { x: M.l, y: d.y }, end: { x: M.l + 230, y: d.y }, thickness: 0.6, color: INK() });
        d.y -= 11; d.text('unterschrieben am ' + new Date(s.zeit).toLocaleString('de-DE', { dateStyle: 'medium', timeStyle: 'short' }) + ' auf dem Gerät des Schätzers', M.l, d.y, { size: 8, color: MUTE });
        d.y -= 12;
      }
    }

    // ---------- Anhang ----------
    d.newPage();
    d.h1('Anhang zum Schätzprotokoll – Az. ' + (f.az || ''), 'Kultur: ' + (f.kultur || '–') + '   |   Wildart: ' + (f.wildart || '–') + '   |   Schlag: ' + (f.schlag || '–'));
    d.h2('A) Ertrag und Preis');
    const rows = [['Methode', f.ertrag.methode || '–']];
    if (f.ertrag.methode === 'Probedrusch') rows.push(['Probedrusch Mittelwert', r.probenMittel != null ? fmt(r.probenMittel) + ' t/ha' : '–']);
    rows.push(['Angesetzter Ertrag', r.ertrag != null ? fmt(r.ertrag) + ' t/ha' : '–'], ['Preis', r.preis != null ? eur(r.preis) + ' je t' : '–']);
    if (r.gruppe === 'Getreide') rows.push(['Strohertrag / Strohpreis', (ctx.num(f.ertrag.stroh) != null ? fmt(ctx.num(f.ertrag.stroh)) + ' t/ha' : '–') + ' / ' + (r.strohpreis != null ? eur(r.strohpreis) + ' je t' : '–')]);
    rows.push(['Ersparte Kosten', r.erspartHa != null ? eur(r.erspartHa) + ' je ha' : '–'], ['Schadfläche wirksam', fmt(r.ha, 4) + ' ha']);
    d.kv(rows);
    d.sumRows([['Wert Korn / Aufwuchs', eur(r.wert)], ...(r.gruppe === 'Getreide' ? [['Wert Stroh', eur(r.stroh)]] : []), ['Ersparte Kosten', eur(-r.erspart)],
      ['Ertragsschaden netto', eur(r.ertragsschaden), { total: true }]]);
    const pr = f.ertrag.proben.map((p, i) => [ctx.num(p.kg), ctx.num(p.m2), i]).filter(([kg, m2]) => kg != null && m2);
    if (f.ertrag.methode === 'Probedrusch' && pr.length) {
      d.para('Probedrusch-Einzelwerte', { bold: true, size: 9.5, after: 2 });
      d.table([{ t: 'Probe', w: 1 }, { t: 'kg', w: 1, align: 'right' }, { t: 'm²', w: 1, align: 'right' }, { t: 't/ha', w: 1, align: 'right' }],
        pr.map(([kg, m2, i]) => ['Probe ' + (i + 1), fmt(kg), fmt(m2, 1), fmt(kg / m2 * 10)]));
    }
    d.h2('B) Schadstellen');
    const gl = r.gruppe === 'Grünland';
    const sRows = ctx.stellen.map(s => [String(s.nr), s.mass, fmt(s.flaeche, 1), s.verlust, fmt(s.eff, 1), ...(gl ? [s.grad] : []), s.gps, s.fotos, s.bem]);
    sRows.push({ total: true, bold: true, cells: ['Summe', '', fmt(r.flaeche, 1), '', fmt(r.eff, 1), ...(gl ? [''] : []), '', '', ''] });
    d.table([{ t: 'Nr.', w: 8 }, { t: 'Maße / Art', w: 15 }, { t: 'Fläche m²', w: 10, align: 'right' }, { t: 'Verlust', w: 8, align: 'right' }, { t: 'wirksam m²', w: 10, align: 'right' },
      ...(gl ? [{ t: 'Grad', w: 8 }] : []), { t: 'GPS', w: 17 }, { t: 'Fotos', w: 8 }, { t: 'Bemerkung', w: 16 }], sRows, { size: 8 });
    if (gl && (r.grad.leicht || r.grad.mittel || r.grad.schwer)) {
      d.table([{ t: 'Schadgrad', w: 1 }, { t: 'Fläche m²', w: 1, align: 'right' }, { t: 'Fläche ha', w: 1, align: 'right' }, { t: 'Empfohlene Maßnahme', w: 2 }],
        ['leicht', 'mittel', 'schwer'].map(g => [g, fmt(r.grad[g], 1), fmt(r.grad[g] / 10000, 4), ctx.grad[g]]));
    }
    d.h2('C) Wiederherstellung');
    if (f.wieder.length) {
      d.table([{ t: 'Arbeitsgang', w: 30 }, { t: 'Quelle', w: 18 }, { t: 'Fläche ha', w: 11, align: 'right' }, { t: 'Satz', w: 23, align: 'right' }, { t: 'Betrag €', w: 18, align: 'right' }],
        [...f.wieder.map((w, i) => [w.gang, w.quelle || 'KTBL/MR-Satz', fmt(ctx.num(w.ha), 4), ctx.wSatz(w), fmt(r.wieder[i])]),
          { total: true, bold: true, cells: ['Summe', '', '', '', fmt(r.wiederSumme)] }]);
    } else d.para('Keine Wiederherstellungsarbeiten angesetzt.', { color: MUTE, size: 9.5 });
    d.h2('D) Saatgut und Material');
    if (f.saat.length) {
      d.table([{ t: 'Saatgut', w: 30 }, { t: 'Methode', w: 18 }, { t: 'Fläche ha', w: 11, align: 'right' }, { t: 'Ansatz', w: 23, align: 'right' }, { t: 'Betrag €', w: 18, align: 'right' }],
        [...f.saat.map((s, i) => [s.sorte, s.methode || 'kg/ha × Preis', fmt(ctx.num(s.ha), 4), ctx.sSatz(s), fmt(r.saat[i])]),
          { total: true, bold: true, cells: ['Summe', '', '', '', fmt(r.saatSumme)] }]);
    } else d.para('Kein Saatgut / Material angesetzt.', { color: MUTE, size: 9.5 });
    d.h2('E) Zusammenfassung');
    d.sumRows([['Ertragsschaden netto', eur(r.ertragsschaden)], ['Wiederherstellung', eur(r.wiederSumme)], ['Saatgut und Material', eur(r.saatSumme)],
      ['Summe Schaden netto', eur(r.summe)], ['abzüglich Selbstbehalt (Pachtvertrag)', eur(-r.sb)], ['Kürzung auf Höchstbetrag (Pachtvertrag)', eur(-r.kuerzung)],
      ['Schadensersatz netto', eur(r.netto)], ['zzgl. Durchschnittssatz (nur pauschalierend)', eur(r.zuschlag)],
      ['SCHADENSERSATZ GESAMT', eur(r.gesamt), { total: true }]]);

    // Fotoanlage
    if (ctx.fotos && ctx.fotos.length) {
      d.newPage();
      d.h1('Fotoanlage – Az. ' + (f.az || ''), ctx.fotos.length + ' Foto(s), nummeriert wie in der Schadstellen-Liste');
      const colW = (d.W - 16) / 2, maxH = 190;
      for (let i = 0; i < ctx.fotos.length; i += 2) {
        const pair = ctx.fotos.slice(i, i + 2);
        const ims = [];
        for (const p of pair) ims.push(await d.image(p.bytes, 'jpg', colW, maxH));
        const h = Math.max(...ims.map(m => m.h)) + 30;
        d.need(h);
        ims.forEach((m, j) => {
          const x = M.l + j * (colW + 16);
          m.draw(x, d.y - m.h);
          const p = pair[j];
          d.text('Foto ' + p.nr + ' – Schadstelle ' + p.stelle, x, d.y - m.h - 11, { size: 8.5, bold: true });
          d.text([p.zeit, p.gps].filter(Boolean).join('  ·  '), x, d.y - m.h - 22, { size: 7.5, color: MUTE });
        });
        d.y -= h + 8;
      }
    }
    return d.finish();
  }

  /* ---------------- Schätzerrechnung ---------------- */
  async function rechnung(f, S, ctx) {
    const st = S.stamm, rc = f.rech;
    const d = await Doc.create({ title: 'Rechnung ' + (rc.nr || ''), author: st.name,
      footer: [st.name, st.strasse && st.ort ? st.strasse + ', ' + st.ort : '', st.steuernr ? 'St.-Nr. ' + st.steuernr : ''].filter(Boolean).join('  ·  ') });
    const sl = senderLines(st);
    d.text(sl[0] || '', M.l, d.y - 10, { size: 12, bold: true });
    d.text([sl[1], st.telefon && 'Tel. ' + st.telefon, st.email].filter(Boolean).join('  |  '), M.l, d.y - 24, { size: 9, color: MUTE });
    d.y -= 34; d.rule(INK, 1);
    // Empfänger + Rechnungsdaten
    const top = d.y - 26;
    d.text(sl.join(', '), M.l, top + 12, { size: 7, color: MUTE });
    const emp = ctx.empfaenger(f);
    d.wrap([emp.name, emp.anschrift].filter(Boolean).join('\n').replace(/, /g, '\n'), 10.5, 250).forEach((ln, i) => d.text(ln, M.l, top - 4 - i * 14, { size: 10.5 }));
    const meta = [['Rechnungs-Nr.', rc.nr], ['Rechnungsdatum', dDE(rc.datum)], ['Aktenzeichen', f.az], ['Leistungszeitraum', ctx.zeitraum]];
    meta.forEach(([k, v], i) => { d.text(k, 360, top - 4 - i * 15, { size: 9, color: MUTE }); d.text(v || '–', A4[0] - M.r, top - 4 - i * 15, { size: 9.5, bold: i === 0, align: 'right' }); });
    d.y = top - 90;
    d.h1('Rechnung über Wildschadensschätzung');
    d.para('Schadensfall: ' + [f.gesch.name, f.ersatz.name].filter(Boolean).join(' / ') + ', Gemarkung ' + (f.gemarkung || '–') + ', Flur ' + (f.flur || '–') + ', Flurstück ' + (f.flurstueck || '–'), { size: 9.5 });
    d.gap(6);
    const z = ctx.zeiten;
    d.table([{ t: 'Datum', w: 12 }, { t: 'Tätigkeit', w: 22 }, { t: 'Beschreibung', w: 32 }, { t: 'Std.', w: 8, align: 'right' }, { t: '€/h', w: 10, align: 'right' }, { t: 'Betrag €', w: 14, align: 'right' }],
      z.map(x => [dDE(x.datum), x.taetigkeit, x.beschr, fmt(x.std), fmt(st.stundensatz), fmt(x.std * st.stundensatz)]));
    const b = ctx.betraege;
    const rows = [['Zwischensumme Honorar', eur(b.honorarRoh)]];
    if (b.kuerzung > 0.004) rows.push(['Kürzung auf Höchstbetrag ' + fmt(st.tagesmax, 0) + ' € je Kalendertag', eur(-b.kuerzung)]);
    rows.push(['Fahrtkosten: ' + fmt(b.km, 0) + ' km × ' + fmt(st.kmgeld) + ' €', eur(b.fahrt)], ['Nettobetrag', eur(b.netto)]);
    if (b.regel) rows.push(['Umsatzsteuer ' + fmt(st.ust * 100, 0) + ' %', eur(b.ust)]);
    rows.push(['Rechnungsbetrag', eur(b.brutto), { total: true }]);
    d.sumRows(rows);
    if (!b.regel) d.para('Gemäß § 19 UStG wird keine Umsatzsteuer berechnet (Kleinunternehmer).', { size: 9 });
    d.gap(10);
    d.para('Bitte überweisen Sie den Betrag innerhalb von 14 Tagen unter Angabe der Rechnungsnummer auf:', { size: 10 });
    d.kv([['IBAN', st.iban], ['Bank', st.bank], ['Steuernummer', st.steuernr]], { labelW: 100 });
    return d.finish();
  }

  global.PDFX = { protokoll, rechnung, safe };
})(typeof window !== 'undefined' ? window : globalThis);
