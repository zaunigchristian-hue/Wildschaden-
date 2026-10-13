/* Stufe 3: Lageplan der Schadstellen.
   Zeichnet Schadstellen (Punkte und GPS-Umgänge) auf OpenStreetMap-Kacheln.
   Ohne Internet: gleiche Zeichnung ohne Kartenhintergrund. */
(function (global) {
  'use strict';
  const TILE = 256;
  const px = (lon, lat, z) => {
    const n = Math.pow(2, z) * TILE, r = lat * Math.PI / 180;
    return [(lon + 180) / 360 * n, (1 - Math.log(Math.tan(r) + 1 / Math.cos(r)) / Math.PI) / 2 * n];
  };
  const tileCache = new Map();
  function loadTile(z, x, y, ms) {
    const key = z + '/' + x + '/' + y;
    if (tileCache.has(key)) return tileCache.get(key);
    const p = new Promise(res => {
      const img = new Image(); img.crossOrigin = 'anonymous';
      const t = setTimeout(() => { img.src = ''; res(null); }, ms);
      img.onload = () => { clearTimeout(t); res(img); };
      img.onerror = () => { clearTimeout(t); tileCache.delete(key); res(null); };
      img.src = 'https://tile.openstreetmap.org/' + key + '.png';
    });
    tileCache.set(key, p); return p;
  }

  /* stellen: [{nr, lat, lon, poly:[[lat,lon],…]}] */
  function features(stellen) {
    const out = [];
    stellen.forEach((s, i) => {
      const nr = s.nr || i + 1;
      const poly = (s.poly || []).filter(p => isFinite(p[0]) && isFinite(p[1]));
      const lat = Number(String(s.lat).replace(',', '.')), lon = Number(String(s.lon).replace(',', '.'));
      if (poly.length >= 3) {
        const c = [poly.reduce((a, p) => a + p[0], 0) / poly.length, poly.reduce((a, p) => a + p[1], 0) / poly.length];
        out.push({ nr, poly, c });
      } else if (s.lat !== '' && s.lat != null && isFinite(lat) && isFinite(lon) && lat && lon) out.push({ nr, c: [lat, lon] });
    });
    return out;
  }

  async function render(stellen, W, H, o = {}) {
    const fs = features(stellen);
    if (!fs.length) return null;
    const pts = fs.flatMap(f => f.poly || [f.c]);
    const lats = pts.map(p => p[0]), lons = pts.map(p => p[1]);
    const box = [Math.min(...lons), Math.min(...lats), Math.max(...lons), Math.max(...lats)];
    let z = 19;
    const pad = 70;
    while (z > 5) {
      const a = px(box[0], box[3], z), b = px(box[2], box[1], z);
      if (b[0] - a[0] <= W - 2 * pad && b[1] - a[1] <= H - 2 * pad) break;
      z--;
    }
    z = Math.max(5, Math.min(z + (o.zoomDelta || 0), 19));
    if (pts.length === 1 && !o.zoomDelta) z = Math.min(z, 18);
    const cLon = (box[0] + box[2]) / 2, cLat = (box[1] + box[3]) / 2;
    const [cx, cy] = px(cLon, cLat, z);
    const ox = cx - W / 2, oy = cy - H / 2;
    const cv = document.createElement('canvas'); cv.width = W; cv.height = H;
    const g = cv.getContext('2d');
    g.fillStyle = '#ece9df'; g.fillRect(0, 0, W, H);
    // Kacheln
    let ok = 0, need = 0;
    if (o.tiles !== false && navigator.onLine !== false) {
      const jobs = [];
      for (let tx = Math.floor(ox / TILE); tx <= Math.floor((ox + W) / TILE); tx++)
        for (let ty = Math.floor(oy / TILE); ty <= Math.floor((oy + H) / TILE); ty++) {
          const m = Math.pow(2, z); if (ty < 0 || ty >= m) continue;
          need++;
          jobs.push(loadTile(z, ((tx % m) + m) % m, ty, o.timeout || 7000).then(img => { if (img) { g.drawImage(img, tx * TILE - ox, ty * TILE - oy); ok++; } }));
        }
      await Promise.all(jobs);
    }
    if (!ok) { // Raster als Ersatz
      g.strokeStyle = 'rgba(0,0,0,.08)'; g.lineWidth = 1;
      for (let x = (-ox % 64 + 64) % 64; x < W; x += 64) { g.beginPath(); g.moveTo(x, 0); g.lineTo(x, H); g.stroke(); }
      for (let y = (-oy % 64 + 64) % 64; y < H; y += 64) { g.beginPath(); g.moveTo(0, y); g.lineTo(W, y); g.stroke(); }
    }
    const P = (lat, lon) => { const [x, y] = px(lon, lat, z); return [x - ox, y - oy]; };
    const S = o.scale || W / 900; // Strichstärken skalieren
    // Flächen
    for (const f of fs) if (f.poly) {
      g.beginPath(); f.poly.forEach((p, i) => { const [x, y] = P(p[0], p[1]); i ? g.lineTo(x, y) : g.moveTo(x, y); }); g.closePath();
      g.fillStyle = 'rgba(200,30,30,.28)'; g.fill(); g.strokeStyle = '#c01818'; g.lineWidth = 3 * S; g.stroke();
    }
    // Nummern
    for (const f of fs) {
      const [x, y] = P(f.c[0], f.c[1]), r = 14 * S;
      g.beginPath(); g.arc(x, y, r, 0, 7); g.fillStyle = '#c01818'; g.fill(); g.lineWidth = 2.5 * S; g.strokeStyle = '#fff'; g.stroke();
      g.fillStyle = '#fff'; g.font = `bold ${Math.round(15 * S)}px sans-serif`; g.textAlign = 'center'; g.textBaseline = 'middle'; g.fillText(String(f.nr), x, y + 1);
    }
    // Maßstab
    const mpp = 156543.03392 * Math.cos(cLat * Math.PI / 180) / Math.pow(2, z);
    const nice = [5, 10, 20, 25, 50, 100, 200, 250, 500, 1000, 2000, 5000, 10000];
    const len = nice.find(n => n / mpp >= 90 * S) || 10000, lpx = len / mpp;
    const bx = 16 * S, by = H - 20 * S;
    g.fillStyle = 'rgba(255,255,255,.85)'; g.fillRect(bx - 8 * S, by - 24 * S, lpx + 16 * S, 34 * S);
    g.fillStyle = '#111'; g.fillRect(bx, by, lpx, 4 * S);
    g.font = `${Math.round(13 * S)}px sans-serif`; g.textAlign = 'left'; g.textBaseline = 'alphabetic';
    g.fillText(len >= 1000 ? len / 1000 + ' km' : len + ' m', bx, by - 6 * S);
    // Nordpfeil
    const nx = W - 30 * S, ny = 34 * S;
    g.fillStyle = 'rgba(255,255,255,.85)'; g.beginPath(); g.arc(nx, ny + 2 * S, 22 * S, 0, 7); g.fill();
    g.fillStyle = '#111'; g.beginPath(); g.moveTo(nx, ny - 14 * S); g.lineTo(nx - 8 * S, ny + 10 * S); g.lineTo(nx, ny + 5 * S); g.lineTo(nx + 8 * S, ny + 10 * S); g.closePath(); g.fill();
    g.font = `bold ${Math.round(11 * S)}px sans-serif`; g.textAlign = 'center'; g.fillText('N', nx, ny + 21 * S);
    // Quelle
    const attr = ok ? '© OpenStreetMap-Mitwirkende' : 'ohne Kartenhintergrund (offline)';
    g.font = `${Math.round(11 * S)}px sans-serif`; const aw = g.measureText(attr).width;
    g.fillStyle = 'rgba(255,255,255,.85)'; g.fillRect(W - aw - 14 * S, H - 20 * S, aw + 14 * S, 20 * S);
    g.fillStyle = '#333'; g.textAlign = 'right'; g.fillText(attr, W - 7 * S, H - 6 * S);
    return { canvas: cv, online: ok > 0, complete: ok === need, zoom: z, count: fs.length };
  }

  global.WSMAP = { render, features };
})(typeof window !== 'undefined' ? window : globalThis);
