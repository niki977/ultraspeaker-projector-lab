/* The Ultraspeaker Projector Lab – pattern di test.
   Ogni pattern è descritto una sola volta come elenco di forme (in pollici, altezza slide 7,5")
   e poi disegnato in due modi: SVG (anteprima e finestra a schermo intero) e forme native PowerPoint (PptxGenJS). */
(function (root) {
  "use strict";

  const H = 7.5;
  const MARK = "USPL:";
  const PT = 1 / 72; // 1 punto in pollici
  const hex = (r, g, b) => [r, g, b].map((v) => Math.round(Math.max(0, Math.min(255, v))).toString(16).padStart(2, "0")).join("").toUpperCase();
  const grayHex = (p) => { const v = Math.round((p / 100) * 255); return hex(v, v, v); };
  const mix = (a, b, t) => a.map((v, i) => v + (b[i] - v) * t);

  /* Formati supportati: la larghezza cambia, l'altezza resta 7,5" */
  const FORMATS = {
    "4:3": 10,
    "16:9": 13.333,
    "3:1": 22.5,
  };
  function formatFromRatio(r) {
    let best = "16:9", d = 9;
    Object.keys(FORMATS).forEach((k) => { const dd = Math.abs(FORMATS[k] / H - r); if (dd < d) { d = dd; best = k; } });
    return best;
  }

  /* ---------- Mattoncini ---------- */
  function Scene(bg) { this.bg = bg; this.items = []; }
  Scene.prototype.rect = function (x, y, w, h, fill, line, lw) { this.items.push({ k: "rect", x, y, w, h, fill, line, lw }); return this; };
  Scene.prototype.ellipse = function (x, y, w, h, fill, line, lw) { this.items.push({ k: "ellipse", x, y, w, h, fill, line, lw }); return this; };
  Scene.prototype.circle = function (cx, cy, r, fill, line, lw) { return this.ellipse(cx - r, cy - r, 2 * r, 2 * r, fill, line, lw); };
  Scene.prototype.pie = function (cx, cy, r, a0, a1, fill) { this.items.push({ k: "pie", cx, cy, r, a0, a1, fill }); return this; };
  // Triangolo isoscele con la punta in alto, poi ruotato in senso orario attorno al centro
  Scene.prototype.tri = function (cx, cy, w, h, rot, fill, line, lw) { this.items.push({ k: "tri", x: cx - w / 2, y: cy - h / 2, w, h, rot, fill, line, lw }); return this; };
  Scene.prototype.text = function (x, y, w, h, text, size, color, o) {
    this.items.push(Object.assign({ k: "text", x, y, w, h, text, size, color, align: "left", valign: "middle", bold: false }, o || {}));
    return this;
  };
  Scene.prototype.image = function (x, y, w, h, data) { this.items.push({ k: "image", x, y, w, h, data }); return this; };
  // Linea orizzontale o verticale disegnata come rettangolo sottile (spessore in punti)
  Scene.prototype.hline = function (x, y, w, pt, color) { const t = pt * PT; return this.rect(x, y - t / 2, w, t, color); };
  Scene.prototype.vline = function (x, y, h, pt, color) { const t = pt * PT; return this.rect(x - t / 2, y, t, h, color); };

  // Titolo piccolo in alto e suggerimento in basso: stanno nei margini e non toccano le aree di misura
  function labels(sc, W, L, key, color, opt) {
    const o = opt || {};
    const tt = L["p." + key + ".title"], hh = L["p." + key + ".hint"];
    if (tt && !o.noTitle) sc.text(0.35, o.titleY != null ? o.titleY : 0.18, W - 0.7, 0.34, tt, 12, color, { bold: true, align: o.titleAlign || "left", spacing: 2 });
    if (hh && !o.noHint) sc.text(0.35, o.hintY != null ? o.hintY : H - 0.52, W - 0.7, 0.34, hh, 12, color, { align: o.hintAlign || "left" });
  }

  // Logo The Ultraspeaker piccolo e semitrasparente, come un watermark (cx, cy = centro)
  const WM_RATIO = 802 / 3955;
  function wm(sc, cx, cy, w, dark) {
    const src = root.PL_WM && (dark ? root.PL_WM.dark : root.PL_WM.light);
    if (!src) return;
    const h = w * WM_RATIO;
    sc.image(cx - w / 2, cy - h / 2, w, h, src);
  }
  const WM_W = 1.35;
  // Posizione standard: in basso a destra, sulla riga del suggerimento
  const wmCorner = (sc, W, dark) => wm(sc, W - 0.35 - WM_W / 2, H - 0.35, WM_W, dark);

  /* ---------- Pattern ---------- */
  const P = {};

  // 1–2. Screen Test di The Ultraspeaker (versione chiara e scura): bordi, proporzioni, trapezio, fuoco, RGB
  function screenTest(W, dark) {
    const bg = dark ? "000000" : "FFFFFF", fg = dark ? "FFFFFF" : "000000";
    const sc = new Scene(bg);
    const s = Math.min(1, W / 13.333);
    // Bordo tratteggiato sul filo della slide: deve vedersi intero su tutti e quattro i lati
    const t = 0.197;
    const nx = Math.max(8, Math.round(W / 0.404)), px = W / nx;
    for (let i = 0; i < nx; i += 2) { sc.rect(i * px, 0, px, t, fg); sc.rect(i * px, H - t, px, t, fg); }
    const ny = Math.round(H / 0.404), py = H / ny;
    for (let i = 0; i < ny; i += 2) { sc.rect(0, i * py, t, py, fg); sc.rect(W - t, i * py, t, py, fg); }
    // Mirini agli angoli (fuoco ai bordi)
    const D = 1.969 * s, inset = 0.39;
    [[inset, inset], [W - inset - D, inset], [inset, H - inset - D], [W - inset - D, H - inset - D]].forEach(([x, y]) => {
      const cx = x + D / 2, cy = y + D / 2, r = D * 0.379, th = 0.079 * s, sq = 0.542 * D;
      sc.circle(cx, cy, r, fg);
      sc.rect(x, cy - th / 2, D, th, fg).rect(cx - th / 2, y, th, D, fg);
      sc.rect(cx - r, cy - th / 2, 2 * r, th, bg).rect(cx - th / 2, cy - r, th, 2 * r, bg);
      sc.rect(cx - sq / 2, cy - sq / 2, sq, sq, null, bg, 0.75);
    });
    // Riquadro RGB centrale
    const B = 3.543 * s, cyM = 3.754;
    sc.rect(W / 2 - B / 2, cyM - B / 2, B, B, bg, fg, 3);
    if (root.PL_VENN) sc.image(W / 2 - B * 0.46, cyM - B * 0.46, B * 0.92, B * 0.92, root.PL_VENN);
    // Triangoli laterali e in basso
    const tw = 2.876 * s, th2 = 2.479 * s, k = 0.657, off = 0.0655;
    const cxL = 0.396 + th2 / 2, gray = dark ? "595959" : "A6A6A6";
    sc.tri(cxL, 3.75, tw, th2, 90, bg, gray, 0.75).tri(cxL - off * th2, 3.75, tw * k, th2 * k, 90, "00FF00");
    sc.tri(W - cxL, 3.75, tw, th2, 270, bg, gray, 0.75).tri(W - cxL + off * th2, 3.75, tw * k, th2 * k, 270, "0000FF");
    const bw = 1.335 * s, bh = 1.151 * s, cyB = 6.54;
    sc.tri(W / 2, cyB, bw, bh, 0, bg, gray, 0.75).tri(W / 2, cyB + off * bh, bw * k, bh * k, 0, "FF0000");
    // Quadrati colore (devono essere quadrati: se sembrano rettangoli le proporzioni sono sbagliate)
    const q = 1.151 * s, gx = (0.396 + th2 + (W / 2 - B / 2)) / 2;
    sc.rect(W / 2 - q / 2, 0.383, q, q, "FF0000");
    sc.rect(gx - q / 2, cyM - q / 2, q, q, "00FF00");
    sc.rect(W - gx - q / 2, cyM - q / 2, q, q, "0000FF");
    // Watermark tra il triangolo rosso e il mirino in basso a destra
    const x0 = W / 2 + bw / 2, x1 = W - inset - D, ww = Math.min(WM_W * s, (x1 - x0) * 0.7);
    wm(sc, (x0 + x1) / 2, 6.98, ww, dark);
    return sc;
  }
  P.screenLight = (W) => screenTest(W, false);
  P.screenDark = (W) => screenTest(W, true);

  // 3. Griglia: linee dritte, celle quadrate, cerchi rotondi
  P.grid = (W, L) => {
    const sc = new Scene("000000");
    const cell = H / 12, c = "BFBFBF";
    for (let x = W / 2, i = 0; x < W; x += cell, i++) { sc.vline(x, 0, H, 1, c); if (i) sc.vline(W - x, 0, H, 1, c); }
    for (let y = 0; y <= H + 0.001; y += cell) sc.hline(0, y, W, 1, c);
    const e = 2 * PT; // bordo esterno
    sc.rect(0, 0, W, e, "FFFFFF").rect(0, H - e, W, e, "FFFFFF").rect(0, 0, e, H, "FFFFFF").rect(W - e, 0, e, H, "FFFFFF");
    sc.circle(W / 2, H / 2, H * 0.46, null, "FFFFFF", 2.5);
    const rc = cell * 1.5;
    [[cell * 2, cell * 2], [W - cell * 2, cell * 2], [cell * 2, H - cell * 2], [W - cell * 2, H - cell * 2]].forEach(([x, y]) => sc.circle(x, y, rc, null, "FFFFFF", 2));
    sc.rect(W / 2 - 0.35, H / 2 - 1.5 * PT, 0.7, 3 * PT, "FFFFFF").rect(W / 2 - 1.5 * PT, H / 2 - 0.35, 3 * PT, 0.7, "FFFFFF");
    // Etichetta su fondo nero, dentro una cella in alto
    const lw = Math.min(W - 1.2, 6.2);
    sc.rect(W / 2 - lw / 2, cell * 0.2, lw, cell * 0.6, "000000");
    sc.text(W / 2 - lw / 2, cell * 0.2, lw, cell * 0.6, L["p.grid.title"] || "", 12, "D9D9D9", { bold: true, align: "center", spacing: 2 });
    sc.rect(W / 2 - lw / 2, H - cell * 0.8, lw, cell * 0.6, "000000");
    sc.text(W / 2 - lw / 2, H - cell * 0.8, lw, cell * 0.6, L["p.grid.hint"] || "", 12, "D9D9D9", { align: "center" });
    sc.rect(W / 2 - 0.85, H - cell * 1.65, 1.7, cell * 0.72, "000000");
    wm(sc, W / 2, H - cell * 1.29, WM_W, true);
    return sc;
  };

  // 4. Fuoco e leggibilità: stelle di Siemens, gruppi di linee sottili, scala dei corpi di testo
  function star(sc, cx, cy, r, n) {
    sc.circle(cx, cy, r, "FFFFFF", "000000", 0.75);
    const step = 360 / (2 * n);
    for (let i = 0; i < n; i++) sc.pie(cx, cy, r, i * 2 * step, i * 2 * step + step, "000000");
    sc.circle(cx, cy, r * 0.05, "FFFFFF");
  }
  P.focus = (W, L) => {
    const sc = new Scene("FFFFFF");
    const R = 1.5, cy = 2.95;
    star(sc, W / 2, cy, R, 24);
    const rc = 0.6, m = 0.3 + rc;
    [[m, m], [W - m, m], [m, H - m], [W - m, H - m]].forEach(([x, y]) => star(sc, x, y, rc, 16));
    // Gruppi di linee: 0,5 – 1 – 2 punti (verticali a sinistra, orizzontali a destra)
    const gw = 1.25, gx = W / 2 - R - 0.35 - gw, gy = cy - 0.9;
    [0.5, 1, 2].forEach((pt, gi) => {
      const bx = gx + gi * (gw / 3), t = pt * PT;
      for (let i = 0; i < 6; i++) sc.vline(bx + 0.06 + i * t * 2.2, gy, 1.8, pt, "000000");
    });
    const hx = W / 2 + R + 0.35;
    [0.5, 1, 2].forEach((pt, gi) => {
      const by = gy + gi * 0.62, t = pt * PT;
      for (let i = 0; i < 6; i++) sc.hline(hx, by + 0.05 + i * t * 2.2, gw, pt, "000000");
    });
    // Scala dei corpi: il più piccolo che si legge dall'ultima fila è il minimo per le tue slide
    const sample = L["p.focus.sample"] || "Teaching is Learning 0123";
    let y = 4.75;
    const tx = Math.max(1.75, W / 2 - 3.1);
    [12, 16, 20, 24, 32].forEach((pt) => {
      const h = (pt * 1.3) / 72;
      sc.text(tx, y, 0.8, h, pt + " pt", 11, "7F7F7F", { valign: "middle" });
      sc.text(tx + 0.8, y, W - 2 * tx - 0.8 + 0.4, h, sample, pt, "000000", { valign: "middle" });
      y += h + 0.02;
    });
    labels(sc, W, L, "focus", "595959", { titleAlign: "center", noHint: true });
    wm(sc, W / 2, H - 0.32, WM_W, false);
    return sc;
  };

  // Fila di riquadri numerati
  function patches(sc, W, values, bg, y, h, labelColor, colorOf) {
    const n = values.length, x0 = 0.6, gap = 0.08;
    const pw = (W - 2 * x0 - gap * (n - 1)) / n;
    values.forEach((v, i) => {
      const x = x0 + i * (pw + gap);
      sc.rect(x, y, pw, h, colorOf ? colorOf(v) : grayHex(v));
      sc.text(x, y + h + 0.08, pw, 0.3, String(v).replace(".", ",") + "%", 13, labelColor, { align: "center" });
    });
  }
  // 5. Livello del nero (LUMINOSITÀ)
  P.black = (W, L) => {
    const sc = new Scene("000000");
    patches(sc, W, [0, 1, 2, 3, 4, 5, 6, 8, 10, 15, 20], "000000", 2.3, 2.5, "8C8C8C");
    labels(sc, W, L, "black", "8C8C8C");
    wmCorner(sc, W, true);
    return sc;
  };
  // 6. Livello del bianco (CONTRASTO)
  P.white = (W, L) => {
    const sc = new Scene("FFFFFF");
    patches(sc, W, [80, 85, 88, 90, 92, 94, 95, 96, 97, 98, 99], "FFFFFF", 2.3, 2.5, "595959");
    labels(sc, W, L, "white", "595959");
    wmCorner(sc, W, false);
    return sc;
  };
  // 7. Scala di grigi: gradini neutri (temperatura colore e gamma)
  P.gray = (W, L) => {
    const sc = new Scene("000000");
    const x0 = 0.6, w = W - 2 * x0;
    for (let i = 0; i <= 10; i++) {
      const bw = w / 11, v = i * 10;
      sc.rect(x0 + i * bw, 0.75, bw + 0.005, 3.1, grayHex(v));
      sc.text(x0 + i * bw, 3.4, bw, 0.36, v + "%", 12, v >= 50 ? "000000" : "BFBFBF", { align: "center" });
    }
    for (let i = 0; i < 32; i++) {
      const bw = w / 32, v = Math.round(((i + 1) / 32) * 255);
      sc.rect(x0 + i * bw, 4.25, bw + 0.005, 1.35, hex(v, v, v));
    }
    // Tre grigi grandi affiancati: devono avere lo stesso colore neutro
    const pw = (w - 0.2) / 3;
    [25, 50, 75].forEach((p, i) => sc.rect(x0 + i * (pw + 0.1), 5.8, pw, 0.95, grayHex(p)));
    labels(sc, W, L, "gray", "A6A6A6", { hintY: H - 0.48 });
    wmCorner(sc, W, true);
    return sc;
  };
  // 8. Barre colore SMPTE (75%)
  P.smpte = (W) => {
    const sc = new Scene("000000");
    const top = ["C0C0C0", "C0C000", "00C0C0", "00C000", "C000C0", "C00000", "0000C0"];
    const mid = ["0000C0", "131313", "C000C0", "131313", "00C0C0", "131313", "C0C0C0"];
    const bw = W / 7, h1 = H * 0.67, h2 = H * 0.08;
    top.forEach((c, i) => sc.rect(i * bw, 0, bw + 0.005, h1, c));
    mid.forEach((c, i) => sc.rect(i * bw, h1, bw + 0.005, h2, c));
    const y3 = h1 + h2, h3 = H - y3, qw = (W * 5) / 28;
    sc.rect(0, y3, qw, h3, "00214C").rect(qw, y3, qw, h3, "FFFFFF").rect(2 * qw, y3, qw, h3, "32006A").rect(3 * qw, y3, W - 3 * qw, h3, "131313");
    const px = 5 * bw, pw = bw / 3;
    sc.rect(px, y3, pw, h3, "090909").rect(px + pw, y3, pw, h3, "131313").rect(px + 2 * pw, y3, pw + 0.005, h3, "1D1D1D");
    // Watermark nel riquadro nero in basso a destra
    wm(sc, 6.5 * bw, H - 0.4, Math.min(WM_W, bw * 0.8), true);
    return sc;
  };
  // 9. Rampe colore a 32 gradini (saturazione e colori che si "impastano")
  P.ramps = (W, L) => {
    const sc = new Scene("000000");
    const x0 = Math.max(0.6, W * 0.09), x1 = W - 0.6, w = x1 - x0, bw = w / 32;
    const rows = [[0, 0, 1], [0, 1, 0], [0, 1, 1], [1, 0, 0], [1, 0, 1], [1, 1, 0], [1, 1, 1]];
    [1, 2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32].forEach((n) => sc.text(x0 + (n - 1) * bw - 0.1, 0.58, bw + 0.2, 0.26, String(n), 10, "A6A6A6", { align: "center" }));
    const rh = 0.6, g = 0.2;
    rows.forEach((c, r) => {
      const y = 0.95 + r * (rh + g);
      for (let i = 0; i < 32; i++) {
        const v = ((i + 1) / 32) * 255;
        sc.rect(x0 + i * bw, y, bw + 0.005, rh, hex(c[0] * v, c[1] * v, c[2] * v));
      }
    });
    labels(sc, W, L, "ramps", "A6A6A6", { hintY: H - 0.55 });
    wmCorner(sc, W, true);
    return sc;
  };
  // 10. Riferimenti clinici: bianchi dello smalto, tessuti molli, ombre del cavo orale
  P.dental = (W, L) => {
    const sc = new Scene("000000");
    const x0 = 0.6, w = W - 1.2, n = 13, gap = 0.06, pw = (w - gap * (n - 1)) / n;
    const rows = [
      ["p.dental.whites", [255, 255, 255], [251, 248, 240], [212, 193, 152]],
      ["p.dental.soft", null, [248, 206, 204], [112, 16, 26]],
      ["p.dental.shadows", null, [86, 34, 38], [0, 0, 0]],
    ];
    rows.forEach(([key, first, a, b], r) => {
      const y = 0.95 + r * 2.0;
      sc.text(x0, y - 0.02, w, 0.32, L[key] || "", 12, "BFBFBF", { bold: true, spacing: 1 });
      for (let i = 0; i < n; i++) {
        let c;
        if (first && i === 0) c = first;
        else { const j = first ? i - 1 : i, m = first ? n - 2 : n - 1; c = mix(a, b, j / m); }
        sc.rect(x0 + i * (pw + gap), y + 0.36, pw, 1.3, hex(c[0], c[1], c[2]));
      }
    });
    labels(sc, W, L, "dental", "A6A6A6", { titleY: 0.2, hintY: H - 0.55 });
    wmCorner(sc, W, true);
    return sc;
  };
  // 11. Spazio per una foto clinica dell'utente
  P.photo = (W, L) => {
    const sc = new Scene("000000");
    sc.rect(W * 0.08, 0.9, W * 0.84, H - 1.8, null, "404040", 1.5);
    sc.text(W * 0.1, H / 2 - 0.5, W * 0.8, 1.0, L["p.photo.placeholder"] || "", 16, "7F7F7F", { align: "center" });
    labels(sc, W, L, "photo", "8C8C8C");
    wmCorner(sc, W, true);
    return sc;
  };
  // 12–14. Campi pieni per uniformità, macchie, pixel difettosi
  const flat = (bg, lc, dark) => (W, L, key) => {
    const sc = new Scene(bg);
    sc.text(0.35, H - 0.49, 4.0, 0.28, L["p." + key + ".title"] || "", 9, lc);
    wmCorner(sc, W, dark);
    return sc;
  };
  P.flatWhite = flat("FFFFFF", "D9D9D9", false);
  P.flatGray = flat("808080", "999999", true);
  P.flatBlack = flat("000000", "262626", true);

  /* Ordine delle slide e passo della procedura a cui appartengono */
  const ORDER = [
    ["screenLight", "geometry"], ["screenDark", "geometry"], ["grid", "geometry"],
    ["focus", "focus"],
    ["black", "black"], ["white", "white"],
    ["gray", "gray"],
    ["smpte", "color"], ["ramps", "color"],
    ["dental", "clinical"], ["photo", "clinical"],
    ["flatWhite", "uniform"], ["flatGray", "uniform"], ["flatBlack", "uniform"],
  ];

  function build(key, W, L) { return P[key](W, L || {}, key); }

  /* ---------- Disegno SVG ---------- */
  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const f = (n) => Math.round(n * 10000) / 10000;
  function paint(fill, line, lw) {
    return `fill="${fill ? "#" + fill : "none"}"` + (line ? ` stroke="#${line}" stroke-width="${f((lw || 1) * PT)}"` : "");
  }
  function toSVG(sc, W, opt) {
    const o = opt || {};
    let out = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${f(W)} ${H}" preserveAspectRatio="${o.par || "xMidYMid meet"}"${o.attrs || ""}>`;
    out += `<rect x="0" y="0" width="${f(W)}" height="${H}" fill="#${sc.bg}"/>`;
    sc.items.forEach((it) => {
      if (it.k === "rect") out += `<rect x="${f(it.x)}" y="${f(it.y)}" width="${f(it.w)}" height="${f(it.h)}" ${paint(it.fill, it.line, it.lw)} shape-rendering="${it.line ? "auto" : "crispEdges"}"/>`;
      else if (it.k === "ellipse") out += `<ellipse cx="${f(it.x + it.w / 2)}" cy="${f(it.y + it.h / 2)}" rx="${f(it.w / 2)}" ry="${f(it.h / 2)}" ${paint(it.fill, it.line, it.lw)}/>`;
      else if (it.k === "pie") {
        const a0 = (it.a0 * Math.PI) / 180, a1 = (it.a1 * Math.PI) / 180;
        const x0 = it.cx + it.r * Math.cos(a0), y0 = it.cy + it.r * Math.sin(a0), x1 = it.cx + it.r * Math.cos(a1), y1 = it.cy + it.r * Math.sin(a1);
        out += `<path d="M${f(it.cx)} ${f(it.cy)}L${f(x0)} ${f(y0)}A${f(it.r)} ${f(it.r)} 0 ${it.a1 - it.a0 > 180 ? 1 : 0} 1 ${f(x1)} ${f(y1)}Z" fill="#${it.fill}"/>`;
      } else if (it.k === "tri") {
        const cx = it.x + it.w / 2, cy = it.y + it.h / 2;
        out += `<polygon points="${f(cx)},${f(it.y)} ${f(it.x + it.w)},${f(it.y + it.h)} ${f(it.x)},${f(it.y + it.h)}" transform="rotate(${it.rot || 0} ${f(cx)} ${f(cy)})" ${paint(it.fill, it.line, it.lw)} stroke-linejoin="miter"/>`;
      } else if (it.k === "image") out += `<image x="${f(it.x)}" y="${f(it.y)}" width="${f(it.w)}" height="${f(it.h)}" href="${it.data}" preserveAspectRatio="xMidYMid meet"/>`;
      else if (it.k === "text") {
        const fs = it.size * PT;
        const x = it.align === "center" ? it.x + it.w / 2 : it.align === "right" ? it.x + it.w : it.x;
        const anchor = it.align === "center" ? "middle" : it.align === "right" ? "end" : "start";
        const lines = String(it.text).split("\n"), lh = fs * 1.25;
        const y = it.valign === "top" ? it.y + fs * 0.9 : it.y + it.h / 2 + fs * 0.35 - ((lines.length - 1) * lh) / 2;
        out += `<text x="${f(x)}" y="${f(y)}" font-family="Arial, Helvetica, sans-serif" font-size="${f(fs)}" fill="#${it.color}" text-anchor="${anchor}"${it.bold ? ' font-weight="700"' : ""}${it.spacing ? ` letter-spacing="${f(it.spacing * PT)}"` : ""}>` +
          lines.map((l, i) => `<tspan x="${f(x)}" dy="${i ? f(lh) : 0}">${esc(l)}</tspan>`).join("") + `</text>`;
      }
    });
    return out + "</svg>";
  }

  /* ---------- Disegno PowerPoint (PptxGenJS) ---------- */
  function lineOpt(it) { return it.line ? { color: it.line, width: it.lw || 1 } : { type: "none" }; }
  function fillOpt(c) { return c ? { color: c } : { type: "none" }; }
  function addToSlide(pres, slide, sc, key) {
    const S = pres.ShapeType || pres.shapes;
    slide.background = { color: sc.bg };
    // Forma invisibile con il nome "USPL:<pattern>": serve a ritrovare le slide di test in qualsiasi presentazione
    if (key) slide.addShape(S.rect, { x: 0, y: 0, w: 0.01, h: 0.01, fill: { type: "none" }, line: { type: "none" }, objectName: MARK + key });
    sc.items.forEach((it) => {
      if (it.k === "rect") slide.addShape(S.rect, { x: it.x, y: it.y, w: it.w, h: it.h, fill: fillOpt(it.fill), line: lineOpt(it) });
      else if (it.k === "ellipse") slide.addShape(S.ellipse, { x: it.x, y: it.y, w: it.w, h: it.h, fill: fillOpt(it.fill), line: lineOpt(it) });
      else if (it.k === "pie") slide.addShape(S.pie, { x: it.cx - it.r, y: it.cy - it.r, w: 2 * it.r, h: 2 * it.r, angleRange: [it.a0, it.a1], fill: { color: it.fill }, line: { type: "none" } });
      else if (it.k === "tri") slide.addShape(S.triangle, { x: it.x, y: it.y, w: it.w, h: it.h, rotate: it.rot || 0, fill: fillOpt(it.fill), line: lineOpt(it) });
      else if (it.k === "image") slide.addImage({ data: it.data.replace(/^data:/, ""), x: it.x, y: it.y, w: it.w, h: it.h });
      else if (it.k === "text") slide.addText(it.text, { x: it.x, y: it.y, w: it.w, h: it.h, fontFace: "Arial", fontSize: it.size, color: it.color, bold: !!it.bold, align: it.align, valign: it.valign, margin: 0, charSpacing: it.spacing || 0, fit: "none", wrap: String(it.text).indexOf("\n") >= 0 });
    });
  }
  /* Crea la presentazione di test. keys: elenco dei pattern; notes(key) restituisce il testo delle note */
  function buildDeck(PptxGen, W, L, keys, notes) {
    const pres = new PptxGen();
    pres.defineLayout({ name: "PL", width: W, height: H });
    pres.layout = "PL";
    pres.author = "The Ultraspeaker Projector Lab";
    pres.title = "The Ultraspeaker Projector Lab – Screen Test";
    (keys || ORDER.map((o) => o[0])).forEach((key) => {
      const slide = pres.addSlide();
      addToSlide(pres, slide, build(key, W, L), key);
      const n = notes && notes(key);
      if (n) slide.addNotes(n);
    });
    return pres;
  }

  root.PL_PATTERNS = { H, MARK, FORMATS, ORDER, formatFromRatio, build, toSVG, buildDeck, grayHex };
  if (typeof module !== "undefined") module.exports = root.PL_PATTERNS;
})(typeof self !== "undefined" ? self : globalThis);
