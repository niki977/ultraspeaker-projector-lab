/* The Ultraspeaker Projector Lab — report in PDF (A4, stile delle guide d'utilizzo).
   Riceve un modello già tradotto (vedi reportModel() in app.js) e restituisce un documento jsPDF.
   Usato dalla pagina principale (versione web) e da report.html (pannello di PowerPoint, nel browser). */
(function () {
  "use strict";
  const W = 210, H = 297, M = 20, BOTTOM = H - 24;
  const C = { black: "#000000", graphite: "#3A3A3A", ash: "#8E8E8E", silver: "#D0D0D0", house: "#F3F3F3", white: "#FFFFFF", ok: "#7DDC8A", warn: "#FFCF80" };
  const PT = 0.3528; // 1 pt in mm

  async function fileB64(url) {
    const buf = await (await fetch(url)).arrayBuffer();
    let s = ""; const b = new Uint8Array(buf);
    for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000));
    return btoa(s);
  }

  async function build(m, base) {
    base = base || "";
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ unit: "mm", format: "a4", compress: true });
    const [fr, fb, fi, logo, arena] = await Promise.all([
      fileB64(base + "fonts/pdf/champagne-limousines-regular.ttf"),
      fileB64(base + "fonts/pdf/champagne-limousines-bold.ttf"),
      fileB64(base + "fonts/pdf/champagne-limousines-italic.ttf"),
      fileB64(base + "assets/report-logo.png"),
      fileB64(base + "assets/report-arena.png"),
    ]);
    doc.addFileToVFS("cl-r.ttf", fr); doc.addFont("cl-r.ttf", "CL", "normal");
    doc.addFileToVFS("cl-b.ttf", fb); doc.addFont("cl-b.ttf", "CL", "bold");
    doc.addFileToVFS("cl-i.ttf", fi); doc.addFont("cl-i.ttf", "CL", "italic");
    const L = m.L;
    doc.setProperties({ title: `${L.footer} — ${m.venue || L.untitled}`, author: "The Ultraspeaker", creator: "The Ultraspeaker Projector Lab", subject: L.kicker });
    doc.setLineHeightFactor(1.45);

    // ---------- strumenti ----------
    const font = (style, size, color) => { doc.setFont("CL", style); doc.setFontSize(size); doc.setTextColor(color); };
    // etichetta in maiuscolo spaziata (0,32 em), come .g-label
    const label = (txt, x, y, color, size, align) => {
      size = size || 7.5; font("bold", size, color || C.ash);
      const cs = size * 0.32 * PT, s = String(txt).toUpperCase();
      doc.setCharSpace(cs);
      let xx = x;
      if (align === "right") xx = x - (doc.getTextWidth(s) + cs * (s.length - 1));
      if (align === "center") xx = x - (doc.getTextWidth(s) + cs * (s.length - 1)) / 2;
      doc.text(s, xx, y);
      doc.setCharSpace(0);
      return doc.getTextWidth(s) + cs * (s.length - 1);
    };
    const line = (x1, y1, x2, color, w) => { doc.setDrawColor(color); doc.setLineWidth(w || 0.2); doc.line(x1, y1, x2, y1); };
    let y = 0;
    function newPage() { doc.addPage(); y = 22; }
    const ensure = (h) => { if (y + h > BOTTOM) { newPage(); return true; } return false; };

    // ---------- intestazione nera ----------
    doc.setFillColor(C.black); doc.rect(0, 0, W, 74, "F");
    doc.addImage("data:image/png;base64," + logo, "PNG", M, 13, 62, 62 * 486 / 1200);
    label(L.kicker, M, 48, C.ash);
    font("normal", 24, C.white);
    const title = doc.splitTextToSize(m.venue || L.untitled, W - 2 * M)[0];
    doc.text(title, M, 58.5);
    label(m.date, M, 67, C.silver);
    label(`${L.formatL} ${m.format}`, W - M, 67, C.silver, 7.5, "right");
    y = 88;

    // ---------- tabella generica ----------
    // cols: [{w, head, align}] ; rows: array di array di stringhe ; draw(row, i, x, y) per celle speciali
    function table(head, cols, rows, opts) {
      opts = opts || {};
      const fs = 9.5, lh = fs * PT * 1.45, pad = 2.4;
      const drawHead = () => {
        let x = M;
        cols.forEach((c, i) => { if (head[i]) label(head[i], c.align === "right" ? x + c.w : x, y, C.ash, 6.8, c.align); x += c.w; });
        y += 2; line(M, y, W - M, C.black, 0.3); y += pad;
      };
      ensure(18); drawHead();
      rows.forEach((r, ri) => {
        font("normal", fs, C.black);
        const cells = r.map((txt, i) => (opts.custom && opts.custom[i]) ? [""] : doc.splitTextToSize(String(txt == null ? "" : txt), cols[i].w - 3));
        const h = Math.max(...cells.map((c) => c.length)) * lh + pad * 1.6;
        if (ensure(h + 2)) drawHead();
        let x = M;
        cols.forEach((c, i) => {
          if (opts.custom && opts.custom[i]) opts.custom[i](r, x, y + lh * 0.75);
          else {
            font(c.bold ? "bold" : "normal", fs, c.color || C.black);
            doc.text(cells[i], c.align === "right" ? x + c.w : x, y + lh * 0.75, { align: c.align || "left" });
          }
          x += c.w;
        });
        y += h; line(M, y - pad * 0.9, W - M, C.silver, 0.2);
      });
      y += 6;
    }
    const section = (txt) => { ensure(26); label(txt, M, y); y += 6; };

    // ---------- esito dei passi ----------
    const RES = { ok: [L.ok, C.ok], fix: [L.fix, C.warn], skip: [L.skip, C.silver] };
    section(L.steps);
    const count = { ok: 0, fix: 0, skip: 0 };
    m.steps.forEach((s) => { if (count[s[2]] != null) count[s[2]]++; });
    // riepilogo con pallini
    let sx = M;
    ["ok", "fix", "skip"].forEach((k) => {
      doc.setFillColor(RES[k][1]); doc.setDrawColor(C.black); doc.setLineWidth(0.2); doc.circle(sx + 1.6, y - 1.2, 1.6, "FD");
      font("bold", 11, C.black); const n = String(count[k]); doc.text(n, sx + 5, y);
      const nw = doc.getTextWidth(n);
      font("normal", 9.5, C.graphite); doc.text(RES[k][0], sx + 6.5 + nw, y);
      sx += 12 + nw + doc.getTextWidth(RES[k][0]);
    });
    y += 8;
    table(["", L.step, L.result], [{ w: 12, bold: true, color: C.ash }, { w: 108 }, { w: 50 }],
      m.steps.map((s) => [String(s[0]).padStart(2, "0"), s[1], s[2]]),
      { custom: { 2: (r, x, yy) => {
        const k = r[2], rs = RES[k];
        if (rs) { doc.setFillColor(rs[1]); doc.setDrawColor(C.black); doc.setLineWidth(0.2); doc.circle(x + 1.5, yy - 1.1, 1.5, "FD"); }
        else { doc.setDrawColor(C.ash); doc.setLineWidth(0.25); doc.circle(x + 1.5, yy - 1.1, 1.5, "S"); }
        font("normal", 9.5, rs ? C.black : C.ash); doc.text(rs ? rs[0] : L.none, x + 5, yy);
      } } });

    // ---------- risposte ----------
    if (m.answers.length) {
      section(L.answers);
      table([L.question, L.answer], [{ w: 78, color: C.graphite }, { w: 92 }], m.answers);
    }
    // ---------- valori del proiettore ----------
    if (m.values.length) {
      section(L.values);
      table([L.setting, L.before, L.after], [{ w: 78, color: C.graphite }, { w: 46, color: C.graphite }, { w: 46, bold: true }],
        m.values.map((v) => [v[0], v[1] || "–", v[2] || v[1] || "–"]));
    }
    // ---------- note ----------
    if (m.note) {
      font("normal", 10, C.black);
      const tl = doc.splitTextToSize(m.note, W - 2 * M - 10), lh = 10 * PT * 1.5;
      section(L.note);
      let i = 0;
      while (i < tl.length) {
        const fit = Math.max(1, Math.min(tl.length - i, Math.floor((BOTTOM - y - 10) / lh)));
        const part = tl.slice(i, i + fit), bh = part.length * lh + 8;
        doc.setDrawColor(C.black); doc.setLineWidth(0.3); doc.roundedRect(M, y, W - 2 * M, bh, 2.6, 2.6, "S");
        font("normal", 10, C.black); doc.text(part, M + 5, y + 4 + lh * 0.75);
        y += bh + 6; i += fit;
        if (i < tl.length) newPage();
      }
    }

    // ---------- chiusura nera in fondo all'ultima pagina (come il retro delle guide, in piccolo) ----------
    const BAND = 46;
    if (y > H - BAND - 6) newPage();
    const by = H - BAND;
    doc.setFillColor(C.black); doc.rect(0, by, W, BAND, "F");
    doc.addImage("data:image/png;base64," + arena, "PNG", W / 2 - 7, by + 7, 14, 14 * 279 / 400);
    font("normal", 11, C.white);
    const c1 = "Knowledge is not enough. ", w1 = doc.getTextWidth(c1);
    font("italic", 11, C.white);
    const c2 = "From expert to Ultraspeaker.", w2 = doc.getTextWidth(c2);
    font("normal", 11, C.white); doc.text(c1, W / 2 - (w1 + w2) / 2, by + 24);
    font("italic", 11, C.white); doc.text(c2, W / 2 - (w1 + w2) / 2 + w1, by + 24);
    const cw = label(L.cta, W / 2, by + 32, C.white, 7, "center");
    line(W / 2 - cw / 2, by + 33.5, W / 2 + cw / 2, C.white, 0.25);
    doc.link(W / 2 - cw / 2, by + 28, cw, 7, { url: m.mail });

    // ---------- piede su ogni pagina ----------
    const n = doc.getNumberOfPages();
    for (let p = 1; p <= n; p++) {
      doc.setPage(p);
      label(L.footer, M, H - 9, C.ash, 6.5);
      label(String(p).padStart(2, "0"), W - M, H - 9, C.ash, 6.5, "right");
    }
    return doc;
  }

  // nome del file solo con caratteri semplici: alcuni browser scartano il nome se contiene accenti o trattini lunghi
  const plain = (x) => x.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/ß/g, "ss").replace(/[\u2010-\u2015]/g, "-")
    .replace(/[^\x20-\x7e]+/g, "").replace(/[\\/:*?"<>|]+/g, " ").replace(/\s+/g, " ").trim();
  const fileName = (m) => `The Ultraspeaker Projector Lab - ${plain(m.L.file)} - ${plain(m.venue || m.L.untitled) || "Report"} - ${m.iso}.pdf`;
  // salvataggio con un link vero nella pagina (più affidabile di doc.save su Safari)
  function save(doc, name) {
    const url = URL.createObjectURL(doc.output("blob"));
    const a = document.createElement("a"); a.href = url; a.download = name; a.rel = "noopener";
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }
  window.PLReport = { build, fileName, save };
})();
