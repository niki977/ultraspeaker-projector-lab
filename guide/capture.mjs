// The Ultraspeaker Projector Lab — screenshot della guida d'utilizzo.
// Uso, dalla radice del repository:   node guide/capture.mjs
// Serve il Lab in locale, lo apre con Playwright e salva in guide/img/<lingua>/ gli screenshot di ogni capitolo
// (scala 2×: 1440 px web, 390 px telefono, 350 px pannello simulato con la classe us-addin), più pins.json
// con la posizione dei richiami. Dati di esempio neutri: nessun dato personale, paziente o volto.
import { createServer } from "node:http";
import { createRequire } from "node:module";
import { readFile, mkdir, writeFile } from "node:fs/promises";
import { extname, join, resolve } from "node:path";

let chromium;
try { ({ chromium } = await import("playwright")); }
catch { ({ chromium } = createRequire(import.meta.url)("playwright")); }

const ROOT = resolve(process.cwd());
const LANGS = ["it", "en", "es", "fr", "de"];
const TYPES = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".svg": "image/svg+xml", ".png": "image/png", ".woff2": "font/woff2", ".json": "application/json", ".webmanifest": "application/manifest+json" };

// Server statico minimo: niente dipendenze
const server = createServer(async (req, res) => {
  try {
    const path = decodeURIComponent(new URL(req.url, "http://x").pathname);
    const file = join(ROOT, path === "/" ? "index.html" : path);
    if (!file.startsWith(ROOT)) throw new Error("fuori dalla cartella");
    res.writeHead(200, { "content-type": TYPES[extname(file)] || "application/octet-stream" });
    res.end(await readFile(file));
  } catch { res.writeHead(404); res.end(); }
});
await new Promise((r) => server.listen(0, r));
const BASE = `http://localhost:${server.address().port}/`;

// PowerPoint simulato per le immagini del pannello: presentazione 16:9 con tre slide
const MOCK = `(() => {
  let n = 300; const KEYS = ["screenLight","screenDark","grid","focus","black","white","gray","cast","smpte","ramps","dental","photo","flatWhite","flatGray","flatBlack"];
  const slide = (name) => ({ id: String(n++), shapes: { load() {}, items: name ? [{ name }] : [] }, delete() { const i = S.indexOf(this); if (i >= 0) S.splice(i, 1); },
    getImageAsBase64() { const b = new Uint8Array(24); b.set([137,80,78,71,13,10,26,10,0,0,0,13,73,72,68,82]); b.set([0,0,6,64,0,0,3,132], 16); return { value: btoa(String.fromCharCode(...b)) }; } });
  const S = [slide(), slide(), slide()];
  const slides = { load() {}, get items() { return S.slice(); } };
  window.PowerPoint = { createPresentation: async () => {}, run: async (f) => f({ sync: async () => {}, presentation: { slides,
    insertSlidesFromBase64(b, o) { S.push(...KEYS.map((k) => slide("USPL:" + k))); }, setSelectedSlides() {} } }) };
  const settings = { d: {}, get(k) { return this.d[k]; }, set(k, v) { this.d[k] = v; }, saveAsync(cb) { cb && cb({}); } };
  window.Office = { onReady(cb) { setTimeout(() => cb({ host: "PowerPoint" }), 20); }, HostType: { PowerPoint: "PowerPoint" },
    AsyncResultStatus: { Succeeded: "succeeded" }, GoToType: { Slide: "slide" }, CoercionType: { Image: "image" }, EventType: {},
    context: { displayLanguage: "en-GB", requirements: { isSetSupported: () => true }, document: { settings, setSelectedDataAsync(d, o, cb) { cb({ status: "succeeded" }); } },
      ui: { displayDialogAsync() {}, openBrowserWindow() {} } } };
})();`;

// Data fissa (2 novembre 2026, 10:00) e timer di riscaldamento a 11 minuti dalla fine
const NOW = Date.UTC(2026, 10, 2, 9, 0, 0);
const SAMPLE = {
  venue: { it: "Sala congressi A", en: "Conference room A", es: "Sala de congresos A", fr: "Salle de congrès A", de: "Kongresssaal A" },
  note: { it: "Modalità sRGB, luci di sala al 50%.", en: "sRGB mode, room lights at 50%.", es: "Modo sRGB, luces de sala al 50 %.", fr: "Mode sRGB, lumières de salle à 50 %.", de: "sRGB-Modus, Saallicht auf 50 %." },
};

const browser = await chromium.launch();
async function open(lang, { width, height = 900, office = false, storage = {} }) {
  const ctx = await browser.newContext({ viewport: { width, height }, deviceScaleFactor: 2, locale: "en-GB" });
  await ctx.addInitScript(({ now, storage }) => {
    // schermo Full HD: nella versione web il formato «automatico» è quello dello schermo
    Object.defineProperty(screen, "width", { get: () => 1920 }); Object.defineProperty(screen, "height", { get: () => 1080 });
    const D = Date; class F extends D { constructor(...a) { super(...(a.length ? a : [now])); } static now() { return now; } } window.Date = F;
    if (!sessionStorage.getItem("seeded")) { localStorage.clear(); for (const k in storage) localStorage.setItem(k, storage[k]); sessionStorage.setItem("seeded", "1"); }
  }, { now: NOW, storage });
  await ctx.route("**/appsforoffice.microsoft.com/**", (r) => r.fulfill({ contentType: "text/javascript", body: office ? MOCK : "" }));
  const page = await ctx.newPage();
  await page.goto(`${BASE}index.html?lang=${lang}`);
  await page.waitForTimeout(office ? 700 : 400);
  await page.evaluate(() => document.fonts.ready);
  return page;
}
const click = (p, s) => p.$eval(s, (e) => e.click());
const step = async (p, i) => { await click(p, `#stepper button:nth-child(${i})`); await p.waitForTimeout(80); };
const box = (p, s) => p.$eval(s, (e) => { const r = e.getBoundingClientRect(); return { x: r.x, y: r.y + scrollY, w: r.width, h: r.height }; });
// Ritaglio dalla pagina intera, dal primo all'ultimo elemento indicato, con un margine
async function shot(p, file, sels, pad = 12) {
  // nei ritagli la barra fissa Indietro / Avanti coprirebbe il contenuto: torna al suo posto in fondo alla pagina
  await p.addStyleTag({ content: ".pl-foot{position:static!important}" });
  const bs = await Promise.all(sels.map((s) => box(p, s)));
  const y0 = Math.max(0, Math.min(...bs.map((b) => b.y)) - pad), y1 = Math.max(...bs.map((b) => b.y + b.h)) + pad;
  const vw = await p.evaluate(() => innerWidth);
  await p.screenshot({ path: file, fullPage: true, clip: { x: 0, y: y0, width: vw, height: y1 - y0 } });
  return { y0, h: y1 - y0, w: vw };
}

for (const lang of LANGS) {
  const dir = join(ROOT, "guide/img", lang); await mkdir(dir, { recursive: true });
  const pins = {};
  const venues = JSON.stringify([{ venue: SAMPLE.venue[lang], date: "2026-10-14T09:00:00Z", fmt: "16:9",
    values: { brightness: { a: "50", b: "47" }, contrast: { a: "50", b: "44" }, temp: { a: "7500 K", b: "6500 K" }, mode: { a: "Bright", b: "sRGB" } }, answers: {}, note: SAMPLE.note[lang] }]);

  // 01-web: versione web, intestazione, Screen Test, passi e inizio del passo 1 (richiami in pins.json)
  {
    const p = await open(lang, { width: 1440, height: 620 });
    await p.evaluate(() => window.scrollTo(0, 0));
    // solo la colonna della pagina (760 px) con un po' di margine: ai lati c'è solo fondo nero
    const col = await box(p, ".us-page"), cx = col.x - 40, cw = col.w + 80, vh = 620;
    await p.screenshot({ path: join(dir, "01-web.png"), clip: { x: cx, y: 0, width: cw, height: vh } });
    const at = async (s, dx = 0.5, dy = 0.5) => { const b = await box(p, s); return [+(100 * (b.x - cx + b.w * dx) / cw).toFixed(1), +(100 * (b.y + b.h * dy) / vh).toFixed(1)]; };
    pins.web = { logo: await at(".us-logo", 0.95, 0.15), tools: await at(".us-header-tools", 0.5, 1.15), setup: await at("#fmtSeg", 1.08, 0.5),
      stepper: await at("#stepper", 0.99, -0.35), restart: await at("#restartBtn", -0.05, 0.5), step: await at("#step .us-title", 0.5, -0.2), foot: await at("#nextBtn", 0.06, 0.5) };
    await p.context().close();
  }
  // 01-panel: pannello di PowerPoint dopo l'inserimento delle slide di test, passo 2 con le anteprime
  {
    const p = await open(lang, { width: 350, height: 1000, office: true });
    await click(p, "#insertBtn"); await p.waitForFunction(() => !document.querySelector("#insertBtn").disabled, null, { timeout: 30000 });
    await p.waitForTimeout(400); await step(p, 2); await p.waitForTimeout(3200); // la scheda si chiude da sola; si aspetta che sparisca l'avviso
    await p.evaluate(() => window.scrollTo(0, 0));
    await p.screenshot({ path: join(dir, "01-panel.png"), clip: { x: 0, y: 0, width: 350, height: 540 } });
    await p.context().close();
  }
  // 03-black e 03-cast: domande «Cosa vedi?» con la risposta e la correzione (telefono)
  {
    const p = await open(lang, { width: 390 });
    await step(p, 4); await click(p, '[data-q="black"] [data-v="6"]');
    await shot(p, join(dir, "03-black.png"), ['[data-q="black"]']);
    await step(p, 6);
    for (const [g, v] of [["low", 3], ["mid", 3], ["high", 3]]) await click(p, `[data-g="${g}"][data-v="${v}"]`);
    await shot(p, join(dir, "03-cast.png"), ['[data-q="cast"]']);
    await p.fill("#set-temp-a", "7500 K"); await p.fill("#set-temp-b", "6500 K"); await p.fill("#set-gamma-a", "2.4"); await p.fill("#set-gamma-b", "2.2");
    await p.$eval("#set-temp-b", (e) => e.blur());
    await shot(p, join(dir, "03-values.png"), [".pl-sets"]);
    await p.context().close();
  }
  // 04-setup: scheda Screen Test della versione web; 04-pattern: Screen Test a schermo intero con la barra
  {
    const p = await open(lang, { width: 1440 });
    const r = await box(p, "#setup"), c = { x: r.x - 10, y: r.y - 10, width: r.w + 20, height: r.h + 20 };
    await p.screenshot({ path: join(dir, "04-setup.png"), fullPage: true, clip: c });
    // richiamo sul bordo sinistro del controllo, per non coprirne il testo
    const at = async (s) => { const b = await box(p, s); return [+(100 * (b.x - 10 - c.x) / c.width).toFixed(1), +(100 * (b.y + b.h / 2 - c.y) / c.height).toFixed(1)]; };
    pins.setup = { fmt: await at("#fmtSeg [data-fmt='auto']"), window: await at("#windowBtn"), download: await at("#downloadBtn") };
    await p.context().close();
    const s = await browser.newPage({ viewport: { width: 1440, height: 810 }, deviceScaleFactor: 1 });
    await s.goto(`${BASE}show.html?lang=${lang}&k=screenLight`); await s.waitForTimeout(500);
    await s.evaluate(() => document.fonts.ready);
    await s.screenshot({ path: join(dir, "04-pattern.png") });
    await s.close();
  }
  // 05-venue: passo 1 con il timer in corso e una sala salvata; 05-report: report del passo 10 (telefono)
  {
    const storage = { "ultraspeaker.projector.venues": venues, "ultraspeaker.projector.timer": String(NOW + (11 * 60 + 24) * 1000) };
    const p = await open(lang, { width: 390, storage });
    await p.fill("#venue", SAMPLE.venue[lang]);
    await p.selectOption("#venuePick", "0"); await p.waitForTimeout(100);
    await shot(p, join(dir, "05-venue.png"), ["#venue", "#venueInfo"]);
    await step(p, 2); await click(p, '[data-r="ok"]');
    await step(p, 4); await click(p, '[data-q="black"] [data-v="3"]');
    await step(p, 5); await click(p, '[data-q="white"] [data-v="98"]');
    await p.fill("#set-contrast-a", "50"); await p.fill("#set-contrast-b", "44");
    await step(p, 10); await p.fill("#note", SAMPLE.note[lang]); await p.$eval("#note", (e) => e.dispatchEvent(new Event("input")));
    await shot(p, join(dir, "05-report.png"), ["#report", "#saveVenueBtn"]);
    await p.context().close();
  }
  await writeFile(join(dir, "pins.json"), JSON.stringify(pins, null, 1));
  console.log(`ok  guide/img/${lang}/`);
}
await browser.close();
server.close();
