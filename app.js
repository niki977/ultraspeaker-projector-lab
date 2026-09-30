/* The Ultraspeaker Projector Lab – add-in PowerPoint che guida la regolazione del proiettore */
(function () {
  "use strict";

  /* ---------- Utilità ---------- */
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const store = {
    get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* archiviazione non disponibile */ } },
  };
  let toastT;
  function toast(msg) {
    const el = $("#toast"); el.textContent = msg; el.classList.add("show");
    clearTimeout(toastT); toastT = setTimeout(() => el.classList.remove("show"), 2800);
  }
  const wait = (ms) => new Promise((r) => setTimeout(r, ms));
  const PLP = window.PL_PATTERNS;
  const ICON_PLAY = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M8 5l11 7-11 7z"/></svg>';

  /* ---------- Lingua ---------- */
  const I18N = window.PL_I18N;
  const LANGS = I18N.langs.map((l) => l[0]).filter((l) => I18N[l]);
  let LANG = "it";
  const pack = () => I18N[LANG] || I18N.it;
  const t = (k, vars) => {
    let s = (pack()[k] != null ? pack()[k] : I18N.it[k] != null ? I18N.it[k] : k);
    if (vars) Object.keys(vars).forEach((v) => { s = s.replace("{" + v + "}", vars[v]); });
    return s;
  };
  const STEPS = () => pack().steps || I18N.it.steps;
  function detectLang() {
    const saved = store.get("projectorlab.lang", null);
    if (saved && LANGS.includes(saved)) return saved;
    let l = "";
    try { if (inOffice && Office.context && Office.context.displayLanguage) l = Office.context.displayLanguage; } catch (e) { /* ignora */ }
    if (!l) l = (navigator.languages && navigator.languages[0]) || navigator.language || "it";
    l = l.slice(0, 2).toLowerCase();
    return LANGS.includes(l) ? l : "en";
  }
  // Etichette per i pattern nella lingua corrente
  function labels() {
    const out = {};
    Object.keys(I18N.it).forEach((k) => { if (k.startsWith("p.")) out[k] = t(k); });
    return out;
  }

  /* ---------- Stato ---------- */
  let inOffice = false, mode = "boot"; // web | office
  let api = { tags: false, select: false, image: false, insert: false, dialog2: false, create: false };
  const saved = store.get("projectorlab.state", {});
  const S = {
    step: saved.step || 0,
    results: saved.results || {},
    checks: saved.checks || {},
    venue: saved.venue || "",
    answers: saved.answers || {},   // risposte alle domande "Cosa vedi?"
    values: saved.values || {},     // valori del proiettore: { chiave: { a: prima, b: dopo } }
    note: saved.note || "",
    fmt: "auto",
    detected: null,
    where: "end",
    map: {},          // pattern -> id della slide inserita
    current: null,    // pattern mostrato ora
  };
  const persist = () => store.set("projectorlab.state", { step: S.step, results: S.results, checks: S.checks, venue: S.venue, note: S.note, answers: S.answers, values: S.values });
  let busy = false;
  let dialog = null, showWin = null;

  const fmtKey = () => (S.fmt === "auto" ? S.detected || "16:9" : S.fmt);
  // Formato scelto diverso da quello della presentazione: lo Screen Test va in una nuova presentazione
  const needsNewPres = () => mode === "office" && api.create && !!S.presFmt && fmtKey() !== S.presFmt;
  const fmtW = () => PLP.FORMATS[fmtKey()];
  const hasTests = () => Object.keys(S.map).length > 0;
  const stepOfPattern = (key) => STEPS().findIndex((s) => (s.slides || []).includes(key));

  /* ---------- Stato in alto ---------- */
  function setStatus(kind) {
    const el = $("#status");
    el.className = "status" + (kind === "busy" ? " busy" : mode === "office" ? "" : " off");
    $("#statusText").textContent = kind === "busy" ? t("status.busy") : mode === "office" ? t("status.office") : t("status.web");
  }
  function setBusy(b) {
    busy = b; setStatus(b ? "busy" : "");
    ["#insertBtn", "#removeBtn", "#downloadBtn"].forEach((s) => { const e = $(s); if (e) e.disabled = b; });
  }

  /* ---------- Scheda Screen Test ---------- */
  function renderSetup() {
    $("#setupText").textContent = mode === "office" ? t("setup.text") : t("setup.textWeb");
    $$(".office-only").forEach((e) => { e.hidden = mode !== "office"; });
    $$(".web-only").forEach((e) => { e.hidden = mode === "office"; });
    $("#windowLabel").textContent = mode === "office" ? t("window") : t("fullscreen");
    $("#windowHint").hidden = mode !== "office";
    // Formati
    const seg = $("#fmtSeg"); seg.innerHTML = "";
    const opts = ["auto"].concat(Object.keys(PLP.FORMATS));
    opts.forEach((o) => {
      const b = document.createElement("button");
      b.type = "button"; b.className = "pill"; b.dataset.fmt = o;
      b.setAttribute("aria-pressed", String(S.fmt === o));
      b.innerHTML = o === "auto" ? `${esc(S.detected || "16:9")}<small>${esc(t("format.auto"))}</small>` : esc(o);
      b.onclick = () => { S.fmt = o; renderSetup(); renderStep(); sendToWindow({ f: o === "auto" ? "" : o }); };
      seg.appendChild(b);
    });
    $$("#whereSeg .pill").forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.where === S.where)));
    $("#fmtHint").textContent = t("fmtHint");
    const n = Object.keys(S.map).length;
    const np = needsNewPres();
    $("#insertLabel").textContent = np ? t("newPres", { f: fmtKey() }) : n ? t("update") : t("insert");
    $("#newPresHint").hidden = !np;
    $("#newPresHint").textContent = np ? t("newPresHint", { p: S.presFmt, f: fmtKey() }) : "";
    $("#whereRow").hidden = mode !== "office" || np;
    $("#removeBtn").hidden = !n || mode !== "office";
    $("#setupDone").hidden = !n;
    $("#setupDoneText").textContent = t("insertedN", { n });
  }
  function toggleSetup(open) {
    const card = $("#setup");
    const o = open != null ? open : card.classList.contains("closed");
    card.classList.toggle("closed", !o);
    $("#setupToggle").setAttribute("aria-expanded", String(o));
  }

  /* ---------- Note del relatore per ogni slide di test ---------- */
  function notesFor(key) {
    const i = stepOfPattern(key);
    if (i < 0) return "";
    const s = STEPS()[i];
    let n = `${t("step")} ${i + 1} ${t("of")} ${STEPS().length} · ${s.title}\n${s.goal}\n\n${t("look")}\n`;
    n += (s.look || []).map((x) => "• " + x).join("\n");
    n += `\n\n${t("adjust")}\n` + (s.adjust || []).map((x) => "• " + x).join("\n");
    return n;
  }

  /* ---------- PowerPoint ---------- */
  function supported(set, v) { try { return Office.context.requirements.isSetSupported(set, v); } catch (e) { return false; } }

  async function detectFormat() {
    if (!api.image) return;
    try {
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides; slides.load("items/id"); await ctx.sync();
        if (!slides.items.length) return;
        const img = slides.items[0].getImageAsBase64({ width: 320 });
        await ctx.sync();
        const bin = atob(img.value.slice(0, 64));
        const w = (bin.charCodeAt(16) << 24) | (bin.charCodeAt(17) << 16) | (bin.charCodeAt(18) << 8) | bin.charCodeAt(19);
        const h = (bin.charCodeAt(20) << 24) | (bin.charCodeAt(21) << 16) | (bin.charCodeAt(22) << 8) | bin.charCodeAt(23);
        if (w > 0 && h > 0) S.detected = S.presFmt = PLP.formatFromRatio(w / h);
      });
    } catch (e) { /* formato non rilevabile: resta 16:9 o la scelta manuale */ }
  }

  // Ritrova le slide di test: dai tag (se disponibili) o dalle impostazioni salvate nel documento
  async function loadMap() {
    if (mode !== "office") return;
    const map = {};
    try {
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides; slides.load("items/id"); await ctx.sync();
        const ids = new Set(slides.items.map((s) => s.id));
        if (api.tags) {
          const tags = slides.items.map((s) => { const tg = s.tags.getItemOrNullObject("USPL"); tg.load("value"); return tg; });
          await ctx.sync();
          tags.forEach((tg, i) => { if (!tg.isNullObject && tg.value) map[tg.value.toLowerCase()] = slides.items[i].id; });
          // i tag vengono salvati in maiuscolo: riporta le chiavi ai nomi dei pattern
          Object.keys(map).forEach((k) => {
            const real = PLP.ORDER.map((o) => o[0]).find((o) => o.toLowerCase() === k);
            if (real && real !== k) { map[real] = map[k]; delete map[k]; }
          });
        }
        if (!Object.keys(map).length) {
          const saved = (Office.context.document.settings.get("uspl") || {});
          Object.keys(saved).forEach((k) => { if (ids.has(saved[k])) map[k] = saved[k]; });
        }
        // Slide create in una nuova presentazione o da un file scaricato: la prima forma si chiama "USPL:<pattern>"
        if (!Object.keys(map).length && api.tags) {
          const firsts = slides.items.map((s) => { const c = s.shapes; c.load({ select: "name", top: 1 }); return c; });
          await ctx.sync();
          firsts.forEach((c, i) => {
            const nm = c.items[0] && c.items[0].name;
            if (nm && nm.indexOf(PLP.MARK) === 0) map[nm.slice(PLP.MARK.length)] = slides.items[i].id;
          });
        }
      });
    } catch (e) { /* nessuna slide di test */ }
    S.map = map;
  }
  function saveMapSetting() {
    try {
      Office.context.document.settings.set("uspl", S.map);
      Office.context.document.settings.saveAsync(() => {});
    } catch (e) { /* impostazioni non disponibili */ }
  }

  async function deleteTests() {
    if (!hasTests()) return 0;
    let n = 0;
    await PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides; slides.load("items/id"); await ctx.sync();
      const want = new Set(Object.values(S.map));
      slides.items.forEach((s) => { if (want.has(s.id)) { s.delete(); n++; } });
      await ctx.sync();
    });
    S.map = {}; saveMapSetting();
    return n;
  }

  async function insertTests() {
    if (busy) return;
    if (typeof PptxGenJS === "undefined") { toast(t("err", { m: "PptxGenJS" })); return; }
    setBusy(true);
    try {
      if (S.fmt === "auto") await detectFormat();
      const W = fmtW();
      const pres = PLP.buildDeck(PptxGenJS, W, labels(), null, notesFor);
      const b64 = await pres.write({ outputType: "base64" });
      if (needsNewPres()) {
        await PowerPoint.createPresentation(b64);
        toast(t("opened", { f: fmtKey() }));
        return;
      }
      await deleteTests();
      const keys = PLP.ORDER.map((o) => o[0]);
      const map = {};
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides; slides.load("items/id"); await ctx.sync();
        const before = new Set(slides.items.map((s) => s.id));
        const opts = { formatting: "KeepSourceFormatting" };
        if (S.where === "end" && slides.items.length) opts.targetSlideId = slides.items[slides.items.length - 1].id;
        ctx.presentation.insertSlidesFromBase64(b64, opts);
        await ctx.sync();
        const after = ctx.presentation.slides; after.load("items/id"); await ctx.sync();
        const fresh = after.items.filter((s) => !before.has(s.id));
        fresh.forEach((s, i) => {
          const key = keys[i]; if (!key) return;
          map[key] = s.id;
          if (api.tags) s.tags.add("USPL", key);
        });
        await ctx.sync();
      });
      S.map = map; saveMapSetting();
      renderSetup(); renderStep();
      toast(t("inserted"));
      toggleSetup(false);
      const first = (STEPS()[S.step].slides || [])[0];
      if (first) goPattern(first);
    } catch (e) {
      toast(t("err", { m: (e && e.message) || e }));
    } finally { setBusy(false); }
  }

  async function removeTests() {
    if (busy) return;
    setBusy(true);
    try {
      const n = await deleteTests();
      renderSetup(); renderStep();
      toast(n ? t("removed") : t("noTest"));
    } catch (e) { toast(t("err", { m: (e && e.message) || e })); } finally { setBusy(false); }
  }

  async function selectSlide(id) {
    try {
      if (api.select) {
        await PowerPoint.run(async (ctx) => { ctx.presentation.setSelectedSlides([id]); await ctx.sync(); });
        return true;
      }
    } catch (e) { /* prova il metodo classico */ }
    return new Promise((res) => {
      try {
        Office.context.document.goToByIdAsync(parseInt(id, 10), Office.GoToType.Slide, (r) => res(r.status === Office.AsyncResultStatus.Succeeded));
      } catch (e) { res(false); }
    });
  }

  /* ---------- Finestra dei pattern ---------- */
  function showUrl(key) {
    const u = new URL("show.html", location.href);
    u.searchParams.set("lang", LANG);
    if (key) u.searchParams.set("k", key);
    if (S.fmt !== "auto") u.searchParams.set("f", S.fmt);
    return u.toString();
  }
  /* ---------- Vista pattern integrata (browser senza popup, anteprima) ---------- */
  const inl = { open: false, key: "screenLight", photo: null, fixed: "" };
  const KEYS = PLP.ORDER.map((o) => o[0]);
  function inlineRender() {
    const box = $("#viewer"); if (!box) return;
    const r = Math.min(3, Math.max(1, innerWidth / Math.max(1, innerHeight)));
    const W = inl.fixed ? PLP.FORMATS[inl.fixed] : r * PLP.H;
    $("#vStage").innerHTML = PLP.toSVG(PLP.build(inl.key, W, labels()), W, { par: inl.fixed ? "xMidYMid meet" : "none" });
    const ph = $("#vPhoto"); ph.hidden = !(inl.key === "photo" && inl.photo);
    if (inl.photo) ph.querySelector("img").src = inl.photo;
    $("#vName").textContent = t("p." + inl.key + ".name");
    $("#vCount").textContent = (KEYS.indexOf(inl.key) + 1) + "/" + KEYS.length;
    $("#vHelp").textContent = t("win.hint");
    $("#vFsLabel").textContent = t("win.fs");
  }
  let vIdle;
  function vWake() { const v = $("#viewer"); v.classList.remove("idle"); clearTimeout(vIdle); vIdle = setTimeout(() => v.classList.add("idle"), 2500); }
  function vGo(d) { inl.key = KEYS[(KEYS.indexOf(inl.key) + d + KEYS.length) % KEYS.length]; inlineRender(); vWake(); followPattern(inl.key); }
  function vFs() {
    const d = document;
    if (d.fullscreenElement || d.webkitFullscreenElement) { (d.exitFullscreen || d.webkitExitFullscreen).call(d); return; }
    const e = $("#viewer"), f = e.requestFullscreen || e.webkitRequestFullscreen;
    if (f) { try { const p = f.call(e); if (p && p.catch) p.catch(() => {}); } catch (x) { /* tutto schermo non disponibile */ } }
  }
  function closeInline() {
    const d = document;
    if (d.fullscreenElement) { try { d.exitFullscreen(); } catch (x) { /* niente */ } }
    inl.open = false; $("#viewer").hidden = true; document.body.style.overflow = "";
  }
  function openInline(key) {
    inl.key = key || inl.key; inl.fixed = S.fmt === "auto" ? "" : S.fmt; inl.open = true;
    $("#viewer").hidden = false; document.body.style.overflow = "hidden";
    inlineRender(); vWake(); vFs();
  }
  function wireInline() {
    $("#vPrev").onclick = () => vGo(-1);
    $("#vNext").onclick = () => vGo(1);
    $("#vFs").onclick = vFs;
    $("#vClose").onclick = closeInline;
    $("#vStage").ondblclick = vFs;
    $("#viewer").onmousemove = vWake;
    addEventListener("resize", () => { if (inl.open) inlineRender(); });
    addEventListener("keydown", (e) => {
      if (!inl.open) return;
      if (["ArrowRight", "PageDown", " ", "Enter"].includes(e.key)) { e.preventDefault(); vGo(1); }
      else if (["ArrowLeft", "PageUp", "Backspace"].includes(e.key)) { e.preventDefault(); vGo(-1); }
      else if (e.key === "f" || e.key === "F") vFs();
      else if (e.key === "Escape" && !document.fullscreenElement) closeInline();
    });
  }

  function sendToWindow(msg) {
    if (inl.open) {
      if (msg.f != null) inl.fixed = PLP.FORMATS[msg.f] ? msg.f : "";
      if (msg.photo) inl.photo = msg.photo;
      if (msg.k) inl.key = msg.k;
      inlineRender(); return true;
    }
    if (dialog) {
      if (api.dialog2) { try { dialog.messageChild(JSON.stringify(msg)); return true; } catch (e) { /* riapri */ } }
      try { dialog.close(); } catch (e) { /* già chiusa */ }
      dialog = null;
      openWindow(msg.k);
      return true;
    }
    if (showWin && !showWin.closed) { showWin.postMessage(Object.assign({ uspl: 1 }, msg), location.origin === "null" ? "*" : location.origin); return true; }
    return false;
  }
  function openWindow(key) {
    key = key || (STEPS()[S.step].slides || [])[0] || "screenLight";
    if (mode === "office") {
      if (dialog) { sendToWindow({ k: key }); return; }
      try {
        Office.context.ui.displayDialogAsync(showUrl(key), { height: 60, width: 60, displayInIframe: false, promptBeforeOpen: false }, (res) => {
          if (res.status !== Office.AsyncResultStatus.Succeeded) { toast(t("err", { m: res.error && res.error.message })); return; }
          dialog = res.value;
          dialog.addEventHandler(Office.EventType.DialogEventReceived, () => { dialog = null; });
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (a) => {
            try { const m = JSON.parse(a.message); if (m.k) followPattern(m.k); } catch (e) { /* messaggio non valido */ }
          });
        });
      } catch (e) { toast(t("err", { m: e.message })); }
    } else {
      if (showWin && !showWin.closed) { sendToWindow({ k: key }); showWin.focus(); return; }
      if (window.PL_INLINE_VIEW) { openInline(key); return; }
      try { showWin = window.open(showUrl(key), "uspl-projector-lab", "popup,width=1280,height=720"); } catch (e) { showWin = null; }
      if (!showWin) openInline(key); // popup bloccato: vista integrata nella pagina
    }
  }
  window.addEventListener("message", (e) => {
    const m = e.data;
    if (m && m.uspl && m.from === "show" && m.k) followPattern(m.k);
  });
  // La finestra dei pattern ha cambiato pattern: il pannello si porta al passo giusto
  function followPattern(key) {
    S.current = key;
    const i = stepOfPattern(key);
    if (i >= 0 && i !== S.step) { S.step = i; persist(); renderStep(); } else markThumbs();
  }

  // Mostra un pattern: slide nella presentazione, finestra dei pattern o anteprima
  async function goPattern(key) {
    S.current = key; markThumbs();
    const winOpen = !!dialog || inl.open || (showWin && !showWin.closed);
    if (winOpen) sendToWindow({ k: key });
    if (mode === "office" && S.map[key]) { await selectSlide(S.map[key]); return; }
    if (!winOpen && mode !== "office") openWindow(key);
    else if (!winOpen && mode === "office" && !hasTests()) toast(t("notInserted"));
  }

  /* ---------- Foto clinica ---------- */
  function pickPhoto() { $("#photoIn").value = ""; $("#photoIn").click(); }
  function readFile(file) {
    return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(file); });
  }
  function imgSize(url) {
    return new Promise((res) => { const i = new Image(); i.onload = () => res([i.naturalWidth, i.naturalHeight]); i.onerror = () => res([4, 3]); i.src = url; });
  }
  async function onPhoto(file) {
    if (!file) return;
    const url = await readFile(file);
    if (mode !== "office") {
      if (!inl.open && !(showWin && !showWin.closed)) openWindow("photo");
      await wait(showWin && !inl.open ? 700 : 0);
      sendToWindow({ k: "photo", photo: url });
      return;
    }
    if (!S.map.photo) { toast(t("notInserted")); return; }
    await selectSlide(S.map.photo);
    await wait(350);
    const [iw, ih] = await imgSize(url);
    const W = fmtW() * 72, H = 7.5 * 72;
    const bx = W * 0.08, by = 0.9 * 72, bw = W * 0.84, bh = H - 1.8 * 72;
    const k = Math.min(bw / iw, bh / ih), w = iw * k, h = ih * k;
    Office.context.document.setSelectedDataAsync(url.split(",")[1], {
      coercionType: Office.CoercionType.Image,
      imageLeft: bx + (bw - w) / 2, imageTop: by + (bh - h) / 2, imageWidth: w, imageHeight: h,
    }, (r) => toast(r.status === Office.AsyncResultStatus.Succeeded ? t("photoDone") : t("err", { m: r.error && r.error.message })));
  }

  /* ---------- Download (versione browser) ---------- */
  async function download() {
    if (typeof PptxGenJS === "undefined") return;
    setBusy(true);
    try {
      const pres = PLP.buildDeck(PptxGenJS, fmtW(), labels(), null, notesFor);
      await pres.writeFile({ fileName: `The Ultraspeaker Projector Lab - Screen Test ${fmtKey().replace(":", "x")}.pptx` });
    } catch (e) { toast(t("err", { m: e.message })); } finally { setBusy(false); }
  }

  /* ---------- Timer di riscaldamento ---------- */
  let timerEnd = store.get("projectorlab.timer", 0), timerT = null;
  function timerLabel() {
    const left = timerEnd - Date.now();
    if (!timerEnd) return t("timer.start");
    if (left <= 0) return t("timer.done");
    const m = Math.floor(left / 60000), s = Math.floor((left % 60000) / 1000);
    return t("timer.run", { t: m + ":" + String(s).padStart(2, "0") });
  }
  function tickTimer() {
    const b = $("#timerBtn"); if (b) b.querySelector("span").textContent = timerLabel();
    if (timerEnd && Date.now() >= timerEnd) { clearInterval(timerT); timerT = null; }
  }
  function startTimer() {
    timerEnd = Date.now() + 10 * 60000; store.set("projectorlab.timer", timerEnd);
    clearInterval(timerT); timerT = setInterval(tickTimer, 1000); tickTimer();
  }

  /* ---------- Report ---------- */
  function reportText() {
    const d = new Date();
    const lines = [t("report.title"), `${t("report.date")}: ${d.toLocaleDateString(LANG)} ${d.toLocaleTimeString(LANG, { hour: "2-digit", minute: "2-digit" })}`];
    if (S.venue) lines.push(`${t("venue")}: ${S.venue}`);
    lines.push(`${t("report.format")}: ${fmtKey()}`, "");
    STEPS().forEach((s, i) => {
      if (s.id === "prep" || s.id === "final") return;
      const r = S.results[s.id];
      const mark = r === "ok" ? "✓" : r === "fix" ? "!" : r === "skip" ? "–" : "·";
      lines.push(`${mark} ${i + 1}. ${s.title}${r ? " — " + t("res." + r) : ""}`);
    });
    const ans = [];
    ["black", "white", "cast", "ramps", "yellow", "gum", "enamel"].forEach((q) => {
      const v = S.answers[q]; if (v == null) return;
      let txt;
      if (q === "cast") txt = CAST_GROUPS.map((g) => v[g] || "–").join(" / ");
      else if (q === "ramps") txt = v.map((k) => (k === "none" ? t("dx.ramps.none") : t("c." + k))).join(", ");
      else txt = optLabel(q, v);
      if (txt) ans.push(`${t("q." + q)}: ${txt}`);
    });
    if (ans.length) lines.push("", t("rep.answers") + ":", ...ans.map((x) => "  " + x));
    const vals = valueLines(S.values);
    if (vals.length) lines.push("", t("rep.values") + ":", ...vals.map((x) => "  " + x));
    if (S.note) lines.push("", `${t("report.note")}: ${S.note}`);
    return lines.join("\n");
  }
  async function copyReport() {
    const txt = reportText();
    try { await navigator.clipboard.writeText(txt); toast(t("report.copied")); return; } catch (e) { /* uso il metodo classico */ }
    const ta = document.createElement("textarea"); ta.value = txt; document.body.appendChild(ta); ta.select();
    try { document.execCommand("copy"); toast(t("report.copied")); } catch (e) { /* niente */ }
    ta.remove();
  }

  /* ---------- Domande guidate "Cosa vedi?" e diagnosi ---------- */
  const BLACK_OPTS = ["1", "2", "3", "4", "5", "6", "8", "10", "15", "20", "g0", "none"];
  const WHITE_OPTS = ["80", "85", "88", "90", "92", "94", "95", "96", "97", "98", "99", "none"];
  const RAMP_KEYS = ["blue", "green", "cyan", "red", "magenta", "yellow", "gray"];
  const CAST_GROUPS = ["low", "mid", "high"];
  // Domande per passo e valori del proiettore da annotare
  const STEP_Q = { black: ["black"], white: ["white"], gray: ["cast"], color: ["ramps", "yellow"], clinical: ["enamel", "gum"] };
  const STEP_SET = {
    geometry: ["keystone"], focus: ["focus"], black: ["brightness", "range"], white: ["contrast", "mode"],
    gray: ["temp", "gamma", "gain", "offset"], color: ["saturation", "tint"],
  };
  const SET_ORDER = ["keystone", "focus", "brightness", "range", "contrast", "mode", "temp", "gamma", "gain", "offset", "saturation", "tint"];
  const PICK = {
    yellow: ["ok", "green", "orange"],
    gum: ["ok", "orange", "violet", "sat", "dull"],
    enamel: ["all", "most", "some", "few"],
  };
  const optLabel = (q, v) => {
    if (q === "black") return v === "g0" ? t("dx.black.g0") : v === "none" ? t("dx.none") : v + "%";
    if (q === "white") return v === "none" ? t("dx.none") : v + "%";
    return t("dx." + q + "." + v);
  };
  const cName = (k) => t("c." + k);

  function rangeCrush() {
    const b = S.answers.black, w = S.answers.white;
    return ["6", "8", "10", "15", "20", "none"].includes(b) && ["94", "92", "90", "88", "85", "80", "none"].includes(w);
  }
  function castOf(i) { const o = PLP.castOffset(+i); return { y: -o.y, m: -o.m }; }
  function castName(c) {
    const parts = [];
    if (c.m > 0) parts.push(t("dir.magenta")); if (c.m < 0) parts.push(t("dir.green"));
    if (c.y > 0) parts.push(t("dir.yellow")); if (c.y < 0) parts.push(t("dir.blue"));
    return parts.join(t("dir.and"));
  }
  function castActs(c) {
    const a = [];
    if (c.y < 0) a.push(t("act.blue")); if (c.y > 0) a.push(t("act.yellow"));
    if (c.m < 0) a.push(t("act.green")); if (c.m > 0) a.push(t("act.magenta"));
    return a.join("; ");
  }
  function castTempTint(c) {
    const a = [];
    if (c.y < 0) a.push(t("fix.temp.warm")); if (c.y > 0) a.push(t("fix.temp.cool"));
    if (c.m < 0) a.push(t("fix.tint.magenta")); if (c.m > 0) a.push(t("fix.tint.green"));
    return a;
  }
  // Ogni diagnosi restituisce { sev: "ok" | "fix", lines: [testo, ...] }
  function diagnose(q) {
    const v = S.answers[q];
    if (v == null || (Array.isArray(v) && !v.length)) return null;
    if (q === "black") {
      const k = v === "g0" ? "g0" : v === "1" ? "1" : ["2", "3"].includes(v) ? "ok" : ["4", "5"].includes(v) ? "low1" : ["6", "8"].includes(v) ? "low2" : "low3";
      const r = { sev: k === "ok" || k === "1" ? "ok" : "fix", lines: [t("dx.black.r." + k)] };
      if (rangeCrush()) r.range = true;
      return r;
    }
    if (q === "white") {
      const k = ["98", "99"].includes(v) ? "ok" : v === "97" ? "97" : ["95", "96"].includes(v) ? "mid" : ["92", "94"].includes(v) ? "high" : "bad";
      const r = { sev: k === "ok" || k === "97" ? "ok" : "fix", lines: [t("dx.white.r." + k)], recheck: k !== "ok" };
      if (rangeCrush()) r.range = true;
      return r;
    }
    if (q === "cast") {
      const got = CAST_GROUPS.filter((g) => v[g]);
      if (!got.length) return null;
      const casts = {}; got.forEach((g) => { casts[g] = castOf(v[g]); });
      const nz = got.filter((g) => casts[g].y || casts[g].m);
      const lines = [];
      if (got.length < 3) lines.push(t("dx.cast.partial"));
      if (!nz.length) { lines.push(t("dx.cast.neutral")); return { sev: got.length < 3 ? "info" : "ok", lines }; }
      const key = (c) => c.y + "," + c.m;
      const same = got.length === 3 && nz.length === 3 && CAST_GROUPS.every((g) => key(casts[g]) === key(casts.low));
      if (same) {
        lines.push(t("dx.cast.global", { cast: castName(casts.low) }));
        castTempTint(casts.low).forEach((x) => lines.push(x));
        lines.push(t("fix.rgb.both", { act: castActs(casts.low) }));
      } else {
        nz.forEach((g) => lines.push(t("dx.cast.group", { group: t("dx.cast." + g), cast: castName(casts[g]) })));
        if (casts.high && (casts.high.y || casts.high.m)) lines.push(t("fix.rgb.gain", { act: castActs(casts.high) }));
        if (casts.low && (casts.low.y || casts.low.m)) lines.push(t("fix.rgb.offset", { act: castActs(casts.low) }));
        if (nz.length === 1 && nz[0] === "mid") castTempTint(casts.mid).forEach((x) => lines.push(x));
        if (casts.low && casts.high && (casts.low.y || casts.low.m) && (casts.high.y || casts.high.m) && key(casts.low) !== key(casts.high)) lines.push(t("dx.cast.mixed"));
      }
      return { sev: "fix", lines };
    }
    if (q === "ramps") {
      if (v.includes("none")) return { sev: "ok", lines: [t("dx.ramps.ok")] };
      if (v.includes("gray")) return { sev: "fix", lines: [t("dx.ramps.gray")], back: 5 };
      if (v.length >= 3) return { sev: "fix", lines: [t("dx.ramps.many")] };
      return { sev: "fix", lines: [t("dx.ramps.one", { c: v.map(cName).join(", ") })] };
    }
    if (q === "yellow") return { sev: v === "ok" ? "ok" : "fix", lines: [t("dx.yellow.r." + v)] };
    if (q === "gum") return { sev: v === "ok" ? "ok" : "fix", lines: [t("dx.gum.r." + v)] };
    if (q === "enamel") {
      const k = v === "all" ? "ok" : v;
      return { sev: k === "ok" ? "ok" : "fix", lines: [t("dx.enamel.r." + k)], back: ["some", "few"].includes(v) ? 5 : 0 };
    }
    return null;
  }
  // L'esito del passo segue le risposte: tutto a posto solo quando tutte le risposte lo sono
  function autoResult(stepId) {
    const qs = STEP_Q[stepId]; if (!qs) return;
    const ds = qs.map(diagnose);
    if (ds.some((d) => d && d.sev === "fix")) S.results[stepId] = "fix";
    else if (ds.every((d) => d && d.sev === "ok")) S.results[stepId] = "ok";
  }
  function answer(q, v) { S.answers[q] = v; autoResult(STEPS()[S.step].id); persist(); renderStep(); }

  function resultBox(d) {
    if (!d) return "";
    let h = `<div class="dxr ${d.sev}" role="status">` + d.lines.map((x) => `<p>${esc(x)}</p>`).join("") + `</div>`;
    if (d.range) h += `<div class="dxr range"><p><b>${esc(t("dx.range.h"))}</b></p><p>${esc(t("dx.range.crush"))}</p></div>`;
    if (d.recheck) h += `<p class="note">${esc(t("dx.recheckNote"))}</p><button type="button" class="btn sec small dx-go" data-go="3">${esc(t("dx.recheckBlack"))}</button>`;
    if (d.back) h += `<button type="button" class="btn sec small dx-go" data-go="${d.back - 1}">${esc(t("dx.backTo", { n: d.back }))}</button>`;
    return h;
  }
  function renderQuestion(q) {
    const v = S.answers[q];
    let h = `<div class="dxq" data-q="${q}">`;
    if (q === "black" || q === "white" || PICK[q]) {
      const opts = q === "black" ? BLACK_OPTS : q === "white" ? WHITE_OPTS : PICK[q];
      h += `<p class="dxq-t">${esc(t("dx." + q + ".q"))}</p><div class="opts q-${q}${PICK[q] ? " words" : ""}" role="group">` +
        opts.map((o) => `<button type="button" class="opt" data-v="${o}" aria-pressed="${v === o}">${esc(optLabel(q, o))}</button>`).join("") + `</div>`;
    } else if (q === "cast") {
      const cur = v || {};
      h += `<p class="dxq-t">${esc(t("dx.cast.q"))}</p><div class="castq">` + CAST_GROUPS.map((g) =>
        `<div class="castg"><span>${esc(t("dx.cast." + g))}</span><div class="pad" role="group" aria-label="${esc(t("dx.cast." + g))}">` +
        [1, 2, 3, 4, 5, 6, 7, 8, 9].map((i) => `<button type="button" class="opt" data-g="${g}" data-v="${i}" aria-pressed="${String(cur[g]) === String(i)}">${i}</button>`).join("") +
        `</div></div>`).join("") + `</div>`;
    } else if (q === "ramps") {
      const cur = v || [];
      h += `<p class="dxq-t">${esc(t("dx.ramps.q"))}</p><div class="opts words" role="group">` +
        RAMP_KEYS.map((k) => `<button type="button" class="opt" data-m="${k}" aria-pressed="${cur.includes(k)}">${esc(t("c." + k))}</button>`).join("") +
        `<button type="button" class="opt" data-m="none" aria-pressed="${cur.includes("none")}">${esc(t("dx.ramps.none"))}</button></div>`;
    }
    return h + resultBox(diagnose(q)) + `</div>`;
  }
  function renderSettings(keys) {
    return `<section class="sec-block"><div class="h">${esc(t("set.h"))}</div><div class="sets">` +
      `<div class="set-row head"><span></span><span>${esc(t("set.before"))}</span><span></span><span>${esc(t("set.after"))}</span></div>` +
      keys.map((k) => {
        const val = S.values[k] || {};
        return `<div class="set-row"><label for="set-${k}-a">${esc(t("set." + k))}</label>` +
          `<input type="text" id="set-${k}-a" data-set="${k}" data-side="a" value="${esc(val.a || "")}" inputmode="text" autocomplete="off">` +
          `<span aria-hidden="true">→</span>` +
          `<input type="text" id="set-${k}-b" data-set="${k}" data-side="b" value="${esc(val.b || "")}" inputmode="text" autocomplete="off" aria-label="${esc(t("set." + k) + " " + t("set.after"))}"></div>`;
      }).join("") + `</div></section>`;
  }
  function wireDx(art) {
    $$(".dxq", art).forEach((box) => {
      const q = box.dataset.q;
      $$(".opt", box).forEach((b) => {
        b.onclick = () => {
          if (q === "cast") { const cur = Object.assign({}, S.answers.cast || {}); cur[b.dataset.g] = +b.dataset.v; answer("cast", cur); return; }
          if (q === "ramps") {
            let cur = (S.answers.ramps || []).slice(); const m = b.dataset.m;
            if (m === "none") cur = cur.includes("none") ? [] : ["none"];
            else { cur = cur.filter((x) => x !== "none"); cur = cur.includes(m) ? cur.filter((x) => x !== m) : cur.concat(m); }
            answer("ramps", cur); return;
          }
          answer(q, b.dataset.v);
        };
      });
    });
    $$(".dx-go", art).forEach((b) => { b.onclick = () => goStep(+b.dataset.go); });
    $$("[data-set]", art).forEach((inp) => {
      inp.oninput = () => {
        const k = inp.dataset.set, v = S.values[k] || {};
        v[inp.dataset.side] = inp.value.trim(); S.values[k] = v; persist();
        const rep = $("#report"); if (rep) rep.textContent = reportText();
      };
    });
  }

  /* ---------- Registro delle sale ---------- */
  const venues = () => store.get("projectorlab.venues", []);
  function saveVenue() {
    const name = (S.venue || "").trim();
    if (!name) { toast(t("venue.need")); return; }
    const list = venues().filter((x) => x.venue.toLowerCase() !== name.toLowerCase());
    list.unshift({ venue: name, date: new Date().toISOString(), fmt: fmtKey(), values: S.values, answers: S.answers, note: S.note });
    store.set("projectorlab.venues", list.slice(0, 40));
    toast(t("venue.saved"));
  }
  function valueLines(values) {
    return SET_ORDER.filter((k) => values[k] && (values[k].a || values[k].b)).map((k) => {
      const v = values[k];
      return `${t("set." + k)}: ${v.a && v.b && v.a !== v.b ? v.a + " → " + v.b : v.b || v.a}`;
    });
  }
  function renderVenues() {
    const list = venues();
    if (!list.length) return "";
    return `<section class="sec-block"><div class="h">${esc(t("venue.past"))}</div>
      <select id="venuePick"><option value="">${esc(t("venue.pick"))}</option>` +
      list.map((x, i) => `<option value="${i}">${esc(x.venue)} · ${esc(new Date(x.date).toLocaleDateString(LANG))}</option>`).join("") +
      `</select><div id="venueInfo"></div></section>`;
  }
  function showVenue(i) {
    const box = $("#venueInfo"); if (!box) return;
    const x = venues()[i];
    if (!x) { box.innerHTML = ""; return; }
    const lines = valueLines(x.values || {});
    box.innerHTML = `<div class="hint"><p class="vtitle">${esc(t("venue.last", { d: new Date(x.date).toLocaleDateString(LANG) }))}</p>` +
      (lines.length ? `<ul class="list adj">${lines.map((l) => `<li>${esc(l)}</li>`).join("")}</ul>` : `<p>${esc(t("venue.empty"))}</p>`) +
      (x.note ? `<p class="vnote">${esc(x.note)}</p>` : "") + `</div>` +
      `<div class="btns">` + (lines.length ? `<button type="button" class="btn sec small" id="venueUse">${esc(t("venue.use"))}</button>` : "") +
      `<button type="button" class="btn sec small danger" id="venueDel">${esc(t("venue.del"))}</button></div>`;
    const use = $("#venueUse");
    if (use) use.onclick = () => {
      Object.keys(x.values || {}).forEach((k) => { const v = x.values[k]; const start = v.b || v.a; if (start) S.values[k] = { a: start, b: "" }; });
      S.venue = x.venue; persist(); toast(t("venue.used")); renderStep();
    };
    $("#venueDel").onclick = () => { const l = venues(); l.splice(i, 1); store.set("projectorlab.venues", l); renderStep(); };
  }

  /* ---------- Passi ---------- */
  function renderStepper() {
    const box = $("#stepper"); box.innerHTML = "";
    STEPS().forEach((s, i) => {
      const b = document.createElement("button");
      b.type = "button"; b.className = "dot " + (S.results[s.id] || "");
      b.textContent = String(i + 1);
      b.title = s.title; b.setAttribute("aria-label", `${t("step")} ${i + 1}: ${s.title}`);
      if (i === S.step) b.setAttribute("aria-current", "step");
      b.onclick = () => goStep(i);
      box.appendChild(b);
    });
    $("#progress").style.width = ((S.step + 1) / STEPS().length) * 100 + "%";
  }
  function markThumbs() { $$(".thumb").forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.k === S.current))); }

  function list(items, cls) { return `<ul class="list ${cls || ""}">${items.map((x) => `<li>${esc(x)}</li>`).join("")}</ul>`; }
  function renderStep() {
    renderStepper();
    const steps = STEPS(), s = steps[S.step], n = S.step + 1;
    const W = fmtW();
    let h = `<header class="step-head"><div class="h">${esc(t("step"))} ${n} ${esc(t("of"))} ${steps.length}</div><h2>${esc(s.title)}</h2><p class="goal">${esc(s.goal)}</p></header>`;

    if (s.id === "prep") {
      h += `<div class="field" style="grid-template-columns:1fr"><span class="h" style="font-size:10.5px">${esc(t("venue"))}</span><input type="text" id="venue" value="${esc(S.venue)}" placeholder="${esc(t("venue.ph"))}"></div>`;
      h += `<button type="button" class="btn sec small" id="timerBtn"><svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="13" r="8"/><path d="M12 9v4l3 2M9 2h6"/></svg><span class="timer"></span></button>`;
      h += renderVenues();
    }
    if (s.slides) {
      const tests = mode === "office" && hasTests();
      h += `<div class="thumbs" style="--ar:${W}/7.5">` + s.slides.map((k) => {
        const svg = PLP.toSVG(PLP.build(k, W, labels()), W, { attrs: ' width="100%" height="100%" aria-hidden="true"' });
        return `<button type="button" class="thumb" data-k="${k}"><div class="img">${svg}</div><span>${ICON_PLAY}${esc(t("p." + k + ".name"))}</span></button>`;
      }).join("") + `</div>`;
      if (mode === "office") h += `<p class="note">${esc(tests ? t("startShow") : t("notInserted"))}</p>`;
    }
    if (s.checks) {
      const done = S.checks[s.id] || [];
      h += `<section class="sec-block"><div class="h">${esc(t("checklist"))}</div><div class="checks">` +
        s.checks.map((c, i) => `<label class="check"><input type="checkbox" data-ci="${i}"${done[i] ? " checked" : ""}><span>${esc(c)}</span></label>`).join("") + `</div></section>`;
    }
    if (s.tips) h += s.tips.map((x) => `<p class="hint">${esc(x)}</p>`).join("");
    if (s.look) h += `<section class="sec-block"><div class="h">${esc(t("look"))}</div>${list(s.look)}</section>`;
    if (STEP_Q[s.id]) {
      h += `<section class="sec-block dx"><div class="h">${esc(t("dx.h"))}</div><p class="note">${esc(t("dx.hint"))}</p>` +
        STEP_Q[s.id].map(renderQuestion).join("") + `</section>`;
    }
    if (s.adjust) h += `<section class="sec-block"><div class="h">${esc(t("adjust"))}</div>${list(s.adjust, "adj")}</section>`;
    if (STEP_SET[s.id]) h += renderSettings(STEP_SET[s.id]);
    if (s.id === "clinical") {
      h += `<button type="button" class="btn sec" id="photoBtn"><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="5" width="18" height="14" rx="2"/><circle cx="9" cy="10" r="1.6"/><path d="M4 18l5-5 4 4 3-3 4 4"/></svg><span>${esc(mode === "office" ? t("photo") : t("photoWeb"))}</span></button>`;
    }
    if (s.issues) {
      h += `<section class="sec-block"><div class="h">${esc(t("issues"))}</div><div>` +
        s.issues.map(([q, a]) => `<details class="issue"><summary>${esc(q)}</summary><p>${esc(a)}</p></details>`).join("") + `</div></section>`;
    }
    if (s.id !== "prep" && s.id !== "final") {
      const r = S.results[s.id];
      h += `<section class="sec-block"><div class="h">${esc(t("result"))}</div><div class="result">` +
        ["ok", "fix", "skip"].map((k) => `<button type="button" class="rb ${k}" data-r="${k}" aria-pressed="${r === k}"><i></i>${esc(t("res." + k))}</button>`).join("") + `</div></section>`;
    }
    if (s.cant) h += `<section class="sec-block"><div class="h">${esc(t("cant.h"))}</div>${list(s.cant, "adj")}</section>`;
    if (s.id === "final") {
      h += `<section class="sec-block"><div class="h">${esc(t("report"))}</div>
        <textarea id="note" placeholder="${esc(t("report.note.ph"))}">${esc(S.note)}</textarea>
        <pre class="report" id="report"></pre>
        <div class="btns"><button type="button" class="btn sec small" id="copyBtn"><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="8" y="8" width="12" height="12" rx="2"/><path d="M16 8V5a1 1 0 0 0-1-1H5a1 1 0 0 0-1 1v10a1 1 0 0 0 1 1h3"/></svg><span>${esc(t("report.copy"))}</span></button>` +
        `<button type="button" class="btn sec small" id="saveVenueBtn"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 4h11l3 3v13H5zM8 4v5h7V4M8 20v-6h8v6"/></svg><span>${esc(t("venue.save"))}</span></button>` +
        (mode === "office" && hasTests() ? `<button type="button" class="btn sec small danger" id="removeBtn2"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/></svg><span>${esc(t("remove"))}</span></button>` : "") +
        `</div></section>`;
    }
    const art = $("#step"); art.innerHTML = h;

    // Collegamenti
    $$(".thumb", art).forEach((b) => { b.onclick = () => goPattern(b.dataset.k); });
    markThumbs();
    $$(".check input", art).forEach((c) => {
      c.onchange = () => { const a = S.checks[s.id] || []; a[+c.dataset.ci] = c.checked; S.checks[s.id] = a; persist(); };
    });
    $$(".rb", art).forEach((b) => {
      b.onclick = () => {
        S.results[s.id] = S.results[s.id] === b.dataset.r ? undefined : b.dataset.r;
        if (!S.results[s.id]) delete S.results[s.id];
        persist(); renderStep();
      };
    });
    const v = $("#venue", art); if (v) v.oninput = () => { S.venue = v.value; persist(); };
    wireDx(art);
    const vp = $("#venuePick", art); if (vp) vp.onchange = () => showVenue(vp.value === "" ? -1 : +vp.value);
    const sv = $("#saveVenueBtn", art); if (sv) sv.onclick = saveVenue;
    const tb = $("#timerBtn", art);
    if (tb) { tb.onclick = startTimer; tickTimer(); if (timerEnd && Date.now() < timerEnd && !timerT) timerT = setInterval(tickTimer, 1000); }
    const pb = $("#photoBtn", art); if (pb) pb.onclick = pickPhoto;
    const note = $("#note", art); if (note) note.oninput = () => { S.note = note.value; persist(); $("#report").textContent = reportText(); };
    const rep = $("#report", art); if (rep) rep.textContent = reportText();
    const cb = $("#copyBtn", art); if (cb) cb.onclick = copyReport;
    const rb2 = $("#removeBtn2", art); if (rb2) rb2.onclick = removeTests;

    $("#prevBtn").disabled = S.step === 0;
    $("#nextBtn").textContent = S.step === steps.length - 1 ? t("report.copy") : t("next");
  }

  function goStep(i) {
    const steps = STEPS();
    S.step = Math.max(0, Math.min(steps.length - 1, i)); persist();
    renderStep();
    window.scrollTo({ top: 0, behavior: "smooth" });
    const first = (steps[S.step].slides || [])[0];
    const winOpen = !!dialog || inl.open || (showWin && !showWin.closed);
    if (first && (winOpen || (mode === "office" && S.map[first]))) goPattern(first);
  }

  /* ---------- Avvio ---------- */
  function applyStatic() {
    document.documentElement.lang = LANG;
    $$("[data-i18n]").forEach((el) => { el.textContent = t(el.dataset.i18n); });
    $$("[data-i18n-aria]").forEach((el) => el.setAttribute("aria-label", t(el.dataset.i18nAria)));
    // Le pagine di supporto si aprono nella lingua del pannello
    $$("a[data-page]").forEach((el) => { el.href = el.dataset.page + "?lang=" + LANG; });
    setStatus(busy ? "busy" : "");
    renderSetup(); renderStep();
  }
  function buildLang() {
    const sel = $("#lang"); sel.innerHTML = "";
    I18N.langs.filter(([k]) => LANGS.includes(k)).forEach(([k, name]) => {
      const o = document.createElement("option"); o.value = k; o.textContent = name; sel.appendChild(o);
    });
    sel.value = LANG;
    sel.onchange = () => { LANG = sel.value; store.set("projectorlab.lang", LANG); applyStatic(); sendToWindow({ lang: LANG }); };
  }
  function wire() {
    $("#setupToggle").onclick = () => toggleSetup();
    $("#insertBtn").onclick = insertTests;
    $("#removeBtn").onclick = removeTests;
    $("#windowBtn").onclick = () => openWindow(S.current);
    $("#downloadBtn").onclick = download;
    $$("#whereSeg .pill").forEach((b) => { b.onclick = () => { S.where = b.dataset.where; renderSetup(); }; });
    $("#prevBtn").onclick = () => goStep(S.step - 1);
    $("#nextBtn").onclick = () => { if (S.step === STEPS().length - 1) copyReport(); else goStep(S.step + 1); };
    $("#photoIn").onchange = (e) => onPhoto(e.target.files && e.target.files[0]);
    // In PowerPoint i link si aprono nel browser di sistema
    $$(".links a").forEach((el) => el.addEventListener("click", (e) => {
      if (!inOffice) return;
      try {
        if (Office.context.requirements.isSetSupported("OpenBrowserWindowApi", "1.1")) { e.preventDefault(); Office.context.ui.openBrowserWindow(el.href); }
      } catch (x) { /* resta il link normale */ }
    }));
  }

  async function start(office) {
    inOffice = office;
    mode = office ? "office" : "web";
    document.body.classList.toggle("web", !office);
    LANG = detectLang();
    buildLang(); wire(); wireInline();
    if (office) {
      api.insert = supported("PowerPointApi", "1.2");
      api.tags = supported("PowerPointApi", "1.3");
      api.select = supported("PowerPointApi", "1.5");
      api.image = supported("PowerPointApi", "1.8");
      api.dialog2 = supported("DialogApi", "1.2");
      api.create = typeof PowerPoint !== "undefined" && typeof PowerPoint.createPresentation === "function";
      if (!api.insert) {
        $("#unsupported").hidden = false;
        mode = "web"; // guida e finestra dei pattern restano disponibili
      }
    }
    if (S.step >= STEPS().length) S.step = 0;
    // Nel browser il formato "automatico" è quello dello schermo
    if (mode !== "office") { try { S.detected = PLP.formatFromRatio(screen.width / screen.height); } catch (e) { S.detected = "16:9"; } }
    applyStatic();
    if (mode === "office") {
      setBusy(true);
      await Promise.all([detectFormat(), loadMap()]);
      setBusy(false);
      applyStatic();
      if (hasTests()) toggleSetup(false);
    }
  }

  let started = false;
  const go = (office) => { if (!started) { started = true; start(office); } };
  if (window.Office && Office.onReady) {
    Office.onReady((info) => go(!!(info && info.host === Office.HostType.PowerPoint)));
  } else go(false);
})();
