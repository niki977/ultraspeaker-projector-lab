/* The Ultraspeaker Projector Lab – add-in PowerPoint che guida la regolazione del proiettore */
(function () {
  "use strict";

  /* ---------- Utilità ---------- */
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  // Memoria del browser: chiavi ultraspeaker.projector.<cosa> (linea guida §10), valori in JSON
  const PREFIX = "ultraspeaker.projector.", OLD = "projectorlab.";
  const store = {
    get(k, d) { try { const v = localStorage.getItem(PREFIX + k); return v ? JSON.parse(v) : d; } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem(PREFIX + k, JSON.stringify(v)); } catch (e) { /* memoria non disponibile */ } },
  };
  // Le versioni precedenti salvavano in projectorlab.*: si copia nella chiave nuova (se manca) e si cancella la vecchia,
  // così nessuno perde sale salvate o procedura in corso
  try {
    ["state", "timer", "venues"].forEach((k) => {
      const old = localStorage.getItem(OLD + k);
      if (old == null) return;
      if (localStorage.getItem(PREFIX + k) == null) localStorage.setItem(PREFIX + k, old);
      localStorage.removeItem(OLD + k);
    });
  } catch (e) { /* memoria non disponibile */ }
  let toastT;
  function toast(msg) {
    const el = $("#toast"); el.textContent = msg; el.classList.add("show");
    clearTimeout(toastT); toastT = setTimeout(() => el.classList.remove("show"), 2800);
  }
  const wait = (ms) => new Promise((r) => setTimeout(r, ms));
  const PLP = window.PL_PATTERNS;
  const ICON_PLAY = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M8 5l11 7-11 7z"/></svg>';

  /* ---------- Lingua ---------- */
  const I18N = window.PL_I18N, L = window.UltraspeakerLang;
  let LANG = "en";
  // Tutte le lingue hanno le stesse chiavi (stesso elenco dell'italiano)
  const t = (k, vars) => {
    let s = I18N[LANG][k] ?? k;
    if (vars) Object.keys(vars).forEach((v) => { s = s.replace("{" + v + "}", vars[v]); });
    return s;
  };
  const STEPS = () => I18N[LANG].steps;
  // Lingua comune a tutti i Lab: ?lang=, scelta salvata, vecchia chiave del Lab, lingua di Office, sistema, inglese
  function detectLang() {
    const lang = L.set(L.get(OLD + "lang", office ? Office.context.displayLanguage : null));
    try { localStorage.removeItem(OLD + "lang"); } catch (e) { /* memoria non disponibile */ }
    return lang;
  }
  // Etichette per i pattern nella lingua corrente
  function labels() {
    const out = {};
    Object.keys(I18N.it).forEach((k) => { if (k.startsWith("p.")) out[k] = t(k); });
    return out;
  }

  /* ---------- Stato ---------- */
  let office = false; // true dentro PowerPoint, false nel browser
  // Funzioni opzionali di PowerPoint: il manifest richiede già PowerPointApi 1.2 (inserimento slide e nuova presentazione)
  const api = { shapes: false, select: false, image: false, dialog2: false };
  const saved = store.get("state", {});
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
  const persist = () => store.set("state", { step: S.step, results: S.results, checks: S.checks, venue: S.venue, note: S.note, answers: S.answers, values: S.values });
  let busy = false;
  let dialog = null, showWin = null;

  const fmtKey = () => (S.fmt === "auto" ? S.detected || "16:9" : S.fmt);
  // Formato scelto diverso da quello della presentazione: lo Screen Test va in una nuova presentazione
  const needsNewPres = () => office && !!S.presFmt && fmtKey() !== S.presFmt;
  const fmtW = () => PLP.FORMATS[fmtKey()];
  const hasTests = () => Object.keys(S.map).length > 0;
  const stepOfPattern = (key) => STEPS().findIndex((s) => (s.slides || []).includes(key));

  /* ---------- Stato in alto ---------- */
  function setStatus(kind) {
    $("#statusText").textContent = t(kind === "busy" ? "status.busy" : office ? "status.office" : "us.web");
  }
  function setBusy(b) {
    busy = b; setStatus(b ? "busy" : "");
    ["#insertBtn", "#removeBtn", "#downloadBtn"].forEach((s) => { $(s).disabled = b; });
  }

  /* ---------- Scheda Screen Test ---------- */
  function renderSetup() {
    $("#setupText").textContent = t(office ? "setup.text" : "setup.textWeb");
    $$(".office-only").forEach((e) => { e.hidden = !office; });
    $$(".web-only").forEach((e) => { e.hidden = office; });
    $("#windowLabel").textContent = t(office ? "window" : "fullscreen");
    // Formati
    const seg = $("#fmtSeg"); seg.innerHTML = "";
    const opts = ["auto"].concat(Object.keys(PLP.FORMATS));
    opts.forEach((o) => {
      const b = document.createElement("button");
      b.type = "button"; b.className = "us-chip"; b.dataset.fmt = o;
      b.setAttribute("aria-pressed", String(S.fmt === o));
      b.innerHTML = o === "auto" ? `${esc(S.detected || "16:9")}<small>${esc(t("format.auto"))}</small>` : esc(o);
      b.onclick = () => { S.fmt = o; renderSetup(); renderStep(); sendToWindow({ f: o === "auto" ? "" : o }); };
      seg.appendChild(b);
    });
    $$("#whereSeg .us-chip").forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.where === S.where)));
    $("#fmtHint").textContent = t("fmtHint");
    const n = Object.keys(S.map).length;
    const np = needsNewPres();
    $("#insertLabel").textContent = np ? t("newPres", { f: fmtKey() }) : n ? t("update") : t("insert");
    $("#newPresHint").hidden = !np;
    $("#newPresHint").textContent = np ? t("newPresHint", { p: S.presFmt, f: fmtKey() }) : "";
    $("#whereRow").hidden = !office || np;
    $("#removeBtn").hidden = !n || !office;
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

  // Ritrova le slide di test: la prima forma di ogni slide di test si chiama "USPL:<pattern>"
  // (vale anche per un file scaricato o una nuova presentazione). Senza PowerPointApi 1.3 le forme
  // non sono leggibili: si usano gli id salvati nelle impostazioni del documento al momento dell'inserimento.
  async function loadMap() {
    const map = {};
    try {
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides; slides.load("items/id"); await ctx.sync();
        if (api.shapes) {
          const firsts = slides.items.map((s) => { const c = s.shapes; c.load({ select: "name", top: 1 }); return c; });
          await ctx.sync();
          firsts.forEach((c, i) => {
            const nm = c.items[0] && c.items[0].name;
            if (nm && nm.startsWith(PLP.MARK)) map[nm.slice(PLP.MARK.length)] = slides.items[i].id;
          });
        } else {
          const ids = new Set(slides.items.map((s) => s.id)), saved = Office.context.document.settings.get("uspl") || {};
          Object.keys(saved).forEach((k) => { if (ids.has(saved[k])) map[k] = saved[k]; });
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
    setBusy(true);
    try {
      if (S.fmt === "auto") await detectFormat();
      const pres = PLP.buildDeck(PptxGenJS, fmtW(), labels(), notesFor);
      const b64 = await pres.write({ outputType: "base64" });
      if (needsNewPres()) {
        await PowerPoint.createPresentation(b64);
        toast(t("opened", { f: fmtKey() }));
        return;
      }
      await deleteTests();
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
        fresh.forEach((s, i) => { if (PLP.ORDER[i]) map[PLP.ORDER[i]] = s.id; });
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
  /* ---------- Pattern dentro la pagina (anteprima, o popup bloccato): show.html in un riquadro a tutto schermo ---------- */
  const frame = () => $("#viewer");
  let frameLoaded = false;
  const inlineOpen = () => !frame().hidden;
  const winOpen = () => !!dialog || inlineOpen() || (showWin && !showWin.closed);
  const target = location.origin === "null" ? "*" : location.origin;
  function toggleFs() {
    if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
    else frame().requestFullscreen().catch(() => { /* tutto schermo non disponibile */ });
  }
  function openInline(key) {
    const v = frame();
    v.hidden = false; document.body.style.overflow = "hidden";
    // Già caricato (dopo una chiusura): riprende con la foto eventualmente inserita
    if (frameLoaded) v.contentWindow.postMessage({ uspl: 1, k: key, f: S.fmt === "auto" ? "" : S.fmt, lang: LANG }, target);
    else { frameLoaded = true; v.onload = () => v.contentWindow.focus(); v.src = showUrl(key); }
    v.contentWindow.focus(); toggleFs();
  }
  function closeInline() {
    if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
    frame().hidden = true; document.body.style.overflow = "";
  }

  function sendToWindow(msg) {
    if (inlineOpen()) frame().contentWindow.postMessage(Object.assign({ uspl: 1 }, msg), target);
    else if (dialog) {
      // Senza DialogApi 1.2 la finestra non riceve messaggi: si chiude e si riapre sul pattern giusto
      if (api.dialog2) { try { dialog.messageChild(JSON.stringify(msg)); return; } catch (e) { /* riapri */ } }
      try { dialog.close(); } catch (e) { /* già chiusa */ }
      dialog = null;
      openWindow(msg.k);
    } else if (showWin && !showWin.closed) showWin.postMessage(Object.assign({ uspl: 1 }, msg), target);
  }
  function openWindow(key) {
    key = key || (STEPS()[S.step].slides || [])[0] || "screenLight";
    if (office) {
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
  // Messaggi da show.html (finestra o riquadro): pattern cambiato, chiudi, tutto schermo
  window.addEventListener("message", (e) => {
    const m = e.data;
    if (!m || !m.uspl || m.from !== "show") return;
    if (m.k) followPattern(m.k);
    if (m.close) closeInline();
    if (m.fs) toggleFs();
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
    const open = winOpen();
    if (open) sendToWindow({ k: key });
    if (office && S.map[key]) { await selectSlide(S.map[key]); return; }
    if (!open && !office) openWindow(key);
    else if (!open && office && !hasTests()) toast(t("notInserted"));
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
    if (!office) {
      if (!winOpen()) { openWindow("photo"); await wait(700); } // tempo per caricare la finestra
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
    setBusy(true);
    try {
      const pres = PLP.buildDeck(PptxGenJS, fmtW(), labels(), notesFor);
      await pres.writeFile({ fileName: `The Ultraspeaker Projector Lab - Screen Test ${fmtKey().replace(":", "x")}.pptx` });
    } catch (e) { toast(t("err", { m: e.message })); } finally { setBusy(false); }
  }

  /* ---------- Timer di riscaldamento ---------- */
  let timerEnd = store.get("timer", 0), timerT = null;
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
    timerEnd = Date.now() + 15 * 60000; store.set("timer", timerEnd);
    clearInterval(timerT); timerT = setInterval(tickTimer, 1000); tickTimer();
  }

  /* ---------- Report ---------- */
  // Contenuto del report, già tradotto: lo usano il testo d'anteprima e il PDF (report-pdf.js)
  function answerLines() {
    const ans = [];
    ["black", "white", "cast", "ramps", "yellow", "gum", "enamel"].forEach((q) => {
      const v = S.answers[q]; if (v == null) return;
      let txt;
      if (q === "cast") txt = CAST_GROUPS.map((g) => {
        if (v[g] == null) return `${t("dx.cast." + g)}: –`;
        const c = castOf(v[g]), d = [];
        if (c.m > 0) d.push(t("dir.magenta")); if (c.m < 0) d.push(t("dir.green"));
        if (c.y > 0) d.push(t("dir.yellow")); if (c.y < 0) d.push(t("dir.blue"));
        return `${t("dx.cast." + g)}: ${d.join(" + ") || t("rep.neutral")}`;
      }).join(" · ");
      else if (q === "ramps") txt = v.map((k) => (k === "none" ? t("dx.ramps.none") : t("c." + k))).join(", ");
      else txt = optLabel(q, v);
      if (txt) ans.push([t("q." + q), txt]);
    });
    return ans;
  }
  const testSteps = () => STEPS().map((s, i) => [i + 1, s]).filter(([, s]) => s.id !== "prep" && s.id !== "final");
  function reportText() {
    const d = new Date();
    const lines = [t("report.title"), `${t("report.date")}: ${d.toLocaleDateString(LANG)} ${d.toLocaleTimeString(LANG, { hour: "2-digit", minute: "2-digit" })}`];
    if (S.venue) lines.push(`${t("venue")}: ${S.venue}`);
    lines.push(`${t("report.format")}: ${fmtKey()}`, "");
    testSteps().forEach(([n, s]) => {
      const r = S.results[s.id];
      const mark = r === "ok" ? "✓" : r === "fix" ? "!" : r === "skip" ? "–" : "·";
      lines.push(`${mark} ${n}. ${s.title}${r ? " — " + t("res." + r) : ""}`);
    });
    const ans = answerLines().map(([q, a]) => `${q}: ${a}`);
    if (ans.length) lines.push("", t("rep.answers") + ":", ...ans.map((x) => "  " + x));
    const vals = valueLines(S.values);
    if (vals.length) lines.push("", t("rep.values") + ":", ...vals.map((x) => "  " + x));
    if (S.note) lines.push("", `${t("report.note")}: ${S.note}`);
    return lines.join("\n");
  }
  function reportModel() {
    const d = new Date(), loc = LANG === "en" ? "en-GB" : LANG;
    const pad = (n) => String(n).padStart(2, "0");
    const L = {};
    ["kicker", "steps", "step", "result", "question", "answer", "setting", "before", "after", "skip", "none", "footer", "untitled", "file", "ready", "again", "error"].forEach((k) => { L[k] = t("rep." + k); });
    Object.assign(L, { ok: t("res.ok"), fix: t("res.fix"), answers: t("rep.answers"), values: t("rep.values"), note: t("report.note"),
      formatL: t("report.format"), cta: t("us.cta"), download: t("report.pdf") });
    return {
      v: 1, lang: LANG, L, venue: (S.venue || "").trim(), format: fmtKey(),
      date: `${d.toLocaleDateString(loc, { day: "numeric", month: "long", year: "numeric" })} · ${d.toLocaleTimeString(loc, { hour: "2-digit", minute: "2-digit" })}`,
      iso: `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`,
      steps: testSteps().map(([n, s]) => [n, s.title, S.results[s.id] || ""]),
      answers: answerLines(),
      values: SET_ORDER.filter((k) => S.values[k] && (S.values[k].a || S.values[k].b)).map((k) => [t("set." + k), S.values[k].a || "", S.values[k].b || ""]),
      note: (S.note || "").trim(),
      mail: "mailto:info@theultraspeaker.com?subject=" + encodeURIComponent(t("us.cta")),
    };
  }
  const loadScript = (src) => new Promise((ok, ko) => {
    if (document.querySelector(`script[src="${src}"]`)) return ok();
    const el = document.createElement("script"); el.src = src; el.onload = ok; el.onerror = () => ko(new Error(src)); document.head.appendChild(el);
  });
  // Nel pannello di PowerPoint scaricare un file spesso non funziona (soprattutto su Mac):
  // il report si apre nel browser (report.html), che crea il PDF da lì. I dati viaggiano dopo il #, quindi restano sul dispositivo.
  async function downloadReport() {
    const m = reportModel();
    if (office && supported("OpenBrowserWindowApi", "1.1")) {
      const json = new TextEncoder().encode(JSON.stringify(m));
      let bin = ""; json.forEach((b) => { bin += String.fromCharCode(b); });
      Office.context.ui.openBrowserWindow(new URL("report.html", location.href).href + "#" + btoa(bin).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, ""));
      toast(t("report.browser"));
      return;
    }
    setBusy(true);
    try {
      await loadScript("vendor/jspdf.umd.min.js?v=20261104b");
      await loadScript("report-pdf.js?v=20261104b");
      const doc = await window.PLReport.build(m, "");
      window.PLReport.save(doc, window.PLReport.fileName(m));
      toast(t("report.done"));
    } catch (e) { toast(t("err", { m: (e && e.message) || e })); } finally { setBusy(false); }
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
    const NOTE = { ok: "us-note--ok", fix: "us-note--warn", info: "" };
    let h = `<div class="us-note ${NOTE[d.sev]} pl-dxr" role="status">` + d.lines.map((x) => `<p>${esc(x)}</p>`).join("") + `</div>`;
    if (d.range) h += `<div class="us-note us-note--warn pl-dxr"><p><b>${esc(t("dx.range.h"))}</b></p><p>${esc(t("dx.range.crush"))}</p></div>`;
    if (d.recheck) h += `<p class="pl-note">${esc(t("dx.recheckNote"))}</p><button type="button" class="us-btn dx-go" data-go="3">${esc(t("dx.recheckBlack"))}</button>`;
    if (d.back) h += `<button type="button" class="us-btn dx-go" data-go="${d.back - 1}">${esc(t("dx.backTo", { n: d.back }))}</button>`;
    return h;
  }
  function renderQuestion(q) {
    const v = S.answers[q];
    let h = `<div class="us-panel pl-dxq dxq" data-q="${q}">`;
    if (q === "black" || q === "white" || PICK[q]) {
      const opts = q === "black" ? BLACK_OPTS : q === "white" ? WHITE_OPTS : PICK[q];
      h += `<p class="pl-dxq-t">${esc(t("dx." + q + ".q"))}</p><div class="pl-opts q-${q}${PICK[q] ? " words" : ""}" role="group">` +
        opts.map((o) => `<button type="button" class="us-chip" data-v="${o}" aria-pressed="${v === o}">${esc(optLabel(q, o))}</button>`).join("") + `</div>`;
    } else if (q === "cast") {
      const cur = v || {};
      h += `<p class="pl-dxq-t">${esc(t("dx.cast.q"))}</p><div class="pl-castq">` + CAST_GROUPS.map((g) =>
        `<div class="pl-castg"><span class="us-label">${esc(t("dx.cast." + g))}</span><div class="pl-pad" role="group" aria-label="${esc(t("dx.cast." + g))}">` +
        [1, 2, 3, 4, 5, 6, 7, 8, 9].map((i) => `<button type="button" class="us-chip" data-g="${g}" data-v="${i}" aria-pressed="${String(cur[g]) === String(i)}">${i}</button>`).join("") +
        `</div></div>`).join("") + `</div>`;
    } else if (q === "ramps") {
      const cur = v || [];
      h += `<p class="pl-dxq-t">${esc(t("dx.ramps.q"))}</p><div class="pl-opts words" role="group">` +
        RAMP_KEYS.map((k) => `<button type="button" class="us-chip" data-m="${k}" aria-pressed="${cur.includes(k)}">${esc(t("c." + k))}</button>`).join("") +
        `<button type="button" class="us-chip" data-m="none" aria-pressed="${cur.includes("none")}">${esc(t("dx.ramps.none"))}</button></div>`;
    }
    return h + resultBox(diagnose(q)) + `</div>`;
  }
  function renderSettings(keys) {
    return `<section class="pl-block"><p class="us-label">${esc(t("set.h"))}</p><div class="pl-sets">` +
      `<div class="pl-set-row head"><span></span><span class="us-label">${esc(t("set.before"))}</span><span></span><span class="us-label">${esc(t("set.after"))}</span></div>` +
      keys.map((k) => {
        const val = S.values[k] || {};
        return `<div class="pl-set-row"><label for="set-${k}-a">${esc(t("set." + k))}</label>` +
          `<input type="text" class="us-field" id="set-${k}-a" data-set="${k}" data-side="a" value="${esc(val.a || "")}" inputmode="text" autocomplete="off">` +
          `<span aria-hidden="true">→</span>` +
          `<input type="text" class="us-field" id="set-${k}-b" data-set="${k}" data-side="b" value="${esc(val.b || "")}" inputmode="text" autocomplete="off" aria-label="${esc(t("set." + k) + " " + t("set.after"))}"></div>`;
      }).join("") + `</div></section>`;
  }
  function wireDx(art) {
    $$(".dxq", art).forEach((box) => {
      const q = box.dataset.q;
      $$(".us-chip", box).forEach((b) => {
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

  /* ---------- Ricomincia e Azzera tutto: due tocchi, il primo chiede conferma per 4 secondi ---------- */
  function twoTaps(b, label, action) {
    clearTimeout(b.disarm);
    if (b.classList.toggle("armed")) b.disarm = setTimeout(() => { b.classList.remove("armed"); label(); }, 4000);
    else action();
    label();
  }
  function clearRun() {
    S.results = {}; S.checks = {}; S.answers = {}; S.values = {}; S.venue = ""; S.note = ""; S.step = 0;
    timerEnd = 0; clearInterval(timerT); timerT = null;
  }
  function restartLabel() {
    const armed = $("#restartBtn").classList.contains("armed");
    $("#restartLabel").textContent = armed ? t("reset.confirm") : t("reset");
    $$(".restart-final").forEach((x) => { x.classList.toggle("armed", armed); x.querySelector("span").textContent = armed ? t("reset.confirm") : t("reset"); });
  }
  // Ricomincia: le sale salvate, la lingua, il formato e le slide di test restano
  const restart = () => twoTaps($("#restartBtn"), restartLabel, () => {
    clearRun(); store.set("timer", 0);
    persist(); renderStep(); window.scrollTo({ top: 0, behavior: "smooth" });
    toast(t("reset.done"));
  });
  // Azzera tutto (solo versione browser): cancella le chiavi di questo Lab (anche le vecchie, finché esistono);
  // la lingua è comune a tutti i Lab e resta
  function resetAllLabel() {
    $("#resetAllLabel").textContent = t($("#resetAllBtn").classList.contains("armed") ? "resetAll.confirm" : "resetAll");
  }
  const resetAll = () => twoTaps($("#resetAllBtn"), resetAllLabel, () => {
    try { Object.keys(localStorage).filter((k) => k.startsWith(PREFIX) || k.startsWith(OLD)).forEach((k) => localStorage.removeItem(k)); } catch (e) { /* memoria non disponibile */ }
    clearRun(); S.fmt = "auto"; S.where = "end";
    closeInline(); frameLoaded = false; // al prossimo uso il riquadro riparte da zero (senza foto)
    applyStatic(); window.scrollTo({ top: 0, behavior: "smooth" });
    toast(t("resetAll.done"));
  });

  /* ---------- Registro delle sale ---------- */
  const venues = () => store.get("venues", []);
  function saveVenue() {
    const name = (S.venue || "").trim();
    if (!name) { toast(t("venue.need")); return; }
    const list = venues().filter((x) => x.venue.toLowerCase() !== name.toLowerCase());
    list.unshift({ venue: name, date: new Date().toISOString(), fmt: fmtKey(), values: S.values, answers: S.answers, note: S.note });
    store.set("venues", list.slice(0, 40));
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
    return `<section class="pl-block"><p class="us-label">${esc(t("venue.past"))}</p>
      <select class="us-field" id="venuePick"><option value="">${esc(t("venue.pick"))}</option>` +
      list.map((x, i) => `<option value="${i}">${esc(x.venue)} · ${esc(new Date(x.date).toLocaleDateString(LANG))}</option>`).join("") +
      `</select><div id="venueInfo"></div></section>`;
  }
  function showVenue(i) {
    const box = $("#venueInfo"); if (!box) return;
    const x = venues()[i];
    if (!x) { box.innerHTML = ""; return; }
    const lines = valueLines(x.values || {});
    box.innerHTML = `<div class="us-note"><p class="vtitle">${esc(t("venue.last", { d: new Date(x.date).toLocaleDateString(LANG) }))}</p>` +
      (lines.length ? `<ul class="pl-list adj">${lines.map((l) => `<li>${esc(l)}</li>`).join("")}</ul>` : `<p>${esc(t("venue.empty"))}</p>`) +
      (x.note ? `<p class="vnote">${esc(x.note)}</p>` : "") + `</div>` +
      `<div class="pl-btns">` + (lines.length ? `<button type="button" class="us-btn" id="venueUse">${esc(t("venue.use"))}</button>` : "") +
      `<button type="button" class="us-btn is-warn" id="venueDel">${esc(t("venue.del"))}</button></div>`;
    const use = $("#venueUse");
    if (use) use.onclick = () => {
      Object.keys(x.values || {}).forEach((k) => { const v = x.values[k]; const start = v.b || v.a; if (start) S.values[k] = { a: start, b: "" }; });
      S.venue = x.venue; persist(); toast(t("venue.used")); renderStep();
    };
    $("#venueDel").onclick = () => { const l = venues(); l.splice(i, 1); store.set("venues", l); renderStep(); };
  }

  /* ---------- Passi ---------- */
  function renderStepper() {
    const box = $("#stepper"); box.innerHTML = "";
    STEPS().forEach((s, i) => {
      const b = document.createElement("button");
      b.type = "button"; b.className = "pl-dot " + (S.results[s.id] || "");
      b.textContent = String(i + 1);
      b.title = s.title; b.setAttribute("aria-label", `${t("step")} ${i + 1}: ${s.title}`);
      if (i === S.step) b.setAttribute("aria-current", "step");
      b.onclick = () => goStep(i);
      box.appendChild(b);
    });
    $("#progress").style.width = ((S.step + 1) / STEPS().length) * 100 + "%";
  }
  function markThumbs() { $$(".pl-thumb").forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.k === S.current))); }

  function list(items, cls) { return `<ul class="pl-list ${cls || ""}">${items.map((x) => `<li>${esc(x)}</li>`).join("")}</ul>`; }
  function renderStep() {
    renderStepper();
    const steps = STEPS(), s = steps[S.step], n = S.step + 1;
    const W = fmtW();
    let h = `<header class="pl-step-head"><p class="us-label">${esc(t("step"))} ${n} ${esc(t("of"))} ${steps.length}</p><h2 class="us-title">${esc(s.title)}</h2><p class="pl-goal">${esc(s.goal)}</p></header>`;

    if (s.id === "prep") {
      h += `<label class="pl-block"><span class="us-label">${esc(t("venue"))}</span><input type="text" class="us-field" id="venue" value="${esc(S.venue)}" placeholder="${esc(t("venue.ph"))}"></label>`;
      h += `<button type="button" class="us-btn us-btn--block" id="timerBtn"><svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="13" r="8"/><path d="M12 9v4l3 2M9 2h6"/></svg><span class="pl-timer"></span></button>`;
      h += renderVenues();
    }
    if (s.slides) {
      const tests = office && hasTests();
      h += `<div class="pl-thumbs" style="--ar:${W}/7.5">` + s.slides.map((k) => {
        const svg = PLP.toSVG(PLP.build(k, W, labels()), W, { attrs: ' width="100%" height="100%" aria-hidden="true"' });
        return `<button type="button" class="pl-thumb" data-k="${k}"><div class="img">${svg}</div><span>${ICON_PLAY}${esc(t("p." + k + ".name"))}</span></button>`;
      }).join("") + `</div>`;
      if (office) h += `<p class="pl-note">${esc(tests ? t("startShow") : t("notInserted"))}</p>`;
    }
    if (s.checks) {
      const done = S.checks[s.id] || [];
      h += `<section class="pl-block"><p class="us-label">${esc(t("checklist"))}</p><div class="pl-checks">` +
        s.checks.map((c, i) => `<label class="pl-check"><input type="checkbox" data-ci="${i}"${done[i] ? " checked" : ""}><span>${esc(c)}</span></label>`).join("") + `</div></section>`;
    }
    if (s.tips) h += s.tips.map((x) => `<p class="us-note">${esc(x)}</p>`).join("");
    if (s.look) h += `<section class="pl-block"><p class="us-label">${esc(t("look"))}</p>${list(s.look)}</section>`;
    if (STEP_Q[s.id]) {
      h += `<section class="pl-block"><p class="us-label">${esc(t("dx.h"))}</p><p class="pl-note">${esc(t("dx.hint"))}</p>` +
        STEP_Q[s.id].map(renderQuestion).join("") + `</section>`;
    }
    if (s.adjust) h += `<section class="pl-block"><p class="us-label">${esc(t("adjust"))}</p>${list(s.adjust, "adj")}</section>`;
    if (STEP_SET[s.id]) h += renderSettings(STEP_SET[s.id]);
    if (s.id === "clinical") {
      h += `<button type="button" class="us-btn us-btn--block" id="photoBtn"><svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="5" width="18" height="14" rx="2"/><circle cx="9" cy="10" r="1.6"/><path d="M4 18l5-5 4 4 3-3 4 4"/></svg><span>${esc(t(office ? "photo" : "photoWeb"))}</span></button>`;
    }
    if (s.issues) {
      h += `<section class="pl-block"><p class="us-label">${esc(t("issues"))}</p><div>` +
        s.issues.map(([q, a]) => `<details class="pl-issue"><summary>${esc(q)}</summary><p>${esc(a)}</p></details>`).join("") + `</div></section>`;
    }
    if (s.id !== "prep" && s.id !== "final") {
      const r = S.results[s.id];
      h += `<section class="pl-block"><p class="us-label">${esc(t("result"))}</p><div class="pl-result">` +
        ["ok", "fix", "skip"].map((k) => `<button type="button" class="us-btn pl-res ${k}" data-r="${k}" aria-pressed="${r === k}"><i></i>${esc(t("res." + k))}</button>`).join("") + `</div></section>`;
    }
    if (s.cant) h += `<section class="pl-block"><p class="us-label">${esc(t("cant.h"))}</p>${list(s.cant, "adj")}</section>`;
    if (s.id === "final") {
      h += `<section class="pl-block"><p class="us-label">${esc(t("report"))}</p>
        <textarea class="us-field" id="note" placeholder="${esc(t("report.note.ph"))}">${esc(S.note)}</textarea>
        <pre class="pl-report" id="report"></pre>
        <div class="pl-btns"><button type="button" class="us-btn" id="pdfBtn"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M14 3H6a1 1 0 0 0-1 1v16a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V8zM14 3v5h5M12 11v6M9 14l3 3 3-3"/></svg><span>${esc(t("report.pdf"))}</span></button>` +
        `<button type="button" class="us-btn restart-final"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 12a9 9 0 1 0 3-6.7L3 8M3 3v5h5"/></svg><span>${esc(t("reset"))}</span></button>` +
        `<button type="button" class="us-btn" id="saveVenueBtn"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 4h11l3 3v13H5zM8 4v5h7V4M8 20v-6h8v6"/></svg><span>${esc(t("venue.save"))}</span></button>` +
        (office && hasTests() ? `<button type="button" class="us-btn is-warn" id="removeBtn2"><svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/></svg><span>${esc(t("remove"))}</span></button>` : "") +
        `</div></section>`;
    }
    const art = $("#step"); art.innerHTML = h;

    // Collegamenti
    $$(".pl-thumb", art).forEach((b) => { b.onclick = () => goPattern(b.dataset.k); });
    markThumbs();
    $$(".pl-check input", art).forEach((c) => {
      c.onchange = () => { const a = S.checks[s.id] || []; a[+c.dataset.ci] = c.checked; S.checks[s.id] = a; persist(); };
    });
    $$(".pl-res", art).forEach((b) => {
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
    $$(".restart-final", art).forEach((b) => { b.onclick = restart; });
    restartLabel();
    const tb = $("#timerBtn", art);
    if (tb) { tb.onclick = startTimer; tickTimer(); if (timerEnd && Date.now() < timerEnd && !timerT) timerT = setInterval(tickTimer, 1000); }
    const pb = $("#photoBtn", art); if (pb) pb.onclick = pickPhoto;
    const note = $("#note", art); if (note) note.oninput = () => { S.note = note.value; persist(); $("#report").textContent = reportText(); };
    const rep = $("#report", art); if (rep) rep.textContent = reportText();
    const pdf = $("#pdfBtn", art); if (pdf) pdf.onclick = downloadReport;
    const rb2 = $("#removeBtn2", art); if (rb2) rb2.onclick = removeTests;

    $("#prevBtn").disabled = S.step === 0;
    $("#nextBtn").textContent = S.step === steps.length - 1 ? t("report.pdf") : t("next");
  }

  function goStep(i) {
    const steps = STEPS();
    S.step = Math.max(0, Math.min(steps.length - 1, i)); persist();
    renderStep();
    window.scrollTo({ top: 0, behavior: "smooth" });
    const first = (steps[S.step].slides || [])[0];
    if (first && (winOpen() || (office && S.map[first]))) goPattern(first);
  }

  /* ---------- Avvio ---------- */
  function applyStatic() {
    document.documentElement.lang = LANG;
    $$("[data-i18n]").forEach((el) => { el.textContent = t(el.dataset.i18n); });
    $$("[data-i18n-aria]").forEach((el) => el.setAttribute("aria-label", t(el.dataset.i18nAria)));
    // Le pagine di servizio si aprono nella lingua del pannello; «Richiedi un consiglio» con oggetto tradotto
    $$(".us-links a").forEach((el) => { el.href = L.linkWithLang(el.getAttribute("href"), LANG); });
    $("#usCta").href = "mailto:info@theultraspeaker.com?subject=" + encodeURIComponent(t("us.cta"));
    // «Guida»: il PDF nella lingua scelta; su telefono e nel pannello resta solo l'icona, con il nome accessibile
    const g = $("#usGuide"); g.href = L.guideHref("projector", LANG); g.setAttribute("aria-label", t("us.guide.title")); g.title = t("us.guide.title");
    setStatus(busy ? "busy" : "");
    renderSetup(); renderStep(); resetAllLabel();
  }
  function buildLang() {
    L.bindSelect($("#lang"), LANG, (code) => { LANG = code; applyStatic(); sendToWindow({ lang: LANG }); });
  }
  function wire() {
    $("#setupToggle").onclick = () => toggleSetup();
    $("#insertBtn").onclick = insertTests;
    $("#removeBtn").onclick = removeTests;
    $("#windowBtn").onclick = () => openWindow(S.current);
    $("#downloadBtn").onclick = download;
    $$("#whereSeg .us-chip").forEach((b) => { b.onclick = () => { S.where = b.dataset.where; renderSetup(); }; });
    $("#restartBtn").onclick = restart;
    $("#resetAllBtn").onclick = resetAll;
    $("#prevBtn").onclick = () => goStep(S.step - 1);
    $("#nextBtn").onclick = () => { if (S.step === STEPS().length - 1) downloadReport(); else goStep(S.step + 1); };
    $("#photoIn").onchange = (e) => onPhoto(e.target.files && e.target.files[0]);
    // In PowerPoint i link si aprono nel browser di sistema
    $$(".us-links a, .us-logo, .us-guide").forEach((el) => el.addEventListener("click", (e) => {
      if (office && supported("OpenBrowserWindowApi", "1.1")) { e.preventDefault(); Office.context.ui.openBrowserWindow(el.href); }
    }));
  }

  async function start(inPowerPoint) {
    office = inPowerPoint;
    document.body.classList.toggle("us-addin", office);
    LANG = detectLang();
    buildLang(); wire();
    if (office) {
      api.shapes = supported("PowerPointApi", "1.3");
      api.select = supported("PowerPointApi", "1.5");
      api.image = supported("PowerPointApi", "1.8");
      api.dialog2 = supported("DialogApi", "1.2");
    } else S.detected = PLP.formatFromRatio(screen.width / screen.height); // nel browser "automatico" = formato dello schermo
    if (S.step >= STEPS().length) S.step = 0;
    applyStatic();
    if (office) {
      setBusy(true);
      await Promise.all([detectFormat(), loadMap()]);
      setBusy(false);
      applyStatic();
      if (hasTests()) toggleSetup(false);
    }
  }

  if (window.Office) Office.onReady((info) => start(info.host === Office.HostType.PowerPoint));
  else start(false);
})();
