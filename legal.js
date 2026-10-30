/* The Ultraspeaker Projector Lab – lingua delle pagine di supporto, privacy, condizioni e guida.
   Stessa scelta del pannello: ?lang=, poi la lingua salvata, poi quella del browser. */
(function () {
  "use strict";
  const LANGS = [["it", "Italiano"], ["en", "English"], ["es", "Español"], ["fr", "Français"], ["de", "Deutsch"]];
  const CODES = LANGS.map((l) => l[0]);
  const NAV = {
    it: { support: "Supporto", privacy: "Privacy", terms: "Condizioni d’uso", guide: "Guida d’installazione", lang: "Lingua" },
    en: { support: "Support", privacy: "Privacy", terms: "Terms of use", guide: "Installation guide", lang: "Language" },
    es: { support: "Soporte", privacy: "Privacidad", terms: "Condiciones de uso", guide: "Guía de instalación", lang: "Idioma" },
    fr: { support: "Assistance", privacy: "Confidentialité", terms: "Conditions d’utilisation", guide: "Guide d’installation", lang: "Langue" },
    de: { support: "Support", privacy: "Datenschutz", terms: "Nutzungsbedingungen", guide: "Installationsanleitung", lang: "Sprache" },
  };
  const KEY = "projectorlab.lang"; // condivisa con il pannello, salvata come JSON
  const saved = () => { try { return JSON.parse(localStorage.getItem(KEY)); } catch (e) { return null; } };

  function detect() {
    const l = [new URLSearchParams(location.search).get("lang"), saved(), navigator.language.slice(0, 2).toLowerCase()].find((x) => CODES.includes(x));
    return l || "en";
  }
  function apply(lang) {
    document.documentElement.lang = lang;
    document.querySelectorAll("[data-lang]").forEach((b) => { b.hidden = b.dataset.lang !== lang; });
    document.querySelectorAll("nav a[data-nav]").forEach((a) => { a.textContent = NAV[lang][a.dataset.nav]; a.href = a.dataset.href + "?lang=" + lang; });
    // link interni (pagine .html del sito): portano la lingua scelta
    document.querySelectorAll("main a:not([href*=':'])").forEach((a) => { a.setAttribute("href", a.getAttribute("href").split("?")[0] + "?lang=" + lang); });
    document.title = NAV[lang][document.body.dataset.page] + " – The Ultraspeaker Projector Lab";
    sel.value = lang; sel.setAttribute("aria-label", NAV[lang].lang);
    const u = new URL(location.href); u.searchParams.set("lang", lang); history.replaceState(null, "", u);
  }
  const sel = document.getElementById("lang");
  LANGS.forEach(([code, name]) => sel.add(new Option(name, code)));
  sel.addEventListener("change", () => { try { localStorage.setItem(KEY, JSON.stringify(sel.value)); } catch (e) { /* archiviazione non disponibile */ } apply(sel.value); });
  apply(detect());
})();
