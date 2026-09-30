/* The Ultraspeaker Projector Lab – lingua delle pagine di supporto, privacy, condizioni e guida.
   Stessa scelta del pannello: ?lang=, poi la lingua salvata, poi quella del browser. */
(function () {
  "use strict";
  var LANGS = [["it", "Italiano"], ["en", "English"], ["es", "Español"], ["fr", "Français"], ["de", "Deutsch"]];
  var CODES = LANGS.map(function (l) { return l[0]; });
  var NAV = {
    it: { support: "Supporto", privacy: "Privacy", terms: "Condizioni d’uso", guide: "Guida d’installazione", lang: "Lingua" },
    en: { support: "Support", privacy: "Privacy", terms: "Terms of use", guide: "Installation guide", lang: "Language" },
    es: { support: "Soporte", privacy: "Privacidad", terms: "Condiciones de uso", guide: "Guía de instalación", lang: "Idioma" },
    fr: { support: "Assistance", privacy: "Confidentialité", terms: "Conditions d’utilisation", guide: "Guide d’installation", lang: "Langue" },
    de: { support: "Support", privacy: "Datenschutz", terms: "Nutzungsbedingungen", guide: "Installationsanleitung", lang: "Sprache" }
  };
  var KEY = "projectorlab.lang";
  function get() { try { return localStorage.getItem(KEY); } catch (e) { return null; } }
  function set(v) { try { localStorage.setItem(KEY, JSON.stringify(v)); } catch (e) { /* archiviazione non disponibile */ } }
  function saved() { var v = get(); try { v = JSON.parse(v); } catch (e) { /* valore semplice */ } return CODES.indexOf(v) >= 0 ? v : null; }
  function detect() {
    var q = null;
    try { q = new URLSearchParams(location.search).get("lang"); } catch (e) { /* niente */ }
    if (CODES.indexOf(q) >= 0) return q;
    var s = saved(); if (s) return s;
    var l = ((navigator.languages && navigator.languages[0]) || navigator.language || "it").slice(0, 2).toLowerCase();
    return CODES.indexOf(l) >= 0 ? l : "en";
  }
  function apply(lang) {
    document.documentElement.lang = lang;
    var blocks = document.querySelectorAll("[data-lang]");
    for (var i = 0; i < blocks.length; i++) blocks[i].hidden = blocks[i].getAttribute("data-lang") !== lang;
    var page = document.body.getAttribute("data-page");
    var nav = document.querySelectorAll("nav a[data-nav]");
    for (var j = 0; j < nav.length; j++) {
      var k = nav[j].getAttribute("data-nav");
      nav[j].textContent = NAV[lang][k];
      nav[j].href = nav[j].getAttribute("data-href") + "?lang=" + lang;
    }
    var links = document.querySelectorAll("main a[href$='.html']");
    for (var m = 0; m < links.length; m++) {
      var h = links[m].getAttribute("href").split("?")[0];
      links[m].setAttribute("href", h + "?lang=" + lang);
    }
    document.title = NAV[lang][page] + " – The Ultraspeaker Projector Lab";
    var sel = document.getElementById("lang");
    sel.value = lang; sel.setAttribute("aria-label", NAV[lang].lang);
    try { var u = new URL(location.href); u.searchParams.set("lang", lang); history.replaceState(null, "", u.toString()); } catch (e) { /* niente */ }
  }
  var sel = document.getElementById("lang");
  LANGS.forEach(function (l) { var o = document.createElement("option"); o.value = l[0]; o.textContent = l[1]; sel.appendChild(o); });
  sel.addEventListener("change", function () { set(sel.value); apply(sel.value); });
  apply(detect());
})();
