# The Ultraspeaker Projector Lab

Add-in per PowerPoint che guida il relatore nella regolazione del proiettore prima di parlare.
Procedura in 10 passi: preparazione, formato/bordi/geometria, fuoco e leggibilità, livello del nero,
livello del bianco, grigi e temperatura colore, colore e saturazione, riferimenti clinici (dentali),
uniformità, prova finale con report. Interfaccia, etichette e note in italiano, inglese, spagnolo, francese e tedesco.

## Come funziona

1. **Inserisci le slide di test**: il pannello crea 15 pattern come forme native di PowerPoint, nel formato
   della presentazione (4:3, 16:9 o 3:1, rilevato da solo; si può scegliere a mano il formato dello schermo, e se è diverso da quello delle slide lo Screen Test si apre in una nuova presentazione) e li inserisce in fondo o all’inizio.
   Ogni slide ha nelle note del relatore cosa guardare e cosa regolare.
2. Collega il proiettore in modalità **Estendi** e avvia la presentazione dalla prima slide di test:
   il pubblico (e tu) vedi il pattern, la Visualizzazione Relatore ti mostra le istruzioni.
3. Nel pannello, per ogni passo: anteprime dei pattern (un clic porta alla slide), cosa devi vedere,
   come regolare, problemi tipici con soluzione, esito (tutto a posto / da rivedere / salta).
   Nei passi 4–8 le **domande guidate** («Cosa vedi?») danno la correzione precisa: menu, direzione e numero di passi.
   Nero e bianco insieme riconoscono il problema del range HDMI; i **Grigi a confronto** (ring-around 3×3 su ombre,
   mezzitoni e luci) rivelano la dominante e indicano temperatura/tinta o Gain/Offset RGB. Ogni passo ha i campi
   **valori prima → dopo**, che finiscono nel report e nel **registro delle sale** (richiamabile al passo 1).
4. Passo 10: checklist finale, cosa chiedere al tecnico se non puoi toccare il proiettore, **Scarica il report (PDF)**
   e **Rimuovi le slide di test**. Il report è un PDF A4 nello stile delle guide (intestazione nera, esiti dei passi,
   risposte, valori prima → dopo, note, chiusura con l’Arena). Nella versione web si scarica subito; nel pannello di
   PowerPoint, dove i download spesso non funzionano, si apre `report.html` nel browser con i dati dopo il `#`.

In alternativa **Apri i pattern in una finestra** mostra gli stessi pattern in una finestra separata da trascinare
sul proiettore (doppio clic = tutto schermo): si adattano al formato reale dello schermo e seguono i passi del pannello.

I pattern derivano dal Screen Test The Ultraspeaker 2.1 (bordo a scacchi, mirini, triangoli, quadrati e riquadro RGB)
più: griglia, stelle di Siemens e scala dei corpi, nero 0–20%, bianco 80–99%, scala di grigi, grigi a confronto per la dominante, barre SMPTE 75%,
rampe colore a 32 gradini, bianchi dello smalto / tessuti molli / ombre del cavo orale, spazio per una foto clinica,
campi pieni bianco/grigio/nero. I pattern portano l’Arena di The Ultraspeaker piccola e semitrasparente in basso a destra, come un watermark (tranne Screen Test chiaro e scuro, Geometria e barre SMPTE).

## Requisiti

PowerPoint di Microsoft 365 (Windows, Mac o web) con PowerPointApi 1.2.
Funzioni in più quando disponibili: lettura delle forme per ritrovare le slide di test (1.3), selezione diretta della slide (1.5),
rilevamento automatico del formato (1.8), finestra dei pattern pilotata dal pannello (DialogApi 1.2).

## Pubblicazione su GitHub Pages

1. Crea su GitHub un repository pubblico chiamato **ultraspeaker-projector-lab** (account `niki977`).
2. Carica tutti i file di questa cartella (anche `assets`, `css`, `fonts`, `guide`, `shared` e `vendor`).
3. In *Settings → Pages* scegli *Deploy from a branch*, ramo `main`, cartella `/ (root)`.
4. Dopo un minuto il pannello è raggiungibile su `https://niki977.github.io/ultraspeaker-projector-lab/`
   (aperto nel browser funziona come guida, con pattern a schermo intero e download dello Screen Test .pptx).

Se usi un altro nome di repository, sostituisci `https://niki977.github.io/ultraspeaker-projector-lab/` in `projector-lab-manifest.xml`.

## Installazione per le prove (sideload)

Vedi `GUIDA-INSTALLAZIONE.html` (Mac, Windows, web). Il pulsante **Projector Lab** compare nella scheda Home,
nel gruppo «The Ultraspeaker», accanto a QR Lab e Image Lab.

## File

| File | Cosa contiene |
|---|---|
| `index.html` | Il pannello e la versione web |
| `css/app.css` | Stile proprio del Lab, caricato dopo `shared/ultraspeaker-lab.css` |
| `app.js` | Procedura, collegamento con PowerPoint (inserimento, navigazione, rimozione, foto), contenuto del report |
| `report-pdf.js` | Impaginazione del report in PDF (A4), usata dalla versione web e da `report.html` |
| `report.html` | Pagina del report aperta nel browser dal pannello di PowerPoint: anteprima e download del PDF |
| `patterns.js` | I 15 pattern, descritti una volta e disegnati in SVG (anteprime e finestra) e in PowerPoint (PptxGenJS) |
| `show.html` | Pattern a schermo intero: finestra di PowerPoint, finestra del browser o riquadro dentro il pannello |
| `i18n.js` | Testi e procedura nelle 5 lingue |
| `venn-data.js` | Immagine del cerchio RGB del Screen Test |
| `wm-data.js` | Arena di The Ultraspeaker semitrasparente (watermark dei pattern) |
| `vendor/pptxgen.bundle.js` | PptxGenJS 3.12 (licenza MIT) per creare le slide |
| `vendor/jspdf.umd.min.js` | jsPDF 4.2.1 (licenza MIT) per creare il report in PDF, caricato solo quando serve |
| `projector-lab-manifest.xml` | Manifest dell’add-in per PowerPoint |
| `assets/` | Lockup Projector Lab, Arena (chiusura delle pagine), favicon e icone PNG usate da manifest e browser; `report-logo.png` e `report-arena.png` per il PDF |
| `shared/` | Kit comune dei Lab The Ultraspeaker 1.2 (stile, lingua, guide, esempi): non si modifica qui |
| `fonts/` | Champagne & Limousines in woff2, i quattro stili del kit; in `fonts/pdf/` regular, bold e italic in TTF per il report PDF |
| `support.html`, `privacy.html`, `terms.html`, `GUIDA-INSTALLAZIONE.html` | Supporto, privacy, condizioni d’uso e guida, in 5 lingue con selettore come nel pannello |
| `legal.js` | Testi del menu e link con la lingua scelta in queste pagine |
| `guide/` | Guida d’utilizzo: PDF nelle 5 lingue, sorgenti, screenshot e `capture.mjs` |
| `store/LISTING.md` | Testi per la scheda del Marketplace |

## Guida d’utilizzo

Cinque PDF in `guide/` (`projector-lab-guide-it.pdf` … `-de.pdf`), aperti dal pulsante **Guida** accanto al selettore
lingua e linkati da Supporto e Installazione. Sorgenti in `guide/src/guide.<lingua>.html` (modello e stile del kit,
nessuno stile proprio); i nomi dei pulsanti citati sono quelli dell’interfaccia in quella lingua.

Per aggiornarla, dalla radice del repository (serve Playwright con Chromium: `npm install --no-save playwright`
e `npx playwright install chromium`; `node_modules/` è escluso dal repository):

1. `node guide/capture.mjs` rifà tutti gli screenshot in `guide/img/<lingua>/` (web 1440 px, telefono 390 px,
   pannello 350 px, scala 2×) e la posizione dei richiami (`pins.json`), con dati di esempio neutri.
2. Aggiorna i testi in `guide/src/guide.<lingua>.html` se l’interfaccia è cambiata (l’italiano è il riferimento).
3. `node shared/guide/build-guide.mjs` rigenera i cinque PDF.

## Memoria del browser

La lingua sta in `ultraspeaker.lang`, comune a tutti i Lab The Ultraspeaker. Il resto in
`ultraspeaker.projector.state` (procedura in corso), `.timer` (riscaldamento) e `.venues` (sale salvate).
Le chiavi delle versioni precedenti (`projectorlab.*`) vengono copiate e cancellate al primo avvio.
«Azzera tutto» cancella solo le chiavi del Lab.

## Versioni

- **1.1.1** (4 novembre 2026): report in PDF scaricabile al posto di «Copia il report».
- **1.1** (novembre 2026): uniformazione alla linea guida comune dei Lab 1.2: kit `shared/`, intestazione e chiusura
  comuni, pagine di servizio scure, lingua comune con riserva inglese, manifest en-US, guida d’utilizzo PDF.
- **1.0**: prima versione.

## Note per il Marketplace Microsoft

Titolare indicato in tutte le pagine di servizio: Nicholas D. Charles (The Ultraspeaker), info@theultraspeaker.com.
Servono anche screenshot del pannello 1366×768. L’add-in non invia dati a nessun server.
