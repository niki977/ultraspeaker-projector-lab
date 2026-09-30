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
4. Passo 10: checklist finale, cosa chiedere al tecnico se non puoi toccare il proiettore, report copiabile
   e **Rimuovi le slide di test**.

In alternativa **Apri i pattern in una finestra** mostra gli stessi pattern in una finestra separata da trascinare
sul proiettore (doppio clic = tutto schermo): si adattano al formato reale dello schermo e seguono i passi del pannello.

I pattern derivano dal Screen Test The Ultraspeaker 2.1 (bordo a scacchi, mirini, triangoli, quadrati e riquadro RGB)
più: griglia, stelle di Siemens e scala dei corpi, nero 0–20%, bianco 80–99%, scala di grigi, grigi a confronto per la dominante, barre SMPTE 75%,
rampe colore a 32 gradini, bianchi dello smalto / tessuti molli / ombre del cavo orale, spazio per una foto clinica,
campi pieni bianco/grigio/nero. Ogni pattern porta il logo The Ultraspeaker piccolo e semitrasparente, come un watermark.

## Requisiti

PowerPoint di Microsoft 365 (Windows, Mac o web) con PowerPointApi 1.2.
Funzioni in più quando disponibili: etichette sulle slide (1.3, anche per ritrovare le slide per nome), selezione diretta della slide (1.5),
rilevamento automatico del formato (1.8), finestra dei pattern pilotata dal pannello (DialogApi 1.2).

## Pubblicazione su GitHub Pages

1. Crea su GitHub un repository pubblico chiamato **ultraspeaker-projector-lab** (account `niki977`).
2. Carica tutti i file di questa cartella (anche `assets`, `fonts` e `vendor`).
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
| `index.html` | Il pannello (struttura e stile) |
| `app.js` | Procedura, collegamento con PowerPoint (inserimento, navigazione, rimozione, foto), report |
| `patterns.js` | I 14 pattern, descritti una volta e disegnati in SVG (anteprime e finestra) e in PowerPoint (PptxGenJS) |
| `show.html` | Finestra dei pattern a schermo intero |
| `i18n.js` | Testi e procedura nelle 5 lingue |
| `venn-data.js` | Immagine del cerchio RGB del Screen Test |
| `wm-data.js` | Logo The Ultraspeaker semitrasparente (watermark dei pattern) |
| `vendor/pptxgen.bundle.js` | PptxGenJS 3.12 (licenza MIT) per creare le slide |
| `projector-lab-manifest.xml` | Manifest dell’add-in per PowerPoint |
| `assets/` | Lockup Projector Lab (bianco/nero), logetto mirino, logo The Ultraspeaker, icone PNG (16–512 px) |
| `fonts/` | Quicksand in woff2 e licenza OFL |
| `support.html`, `privacy.html`, `terms.html`, `GUIDA-INSTALLAZIONE.html` | Supporto, privacy, condizioni d’uso e guida, in 5 lingue con selettore come nel pannello |
| `legal.css`, `legal.js` | Stile e cambio lingua di queste pagine |
| `store/LISTING.md` | Testi per la scheda del Marketplace |

## Note per il Marketplace Microsoft

Titolare indicato in privacy e condizioni d’uso: Nicholas D. Charles.
Servono anche screenshot del pannello 1366×768. L’add-in non invia dati a nessun server.
