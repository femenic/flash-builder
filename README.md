# ⚡ DPV Flash Update Generator

App client-side per generare flash updates bi-settimanali per la leadership DPV. Nessun backend richiesto — funziona interamente nel browser.

## Come usarla

1. Apri `index.html` nel browser (locale) oppure vai al link di deploy (vedi sotto)
2. **Step 1** — Carica il file Excel con i dati delle tabelle
3. **Step 2** — Scrivi/modifica le narrative per ogni sezione
4. **Step 3** — Rivedi la preview, copia l'HTML o scarica il file

## Deploy

### GitHub Pages

1. Pusha questo repo (o la cartella `flash-update-app/`) su GitHub
2. Vai su **Settings → Pages**
3. Seleziona il branch (es. `main`) e la cartella root (o `/flash-update-app`)
4. Il sito sarà disponibile a `https://<org>.github.io/<repo>/`

### Amazon S3 + CloudFront

```bash
# Sync dei file su un bucket S3 statico
aws s3 sync ./flash-update-app s3://YOUR-BUCKET-NAME/ --delete

# Invalida la cache CloudFront (se configurato)
aws cloudfront create-invalidation --distribution-id YOUR_DIST_ID --paths "/*"
```

### Uso locale (senza deploy)

Basta aprire `index.html` con un browser. Tutto funziona offline tranne il caricamento della libreria SheetJS (da CDN).

## Stack

- HTML / CSS / JavaScript vanilla
- [SheetJS (xlsx)](https://sheetjs.com/) per il parsing Excel — caricato da CDN
- Nessuna dipendenza server-side

## Struttura

```
flash-update-app/
├── index.html    # Pagina principale con layout e form
├── app.js        # Logica applicativa (parsing, generazione, export)
├── styles.css    # Stili UI
└── README.md     # Questo file
```

## Contribuire

Clona il repo, modifica i file, e apri una PR. Non serve `npm install` né build step.
