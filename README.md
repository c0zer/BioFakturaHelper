# BioFakturaHelper

Browser-based tools for invoice and delivery data processing. Hosted on [GitHub Pages](https://c0zer.github.io/BioFakturaHelper/).

## Tools

### 📋 Fakturagranskning (`index.html`)
Upload a `.xlsx` invoice export file to check for ellipsis issues and time-based analysis.

### 📍 Mätplats-berikare (`matplats-berikare.html`)
Upload a Kvalitetsorder Excel file and a semicolon-delimited MPS delivery CSV to automatically fill in empty `Mätplats` and `Mätplatsnamn` values. Outputs an enriched CSV ready for Excel.

Mätplatsnamn data is embedded at build time from the master data register (see below).

## Updating Mätplatser master data

When `PROD_Mätplatser_*.xlsx` is updated in `MasterdataRegister/`, re-embed the data:

```bash
node scripts/update-matplatser.js
```

Then commit and push the updated `matplats-berikare.html`.

## Local setup

```bash
npm install
```

Only needed to run the update script above — the web app itself has no build step.
