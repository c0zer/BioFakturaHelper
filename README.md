# Faktura Flow – Ellips + Pilleps Office Scripts

## Arkitektur

```
Ny fakturafil (.xlsx) laddas upp i triggermapp (SharePoint)
        │
        ▼ Power Automate triggas
        │
        ▼ Action 1: Kör ellips-granskning på den uppladdade filen
        │       └── returnerar JSON (under12, over100, fakturaBestrida)
        ▼ Action 2: Kör las-referensdata på Aktörsregister V2.xlsx
        │       └── returnerar JSON med aktörslista (Hkod, IntNr, Kundnr, Namn)
        ▼ Action 3: Kör pilleps-aktorsanalys på den uppladdade filen
        │       └── parameter: referensDataJson från Action 2
        │       └── returnerar JSON (huvudtabell, kritiskLista, summering)
        ▼ Action 4: Kopiera ResultatMall.xlsx → ny fil med datumnamn
        │       └── spara fil-ID för den nya kopian
        ▼ Action 5: Kör skriv-resultat på den nya kopian
                └── parametrar: ellipsJson + pillepsJson + originFilnamn
                └── skapar ark Ellips-flaggade / Pilleps-analys / Pilleps-kritisk
```

> **Nyckelprincip:** Scriptet körs direkt på den uppladdade fakturafilen.
> Inga tabeller behöver skapas. Inga parametrar skickas.
> Scriptet läser all data automatiskt via `getUsedRange()`.

---

## Steg 1 – Spara alla fyra scripts i ditt Office Scripts-bibliotek

Varje script sparas EN GÅNG i OneDrive → Dokument → Office-skript och är sedan tillgängligt mot valfri fil.

| Fil | Scriptnamn |
|-----|------------|
| `scripts/ellips-granskning.ts` | `ellips-granskning` |
| `scripts/las-referensdata.ts` | `las-referensdata` |
| `scripts/pilleps-aktorsanalys.ts` | `pilleps-aktorsanalys` |
| `scripts/skriv-resultat.ts` | `skriv-resultat` |

**För varje script:**
1. Öppna valfri Excel Online-fil → **Automatisera** → **Ny script**
2. Klistra in innehållet från respektive `.ts`-fil
3. Namnge scriptet exakt som i tabellen ovan
4. Klicka **Spara**

---

## Steg 2 – Skapa ResultatMall.xlsx på SharePoint

Skapa en **tom Excel-fil** på SharePoint i en mapp som heter t.ex. `Resultat`.
Filen behöver inte innehålla något – `skriv-resultat`-skriptet skapar arken automatiskt.

Notera sökvägen till mappen (används i Action 4).

---

## Steg 3 – Kontrollera Aktörsregister V2.xlsx

Se till att `Aktörsregister V2.xlsx` (eller motsvarande fil) finns på SharePoint och att du
känner till bibliotek + sökväg. Filen behöver ingen speciell tabellformatering.

---

## Steg 4 – Power Automate-flöde

### Trigger
- **Typ:** SharePoint → *When a file is created in a folder*
- **Site Address:** [Din SharePoint-sajt]
- **Library Name:** [Dokumentbiblioteket]
- **Folder:** [Triggermappen där fakturafilerna laddas upp]

---

### Action 1 – Kör Ellips-script (på den uppladdade filen)
- **Connector:** Excel Online (Business) → **Run script**
  - **Location:** SharePoint
  - **Document Library:** [Biblioteket]
  - **File:** `Identifierare` *(dynamiskt innehåll från trigger)*
  - **Script:** `ellips-granskning`
- **Spara output:** döp outputvariabeln till `ellipsResult`

---

### Action 2 – Kör las-referensdata (på Aktörsregistret)
- **Connector:** Excel Online (Business) → **Run script**
  - **Location:** SharePoint
  - **Document Library:** [Biblioteket]
  - **File:** *(välj Aktörsregister V2.xlsx direkt – statisk filreferens)*
  - **Script:** `las-referensdata`
- **Spara output:** döp outputvariabeln till `referensResult`

---

### Action 3 – Kör pilleps-aktorsanalys (på den uppladdade filen)
- **Connector:** Excel Online (Business) → **Run script**
  - **Location:** SharePoint
  - **Document Library:** [Biblioteket]
  - **File:** `Identifierare` *(dynamiskt innehåll från trigger)*
  - **Script:** `pilleps-aktorsanalys`
  - **referensDataJson:** `body/result` från Action 2, serialiserat som JSON-sträng:
    ```
    @{string(outputs('Action_2')?['body/result'])}
    ```
    *(alternativt: använd "JSON till text"-steget om expression-syntaxen inte fungerar)*
- **Spara output:** döp outputvariabeln till `pillepsResult`

---

### Action 4 – Kopiera ResultatMall.xlsx
- **Connector:** SharePoint → **Copy file**
  - **Current Site Address:** [Din SharePoint-sajt]
  - **File Path:** `/Sökväg/Till/Resultat/ResultatMall.xlsx`
  - **Destination Site Address:** [Din SharePoint-sajt]
  - **Destination Folder Path:** `/Sökväg/Till/Resultat/`
  - **New Name:** `Resultat_@{formatDateTime(utcNow(), 'yyyyMMdd_HHmm')}.xlsx`
- **Spara output:** döp outputvariabeln till `kopiResult`
  - Notera att `Id` från detta steg används i Action 5

---

### Action 5 – Kör skriv-resultat (på den nya kopian)
- **Connector:** Excel Online (Business) → **Run script**
  - **Location:** SharePoint
  - **Document Library:** [Biblioteket]
  - **File:** `Id` *(dynamiskt innehåll från Action 4 – kopiResult)*
  - **Script:** `skriv-resultat`
  - **ellipsJson:**
    ```
    @{string(outputs('Action_1')?['body/result'])}
    ```
  - **pillepsJson:**
    ```
    @{string(outputs('Action_3')?['body/result'])}
    ```
  - **originFilnamn:** `Filnamn` *(dynamiskt innehåll från trigger)*

---

## Scripts – översikt

| Script | Fil | Syfte | Parametrar |
|--------|-----|-------|------------|
| `ellips-granskning` | `scripts/ellips-granskning.ts` | Flaggar ärenden med avvikande tid/debitering | Inga |
| `las-referensdata` | `scripts/las-referensdata.ts` | Läser Aktörsregistret | Inga |
| `pilleps-aktorsanalys` | `scripts/pilleps-aktorsanalys.ts` | Identifierar aktörer utan kundnr/fakturakonto | `referensDataJson: string` |
| `skriv-resultat` | `scripts/skriv-resultat.ts` | Skriver resultat till tre ark i Excel | `ellipsJson`, `pillepsJson`, `originFilnamn` |

---

## Kolumnmappning – fakturafilen

| Excel-kolumn | Rubrik i filen                          | Scriptnyckel            |
|-------------|------------------------------------------|-------------------------|
| B           | `Kontonamn`                              | `kontonamn` (Pilleps)   |
| C           | `Ärendenummer`                           | `arendenummer` (Ellips) |
| D           | `Ärenderubrik`                           | `arenderubrik` (Ellips) |
| G           | `Antal minuter fakturering`              | `tid`                   |
| H           | `Debiteringsenheter för fakturering`     | `debiteringsenhet`      |
| K           | `Fakt. Hkod+Intnr`                       | `faktHkodIntnr` (Pilleps matchnyckel) |

## Kolumnmappning – Aktörsregister V2.xlsx

| Excel-kolumn | Rubrik | Beskrivning |
|-------------|--------|-------------|
| A           | `Hkod` | Aktörskod (5-siffrig med ledande nollor, t.ex. "00019") |
| B           | `IntNr` | Internt nummer (5-siffrig med ledande nollor) |
| C           | `Kundnr` | Kundnummer |
| E           | `Namn` | Aktörens namn |

**Matchningsnyckel:** `parseInt(Hkod).toString() + IntNr` = t.ex. `"19" + "00000"` = `"1900000"`
Matchar mot kolumn K (`Fakt. Hkod+Intnr`) i fakturafilen.

Rubrikrad ligger på **rad 3** (rad 1–2 är tomt/titel). `las-referensdata` hittar den automatiskt.

---

## Output-format

### Ellips (JSON från `ellips-granskning`)

```json
{
  "under12": [{ "arendenummer": "12345", "arenderubrik": "...", "tid": 8 }],
  "over100": [{ "arendenummer": "67890", "arenderubrik": "...", "tid": 145 }],
  "fakturaBestrida": [{ "arenderubrik": "Fakturering av tjänst", "arendenummer": "11111" }],
  "debug": { "sheetName": "...", "rowCount": 100 }
}
```

### Pilleps (JSON från `pilleps-aktorsanalys`)

```json
{
  "huvudtabell": [{
    "kontonamn": "SCA Skog AB",
    "faktHkodIntnr": "1900000",
    "matchtyp": "Har både kundnummer och fakturakonto",
    "primarIdentifierare": "1900000",
    "kundnummer": "42023556",
    "fakturakonto": "1900000",
    "kalla": "Aktörsregister",
    "kommentar": ""
  }],
  "kritiskLista": [],
  "summering": {
    "totalt": 150,
    "harBada": 148,
    "harEndastKundnr": 0,
    "harEndastFakturakonto": 2,
    "saknarBada": 0
  }
}
```

---

## Ellips-regler

| Regel | Logik |
|-------|-------|
| Ärenden < 12 min | `tid < 12` → läggs i `under12` |
| Ärenden > 100 min | `tid > 100` → läggs i `over100` |
| Debiteringsenhet > 8,3 | `debiteringsenhet > 8.3` → läggs i `over100` (kontrollregel) |
| Faktura/bestrida-ärenden | Kolumn D matchar `/bestrid\|bestreda\|faktur/i` → läggs i `fakturaBestrida` |

## Pilleps-klassificering

| Klassificering | Innebär |
|----------------|---------|
| Har både kundnummer och fakturakonto | Hittades i Aktörsregister + Fakt.Hkod+IntNr är ifyllt |
| Har kundnummer men saknar fakturakonto | Hittades i register men kol K är tom |
| Har fakturakonto men saknar kundnummer | Kol K ifyllt men ingen match i register |
| Saknar BÅDA | Kol K tom + ingen matchning → **kritisk lista** |

---

## Öppna frågor

Se `.omcp/plans/open-questions.md` för utestående beslut.
