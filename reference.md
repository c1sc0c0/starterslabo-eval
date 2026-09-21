# Evaluatiefiche — cell reference

## Official workbook (`ev_*_YYYYMM.xlsx` or portal export)

Sheets: `Export Summary` (ignore if present), `Infopagina`, month tabs `MM_YY`.

### Infopagina (stable)

| Cell | Meaning |
|------|---------|
| `B7` | Naam star\*ter |
| `B9` | Email |
| `B10` | Telefoon |
| `B11` | Website |
| `B12` | Social |
| `B14` | Consulent contact (if filled) |
| `B15` | Startdatum traject |
| `B16` | Initieel projectidee |
| `B17` | Aangepast projectidee |
| `B24`–`C26` | Coach booking links / phone (template) |

Do not publish or commit real Infopagina values from a filled workbook.

### Month tab structure

Questions live in column C as prompts; **answers** go in the C/D/E cells below.

| # | Topic | Answer cells |
|---|-------|--------------|
| header | Maand / Datum / Naam | `B6` (e.g. `Juli`), `B7` (date), `B8` |
| 1 | Terugblik | `C12`, `C13`, `C14` |
| 2 | Vooruitblik | `C16`, `C17`, `C18` |
| 3 | Ondernemersidee | `C20` |
| 4 | Jij als ondernemer | `C22` |
| 5 | Afgelopen maand | `C25` (actie), `C26` (vaardigheid) |
| 6 | Komende maand | `C28` (actie), `C29` (vaardigheid) |
| 7 | Co-lab | `C32`, optional `C33` |
| 8 | Coach | `E36` ja/nee; `D37` reden indien nee; `C40`/`C41` gespreksdoelen |

Many answer ranges are merged `C*:Q*` — write the top-left cell only.

### Continuity rule

Previous month `C16`/`C17`/`C18` → current month terugblik `C12`/`C13`/`C14`
(status + evidence).

### Placeholder answers JSON

Use fiction placeholders only in docs/tests — never paste a real evaluation:

```json
{
  "B6": "Juli",
  "B7": "2026-08-01",
  "B8": "Alex Example",
  "C12": "Doel 1: behaald. Korte toelichting.",
  "C13": "Doel 2: deels behaald. Wat nog openstaat.",
  "C14": "Doel 3: niet behaald. Waarom.",
  "C16": "Volgende maand: actie A.",
  "C17": "Volgende maand: actie B.",
  "C18": "Volgende maand: actie C.",
  "C20": "Projectidee is helder / bijgestuurd omdat …",
  "C22": "Work/life: … Grenzen: …",
  "C25": "Positieve actie: …",
  "C26": "Vaardigheid: …",
  "C28": "Actie komende maand: …",
  "C29": "Vaardigheid om in te zetten: …",
  "C32": "Heeft niet plaatsgevonden.",
  "E36": "nee",
  "D37": "Reden, bijvoorbeeld geen beschikbaarheid deze maand.",
  "C40": "Indien later wél: wat bespreken met de coach."
}
```

`B7` may be `YYYY-MM-DD`, `DD/MM/YYYY`, or omitted (script uses today).

## Alternate / older templates

Some LABO Excel templates use denser cell layouts or extra tabs (POP,
competenties, vaardigheden). If the user’s file does not match the map above,
inspect with openpyxl and ask before writing. Do not mix cell maps.
