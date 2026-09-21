---
name: starterslabo-eval
description: >-
  Fill Starters Labo monthly evaluation Excel (evaluatiefiche /
  maandevaluatie). Copies the previous month tab, maps
  terugblik←vorige vooruitblik, drafts Dutch first-person answers, writes
  cells with openpyxl. Use when the user mentions evaluatiefiche,
  maandevaluatie, maandelijkse evaluatie, LABO evaluatie, or filling a
  YYYYMM evaluation workbook.
---

# Starters Labo — Maandelijkse evaluatiefiche

Help fill the official LABO monthly self-evaluation Excel for a
Starterslabo LABO-traject deelnemer.

Not affiliated with Starterslabo. Work only on workbooks the user provides.
Never invent personal facts; never commit filled workbooks or answer JSON
with real content to git.

## Files

| Pattern | Role |
|---------|------|
| `ev_*_YYYYMM.xlsx` (or similar export name) | Official workbook (period = month being evaluated) |
| Tab `MM_YY` (e.g. `07_26`) | One tab per calendar month |
| `Infopagina` | Contact / project metadata (rarely change) |

Dependency: `openpyxl` (see `scripts/requirements.txt`).

## Workflow

```
Progress:
- [ ] 1. Locate workbook + previous month tab
- [ ] 2. Interview / draft answers (Dutch, first person)
- [ ] 3. Confirm draft with user (unless they said "gewoon invullen")
- [ ] 4. Copy tab → new MM_YY and write cells
- [ ] 5. Print verification of key cells
```

### 1. Locate & inspect

Resolve this skill’s directory, then inspect workbooks in the **user’s**
working directory (do not hard-code machine paths):

```bash
python3 - <<'PY'
import openpyxl, glob
for p in sorted(glob.glob("ev_*.xlsx") + glob.glob("*evaluatie*.xlsx")):
    wb = openpyxl.load_workbook(p, read_only=True)
    print(p, "→", wb.sheetnames)
    wb.close()
PY
```

- Period `YYYYMM` in the filename = month being evaluated (e.g. `202607` → juli → tab `07_26`).
- **Terugblik doelstellingen** = previous tab’s **vooruitblik** (C16–C18 → C12–C14).
- Read previous tab before drafting.
- Ask the user which file/tab if several match.

### 2. Gather from user (ask what’s missing)

Required narrative (Dutch):

1. **Terugblik** (3 doelen) — behaald / deels / niet + waarom
2. **Vooruitblik** (3 doelen next month)
3. **Ondernemersidee** (helderheid, hulp, kansen)
4. **Jij als ondernemer** (work/life, grenzen, legale/traject)
5. **Positieve actie** + **vaardigheid** (vorige maand)
6. **Actie komende maand** + **vaardigheid**
7. **Co-lab** — wat meegenomen, of “heeft niet plaatsgevonden”
8. **Coach** — ja/nee; indien nee: reden; indien later: timing + agenda

Tone: first person, concrete, honest, short paragraphs. No corporate fluff.
Use the starter’s own name from `Infopagina` / `B8`, not a baked-in default.

### 3. Confirm

Show a compact draft (8 sections). Only write Excel after confirmation,
unless the user already gave final copy and asked to fill.

### 4. Write Excel

Prefer the helper script next to this `SKILL.md`:

```bash
python3 -m venv .venv && .venv/bin/pip install -r scripts/requirements.txt

.venv/bin/python scripts/fill_eval_tab.py \
  --workbook path/to/ev_NAME_YYYYMM.xlsx \
  --tab 07_26 \
  --from-tab 06_26 \
  --answers /tmp/eval_answers.json
```

Answers JSON keys = cell coords (see [reference.md](reference.md)).
Minimum: `B6`, `B7`, `C12`–`C14`, `C16`–`C18`, `C20`, `C22`, `C25`–`C26`,
`C28`–`C29`, `C32`, `E36`, `D37`, and optionally `C40`.

Or inline openpyxl: `copy_worksheet` previous month → rename → set cells →
wrap long text → `save`.

### 5. Verify

Re-open and print the filled cells. Tell the user the path + new tab name.

## Safety / privacy

- Do **not** invent facts (orders, coach meetings, co-lab). Ask or leave a clear TODO.
- Do **not** overwrite an existing month tab without asking.
- Infopagina: only touch if contact/project idea changed.
- Keep previous month tabs intact (audit trail).
- Do **not** commit `*.xlsx`, filled answer JSON, `.env`, or portal credentials.
- This skill does **not** log into the portal; no username/password needed.

## Cell map (official month tabs)

| Section | Cells |
|---------|-------|
| Maand / datum / naam | `B6`, `B7`, `B8` |
| Terugblik doelen 1–3 | `C12`, `C13`, `C14` |
| Vooruitblik doelen 1–3 | `C16`, `C17`, `C18` |
| Ondernemersidee | `C20` |
| Jij als ondernemer | `C22` |
| Afgelopen maand actie / vaardigheid | `C25`, `C26` |
| Komende maand actie / vaardigheid | `C28`, `C29` |
| Co-lab | `C32` (optional `C33`) |
| Coach ja/nee | `E36` (`ja`/`nee`) |
| Reden geen afspraak | `D37` |
| Agenda coachgesprek | `C40` (optional `C41`) |

Full labels + placeholder JSON: [reference.md](reference.md).
