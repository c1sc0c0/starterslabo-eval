# Starters Labo evaluatiefiche skill

Unofficial [Cursor](https://cursor.com) / [Claude](https://claude.ai) skill that
helps fill the monthly LABO **evaluatiefiche** (maandevaluatie) Excel:
interview → Dutch draft → copy previous month tab → write cells with openpyxl.

Not affiliated with Starterslabo. You bring your own workbook from the portal.
This skill does **not** log in and needs **no** username/password.

## Install

Clone this repo **as the skill folder** (it must contain `SKILL.md` at the root):

```bash
# Cursor — this project
git clone https://github.com/c1sc0c0/starterslabo-eval.git .cursor/skills/starterslabo-eval

# Cursor — all projects
git clone https://github.com/c1sc0c0/starterslabo-eval.git ~/.cursor/skills/starterslabo-eval

# Claude Code — this project
git clone https://github.com/c1sc0c0/starterslabo-eval.git .claude/skills/starterslabo-eval

# Claude Code — all projects
git clone https://github.com/c1sc0c0/starterslabo-eval.git ~/.claude/skills/starterslabo-eval
```

Restart the agent chat (or Cursor/Claude Code) so it picks up the skill.

**Claude.ai:** zip this folder and upload it as a custom skill.

## Use

Put your portal export (e.g. `ev_NAME_YYYYMM.xlsx`) in the project, then tell
the agent something like:

> Vul de evaluatiefiche voor 202608. Vorige maand stond in tab 07_26.

The agent will:

1. Inspect sheets / previous-month vooruitblik
2. Ask for missing answers (terugblik, vooruitblik, co-lab, coach, …)
3. Show a short draft for confirmation
4. Run `scripts/fill_eval_tab.py` (or equivalent openpyxl) to add tab `08_26`

## Manual fill

```bash
cd /path/to/starterslabo-eval
python3 -m venv .venv
.venv/bin/pip install -r scripts/requirements.txt

.venv/bin/python scripts/fill_eval_tab.py \
  --workbook /path/to/your-export.xlsx \
  --tab 08_26 \
  --from-tab 07_26 \
  --answers ./answers.json
```

See `reference.md` for the cell map and a **placeholder** JSON shape.
Never commit real filled workbooks or real answer files.

## Privacy

- No credentials in this repo (none required).
- `.gitignore` excludes `*.xlsx`, `answers*.json`, `.env`.
- Do not paste personal evaluation text into public issues/PRs.

## Related

Index: [starterslabo-skills](https://github.com/c1sc0c0/starterslabo-skills).

- [starterslabo-faq](https://github.com/c1sc0c0/starterslabo-faq) — crawl portal FAQ (login required)
- [starterslabo-km](https://github.com/c1sc0c0/starterslabo-km) — kilometervergoeding Excel

## License

MIT for this skill’s code and instructions. Starterslabo’s website, portal,
and evaluation templates remain theirs.
