# TMJ — Quiet Tutor Finder

Finds tutors who have gone **quiet** (no session in three weeks or more) and hands
a person a list they can act on. Built for the TMJ Tutoring dev challenge.

Export is treated as current on **Friday 29 May 2026**.

## What's here

| File | What it is |
|------|------------|
| `quiet_tutors.py` | The script. Reads the two CSVs, matches sessions to the roster, writes the report. |
| `index.html` | A browser version — import the two CSVs and get the same result, no install. |
| `quiet_tutors_report.xlsx` | Example output: one workbook, two tabs (*Quiet Tutors*, *Needs Review*). |
| `tutors.csv` / `sessions.csv` | The input data. |

## Run the script

```bash
pip install openpyxl
python quiet_tutors.py
```

Writes `quiet_tutors_report.xlsx` next to the CSVs.

## Use the web version

Open `index.html` in any browser, or visit the live site (see below). Drop in
`tutors.csv` and `sessions.csv`, pick the "today" date and the quiet window,
and click **Find quiet tutors**. Everything runs locally in the browser —
nothing is uploaded.

## How it works

- **Quiet** = no session in ≥ 21 days, measured from the export date.
- Sessions are matched to the roster by name: exact match first; abbreviated
  names (`J. Smith`, `Sarah L.`) are resolved by initial + surname and
  disambiguated by **subject** (Physics → John Smith; Biology → Sarah Leung).
- A session name with no confident roster match (e.g. `Kevin Tran`) is surfaced
  on a **Needs Review** sheet rather than dropped — every quiet tutor is
  identified by `tutor_id`.

## Live site (GitHub Pages)

After pushing this repo to GitHub:

1. Repo **Settings → Pages**.
2. **Build and deployment → Source: Deploy from a branch**.
3. Branch: **main**, folder: **/ (root)**, then **Save**.
4. After a minute the site is live at
   `https://<your-username>.github.io/<repo-name>/`.
