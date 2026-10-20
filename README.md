# Habit Tracker

Tick your habits in one tap, see your streaks and look back once a week, without an account and without your data leaving your phone.

**Use it free:** https://hmsk.life/tools/habit-tracker/ (or straight to the app: https://hmsk.life/apps/habit-tracker/)

![Habit Tracker](https://hmsk.life/assets/images/tools/habit-tracker.png)

## What it does

- Habits with a name, an optional emoji and a target: every day, or N days a week.
- Today view: one tap (or keys 1 to 9) ticks a habit. Current and best streak per habit; weekly targets count weeks met.
- Month grid (habits × days): tap any past day to fix it. Completion % this week and this month, a 12 week heatmap per habit.
- Weekly review: the week at a glance plus a "what got in the way?" note per week.
- Data stays in the browser (localStorage). JSON backup and restore, CSV export, and Excel (.xlsx) export of your real data,
  written in the browser by a small zip/OOXML writer (`public/xlsx.js`, no library).
- English and German (`?lang=de`), works on phones, add it to your home screen.
- A separate yearly spreadsheet (`tools/make_xlsx.py`) for people who prefer Excel, LibreOffice or Google Sheets.

## Quick start

It is plain static files, no build step and no dependencies.

```sh
python3 -m http.server 8080 --directory public   # or: npm start
# open http://localhost:8080
```

Any static web server works (nginx, Apache, GitHub Pages, Netlify). `public/config.js` sets where the links back to
hmsk.life point; set `site: ''` if you host it next to your own site.

Fonts (Newsreader, Public Sans; SIL Open Font License) are in `public/fonts/`; the shared hmsk.life look is in
`public/letterpress.css`.

Demo data: open `/?demo=1`. It is generated in memory and never touches what you stored.

## The yearly spreadsheet

```sh
pip install openpyxl
python3 tools/make_xlsx.py habit-tracker.xlsx --year 2026
```

One Setup sheet (the year and up to 10 habits with days per week, typed once), an Overview sheet (% per month, year %,
current and best streak) and one sheet per month with tick cells (x or ✓), done/goal/% per habit, daily %, a streak that
carries over between months, blue colour scales (hmsk.life Letterpress colours), frozen headers and A4 landscape print setup. Change the year on
Setup and every month follows. It only uses functions that exist in Excel, LibreOffice and Google Sheets.

The ready file is at https://hmsk.life/assets/downloads/habit-tracker.xlsx.

## Tests

```sh
node --test tests/     # or: npm test
```

Covers dates (local days, DST change days, ISO weeks), streaks and targets, completion %, schema migration and the
zip/xlsx writer (CRC-32 vectors, zip structure, and the exported file opened with python openpyxl when available).

## Data and privacy

Everything is stored in `localStorage` under the key `hmsk-habit-tracker`, as versioned JSON (`v: 2`). Older backups are
migrated on import. Nothing is sent anywhere: no analytics, no cookies, no third-party requests.

## License

MIT, see [LICENSE](LICENSE). Made by Hafiz Muhammad Sharoz Khalid, https://hmsk.life
