# Ready Set Go scoring

Live scoring for Ready Set Go. The pages are hosted on GitHub Pages; the data stays in the Google Sheet, with `Code.gs` (Apps Script) as the backend.

**Live site**
- Scoring page: [utalpfa.github.io/rsg-scoring-web](https://utalpfa.github.io/rsg-scoring-web/)
- Admin page: [utalpfa.github.io/rsg-scoring-web/admin.html](https://utalpfa.github.io/rsg-scoring-web/admin.html)

```
docs/          GitHub Pages site (scoring page, admin page, api.js)
apps-script/   Code.gs + the original Apps Script HTML pages
```

## Why GitHub Pages

Apps Script web apps break for visitors signed in to a school Google Workspace account or to several Google accounts at once. The hosted pages call the web app with plain anonymous requests (`docs/api.js` stands in for `google.script.run`), so the visitor's Google sign-in never gets involved.

## Backend setup

1. Paste `apps-script/Code.gs` into the sheet's Apps Script project (Extensions > Apps Script).
2. Deploy > Manage deployments > edit the existing web app > Version: **New version** > Deploy. Keep Execute as: Me and Access: Anyone. Editing the existing deployment keeps the `/exec` URL that `docs/api.js` points to; a brand-new deployment gets a new URL, which then has to go in `RSG_API` in `docs/api.js`.
3. First time only: run `setup`, fill the `Students` and `Professionals` tabs, and set the admin passcode (Ready Set Go menu > Set admin passcode).

`doPost` only runs the functions listed in `API` in `Code.gs`. The old `/exec` page still works too.

## Rules

Ranking: total desc, then highest single boost, then average. Leaderboard and Live Monitor update on every submission.
Corrections: professionals can undo within 60 s; otherwise void the score on the admin page or set `Voided` = Y in `Scores`.
