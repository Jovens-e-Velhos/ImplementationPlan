# Farol PMO

A browser-based version of the `farol_pmo.py` traffic-light dashboard. Upload
an **Implementation Plan** `.xlsx` and it renders the same report the Python
script used to generate — entirely client-side, no server, no upload to
anywhere. Works great hosted for free on GitHub Pages.

## Deploy to GitHub Pages

1. Copy the three files in this folder (`index.html`, `app.js`, `README.md`)
   into the root of your repository (or into a `/docs` folder — your choice).
2. Push to GitHub.
3. In your repo: **Settings → Pages** → under "Build and deployment", set
   **Source** to "Deploy from a branch", pick the branch (e.g. `main`) and
   the folder (`/ (root)` or `/docs`), then **Save**.
4. GitHub will publish the site at
   `https://<your-username>.github.io/<repo-name>/` within a minute or two.

That's it — no build step, no dependencies to install. The page loads the
SheetJS library from a CDN to parse `.xlsx` files in the browser.

## How it works

- `index.html` — page shell, upload screen, and the dashboard's visual theme
  (same dark/light GitHub-style theme as the original report).
- `app.js` — a JavaScript port of `farol_pmo.py`'s logic: reads the
  `Implementation Plan` sheet, computes stage completion %, applies the same
  traffic-light rules (on track / delayed / overdue / finished / not
  started), and renders the report into the page.

Anyone who opens the page can drop in their own copy of the Implementation
Plan spreadsheet and get an instant dashboard — nothing is sent to a server,
so it's safe to use with confidential project data.

## Updating the report

There's no "master" data file to keep in sync — whoever opens the page just
uploads the latest `.xlsx`. If you want a single shared, always-up-to-date
link instead (where you update the spreadsheet and everyone sees the new
numbers automatically), that needs a small backend or a scheduled step that
publishes the spreadsheet to a fixed URL — let me know if you'd like that
version instead.

## Expected spreadsheet format

Same as the original script expects:
- A sheet named exactly `Implementation Plan`.
- Header row at Excel row 7 with columns `Task ID` (B), `Sub Task ID` (C),
  `Task` (D), `Responsible` (E), `Planned start Date` (F), `Planned End Date`
  (G), `Days till planned finish` (H), `Status` (I), `Remarks` (J).
- Project metadata (`Latest update:`, `Responsible: ...`) somewhere in the
  first 6 rows.
- Milestone tasks with Task ID `602` (Go-Live) and `706` (Project Closure)
  for the header date badges.
