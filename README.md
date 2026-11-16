# FGS Recipe Bundle Simulator

A static web page (no server, no build step) that shows what a participant sees in the Family Grocery Study baseline recipe ranking, for **FGS-S**, **FGS 2.0 – 3 meal** and **FGS 2.0 – 4 meal**. It reads the bundle spreadsheets in `data/` and runs the same filtering and ranking logic as the REDCap JS Injector scripts.

## Publish on GitHub Pages
1. Create a repository and upload everything in this folder (keep the folder structure).
2. Settings → Pages → Build from branch → `main` / root.
3. Open `https://<your-account>.github.io/<repo>/`.

> Pages from a **public** repo are public, and so are the spreadsheets in `data/`. Use a private repo (Pages on private repos needs a paid GitHub plan) or an internal host if the recipes should not be public. The FGS-S file here contains only the two sheets the page needs (the original workbook also has costing sheets – keep those out of the repo).

## Update the bundles
1. Replace the file in `data/` (same file name, same sheet name, same columns: recipe name in column A, `Ingredients`, `Tags`, `Prep/Cook Time`).
2. Commit. The page picks it up on reload.
3. Preview before committing: **Try a different .xlsx…** at the top of the page.

Settings per project (file, sheet, number of main meals, fallback bundle labels) are in `config.json`. FGS-S reads the sheet `Final (Canva Info)` and takes bundle labels from `Final Bundles`.

## Keep the REDCap scripts in step
* **Data & export → Copy var BUNDLES line** gives the dataset the REDCap scripts need, built from the sheet with all patches applied.
* **Checks & compare → choose injector .js** lists every difference between the sheet and the data currently inside a REDCap script.

## Files
* `index.html`, `style.css`, `app.js` – the page.
* `engine.js` – spreadsheet parsing + filter/ranking logic (a port of the injector scripts).
* `data/*.xlsx` – bundle spreadsheets. `data/patches/*.json` – reviewed corrections applied on top of the sheet (add-only: extra allergens, vegetarian/vegan upgrade, protein bucket).
* `lib/xlsx.full.min.js` – SheetJS 0.18.5 (reads the .xlsx in the browser).

## Rules implemented
Current rules (FGS-S and FGS 2.0 3-meal / 4-meal): cook time never filters recipes; participants see main meals only; a bundle is offered only if ALL its main meals fit the family (3 main meals for FGS-S and the 3-meal survey, 4 for the 4-meal survey); no partial bundles or cross-bundling; quicker bundles sort last. The Rules switch also offers the earlier per-recipe rule and, for FGS 2.0, the old cook-time-filter script, for comparison.

## Test the real REDCap script without REDCap
Open `redcap-test/index.html` (same site, same local server). Choose the injector `.js` file, then use the fake survey page like a participant (page 1 = consent + cook time, page 2 = recipe ranking) or press **Run checks**: it fills the form with presets and random families, lets your unmodified script react, and checks what it drew and what it wrote to `bundle_rank_1..6` against the bundle data inside the script. The tester detects from the file you load whether it is FGS-S, FGS 2.0 3-meal or FGS 2.0 4-meal (field suffix, number of main meals, helper name) and builds the matching fake page. It cannot catch differences between this mock markup and live REDCap pages, so still do one real-page check after a change.