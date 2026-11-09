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
Target rules (FGS-S, and FGS 2.0 once updated): cook time never filters recipes; participants see main meals only; a bundle needs at least one surviving main meal; ranking counts main meals; quicker bundles sort last. FGS 2.0 also has a "Current live 2.0 script" switch that reproduces today's cook-time filter for comparison.
