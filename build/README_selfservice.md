# GFC feeds – self-service build (prototype, branch `selfservice-prototype`)

Everything the feeds need now comes from the Google Sheet «GFC – Об'єкти та реклама»; no local files on the Grok box.

| Was (Grok box) | Now (sheet tab) |
|---|---|
| `build_feed.py` META (town, province, coordinates, Prian district, Kyero/Prian types, new-build, pool, energy, specialtxt, features) | «Параметри» |
| `site_feed/meta.py` SITE + COVER (site id, cabinet id, type, area, license, rent/ADR/occupancy/management, «від», hidden, cover) | «Сайт» (row order = site order) |
| `site_feed/{en,es,ru,de,fr,cs}.py` (title/summary/lifestyle/invest/risk) | «Переклади» (1 row per object × language) |
| `images/manifest.json` (order, which photos) | «Фото» + files in `images/GFC-<code>/` |
| Prian/Kyero code lists (`prian_ref/*.xml`) | «Довідник» (reference only) |
| Capri/KOSMOS rows (`kronos/feed_rows/add_rows.py`, TSV) | ordinary rows in «Обєкти» / «Тексти» |
| `EXPECT = 14`, fallback numbers | removed; `guard.py` compares with what is live |

Feed data unchanged: Kyero/Prian XML and gfc_site.json come out byte-identical to the live files (except the timestamps
`feed_generated` / `generated`). `feed_state.json` (per-listing <date>) lives in `build/` and is committed by the workflow.

## Flow (`.github/workflows/build-feeds.yml`)
`fetch_sheet.py` (Apps Script web app JSON → xlsx, only public columns) → `photos.py` (photos_in/ uploads + links in «Фото»)
→ `build_feed.py` → `guard.py` (refuses on 0 listings, drop > 2 or > 25 %, empty price, missing photo file, XSD error;
warns when the price in the sheet is not the one written in a text) → `stage.py` (copy only if content changed)
→ commit to main → `deploy-pages` job.
Triggers: «Run workflow» (inputs: allow_drop, dry_run), daily 05:15 UTC, `repository_dispatch: sheet-updated`, push to `photos_in/**`.

## Going live (not done yet)
1. Import `GFC_selfservice.xlsx` tabs into the Google Sheet (or paste the new tabs) – see INSTRUKTSIIA_UA.md.
2. Owner (Mariia) installs `build/apps_script/Code.gs` as a web app (Execute as: Me, Who has access: Anyone) and gives the /exec URL.
3. Repo → Settings → Secrets and variables → Actions → New repository secret `SHEET_JSON_URL` = that URL.
4. Repo → Settings → Pages → Build and deployment → Source: **GitHub Actions** (commits made with GITHUB_TOKEN do not
   trigger a branch-based Pages build, so the workflow deploys Pages itself).
5. Merge this branch into main. First run: Actions → «Оновити фіди з таблиці» → Run workflow with dry_run = так, then for real.
6. Retire the Grok-box `publish.sh` (or keep it only as emergency fallback – it pushes to main, which is fine with step 4
   only if the workflow is re-run afterwards to deploy).

## Local test
    python3 build/simulate_apps_script.py GFC_selfservice.xlsx work/sheet.json
    python3 build/fetch_sheet.py --json work/sheet.json --out work/GFC.xlsx
    python3 build/photos.py --xlsx work/GFC.xlsx --root .
    python3 build/build_feed.py --xlsx work/GFC.xlsx --out work/out --image-base-url https://asistentkaviktorie-crypto.github.io/gfc-feeds/images
    python3 build/guard.py --new work/out --old . --root . --log work/build.log
