# California New Grad Affordability Visualizer

Interactive D3 dashboard that maps how affordable each California county is for a new graduate, based on median gross rent as a share of typical monthly occupation income.

ECS 163 final project, Team 27. Oscar Pineda, Brian Le, Ayush Lenka, and Sai Mannava.

## What's in the repo

| Path | Purpose |
|------|---------|
| `index.html` | Single-page dashboard entry |
| `main.js` | All D3 logic: data load, county join, choropleth, legend, interactions |
| `style.css` | Layout and theme |
| `Data/` | Raw and preprocessed datasets (about 204 MB total) |
| `scripts/` | Python preprocessors for OEWS workbooks |
| `.gitignore` | Standard ignores |

### Data folders

- `Data/Industry_employmet&wages/` — California OEWS regional workbooks for the 22 major SOC occupation groups. Raw `.xlsx` per region plus four preprocessed JSON outputs.
- `Data/household_burden/` — Burdened-households series from FRED (sourced from the US Census ACS 5-Year). The 30% rent-to-income threshold defines a burdened household.
- `Data/median_gross_rent/` — Median gross rent by county, from US Census ACS 5-Year table B25064 (2024).

## Install

Clone:

```bash
git clone https://github.com/Oscar-Pineda524/California_NewGrad_Visualizer.git
cd California_NewGrad_Visualizer
```

That is the full client install. The dashboard runs in the browser with D3, TopoJSON, and SheetJS loaded from CDN. No `npm`, no build step.

The optional Python preprocessors only need two libraries:

```bash
pip install openpyxl xlrd
```

## Run the dashboard

Browsers block `file://` loads of local CSV, JSON, and XLSX, so the project needs a local HTTP server. Any static server works; `http.server` is the easiest:

```bash
python3 -m http.server 8000
```

Then open <http://localhost:8000> in Chrome or Safari. The dashboard loads:

1. California county boundary topology from the `us-atlas` CDN.
2. Median rent CSV for the latest year.
3. Preprocessed OEWS occupation wage JSON (and falls back to the raw `.xlsx` workbooks if a JSON is missing).
4. Household burden CSV from FRED.

Pick an occupation group from the dropdown. The map recolors each county by the ratio of median rent to typical monthly income for that occupation. Hover or click a county to open the detail panel on the right.

First paint is fast because the runtime payload is the compact JSONs (a few MB). The 200 MB of raw `.xlsx` files in `Data/` are only touched if you re-run the Python preprocessing or hit the xlsx fallback.

## Reproducibility

### Demo data is bundled

Everything needed to load and explore the dashboard sits in `Data/`. A grader cloning the repo and running `python3 -m http.server` should see a working map with no extra fetches.

### Regenerating the OEWS JSON files

To rebuild the normalized OEWS JSON from the raw workbooks, run:

```bash
python3 scripts/preprocess_oews_normalized.py
```

The script reads `Data/Industry_employmet&wages/all-oews-<year>/*.xlsx` for the latest five years and writes:

- `oews_normalized.json` — full long-form dataset
- `oews_latest.json` — most recent year, full schema
- `oews_latest_compact.json` — most recent year, browser-friendly payload
- `oews_groups_5yr_compact.json` — five-year occupation-group rollup

### Converting legacy `.xls` workbooks

The 2020 and 2021 OEWS releases shipped as legacy `.xls`. Convert them in place once:

```bash
python3 scripts/convert_legacy_oews_xls.py
```

### Original data sources

| Dataset | Source |
|---------|--------|
| OEWS occupation wages and employment | California EDD Labor Market Information ([labormarketinfo.edd.ca.gov/data/wages.html](https://labormarketinfo.edd.ca.gov/data/wages.html)) |
| Median gross rent (B25064) | US Census ACS 5-Year, 2024 ([data.census.gov](https://data.census.gov/table?q=B25064)) |
| Burdened households (5-year estimate) | FRED, Federal Reserve Bank of St. Louis ([data list 10679](https://fredaccount.stlouisfed.org/public/datalist/10679)) |
| California county geometry | us-atlas TopoJSON, FIPS prefix 06 ([topojson/us-atlas](https://github.com/topojson/us-atlas)) |

## Browser support

Tested on Chrome 120+ and Safari 17+. JavaScript must be on.

## Team

| Member | Owns |
|--------|------|
| Oscar Pineda | Data preparation, framing, storytelling direction |
| Brian Le | Choropleth, visual encoding, transitions, interactive controls |
| Ayush Lenka | Preprocessing and build pipeline, county detail views |

Final writing, review, and presentation preparation were shared by the team.
