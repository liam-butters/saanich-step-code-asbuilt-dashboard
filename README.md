# Saanich As-Built Step Code Data

A small comparison dashboard built from `Saanich Data from Digital Versions
- Extracted.xlsx`, an independently-collected set of as-built BC Energy Step
Code reports for Saanich. It compares that data against the main
[BC Energy Step Code Dashboard](https://github.com/liam-butters/bc-energy-step-code-dashboard)
(built from the NRCan evaluation database), filtered to Saanich, to check
whether that larger dashboard gives a representative view of this
separately-sourced data.

**Live site:** published via GitHub Pages from this repo (see repo settings for the URL).

## What's here

- `extract_data.py` — reads the "Extracted" tab of the source workbook,
  keeps only "As Built" rows (using each row's "Proposed" column as the
  as-built value, since this workbook has no dedicated As-Built columns),
  and aggregates them into (municipality, year) cells with
  counts/sums/category histograms — the same metrics the main dashboard
  uses. No per-home records are shipped, matching the main dashboard's own
  privacy approach.
- `build_baseline.py` — reads the main dashboard's own `data.json` and
  computes a matching Saanich-filtered baseline (`baseline.json`) to
  compare against (using its "entry" year mode, since this workbook has no
  year-built column, only a report date).
- `data.json` / `baseline.json` — the aggregated datasets the page reads.
- `index.html` — the dashboard itself: a static page (D3.js + vanilla JS)
  with a comparison table, year-trend charts, category-mix charts, and a
  methodology section documenting the column mapping and caveats (including
  metrics with no equivalent in this workbook).

## Refreshing the data

```bash
python extract_data.py "path to Saanich Data from Digital Versions - Extracted.xlsx"
python build_baseline.py "path to the main dashboard's data.json"
```

Commit and push `data.json` / `baseline.json` to update the published site.

## Privacy

`data.json` contains only counts and sums aggregated by municipality and
year — no addresses, names, or other record-level identifiers.
`baseline.json` is derived the same way from the main dashboard's own
already-aggregated data.
