"""
Rebuilds dashboard/data.json from the "Extracted" tab of
"Saanich Data from Digital Versions - Extracted.xlsx" (the sibling folder's
as-built step code workbook for Saanich).

This is a separate, much smaller dataset than the NRCan Database Analysis
dashboard (bc-energy-step-code-dashboard). It exists to check whether that
dashboard, filtered to Saanich, gives a representative view of this
independently-sourced as-built data. See method.html for the column mapping
and caveats.

Unlike the sibling Victoria workbook (step_code_reports.xlsx), this workbook
has no dedicated "As-Built" columns - only "Required" and "Proposed". For a
row whose Report Stage is "As Built", the "Proposed" column holds the as-built
measured value, so that's what's used here.

Output is aggregated into (municipality, year) cells - counts, sums, and
category histograms only, no per-home records - matching the main
dashboard's own privacy approach.

Usage:
    python extract_data.py ["path to Saanich Data from Digital Versions - Extracted.xlsx"]
"""
import datetime
import json
import re
import shutil
import sys
import tempfile
from pathlib import Path

import openpyxl

SHEET_NAME = 'Extracted'
DEFAULT_WORKBOOK = Path(__file__).parent.parent / 'Saanich Data from Digital Versions - Extracted.xlsx'
OUT_PATH = Path(__file__).parent / 'data.json'

# Only as-built reports have a measured "Proposed" value worth comparing; the
# handful of other stages in this workbook are projects not yet built, or
# corrupted/shifted rows (e.g. "Type: Single Detached ...").
REPORT_STAGE_FIELD = 'Report Stage'
WANTED_REPORT_STAGE = 'As Built'

MUNICIPALITY_FIELD = 'Municipality / District'
REPORT_DATE_FIELD = 'Report Date'
BUILDING_TYPE_FIELD = 'Building Type'
STEP_ACHIEVED_FIELD = 'Step Achieved'
FLOOR_AREA_FIELD = 'Conditioned Floor Area (m2)'

# Numeric fields taken from this row's "Proposed" column (the as-built
# measured value, for As-Built-stage rows), mapped onto the same metric keys
# the main NRCan dashboard uses (see method.html).
NUMERIC_FIELDS = {
    'MEUI Proposed': 'meui',
    'TEDI Proposed': 'tedi',
    'ACH@50Pa Proposed': 'ach',
    'NLA@10Pa Proposed': 'nla',
    'NLR Proposed': 'nlr',
    FLOOR_AREA_FIELD: 'floor_area',
    # Unlike the Victoria workbook, this one has a direct % improvement over
    # the CSA F280 reference house - matching the main dashboard's
    # "Calculated MEUI % Improvement" / "Calc'd TEDI % Imp" metrics.
    '% Lower Than Ref / % Improvement Proposed': 'meui_imp',
    'Envelope % Better / Heat Loss Reduction Proposed': 'tedi_imp',
}

# GHGs are taken directly from ZC Total GHG Result. Unlike the sibling
# Victoria workbook - where this same column was unreliable because those
# reports used the 2018-era compliance report format (some rows had values
# three orders of magnitude too high, fixed there by recalculating GHGs from
# as-built electricity/natural gas consumption instead) - this workbook's
# compliance reports are all the modern format, so this column is trusted
# as-is. This workbook also has no fuel-consumption columns to recalculate
# from even if it were needed.
GHG_FIELD = 'ZC Total GHG Result'
GHG_PLAUSIBILITY_CEILING = 100000

MIN_PLAUSIBLE_YEAR = 1990
MAX_PLAUSIBLE_YEAR = datetime.date.today().year + 1
YEAR_RE = re.compile(r'(19|20)\d{2}')

MUNICIPALITY_FIXES = {
    'saanich': 'Saanich',
    'district of saanich': 'Saanich',
    'victoria': 'Victoria',
    'victoria (saanich)': 'Victoria (Saanich)',
    'district of central saanich': 'Central Saanich',
}

BUILDING_TYPE_FIXES = {
    'row house': 'Row House (non-MURB)',
    'laneway house': 'Laneway House',
    'low-rise (murb)': 'Low-Rise MURB',
    'duplex/semi-detached': 'Double/Semi-detached (non-MURB)',
    '(non-murb)': 'Unknown',
}


def to_float(raw):
    if raw is None or raw == '':
        return None
    try:
        return float(raw)
    except (TypeError, ValueError):
        return None


def normalize_municipality(raw):
    if not raw:
        return 'Unknown'
    name = ' '.join(str(raw).strip().split())
    if not name:
        return 'Unknown'
    return MUNICIPALITY_FIXES.get(name.lower(), name)


def normalize_building_type(raw):
    if not raw:
        return 'Unknown'
    name = ' '.join(str(raw).strip().split())
    if not name:
        return 'Unknown'
    return BUILDING_TYPE_FIXES.get(name.lower(), name)


def normalize_year_from_report_date(raw):
    if not raw:
        return None
    m = YEAR_RE.search(str(raw))
    if not m:
        return None
    y = int(m.group(0))
    if y < MIN_PLAUSIBLE_YEAR or y > MAX_PLAUSIBLE_YEAR:
        return None
    return y


def normalize_step_achieved(raw):
    v = to_float(raw)
    if v is None:
        return 'Unknown'
    return str(int(v)) if v == int(v) else str(v)


def load_workbook_with_lock_retry(path):
    try:
        return openpyxl.load_workbook(path, read_only=True, data_only=True)
    except PermissionError:
        tmp = Path(tempfile.gettempdir()) / f'_{path.name}'
        shutil.copy(path, tmp)
        try:
            return openpyxl.load_workbook(tmp, read_only=True, data_only=True)
        finally:
            tmp.unlink(missing_ok=True)


def main():
    wb_path = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_WORKBOOK
    print(f'Reading {wb_path.name} ...')
    wb = load_workbook_with_lock_retry(wb_path)
    ws = wb[SHEET_NAME]

    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    col = {name: i for i, name in enumerate(header) if name}

    missing = [n for n in [REPORT_STAGE_FIELD, MUNICIPALITY_FIELD, REPORT_DATE_FIELD, BUILDING_TYPE_FIELD,
                            STEP_ACHIEVED_FIELD, GHG_FIELD]
               + list(NUMERIC_FIELDS) if n not in col]
    if missing:
        raise SystemExit(f'Columns not found in header: {missing}')

    # Aggregated into (municipality, year) cells - counts/sums/category
    # histograms only, no per-home records.
    cells = {}
    total_in_sheet = 0
    total_kept = 0
    skipped_not_as_built = 0
    muni_counts = {}

    for r in rows:
        if r is None or all(v is None for v in r):
            continue
        total_in_sheet += 1

        def get(name):
            return r[col[name]]

        stage = get(REPORT_STAGE_FIELD)
        stage = stage.strip() if isinstance(stage, str) else stage
        if stage != WANTED_REPORT_STAGE:
            skipped_not_as_built += 1
            continue

        municipality = normalize_municipality(get(MUNICIPALITY_FIELD))
        year = normalize_year_from_report_date(get(REPORT_DATE_FIELD))
        year_key = year if year is not None else 'Unknown'
        building_type = normalize_building_type(get(BUILDING_TYPE_FIELD))
        step_achieved = normalize_step_achieved(get(STEP_ACHIEVED_FIELD))
        # Matches the main dashboard's two-bucket BUILDINGTYPE convention
        # ("House" vs "Multi-unit: whole building"), so the two datasets'
        # building-type mixes can be compared directly despite this sheet's
        # much more granular labels.
        house_type_group = 'Multi-unit: whole building' if 'MURB' in building_type.upper() else 'House'

        muni_counts[municipality] = muni_counts.get(municipality, 0) + 1
        total_kept += 1

        num = {}
        for field, key in NUMERIC_FIELDS.items():
            fv = to_float(get(field))
            if fv is not None:
                num[key] = fv

        floor_area = num.get('floor_area')
        ghg = to_float(get(GHG_FIELD))
        if ghg is not None and ghg > GHG_PLAUSIBILITY_CEILING:
            ghg = None
        if ghg is not None and floor_area:
            num['ghg'] = ghg
            num['ghg_intensity'] = ghg / floor_area

        cell = cells.setdefault((municipality, year_key), {
            'municipality': municipality,
            'year': year_key,
            'count': 0,
            'num': {},
            'cat': {'step_achieved': {}, 'house_type_group': {}, 'building_type': {}},
        })
        cell['count'] += 1
        for key, v in num.items():
            pair = cell['num'].setdefault(key, [0.0, 0])
            pair[0] += v
            pair[1] += 1
        for cat_key, value in (('step_achieved', step_achieved), ('house_type_group', house_type_group), ('building_type', building_type)):
            hist = cell['cat'][cat_key]
            hist[value] = hist.get(value, 0) + 1

    print(f'{total_in_sheet} data rows in sheet, {skipped_not_as_built} skipped (not "{WANTED_REPORT_STAGE}" report stage), '
          f'{total_kept} kept')
    print('Municipalities:', dict(sorted(muni_counts.items(), key=lambda kv: -kv[1])))

    out = {
        'generated_at': datetime.datetime.now().isoformat(timespec='seconds'),
        'source_file': wb_path.name,
        'total_rows': total_kept,
        'cells': list(cells.values()),
    }
    OUT_PATH.write_text(json.dumps(out, separators=(',', ':')), encoding='utf-8')
    print(f'Wrote {OUT_PATH} ({OUT_PATH.stat().st_size / 1000:.1f} KB)')


if __name__ == '__main__':
    main()
