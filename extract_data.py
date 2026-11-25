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
# A row can represent a whole block of identical attached units (e.g. a row
# of townhomes modeled once in HOT2000 and applied to all of them) rather
# than a single house - see Dwelling Units below.
DWELLING_UNITS_FIELD = 'Dwelling Units'

# Conditioned Floor Area and GHGs are reported as totals for the whole row
# (i.e. for all of its Dwelling Units combined), not per-home values - e.g. a
# 6-unit rowhouse block's floor area is the sum of all 6 units' areas, and
# its GHGs are the combined total for all 6. Every other metric (MEUI, TEDI,
# ACH, NLA, NLR, GHG intensity, % improvements) is already a per-home
# intensity/ratio, since it comes from the single HOT2000 model that stands
# for every identical unit in the block. GHGs is divided by Dwelling Units
# unconditionally - same as floor area - so "GHGs" here always means kg
# CO2e per dwelling unit, per year, matching the main dashboard's own
# per-dwelling-unit convention (see method.html).
BLOCK_TOTAL_KEYS = {'floor_area', 'ghg'}

# MEUI, TEDI and GHG intensity are all per-floor-area values, so a smaller
# home shows a higher per-m2 figure than a larger one even with identical
# underlying performance (building heat loss/use has a floor-area-independent
# share, diluted less in a smaller home). To make this spreadsheet's average
# comparable to the main dashboard's despite its very different mix of home
# sizes, fit intensity = a + b/floor_area per metric (weighted least squares,
# weighted by Dwelling Units) from this spreadsheet's own homes, then report
# what that fit predicts at the main dashboard's own average floor area - see
# build_fit()/solve_fit() and method.html.
FIT_METRIC_KEYS = ('meui', 'tedi', 'ghg_intensity')


def new_fit_accumulator():
    return {k: [0.0, 0.0, 0.0, 0.0, 0.0, 0] for k in FIT_METRIC_KEYS}


def accumulate_fit(acc, num, floor_area, units):
    if not floor_area:
        return
    per_unit_area = floor_area / units
    x = 1.0 / per_unit_area
    for key in FIT_METRIC_KEYS:
        y = num.get(key)
        if y is None:
            continue
        a = acc[key]
        a[0] += units
        a[1] += units * x
        a[2] += units * y
        a[3] += units * x * x
        a[4] += units * x * y
        a[5] += 1


def solve_fit(acc):
    out = {}
    for key in FIT_METRIC_KEYS:
        sw, swx, swy, swxx, swxy, n = acc[key]
        if n < 5 or sw <= 0:
            continue
        denom = sw * swxx - swx * swx
        if abs(denom) < 1e-9:
            out[key] = {'a': swy / sw, 'b': 0.0, 'n': n}
        else:
            b = (sw * swxy - swx * swy) / denom
            a = (swy - b * swx) / sw
            out[key] = {'a': a, 'b': b, 'n': n}
    return out

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

# These columns use 0 to mean "not measured/not applicable" rather than
# leaving the cell blank in practice (a real home never has literally zero
# airtightness, floor area, or energy use) - treat 0 as missing for these keys
# so it doesn't drag an average/median toward zero. Deliberately excludes
# meui_imp/tedi_imp: 0% (or negative) improvement over the reference house is
# a legitimate, real value there, not a sentinel.
ZERO_AS_MISSING = {'meui', 'tedi', 'ach', 'nla', 'nlr', 'floor_area'}

# The main NRCan dashboard's own Victoria/Saanich data tops out at 2 Dwelling
# Units per row (it has no large apartment/rowhouse blocks at all) - this
# workbook, by contrast, includes real 4-12 unit MURB/rowhouse buildings.
# Those are modern and all-electric, so their per-unit GHGs is small, and
# because every metric is weighted by Dwelling Units, a handful of such rows
# inject many repeated low values into the comparison (especially the
# median), well beyond anything the NRCan side's own population could ever
# produce. Excluded entirely (every metric, not just GHGs) so both sides
# cover the same building-size range - see method.html.
MAX_DWELLING_UNITS = 2

# NLR is a much less consistently reported airtightness metric than the
# others - requiring it alongside everything else would throw out otherwise-
# complete rows just because of one rarely-filled field. Treated as an edge
# case: a row missing NLR alone still counts for every other metric (it just
# doesn't contribute to the NLR average itself, same as before). But a row
# missing any OTHER metric actually shown on this dashboard is dropped
# entirely, from every metric, not just the one it lacks - so every metric
# here (other than NLR) is computed from the exact same set of rows, rather
# than each silently drawing from a different subset. meui_imp/tedi_imp are
# extracted but never shown on this dashboard (see method.html), so they're
# deliberately left out of this requirement too.
REQUIRED_METRICS = ['meui', 'tedi', 'ach', 'nla', 'floor_area', 'ghg', 'ghg_intensity']

# GHGs are taken directly from ZC Total GHG Result. Unlike the sibling
# Victoria workbook - where this same column was unreliable because those
# reports used the 2018-era compliance report format (some rows had values
# three orders of magnitude too high, fixed there by recalculating GHGs from
# as-built electricity/natural gas consumption instead) - this workbook's
# compliance reports are all the modern format, so this column is trusted
# as-is. This workbook also has no fuel-consumption columns to recalculate
# from even if it were needed.
GHG_FIELD = 'ZC Total GHG Result'
# Whenever one of GHG_FIELD/GHGI_FIELD is blank (or zero - the Zero Carbon
# framework appears to use 0 as a "not computed" sentinel on some rows, same
# as several other columns in this workbook) but the other is a usable
# number, the missing one is derived from it via Conditioned Floor Area
# instead of the row losing both values - see the GHG/GHGI derivation below.
GHGI_FIELD = 'ZC GHG per m2 Result'
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
                            STEP_ACHIEVED_FIELD, GHG_FIELD, GHGI_FIELD, DWELLING_UNITS_FIELD]
               + list(NUMERIC_FIELDS) if n not in col]
    if missing:
        raise SystemExit(f'Columns not found in header: {missing}')

    # Aggregated into (municipality, year) cells - counts/sums/category
    # histograms only, no per-home records.
    cells = {}
    total_in_sheet = 0
    total_kept = 0
    skipped_not_as_built = 0
    skipped_large_murb = 0
    skipped_incomplete = 0
    muni_counts = {}
    fit_acc = new_fit_accumulator()

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

        # Default to 1 unit when Dwelling Units is blank/invalid.
        units = to_float(get(DWELLING_UNITS_FIELD))
        if not units or units < 1:
            units = 1.0

        if units > MAX_DWELLING_UNITS:
            skipped_large_murb += 1
            continue

        num = {}
        for field, key in NUMERIC_FIELDS.items():
            fv = to_float(get(field))
            if fv is not None and not (fv == 0 and key in ZERO_AS_MISSING):
                num[key] = fv

        floor_area = num.get('floor_area')
        ghg = to_float(get(GHG_FIELD))
        if ghg is not None and (ghg == 0 or ghg > GHG_PLAUSIBILITY_CEILING):
            ghg = None
        ghg_intensity = to_float(get(GHGI_FIELD))
        if ghg_intensity == 0:
            ghg_intensity = None
        if floor_area:
            if ghg is None and ghg_intensity is not None:
                ghg = ghg_intensity * floor_area
            elif ghg_intensity is None and ghg is not None:
                ghg_intensity = ghg / floor_area
        if ghg is not None:
            num['ghg'] = ghg
        if ghg_intensity is not None:
            num['ghg_intensity'] = ghg_intensity

        if any(key not in num for key in REQUIRED_METRICS):
            skipped_incomplete += 1
            continue

        muni_counts[municipality] = muni_counts.get(municipality, 0) + units
        total_kept += 1

        accumulate_fit(fit_acc, num, floor_area, units)

        cell = cells.setdefault((municipality, year_key), {
            'municipality': municipality,
            'year': year_key,
            'count': 0,
            'num': {},
            'median': {},
            'cat': {'step_achieved': {}, 'house_type_group': {}, 'building_type': {}},
        })
        # 'count' and every metric are weighted by Dwelling Units, to match
        # the NRCan baseline's convention of counting individual homes. For
        # the two BLOCK_TOTAL_KEYS, the row's value is already a total across
        # `units` homes, so only the denominator is weighted (giving the
        # correct per-home average); every other metric is already a
        # per-home value, so both its sum and denominator are weighted.
        cell['count'] += units
        unit_repeats = max(1, round(units))
        for key, v in num.items():
            pair = cell['num'].setdefault(key, [0.0, 0])
            per_home_v = v / units if key in BLOCK_TOTAL_KEYS else v
            pair[0] += per_home_v * units
            pair[1] += units
            # Raw per-home values (not just the sum), so the dashboard can
            # offer a median alongside the mean - each of this row's `units`
            # identical homes gets its own entry.
            cell['median'].setdefault(key, []).extend([per_home_v] * unit_repeats)
        for cat_key, value in (('step_achieved', step_achieved), ('house_type_group', house_type_group), ('building_type', building_type)):
            hist = cell['cat'][cat_key]
            hist[value] = hist.get(value, 0) + units

    print(f'{total_in_sheet} data rows in sheet, {skipped_not_as_built} skipped (not "{WANTED_REPORT_STAGE}" report stage), '
          f'{skipped_large_murb} skipped (> {MAX_DWELLING_UNITS} Dwelling Units), '
          f'{skipped_incomplete} skipped (missing a metric other than NLR), {total_kept} kept')
    print('Municipalities:', dict(sorted(muni_counts.items(), key=lambda kv: -kv[1])))

    size_fit = solve_fit(fit_acc)
    print('Size-adjustment fit (intensity = a + b/floor_area):', size_fit)

    for cell in cells.values():
        for key in cell['median']:
            cell['median'][key].sort()

    out = {
        'generated_at': datetime.datetime.now().isoformat(timespec='seconds'),
        'source_file': wb_path.name,
        'total_rows': total_kept,
        'cells': list(cells.values()),
        'size_fit': size_fit,
    }
    OUT_PATH.write_text(json.dumps(out, separators=(',', ':')), encoding='utf-8')
    print(f'Wrote {OUT_PATH} ({OUT_PATH.stat().st_size / 1000:.1f} KB)')


if __name__ == '__main__':
    main()
