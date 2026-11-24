"""
Computes the comparison baseline for this dashboard: the main BC Energy Step
Code Dashboard (NRCan Database Analysis), filtered to Saanich, collapsed to
the metrics this smaller as-built dataset can also provide.

Reads the main dashboard's already-built data.json (run its own
extract_data.py first if it needs refreshing) rather than the workbook, since
that cube is already the aggregated source of truth.

Uses the "entry" year mode (year derived from the report's entry/submission
date), since this workbook has no "year built" column - only a Report Date,
which maps onto the main dashboard's own ENTRYDATE-derived year, not
YEARBUILT.

Usage:
    python build_baseline.py ["path to main dashboard's data.json"]
"""
import json
import sys
from pathlib import Path

DEFAULT_MAIN_DATA = (Path(__file__).parent.parent.parent.parent.parent
                      / '2025-16 - NRCan Database Analysis' / 'dashboard' / 'data.json')
OUT_PATH = Path(__file__).parent / 'baseline.json'

CITY = 'Saanich'
YEAR_MODE = 'entry'

# Metric keys shared with this dashboard's extract_data.py (see method.html).
NUMERIC_KEYS = ['meui', 'tedi', 'ach', 'nla', 'nlr', 'floor_area', 'ghg', 'ghg_intensity', 'meui_imp', 'tedi_imp']
# The main dashboard stores Calc'd TEDI % Imp as a 0-1 fraction; this
# workbook's "Envelope % Better / Heat Loss Reduction" is already a
# percentage, so scale the main dashboard's value to match.
PERCENT_SCALE_100 = {'tedi_imp'}
CATEGORY_KEYS = {'step_achieved': 'step_achieved', 'building_type': 'building_type'}


def main():
    main_data_path = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_MAIN_DATA
    print(f'Reading {main_data_path} ...')
    data = json.loads(main_data_path.read_text(encoding='utf-8'))
    cells = data['year_modes'][YEAR_MODE]['cells']
    saanich_cells = [c for c in cells if c['city'] == CITY]

    total_count = sum(c['count'] for c in saanich_cells)
    sums = {k: [0.0, 0] for k in NUMERIC_KEYS}
    cats = {k: {} for k in CATEGORY_KEYS}
    by_year = {}

    for cell in saanich_cells:
        for k in NUMERIC_KEYS:
            v = cell['num'].get(k)
            if v:
                sums[k][0] += v[0]
                sums[k][1] += v[1]
        for out_key, cell_key in CATEGORY_KEYS.items():
            hist = cell['cat'].get(cell_key)
            if hist:
                for cat, n in hist.items():
                    cats[out_key][cat] = cats[out_key].get(cat, 0) + n
        year = cell['year']
        if isinstance(year, int):
            acc = by_year.setdefault(year, {'count': 0, 'num': {k: [0.0, 0] for k in NUMERIC_KEYS}})
            acc['count'] += cell['count']
            for k in NUMERIC_KEYS:
                v = cell['num'].get(k)
                if v:
                    acc['num'][k][0] += v[0]
                    acc['num'][k][1] += v[1]

    def scale(key, value):
        return value * 100 if (value is not None and key in PERCENT_SCALE_100) else value

    overall_avg = {k: scale(k, sums[k][0] / sums[k][1] if sums[k][1] else None) for k in NUMERIC_KEYS}
    overall_n = {k: sums[k][1] for k in NUMERIC_KEYS}
    years_out = []
    for year in sorted(by_year):
        acc = by_year[year]
        row = {'year': year, 'count': acc['count']}
        for k in NUMERIC_KEYS:
            pair = acc['num'][k]
            row[k] = scale(k, pair[0] / pair[1] if pair[1] else None)
        years_out.append(row)

    out = {
        'generated_at': data.get('generated_at'),
        'source_file': data.get('source_file'),
        'city': CITY,
        'count': total_count,
        'avg': overall_avg,
        'n': overall_n,
        'cat': cats,
        'years': years_out,
    }
    OUT_PATH.write_text(json.dumps(out, separators=(',', ':')), encoding='utf-8')
    print(f'Wrote {OUT_PATH}: {total_count} Saanich homes, {len(years_out)} years')
    print('Overall avg:', overall_avg)


if __name__ == '__main__':
    main()
