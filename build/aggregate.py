"""Decode the workbook's embedded Power Pivot Data Model and compute the
aggregates the two Lab 3 PivotCharts display (Sum of Revenue)."""
import json, sys
import pandas as pd
from pbixray import PBIXRay
from pbixray.utils import get_data_slice

src, out = sys.argv[1], sys.argv[2]
m = PBIXRay(src)
sales = m.get_table('Sales')
dates = m.get_table('Dates')
mgr = m.get_table('SalesManager')

# pbixray returns the Region table with 0 rows (row count metadata quirk), so
# read its two dictionaries + the 3-bit packed Region column directly.
dec = m._vertipaq_decoder
def dict_vals(col):
    fn = next(f['FileName'] for f in m._data_model.file_log
              if f['FileName'].endswith(f'Region_14c6a382-ae44-4a42-9f6c-7491bddae4a3.{col}.dictionary'))
    return list(dec._read_dictionary(get_data_slice(m._data_model, fn), 0).values.values())
countries, regions = dict_vals('Country'), dict_vals('Region')
fn = next(f['FileName'] for f in m._data_model.file_log
          if f['FileName'].endswith('Region_14c6a382-ae44-4a42-9f6c-7491bddae4a3.Region.0.idf'))
packed = int.from_bytes(bytes(get_data_slice(m._data_model, fn))[-8:], 'little')
region = pd.DataFrame({'Country': countries,
                       'Region': [regions[(packed >> (3 * i)) & 7] for i in range(len(countries))]})

s = sales.merge(dates[['FullDate', 'Start of Month']], left_on='Date', right_on='FullDate', how='left')
s = s.merge(region, left_on='Retailer country', right_on='Country', how='left')
s = s.merge(mgr[['Country', 'Sales Manager']], left_on='Retailer country', right_on='Country', how='left')
assert s['Region'].notna().all() and s['Sales Manager'].notna().all() and s['Start of Month'].notna().all()

monthly = s.pivot_table(index='Start of Month', columns='Region', values='Revenue', aggfunc='sum')
bymgr = s.pivot_table(index='Sales Manager', columns='Region', values='Revenue', aggfunc='sum')
res = {
    'total': float(s.Revenue.sum()),
    'regions': sorted(region.Region.unique().tolist()),
    'region_map': region.set_index('Country').Region.to_dict(),
    'months': [d.strftime('%Y-%m-%d') for d in monthly.index],
    'monthly': {r: [None if pd.isna(v) else float(v) for v in monthly[r]] for r in monthly.columns},
    'managers': bymgr.index.tolist(),
    'bymgr': {r: [None if pd.isna(v) else float(v) for v in bymgr[r]] for r in bymgr.columns},
}
json.dump(res, open(out, 'w'), indent=1, ensure_ascii=False)
print(res['total'], res['regions'], len(res['months']), res['managers'])
