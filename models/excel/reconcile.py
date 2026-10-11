"""Compare the recalculated workbook against the Python engine's saved monthly output (working case)."""
import json, sys, openpyxl, warnings
warnings.filterwarnings('ignore')
from pathlib import Path
ROOT = Path(__file__).resolve().parents[2]
maps = json.loads((ROOT / 'models/excel/rowmap.json').read_text()); MC0, NM = maps['MC0'], maps['NM']
case = sys.argv[1] if len(sys.argv) > 1 else 'working'
path = sys.argv[2] if len(sys.argv) > 2 else str(ROOT / 'models/excel/PHVS_Crimson_Capital_Model.xlsx')
wb = openpyxl.load_workbook(path, data_only=True)
monthly = json.loads((ROOT / f'models/finalized/{case}_monthly.json').read_text())
res = json.loads((ROOT / 'models/finalized/results.json').read_text())[case]
OFFS = maps.get('OFFS', {})
class _Sheet:
    def __init__(self, logical):
        phys, self.off = OFFS.get(logical, [logical, 0]); self.ws = wb[phys]
    def cell(self, r, c): return self.ws.cell(r + self.off, c)
def row(ws, r): return [ws.cell(r, MC0 + j).value for j in range(NM)]
def cmp(label, got, exp, tol=1e-6):
    bad = [(j, g, e) for j, (g, e) in enumerate(zip(got, exp)) if g is None or abs((g or 0) - e) > tol]
    mx = max(abs((g or 0) - e) for g, e in zip(got, exp) if g is not None) if got else None
    print(f"{'OK ' if not bad else 'BAD'} {label:42s} max|diff|={mx!s:>12} first_bad={bad[0] if bad else ''}")
    return not bad
X, I, V, Fn = _Sheet('XR Engine'), _Sheet('IR Engine'), _Sheet('Revenue'), _Sheet('Fin Monthly')
b = monthly['both']
ok = True
ok &= cmp('XR patient-years', row(X, maps['RX']['tot_py']), [r['xr_patient_years'] for r in b])
ok &= cmp('XR active', row(X, maps['RX']['tot_active']), [r['xr_active'] for r in b])
ok &= cmp('XR starts', row(X, maps['RX']['tot_starts']), [r['xr_starts'] for r in b])
ok &= cmp('IR member-years', row(I, maps['RI']['tot_py']), [r['ir_member_years'] for r in b])
ok &= cmp('IR active', row(I, maps['RI']['tot_active']), [r['ir_access_active'] for r in b])
ok &= cmp('IR rx', row(I, maps['RI']['tot_rx']), [r['ir_rx'] for r in b])
ok &= cmp('IR paid initial (both)', row(V, maps['RV']['IR_A_paidinit']), [r['ir_initial_doses'] for r in b])
ok &= cmp('IR replacement (both)', row(V, maps['RV']['IR_A_repl']), [r['ir_replacement_doses'] for r in b])
ok &= cmp('IR units (both)', row(V, maps['RV']['IR_A_units']), [r['IR_units'] for r in b])
ok &= cmp('XR units', row(V, maps['RV']['XR_units']), [r['XR_units'] for r in b])
ok &= cmp('XR US revenue', row(V, maps['RV']['XR_US']), [r['XR_US'] for r in b])
ok &= cmp('IR US revenue (both)', row(V, maps['RV']['IR_A_US']), [r['IR_US'] for r in b])
ok &= cmp('XR recognised', row(V, maps['RV']['XR_rec']), [r['XR_recognized'] for r in b])
ok &= cmp('IR recognised (both)', row(V, maps['RV']['IR_A_rec']), [r['IR_recognized'] for r in b])
ok &= cmp('IR US revenue (IR only)', row(V, maps['RV']['IR_B_US']), [r['IR_US'] for r in monthly['IR_only']])
for st in ('both', 'IR_only', 'XR_only', 'neither'):
    m = monthly[st]
    for key, ek in (('revenue', 'revenue'), ('rd', 'rd'), ('commercial', 'commercial'), ('ga', 'ga'), ('manufacture', 'manufacture'), ('license', 'license'), ('capex', 'capex'), ('dep', 'depreciation'), ('ebit', 'operating_income'), ('tax', 'tax'), ('ni', 'net_income'), ('dnwc', 'delta_nwc'), ('fcf', 'fcf'), ('fund_net', 'funding_net'), ('cash', 'cash_end'), ('shares', 'shares_end'), ('equity', 'equity_end'), ('dist', 'distribution')):
        ok &= cmp(f'{st}: {key}', row(Fn, maps['RF'][str((st, key))]), [r[ek] for r in m])
W = _Sheet('Valuation')
for tag, d in (('val', 'valuation'), ('tgt', 'target')):
    for ci, st in enumerate(('both', 'IR_only', 'XR_only', 'neither')):
        got = W.cell(maps['VALROW'][str((tag, 'price'))], 3 + ci).value; exp = res['conditional_valuations'][st][d]['price']
        print(f"{'OK ' if abs(got-exp)<1e-6 else 'BAD'} price {tag} {st:8s} wb={got:.9f} engine={exp:.9f}")
        ok &= abs(got - exp) < 1e-6
got = W.cell(maps['VALROW'][str(('val', 'expected'))], 3).value; print(f"value today wb={got:.9f} engine={res['risk_adjusted_value_today']:.9f}")
got = W.cell(maps['VALROW'][str(('tgt', 'expected'))], 3).value; print(f"value target wb={got:.9f} engine={res['expected_12month_value']:.9f}")
print('ALL OK' if ok else 'MISMATCHES')
