"""Re-solve the reverse DCF with the Street path extended through 2038 at Guggenheim's 2036 growth rates.
Reuses the independent cash-flow replica from verify.py (same inputs workbook); writes models/excel/reverse_resolve.json.
Run from the repository root: python3 models/reverse_dcf/resolve_extended.py
"""
import datetime as dt, json, pathlib, zipfile, xml.etree.ElementTree as ET
root = pathlib.Path(__file__).resolve().parents[2]
p = root / 'outputs/01a111f7-26b0-7fc2-b98f-41db76412d40/PHVS_reverse_SOTP_DCF.xlsx'
ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with zipfile.ZipFile(p) as z:
    tabs = {}
    for i, name in enumerate(['Reverse SOTP', 'Inputs', 'DCF', 'Launch comparison', 'Street evidence'], 1):
        tree = ET.fromstring(z.read(f'xl/worksheets/sheet{i}.xml')); tabs[name] = {c.attrib['r']: c for c in tree.findall('.//m:c', ns)}
def n(tab, cell): return float(tabs[tab][cell].find('m:v', ns).text)
a = lambda r: n('Inputs', 'E' + str(r))
ir = [n('Street evidence', chr(69 + i) + '25') for i in range(11)]; xr = [n('Street evidence', chr(69 + i) + '26') for i in range(11)]
rd = [n('Inputs', chr(69 + i) + '46') for i in range(11)]; sg = [n('Inputs', chr(69 + i) + '47') for i in range(11)]
date = dt.date(2026, 10, 7); cash = a(31) - (date - dt.date(2026, 6, 30)).days * a(33) + a(34)
FX, POS_IR, POS_XR = a(29), a(16), a(17)
g_ir, g_xr = ir[10] / ir[9] - 1, xr[10] / xr[9] - 1
def series(extend):
    if not extend: return ir + [ir[-1]] * 2, xr + [xr[-1]] * 2          # plateau 2037-2038 (saved convention)
    return ir + [ir[-1] * (1 + g_ir), ir[-1] * (1 + g_ir) ** 2], xr + [xr[-1] * (1 + g_xr), xr[-1] * (1 + g_xr) ** 2]
def run(scale, extend, want_rows=False):
    IR, XR = series(extend); pv = 0; previous_wc = 0; previous_retention = [1, 1]; rows = []
    def retention(year, entry):
        age = max(0, year - entry + 1)
        if age == 0: return 1
        if age <= 4: return a(55 + int(age))
        return a(59) * (1 + a(21)) ** (age - 4)
    for i, year in enumerate(range(2026, 2056)):
        t = (dt.date(year, 12, 31) - date).days / 365.25; fraction = t if i == 0 else 1
        retained = [retention(year, a(54)), retention(year, a(55))]
        costfactor = 1 if year <= max(a(54), a(55)) else sum(previous_retention) / 2
        previous_retention = retained
        ir_y = (IR[i] if i < 13 else IR[-1] * retained[0]) * scale; xr_y = (XR[i] if i < 13 else XR[-1] * retained[1]) * scale
        rev = ir_y * POS_IR + xr_y * POS_XR
        ebit = rev * (1 - a(25)) - (rd[min(i, 10)] + sg[min(i, 10)] + a(28)) * fraction * costfactor
        tax = max(0, ebit) * a(23) if year >= a(24) else 0
        wc = rev * a(26); fcf = ebit - tax + (a(28) - a(27)) * fraction * costfactor - (wc - previous_wc); previous_wc = wc
        discount = (1 + a(20)) ** (-t); pv += fcf * discount
        rows.append((year, ir_y * FX, xr_y * FX))
    pv += fcf * (1 + a(21)) / (a(20) - a(21)) * discount
    price = (pv + cash + a(36) - a(35)) * a(29) / a(30)
    return (price, rows) if want_rows else price
def solve(target, extend):
    lo, hi = 0, 5
    for _ in range(100):
        mid = (lo + hi) / 2
        if run(mid, extend) < target: lo = mid
        else: hi = mid
    return (lo + hi) / 2
out = {'as_of': '2026-10-10', 'method': 'Independent cash-flow replica of the 7 October reverse DCF (verify.py), re-solved with the Guggenheim path extended through 2038 at its 2036 growth rates instead of held flat.', 'guggenheim_2036_growth': {'IR': g_ir, 'XR': g_xr}, 'discount_rate': a(20), 'variants': {}}
for label, extend in (('plateau', False), ('extended', True)):
    k_spot = solve(a(9), extend); k_pt = solve(a(10) / (1 + a(20)), extend)
    _, rows_s = run(k_spot, extend, True); _, rows_p = run(k_pt, extend, True)
    req = [[y, s_ir, s_xr, s_ir + s_xr, (s_ir + s_xr) * 0.9, p_ir, p_xr, p_ir + p_xr, (p_ir + p_xr) * 0.9] for (y, s_ir, s_xr), (_, p_ir, p_xr) in zip(rows_s, rows_p) if y >= 2027]
    peak_s = max(r[3] for r in req); peak_p = max(r[7] for r in req)
    cum_s = sum(r[3] for r in req if r[0] <= 2036); cum_p = sum(r[7] for r in req if r[0] <= 2036)
    out['variants'][label] = {'k_spot': k_spot, 'k_pt44': k_pt, 'price_spot': a(9), 'price_pt': a(10), 'horizon_pt_years': 1, 'baseline_1x_price': run(1, extend),
        'required_ramps': {'columns': ['Calendar year', 'Spot IR', 'Spot XR', 'Spot combined', 'Spot risk-weighted combined', 'PT44 IR', 'PT44 XR', 'PT44 combined', 'PT44 risk-weighted combined'], 'rows': req},
        'solve_summary': {'columns': ['Metric', 'Spot $31.83', 'PT $44', 'Note'], 'rows': [
            ['Target price USD', a(9), a(10), 'Spot snapshot 7 Oct 2026 vs Guggenheim target (one-year horizon)'],
            ['Target horizon years', 0, 1, 'Broker target discounted one year at 12.5%'],
            ['Common IR/XR revenue scale (x Guggenheim path)', k_spot, k_pt, 'Multiplier on the Street ramp that reprices the target'],
            ['Normalised 1x reference value USD/share', run(1, extend), run(1, extend), 'Value of the unscaled Street path'],
            ['Peak combined worldwide sales USDm', peak_s, peak_p, '2038 peak, success-case' if extend else '2036-2038 plateau, success-case'],
            ['Cumulative sales 2027-2036 USDm', cum_s, cum_p, 'Undiscounted, success-case']]}}
    print(f"{label:9s} k_spot={k_spot:.6f} k_pt44={k_pt:.6f} peak_spot={peak_s:,.1f} cum_spot_27_36={cum_s:,.1f}")
(root / 'models/excel/reverse_resolve.json').write_text(json.dumps(out, indent=1))
print('2036 growth IR/XR', round(g_ir, 4), round(g_xr, 4))
