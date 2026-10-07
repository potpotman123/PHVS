"""Independent cash-flow bisection and exported workbook checks; no authoring."""
import datetime as dt
import json
import pathlib
import zipfile
import xml.etree.ElementTree as ET

root = pathlib.Path(__file__).resolve().parents[2]
p = root / 'outputs/01a111f7-26b0-7fc2-b98f-41db76412d40/PHVS_reverse_SOTP_DCF.xlsx'
ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with zipfile.ZipFile(p) as z:
    tabs = {}
    for i, name in enumerate(['Reverse SOTP','Inputs','DCF','Launch comparison','Street evidence'], 1):
        tree = ET.fromstring(z.read(f'xl/worksheets/sheet{i}.xml'))
        tabs[name] = {c.attrib['r']: c for c in tree.findall('.//m:c',ns)}
        assert not tree.findall('.//m:c[@t="e"]',ns), name
    charts = [n for n in z.namelist() if '/charts/chart' in n and n.endswith('.xml')]
    assert len(charts) == 2
    assert all(b'<c:f>' in z.read(n) or b'<f>' in z.read(n) for n in charts)

def n(tab, cell):
    return float(tabs[tab][cell].find('m:v',ns).text)

a = lambda r: n('Inputs', 'E'+str(r))
ir = [n('Street evidence', chr(69+i)+'25') for i in range(11)]
xr = [n('Street evidence', chr(69+i)+'26') for i in range(11)]
rd = [n('Inputs',chr(69+i)+'46') for i in range(11)]
sg = [n('Inputs',chr(69+i)+'47') for i in range(11)]
date = dt.date(2026,10,7)
cash = a(31)-(date-dt.date(2026,6,30)).days*a(33)+a(34)

def price(scale, terminal=True):
    pv = 0
    previous_wc = 0
    previous_retention = [1,1]
    def retention(year, entry):
        age = max(0,year-entry+1)
        if age==0: return 1
        if age<=4: return a(55+int(age))
        return a(59)*(1+a(21))**(age-4)
    for i, year in enumerate(range(2026,2056)):
        t = (dt.date(year,12,31)-date).days/365.25
        fraction = t if i == 0 else 1
        retained = [retention(year,a(54)),retention(year,a(55))]
        costfactor = 1 if year<=max(a(54),a(55)) else sum(previous_retention)/2
        previous_retention = retained
        rev = ((ir[i] if i<11 else ir[-1]*retained[0])*a(16)+(xr[i] if i<11 else xr[-1]*retained[1])*a(17))*scale
        ebit = rev*(1-a(25))-(rd[min(i,10)]+sg[min(i,10)]+a(28))*fraction*costfactor
        tax = max(0,ebit)*a(23) if year>=a(24) else 0
        wc = rev*a(26)
        fcf = ebit-tax+(a(28)-a(27))*fraction*costfactor-(wc-previous_wc)
        previous_wc = wc
        discount = (1+a(20))**(-t)
        pv += fcf*discount
    if terminal:
        pv += fcf*(1+a(21))/(a(20)-a(21))*discount
    return (pv+cash+a(36)-a(35))*a(29)/a(30)

def solve(target, terminal=True):
    lo, hi = 0, 5
    for _ in range(100):
        mid = (lo+hi)/2
        if price(mid, terminal)<target: lo=mid
        else: hi=mid
    return (lo+hi)/2

v = json.loads((root/'models/reverse_dcf/verification.json').read_text())
cases = [('spot',solve(a(9)),v['base']['irScale']),
         ('broker',solve(a(10)/(1+a(20))),v['tests'][0]['irScale']),
         ('no_terminal',solve(a(9),False),v['tests'][3]['irScale'])]
for name, actual, expected in cases:
    assert abs(actual-expected)<1e-9, (name,actual,expected)
result = {'method':'Independent year-by-year cash flow with no-refund tax and bisection',
          'checks':cases,'cached_formula_errors':0,'native_linked_charts':len(charts),
          'native_Excel_execution':'not performed'}
# Explicit long-sum check of the infinite tail, independently of the workbook formula.
last_fcf=n('DCF','AH80')
discount=n('DCF','AH81')
long_tail=sum(last_fcf*((1+a(21))/(1+a(20)))**j for j in range(1,501))*discount
assert abs(long_tail-n('DCF','E88'))<1e-8
assert abs(n('Street evidence','F81')-285.776/457.7)<1e-12
assert abs(n('Street evidence','G97')-(1571/2201-1))<1e-12
result['500_year_tail_matches_perpetuity']=True
result['retained_revenue_years_1_to_4']=[a(i) for i in range(56,60)]
result['long_run_annual_change']=a(21)
(root/'models/reverse_dcf/independent_verification.json').write_text(json.dumps(result,indent=2))
print(json.dumps(result,indent=2))
