"""Build the Crimson Capital PHVS workbook from models/finalized inputs.
Run from repository root: python3 models/excel/build_model.py
Every forecast cell is a live formula; engine outputs are used only on the Checks,
Sensitivities, Street Bridge and Monte Carlo tabs, where they are labelled as saved results.
"""
import json, math, gzip, datetime as dt, calendar, re, sys, copy
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.properties import PageSetupProperties

ROOT = Path(__file__).resolve().parents[2]
FIN = ROOT / 'models/finalized'
OUT = ROOT / 'models/excel/PHVS_Crimson_Capital_Model.xlsx'
TEAM = 'Crimson Capital'
PRICE, PRICE_DATE = 31.10, '2026-10-10'

def J(name, base=FIN): return json.loads((base / name).read_text())
cfg = J('inputs.json'); hist = J('history.json'); src = J('sources.json')['sources']
results = J('results.json'); fin_tables = J('financial_tables.json')
sens = J('sensitivities.json'); rank = J('driver_ranking.json'); priors = J('monte_carlo_priors.json')
mc = J('analysis.json', ROOT / 'models/monte_carlo/results')
CRIM = json.loads((ROOT / 'models/excel/crimson_case.json').read_text()); CR = json.loads((ROOT / 'models/excel/crimson_results.json').read_text())
MCX = json.loads((ROOT / 'models/excel/mc_expected_path_crimson.json').read_text())
for _k, _v in CRIM['priors'].items(): priors['priors'][_k].update(_v)
P = cfg['parameters']

# ------------------------------------------------------------------ styles
F = lambda **k: Font(name='Arial', size=10, **k)
NAVY = '002060'
f_norm, f_bold = F(), F(bold=True)
f_input = F(color='0000FF'); f_link = F(color='008000'); f_hdr = F(bold=True, color='FFFFFF')
f_title = Font(name='Arial', size=14, bold=True, color=NAVY); f_sub = Font(name='Arial', size=10, italic=False, color='595959')
f_flag = F(bold=True, color='C00000')
fill_hdr = PatternFill('solid', fgColor=NAVY); fill_in = PatternFill('solid', fgColor='FFFF00'); fill_key = PatternFill('solid', fgColor='FCE4D6')
fill_band = PatternFill('solid', fgColor='DDEBF7'); fill_grey = PatternFill('solid', fgColor='F2F2F2')
thin = Side(style='thin', color='7F7F7F'); b_bottom = Border(bottom=thin); b_top = Border(top=thin)
NUM = '#,##0.0;(#,##0.0);"-"'; NUM0 = '#,##0;(#,##0);"-"'; NUM2 = '#,##0.00;(#,##0.00);"-"'; NUM3 = '#,##0.000;(#,##0.000);"-"'
PCT = '0.0%;(0.0%);"-"'; PCT2 = '0.00%'; USD = '$#,##0.00;($#,##0.00);"-"'; USD0 = '$#,##0;($#,##0);"-"'; MULT = '0.000"x"'; DATE = 'mmm-yy'; GEN = 'General'
wb = openpyxl.Workbook(); wb.remove(wb.active)

def sheet(name, tab=None, gridlines=False):
    ws = wb.create_sheet(name); ws.sheet_view.showGridLines = gridlines
    if tab: ws.sheet_properties.tabColor = tab
    ws.page_setup.orientation = 'landscape'; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.oddFooter.left.text = TEAM; ws.oddFooter.center.text = name; ws.oddFooter.right.text = 'Page &P of &N'
    ws.oddHeader.left.text = 'Pharvaris (PHVS)'; ws.oddHeader.right.text = '&D'
    return ws

def put(ws, r, c, v, font=None, fmt=None, fill=None, align=None, border=None):
    cell = ws.cell(r, c, v)
    cell.font = font or (f_link if isinstance(v, str) and v.startswith('=') and '!' in v else f_norm)
    if fmt: cell.number_format = fmt
    if fill: cell.fill = fill
    if align: cell.alignment = align
    if border: cell.border = border
    return cell

def band(ws, r, text, c1=1, c2=12):
    for c in range(c1, c2 + 1):
        ws.cell(r, c).fill = fill_hdr; ws.cell(r, c).font = f_hdr
    ws.cell(r, c1, text)

def title(ws, text, sub=None):
    ws.cell(1, 2, text).font = f_title

def widths(ws, spec):
    for col, w in spec.items(): ws.column_dimensions[col].width = w

def name(nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)

def fix_spacing(s):
    s = re.sub(r'(?<=[A-Za-z]{3})(?=\d)', ' ', s)
    s = re.sub(r'(?<=\d)(?=[A-Za-z]{3})', ' ', s)
    s = re.sub(r'(?<=[a-z])(?=[A-Z][a-z])', ' ', s)
    return s.replace('  ', ' ')

# ------------------------------------------------------------------ calendar constants
START = int(cfg['start_month'][:4]) * 12 + int(cfg['start_month'][5:7]) - 1
END_YEAR = cfg['end_year']; NM = (END_YEAR + 1) * 12 - START
PAD = 12; C0 = 5                      # pad columns E..P, first month column Q
MC0 = C0 + PAD                       # column index of July 2026
def mcol(j): return CL(MC0 + j)
LASTCOL = CL(MC0 + NM - 1); FIRSTPAD = CL(C0); LASTPAD = CL(C0 + PAD - 1)
def rng(ws_name, r): return f"'{ws_name}'!${FIRSTPAD}${r}:${LASTCOL}${r}"      # full row incl pads
def mrng(ws_name, r): return f"'{ws_name}'!${mcol(0)}${r}:${LASTCOL}${r}"     # months only
def ref(ws_name, r, j, abs_row=True): return f"'{ws_name}'!{mcol(j)}${r}" if abs_row else f"'{ws_name}'!{mcol(j)}{r}"

def month_header(ws, start_row=1):
    """rows: 2 abs idx, 3 year, 4 quarter, 5 label, 6 ordinal (pads count from 1)."""
    labels = ['Month index', 'Year', 'Quarter', 'Month', 'Column ordinal']
    for i, l in enumerate(labels): put(ws, 2 + i, 1, l, f_bold)
    for k in range(PAD):
        c = C0 + k; put(ws, 6, c, k + 1, f_norm, NUM0); put(ws, 5, c, 'pre', f_sub)
    for j in range(NM):
        m = START + j; y, mo = m // 12, m % 12 + 1; c = MC0 + j
        put(ws, 2, c, m, f_norm, NUM0); put(ws, 3, c, y, f_norm, '0'); put(ws, 4, c, f'{y}Q{(mo - 1) // 3 + 1}')
        put(ws, 5, c, dt.date(y, mo, 1), f_bold, DATE); put(ws, 6, c, PAD + j + 1, f_norm, NUM0)
    ws.freeze_panes = ws.cell(8, MC0)
    ws.column_dimensions['A'].width = 34; ws.column_dimensions['B'].width = 10; ws.column_dimensions['C'].width = 12; ws.column_dimensions['D'].width = 12
    for k in range(PAD): ws.column_dimensions[CL(C0 + k)].width = 7.5
    for j in range(NM): ws.column_dimensions[mcol(j)].width = 9.5

def mrow(ws, r, label, fn, unit='', fmt=NUM, font=None, pad=0, note=''):
    """fn(j) -> formula or value for month j. Pads get constant `pad`."""
    put(ws, r, 1, label); put(ws, r, 2, unit, f_sub)
    if note: put(ws, r, 3, note, f_sub)
    for k in range(PAD): put(ws, r, C0 + k, pad, f_norm, fmt)
    for j in range(NM):
        v = fn(j); put(ws, r, MC0 + j, v, font or (f_norm if isinstance(v, str) and v.startswith('=') else f_input), fmt)
    return r

def lag(ws_name, r, j, lag_expr):
    """value of row r at month j minus lag (lag is a name/expr), 0 if before model."""
    o = PAD + j + 1
    return f"IF({o}-({lag_expr})<1,0,INDEX({rng(ws_name, r)},1,{o}-({lag_expr})))"

# ================================================================== INPUTS
wsI = sheet('Inputs', 'FFC000'); title(wsI, 'Inputs', 'One row per assumption. Blue = hardcoded input. Change the case selector or type an override; every forecast recalculates.')
put(wsI, 4, 2, 'Base case: the Working column, which is the research mode for every input except two Crimson underwriting choices flagged in the Status column. Type a value in the Override column to change any assumption; Cautious and Upside are the bundled low and high commercial cases.', f_sub)



hdr = ['Assumption', 'Name', 'Units', 'Working', 'Cautious', 'Upside', 'Active', 'Override', 'Low', 'High', 'Status', 'Source', 'Where the number comes from']
for i, h in enumerate(hdr): put(wsI, 7, 2 + i, h, f_hdr, fill=fill_hdr)
groups = [
 ('Population', ['population', 'growth']),
 ('Prevention (XR) adoption', ['category_switch', 'prevention_oral_preference', 'oral_destination_capture', 'untreated_xr_reach', 'oral_launch_share', 'oral_mature_share', 'choice_months', 'switch_mature_multiplier', 'hazard_year6_multiplier']),
 ('Rescue (IR) adoption', ['ir_oral_switch_relative', 'ir_injectable_opportunity', 'ir_no_rescue_opportunity']),
 ('Ekterly launch stock projection', ['ekterly_future_forms', 'ekterly_paid_conversion', 'ekterly_continuation', 'ekterly_unique_fraction', 'ekterly_free_stock_multiplier', 'ltp_no_rescue_fraction']),
 ('Payer access', ['conversion', 'delay', 'retry_months', 'reentry_fraction', 'reentry_delay']),
 ('Retention and paid exposure', ['xr_first_retention', 'xr_mature_retention', 'ir_retention', 'xr_paid_intensity']),
 ('Rescue quantity', ['ir_doses_ltp', 'ir_doses_other', 'ir_initial_doses', 'ir_initial_paid_fraction', 'dual_use', 'nonphvs_ltp_shift', 'ir_quantity_multiplier']),
 ('Net price', ['xr_wac_28days', 'xr_net_fraction', 'xr_price_multiplier', 'ir_wac_dose', 'ir_net_fraction', 'ir_price_multiplier']),
 ('Geography', ['ir_exus_ratio', 'xr_exus_ratio', 'exus_lag', 'exus_phase', 'exus_direct_fraction', 'partner_royalty']),
 ('Costs', ['cost_scale', 'cost_inflation', 'commercial_variable_us', 'commercial_variable_exus', 'manufacture_rate', 'license_rate', 'capex_fraction', 'nwc_fraction', 'tax_rate']),
 ('Funding and FX', ['minimum_cash_m', 'issue_price', 'issue_fee', 'fx']),
 ('Clinical risk, discounting, exclusivity', ['ir_pos', 'xr_pos', 'shared_success', 'discount_rate', 'loe_year', 'loe_price_cap']),
]
THESIS = {'ir_quantity_multiplier', 'category_switch', 'switch_mature_multiplier', 'ir_retention', 'oral_launch_share'}
LABELS = {'population': 'US diagnosed HAE population, July 2026', 'growth': 'Net population growth', 'category_switch': 'Annual category switching opportunity',
 'prevention_oral_preference': 'Oral preference among preventive patients', 'oral_destination_capture': 'PHVS capture of departing oral users', 'untreated_xr_reach': 'Reach into diagnosed untreated patients',
 'oral_launch_share': 'PHVS share of oral choices at launch', 'oral_mature_share': 'PHVS share of oral choices, mature', 'choice_months': 'Months to mature share', 'switch_mature_multiplier': 'Switching opportunity multiplier by month 60',
 'hazard_year6_multiplier': 'Hazard multiplier after year 5 (disabled)', 'ir_oral_switch_relative': 'Ekterly-to-PHVS switch transfer', 'ir_injectable_opportunity': 'Injectable rescue opportunity multiplier', 'ir_no_rescue_opportunity': 'Rescue adoption among patients without rescue',
 'ekterly_future_forms': 'Ekterly start forms per month after Feb 2026', 'ekterly_paid_conversion': 'Ekterly paid conversion of forms', 'ekterly_continuation': 'Ekterly annual continuation', 'ekterly_unique_fraction': 'Unique share of Ekterly forms',
 'ekterly_free_stock_multiplier': 'Extra Ekterly free supply multiplier', 'ltp_no_rescue_fraction': 'Preventive patients without any rescue', 'conversion': 'Paid access conversion', 'delay': 'Access delay', 'retry_months': 'Retry after failed access',
 'reentry_fraction': 'Discontinuers who return', 'reentry_delay': 'Return delay', 'xr_first_retention': 'XR first-year retention', 'xr_mature_retention': 'XR mature annual retention', 'ir_retention': 'IR annual membership continuation',
 'xr_paid_intensity': 'XR paid intensity (refill gaps)', 'ir_doses_ltp': 'Rescue doses per year, prevention patients', 'ir_doses_other': 'Rescue doses per year, no prevention', 'ir_initial_doses': 'Initial rescue stock per new member',
 'ir_initial_paid_fraction': 'Share of initial stock that is paid', 'dual_use': 'Same-drug rescue allocation for XR users', 'nonphvs_ltp_shift': 'Non-PHVS prophylaxis migration by year 10', 'ir_quantity_multiplier': 'Redosing quantity multiplier (1.17/1.38)',
 'xr_wac_28days': 'XR comparator WAC per 28 days', 'xr_net_fraction': 'XR net-to-gross', 'xr_price_multiplier': 'XR efficacy price multiplier', 'ir_wac_dose': 'IR comparator WAC per dose', 'ir_net_fraction': 'IR net-to-gross', 'ir_price_multiplier': 'IR efficacy price multiplier',
 'ir_exus_ratio': 'Ex-US / US demand ratio, IR', 'xr_exus_ratio': 'Ex-US / US demand ratio, XR', 'exus_lag': 'Ex-US lag', 'exus_phase': 'Ex-US rollout phase', 'exus_direct_fraction': 'Ex-US sold directly', 'partner_royalty': 'Partner royalty rate on licensed sales',
 'cost_scale': 'Fixed cost scale', 'cost_inflation': 'Fixed cost inflation after 2031', 'commercial_variable_us': 'Variable commercial cost, US', 'commercial_variable_exus': 'Variable commercial cost, ex-US', 'manufacture_rate': 'Manufacturing and distribution', 'license_rate': 'BRAIN royalty',
 'capex_fraction': 'Variable capex', 'nwc_fraction': 'Net working capital', 'tax_rate': 'Tax rate', 'minimum_cash_m': 'Minimum cash trigger', 'issue_price': 'Illustrative issue price', 'issue_fee': 'Issuance fee', 'fx': 'USD per EUR (opening balance translation)',
 'ir_pos': 'IR approval probability', 'xr_pos': 'XR approval probability', 'shared_success': 'Shared molecule gate', 'discount_rate': 'Discount rate', 'loe_year': 'First full generic year', 'loe_price_cap': 'Cap post-generic price index at 1', 'price_growth': 'Annual net price growth from 2028'}
INPUT_ROW = {}
r = 8
for gname, keys in groups:
    band(wsI, r, gname, 2, 14); r += 1
    for k in keys:
        v = P[k]; val = CRIM['overrides'].get(k, v['value']); cases = {c: (val if c == 'working' else cfg['cases'][c].get(k, v['value'])) for c in ('working', 'cautious', 'upside_challenge')}
        put(wsI, r, 2, LABELS.get(k, k), f_bold if k in THESIS else f_norm, fill=fill_key if k in THESIS else None)
        put(wsI, r, 3, k, f_sub); put(wsI, r, 4, v.get('units', ''), f_sub)
        fmt = PCT if ('fraction' in v.get('units', '') or 'probability' in v.get('units', '') or 'survival' in v.get('units', '') or k in ('growth', 'cost_inflation', 'discount_rate', 'tax_rate', 'dual_use', 'oral_launch_share', 'oral_mature_share')) and not k.endswith('multiplier') else (NUM0 if isinstance(val, int) and not isinstance(val, bool) else NUM3)
        if k in ('xr_wac_28days', 'ir_wac_dose', 'issue_price'): fmt = USD
        if k in ('fx',): fmt = NUM3
        for i, c in enumerate(('working', 'cautious', 'upside_challenge')):
            cv = cases[c]
            if k == 'discount_rate' and c == 'working': put(wsI, r, 5 + i, '=ROUND(wacc+single_asset_premium,3)', f_norm, fmt)
            else: put(wsI, r, 5 + i, (int(cv) if isinstance(cv, bool) else cv), f_input, fmt)
        put(wsI, r, 8, f'=IF($I{r}<>"",$I{r},$E{r})', f_bold, fmt)
        wsI.cell(r, 9).fill = fill_in; wsI.cell(r, 9).font = f_input; wsI.cell(r, 9).number_format = fmt
        b = v.get('bounds') or ['', '']
        put(wsI, r, 10, b[0], f_sub, fmt); put(wsI, r, 11, b[1], f_sub, fmt)
        put(wsI, r, 12, ('Crimson underwriting' if k in CRIM['overrides'] else ('WACC build plus premium (DCF tab)' if k == 'discount_rate' else v.get('status', ''))), f_sub); put(wsI, r, 13, ', '.join(v.get('sources', [])), f_sub)
        put(wsI, r, 14, fix_spacing(v.get('translation', '')), f_sub)
        name('pop_growth' if k == 'growth' else k, f"Inputs!$H${r}"); INPUT_ROW[k] = r; r += 1
    r += 1
# structural inputs
band(wsI, r, 'Net price escalation (workbook only; the Python engine holds net price flat)', 2, 14); r += 1
put(wsI, r, 2, 'Annual net price growth from 2028'); put(wsI, r, 3, 'price_growth', f_sub); put(wsI, r, 4, 'annual fraction', f_sub)
for i in range(3): put(wsI, r, 5 + i, 0, f_input, PCT)
put(wsI, r, 8, f'=IF($I{r}<>"",$I{r},$E{r})', f_bold, PCT); wsI.cell(r, 9).fill = fill_in; wsI.cell(r, 9).font = f_input; wsI.cell(r, 9).number_format = PCT
put(wsI, r, 10, -0.02, f_sub, PCT); put(wsI, r, 11, 0.03, f_sub, PCT); put(wsI, r, 12, 'assumed', f_sub); put(wsI, r, 13, 'JUDGMENT', f_sub)
put(wsI, r, 14, 'Applied to both net prices after the 2027 launch year. Zero in every case: list prices in HAE have risen (Orladeyo WAC up in January 2026) but net prices face deeper rebates as Ekterly, Andembry and Dawnzera compete, so flat net is the working assumption. Non-zero values take the workbook away from the engine and flag the Monte Carlo as stale.', f_sub)
name('price_growth', f'Inputs!$H${r}'); INPUT_ROW['price_growth'] = r; r += 2
band(wsI, r, 'Launch and decision dates', 2, 14); r += 1
for prod in ('IR', 'XR'):
    L = cfg['launches'][prod]
    put(wsI, r, 2, f'{prod} approval decision'); put(wsI, r, 5, dt.date.fromisoformat(L['approval_assumption']), f_input, 'd-mmm-yy'); put(wsI, r, 13, L['source'], f_sub)
    put(wsI, r, 14, 'IR: PDUFA date. XR: assumed approval after a 1H 2027 filing. Decision months set when failed-product costs stop.', f_sub); name(f'approval_{prod}', f'Inputs!$E${r}'); r += 1
    put(wsI, r, 2, f'{prod} commercial launch month'); put(wsI, r, 5, dt.date(int(L['commercial_month'][:4]), int(L['commercial_month'][5:]), 1), f_input, 'mmm-yy'); name(f'launch_{prod}', f'Inputs!$E${r}')
    put(wsI, r, 14, 'First month prescriptions can be written; first paid activation follows the access delay.', f_sub); r += 1
    put(wsI, r, 2, f'{prod} launch delay stress'); put(wsI, r, 4, 'months', f_sub); put(wsI, r, 5, 0, f_input, NUM0); name(f'launch_delay_{prod}', f'Inputs!$E${r}')
    put(wsI, r, 14, 'Shifts both the decision and the commercial launch. Six months is the published stress.', f_sub); r += 1
    put(wsI, r, 2, f'{prod} first full generic year'); put(wsI, r, 5, '=loe_year', f_norm, '0'); name(f'loe_year_{prod}', f'Inputs!$E${r}')
    put(wsI, r, 14, 'Defaults to the shared assumption; editable separately.', f_sub); r += 1
r += 1
band(wsI, r, 'Patient origin weights (share of diagnosed population)', 2, 14); r += 1
POOL_ROW = {}
for k, w in cfg['pool_weights'].items():
    put(wsI, r, 2, {'orladeyo': 'Orladeyo', 'other_prevention': 'Other prevention (injectables, older)', 'acute_only': 'Rescue only, no prevention', 'diagnosed_untreated': 'Diagnosed, untreated'}[k]); put(wsI, r, 3, k, f_sub)
    put(wsI, r, 5, w, f_input, PCT); put(wsI, r, 13, 'POP, MIX', f_sub); name(f'w_{k}', f'Inputs!$E${r}'); POOL_ROW[k] = r; r += 1
put(wsI, r, 2, 'Check: weights sum to 100%'); put(wsI, r, 5, f'=SUM(E{r-4}:E{r-1})', f_norm, PCT); r += 2
band(wsI, r, 'Explicit annual budgets, USD millions (fixed costs before scale, inflation and runoff)', 2, 14); r += 1
put(wsI, r, 2, 'Year', f_bold); put(wsI, r, 5, 'R&D', f_bold); put(wsI, r, 6, 'Commercial fixed', f_bold); put(wsI, r, 7, 'G&A', f_bold); r += 1
COST_ROW0 = r
for y, s in cfg['cost_schedule_usdm'].items():
    put(wsI, r, 2, int(y), f_norm, '0'); put(wsI, r, 5, s['rd'], f_input, NUM); put(wsI, r, 6, s['commercial_fixed'], f_input, NUM); put(wsI, r, 7, s['ga'], f_input, NUM); r += 1
name('cost_years', f'Inputs!$B${COST_ROW0}:$B${r-1}'); name('cost_rd', f'Inputs!$E${COST_ROW0}:$E${r-1}'); name('cost_comm', f'Inputs!$F${COST_ROW0}:$F${r-1}'); name('cost_ga', f'Inputs!$G${COST_ROW0}:$G${r-1}')
put(wsI, r, 14, '2026 is the annualised H1 run-rate; 2027-2028 reflect launch build; 2031 is held and inflated thereafter. KalVista launch SG&A and BioCryst mature costs are the reasonableness anchors (COST, COST2).', f_sub); r += 2
band(wsI, r, 'Financial policy', 2, 14); r += 1
POLICY_LABEL = {'rd_IR_fraction': 'R&D avoided if IR fails', 'single_product_platform_fraction': 'Fixed cost kept with one product', 'winddown_platform_fraction': 'Fixed cost kept during wind-down', 'winddown_months': 'Months to liquidation after last failure', 'remaining_milestone_eurm': 'Remaining BRAIN milestone, EUR m', 'capex_fixed_usdm': 'Fixed capex per year, USD m', 'depreciation_months': 'Depreciation life, months', 'tax_floor_usdm': 'Cash tax floor per year, USD m', 'nol_utilization_cap': 'NOL utilisation cap', 'cash_raise_target_multiple': 'Raise cash to this multiple of minimum'}
for k, v in cfg['financial_policy'].items():
    if not isinstance(v, dict): continue
    put(wsI, r, 2, POLICY_LABEL[k]); put(wsI, r, 3, k, f_sub); put(wsI, r, 5, v['value'], f_input, PCT if v['value'] < 1 and 'months' not in k else NUM2)
    put(wsI, r, 10, v['bounds'][0], f_sub); put(wsI, r, 11, v['bounds'][1], f_sub); put(wsI, r, 13, ', '.join(v['source']), f_sub); put(wsI, r, 14, fix_spacing(v['definition']), f_sub)
    name(k, f'Inputs!$E${r}'); r += 1
r += 1
band(wsI, r, 'Post-generic price and quantity framework (50/50 pooled study and Kuvan)', 2, 14); r += 1
z = cfg['loe']
put(wsI, r, 2, 'Weight on pooled brand study'); put(wsI, r, 5, z['study_weight'], f_input, PCT); name('study_weight', f'Inputs!$E${r}'); put(wsI, r, 14, 'User choice: half the Health Affairs brand price x volume path, half the Kuvan revenue path.', f_sub); r += 1
put(wsI, r, 2, 'Post-year-4 quantity decline'); put(wsI, r, 5, z['tail_growth'], f_input, PCT); name('tail_growth', f'Inputs!$E${r}'); put(wsI, r, 14, 'Kuvan 2021-24 revenue CAGR, assigned to quantity with price flat.', f_sub); r += 1
put(wsI, r, 2, 'Exposure year', f_bold); [put(wsI, r, 5 + i, i + 1, f_bold, '0') for i in range(4)]; r += 1
LOE_ARR = {}
for key, lab in (('study_p', 'Study brand net price index'), ('study_q', 'Study brand volume index'), ('rare_p', 'Kuvan price index (assumed flat)'), ('rare_revenue', 'Kuvan revenue retained')):
    put(wsI, r, 2, lab); [put(wsI, r, 5 + i, z[key][i], f_input, NUM3) for i in range(4)]; name(key, f'Inputs!$E${r}:$H${r}'); put(wsI, r, 13, 'LOE, KUVAN', f_sub); r += 1
r += 1
band(wsI, r, 'Segmentation constants (observed anchors and ratios)', 2, 14); r += 1
sc = cfg['segmented_constants']
for key, lab, fmt, note in (('reference_paid_conversion', 'Normalisation paid conversion', PCT, 'Completed-start proxies are divided by 80% once; actual access conversion is applied separately.'),
                            ('rescue_oral_preference', 'Oral rescue preference (survey)', PCT, 'RESCUEPREF survey; applies to patients with no rescue.'),
                            ('ekterly_first8_forms', 'Ekterly forms, first 8 months', NUM0, 'KALVFORM: 1,702 start forms through Feb 2026.'),
                            ('ekterly_launch_addressable', 'Ekterly addressable patients', NUM0, 'Company ~9,000 diagnosed estimate used as the denominator.'),
                            ('entrant_rescue_to_prevention_ratio', 'Entrant ratio, rescue-only to prevention (32/52)', NUM3, 'MIX: historical first-oral trier mix 52/32/17.'),
                            ('entrant_naive_to_rescue_ratio', 'Entrant ratio, untreated to rescue-only (17/32)', NUM3, 'MIX.')):
    put(wsI, r, 2, lab); put(wsI, r, 3, key, f_sub); put(wsI, r, 5, sc[key], f_input, fmt); put(wsI, r, 14, note, f_sub); name(key, f'Inputs!$E${r}'); r += 1
put(wsI, r, 2, 'Ekterly monthly start forms, Jul 2025 to Feb 2026'); put(wsI, r, 3, 'ekterly_forms', f_sub)
for i, v in enumerate(sc['ekterly_monthly_forms_Jul2025_Feb2026']): put(wsI, r, 5 + i, v, f_input, NUM)
name('ekterly_forms_obs', f'Inputs!$E${r}:$L${r}'); put(wsI, r, 14, 'Bimonthly disclosures split evenly across their two months.', f_sub); r += 2
band(wsI, r, 'Valuation dates and market reference', 2, 14); r += 1
put(wsI, r, 2, 'Valuation date'); put(wsI, r, 5, dt.date.fromisoformat(cfg['valuation_date']), f_input, 'd-mmm-yy'); name('valuation_date', f'Inputs!$E${r}'); r += 1
put(wsI, r, 2, 'Target date (12 months)'); put(wsI, r, 5, dt.date.fromisoformat(cfg['target_date']), f_input, 'd-mmm-yy'); name('target_date', f'Inputs!$E${r}'); r += 1
put(wsI, r, 2, 'Share price'); put(wsI, r, 5, PRICE, f_input, USD); name('share_price', f'Inputs!$E${r}'); put(wsI, r, 14, f'NASDAQ close reference supplied by the team, {PRICE_DATE}.', f_sub); r += 1
put(wsI, r, 2, 'Price date'); put(wsI, r, 5, dt.date.fromisoformat(PRICE_DATE), f_input, 'd-mmm-yy'); name('price_date', f'Inputs!$E${r}'); r += 1
put(wsI, r, 2, 'Last explicit forecast year'); put(wsI, r, 5, END_YEAR, f_input, '0'); name('end_year', f'Inputs!$E${r}'); r += 1
widths(wsI, {'A': 2, 'B': 44, 'C': 24, 'D': 16, 'E': 11, 'F': 11, 'G': 11, 'H': 11, 'I': 10, 'J': 9, 'K': 9, 'L': 24, 'M': 18, 'N': 120})
wsI.freeze_panes = 'C8'

# ================================================================== HISTORY
wsH = sheet('History', 'FFC000'); title(wsH, 'Reported financials', 'EUR millions as reported (FY2025 Form 20-F; June 2026 interim). Forecasts start from the 30 June 2026 balance translated once at the ECB rate.')
hp = hist['periods']; cols = ['2024', '2025', 'H1_2026']
put(wsH, 4, 2, 'EUR millions', f_bold); [put(wsH, 4, 3 + i, c.replace('_', ' '), f_hdr, fill=fill_hdr, align=Alignment(horizontal='right')) for i, c in enumerate(cols)]
HIST_ROW = {}; r = 5
for key, lab in (('revenue', 'Revenue'), ('rd', 'Research and development'), ('ga', 'General and administrative'), ('opex', 'Operating expenses'), ('finance_income', 'Finance income (expense)'), ('tax', 'Income tax'), ('net_income', 'Net loss'), ('sbc', 'Share-based compensation (non-cash)'), ('depreciation', 'Depreciation'), ('cfo', 'Cash from operations'), ('capex', 'Capital expenditure'), ('cfi', 'Cash from investing'), ('cff', 'Cash from financing'), ('cash_fx', 'FX on cash'), ('cash_begin', 'Cash, beginning'), ('cash_end', 'Cash, end'), ('total_assets', 'Total assets'), ('liabilities', 'Total liabilities'), ('equity', 'Shareholders equity'), ('shares_end', 'Shares outstanding, millions')):
    put(wsH, r, 2, lab); [put(wsH, r, 3 + i, hp[c][key], f_input, NUM) for i, c in enumerate(cols)]; HIST_ROW[key] = r; r += 1
put(wsH, r, 2, 'Check: cash roll-forward'); [put(wsH, r, 3 + i, f'={CL(3+i)}{HIST_ROW["cash_begin"]}+{CL(3+i)}{HIST_ROW["cfo"]}+{CL(3+i)}{HIST_ROW["cfi"]}+{CL(3+i)}{HIST_ROW["cff"]}+{CL(3+i)}{HIST_ROW["cash_fx"]}-{CL(3+i)}{HIST_ROW["cash_end"]}', f_norm, NUM3) for i in range(3)]; r += 2
band(wsH, r, 'Opening balance, 30 June 2026', 2, 6); r += 1
o = hist['opening_2026_06_30']; OPEN = {}
for key, lab, fmt in (('cash_eurm', 'Cash, EUR m', NUM), ('nwc_eurm', 'Operating net working capital, EUR m', NUM), ('ppe_and_rou_eurm', 'PP&E and right-of-use assets, EUR m', NUM), ('other_net_assets_eurm', 'Other net assets, EUR m', NUM), ('equity_eurm', 'Equity, EUR m', NUM), ('term_debt_eurm', 'Term debt, EUR m', NUM), ('shares_m', 'Ordinary shares, millions', NUM3), ('options_m', 'Options outstanding, millions', NUM3), ('option_strike_eur', 'Weighted option strike, EUR', NUM2), ('rsus_m', 'RSUs outstanding, millions', NUM3)):
    put(wsH, r, 2, lab); put(wsH, r, 3, o[key], f_input, fmt); name('open_' + key.replace('_eurm', '').replace('_m', '_m').replace('_eur', '_eur'), f'History!$C${r}'); OPEN[key] = r; r += 1
put(wsH, r, 2, 'Check: cash + NWC + PP&E + other = equity (EUR m)'); put(wsH, r, 3, f'=C{OPEN["cash_eurm"]}+C{OPEN["nwc_eurm"]}+C{OPEN["ppe_and_rou_eurm"]}+C{OPEN["other_net_assets_eurm"]}-C{OPEN["equity_eurm"]}', f_norm, NUM3); r += 1
put(wsH, r, 2, 'Sources: F25 (20-F, 2 April 2026), H126 (interim, 12 August 2026). FX 1.1186 USD/EUR, ECB 8 October 2026.', f_sub)
widths(wsH, {'A': 2, 'B': 44, 'C': 14, 'D': 14, 'E': 14})

# ================================================================== LOE
wsL = sheet('LOE', '7F7F7F'); title(wsL, 'Post-generic economics', 'Revenue, price and quantity indices by calendar year. Blend of the pooled brand study and the Kuvan revenue path; price = revenue / quantity.')
years = list(range(2026, END_YEAR + 1))
put(wsL, 4, 2, 'Year', f_bold)
heads = ['IR price index', 'IR quantity index', 'IR revenue index', 'XR price index', 'XR quantity index', 'XR revenue index']
for i, h in enumerate(heads): put(wsL, 4, 3 + i, h, f_hdr, fill=fill_hdr)
LOE_ROW0 = 5
for i, y in enumerate(years):
    rr = LOE_ROW0 + i; put(wsL, rr, 2, y, f_norm, '0')
    for pi, prod in enumerate(('IR', 'XR')):
        c = 3 + pi * 3; e = f'loe_year_{prod}'; idx = f'MIN($B{rr}-{e},3)+1'
        rev = f'(study_weight*INDEX(study_p,{idx})*INDEX(study_q,{idx})+(1-study_weight)*INDEX(rare_revenue,{idx}))'
        qty = f'(study_weight*INDEX(study_q,{idx})+(1-study_weight)*INDEX(rare_revenue,{idx})/INDEX(rare_p,{idx}))'
        put(wsL, rr, c, f'=IF($B{rr}<{e},1,IF(loe_price_cap,MIN(1,{rev}/{qty}),{rev}/{qty}))', f_norm, NUM3)
        put(wsL, rr, c + 1, f'=IF($B{rr}<{e},1,{qty}*IF($B{rr}-{e}>3,(1+tail_growth)^($B{rr}-{e}-3),1))', f_norm, NUM3)
        put(wsL, rr, c + 2, f'={CL(c)}{rr}*{CL(c+1)}{rr}', f_norm, NUM3)
LOE_LAST = LOE_ROW0 + len(years) - 1
name('loe_years', f'LOE!$B${LOE_ROW0}:$B${LOE_LAST}')
for i, nm in enumerate(['ir_p_idx', 'ir_q_idx', 'ir_r_idx', 'xr_p_idx', 'xr_q_idx', 'xr_r_idx']): name(nm, f'LOE!${CL(3+i)}${LOE_ROW0}:${CL(3+i)}${LOE_LAST}')
put(wsL, LOE_LAST + 2, 2, 'Years 1-4 after entry take the blended study and Kuvan indices; afterwards price is flat and quantity declines at the Kuvan 2021-24 rate. Protected-year volumes are frozen as the anchor (Revenue tab).', f_sub)
widths(wsL, {'A': 2, 'B': 10, 'C': 14, 'D': 14, 'E': 14, 'F': 14, 'G': 14, 'H': 14})

# ================================================================== CALENDAR
wsC = sheet('Calendar', '7F7F7F'); title(wsC, 'Monthly calendar', 'Timing flags shared by the monthly sheets. July 2026 to December 2055.')
month_header(wsC)
CAL = {}; r = 8
def crow(key, label, fn, unit='', fmt=NUM0, pad=0, note=''):
    global r; CAL[key] = r; mrow(wsC, r, label, fn, unit, fmt, pad=pad, note=note); r += 1
crow('days', 'Days in month', lambda j: f'=DAY(EOMONTH({mcol(j)}$5,0))')
crow('begin', 'Month begins', lambda j: f'={mcol(j)}$5', fmt='d-mmm-yy')
crow('end', 'Month ends', lambda j: f'=EOMONTH({mcol(j)}$5,0)', fmt='d-mmm-yy')
crow('launch_idx_IR', 'IR launch month index', lambda j: '=YEAR(launch_IR)*12+MONTH(launch_IR)-1+launch_delay_IR')
crow('launch_idx_XR', 'XR launch month index', lambda j: '=YEAR(launch_XR)*12+MONTH(launch_XR)-1+launch_delay_XR')
crow('age_IR', 'IR months since launch', lambda j: f'={mcol(j)}$2-{mcol(j)}{CAL["launch_idx_IR"]}')
crow('age_XR', 'XR months since launch', lambda j: f'={mcol(j)}$2-{mcol(j)}{CAL["launch_idx_XR"]}')
crow('on_IR', 'IR launched', lambda j: f'=IF({mcol(j)}{CAL["age_IR"]}>=0,1,0)')
crow('on_XR', 'XR launched', lambda j: f'=IF({mcol(j)}{CAL["age_XR"]}>=0,1,0)')
crow('g', 'Monthly population growth', lambda j: '=0' if j == 0 else '=(1+pop_growth)^(1/12)-1', fmt=PCT2, note='Zero in the anchor month')
crow('dec_IR', 'IR decision month index', lambda j: '=YEAR(approval_IR)*12+MONTH(approval_IR)-1+launch_delay_IR')
crow('dec_XR', 'XR decision month index', lambda j: '=YEAR(approval_XR)*12+MONTH(approval_XR)-1+launch_delay_XR')
crow('pre_dec_IR', 'Before or at IR decision', lambda j: f'=IF({mcol(j)}$2<={mcol(j)}{CAL["dec_IR"]},1,0)')
crow('pre_dec_XR', 'Before or at XR decision', lambda j: f'=IF({mcol(j)}$2<={mcol(j)}{CAL["dec_XR"]},1,0)')
crow('shutdown', 'Liquidation month if both fail', lambda j: f'=MAX({mcol(j)}{CAL["dec_IR"]},{mcol(j)}{CAL["dec_XR"]})+winddown_months')
crow('is_shutdown', 'Liquidation month flag', lambda j: f'=IF({mcol(j)}$2={mcol(j)}{CAL["shutdown"]},1,0)')
crow('ongoing_fail', 'Still operating if both fail', lambda j: f'=IF({mcol(j)}$2<={mcol(j)}{CAL["shutdown"]},1,0)')
crow('sched_year', 'Cost schedule year', lambda j: f'=MIN({mcol(j)}$3,2031)')
crow('infl', 'Fixed cost inflation factor', lambda j: f'=(1+cost_inflation)^MAX(0,MIN({mcol(j)}$3,loe_year-1)-2031)', fmt=NUM3)
crow('runoff', 'Cost runoff (prior-year erosion)', lambda j: f'=(INDEX(ir_r_idx,MATCH(MAX(2026,{mcol(j)}$3-1),loe_years,0))+INDEX(xr_r_idx,MATCH(MAX(2026,{mcol(j)}$3-1),loe_years,0)))/2', fmt=NUM3)
for tag, dname in (('val', 'valuation_date'), ('tgt', 'target_date')):
    crow(f'frac_{tag}', f'Elapsed fraction at {tag} date', lambda j, d=dname: f'=IF({d}<={mcol(j)}{CAL["begin"]},0,IF({d}>{mcol(j)}{CAL["end"]},1,({d}-{mcol(j)}{CAL["begin"]})/{mcol(j)}{CAL["days"]}))', fmt=NUM3)
    crow(f'after_{tag}', f'Month end on or after {tag} date', lambda j, d=dname: f'=IF({mcol(j)}{CAL["end"]}>={d},1,0)')
    crow(f'full_{tag}', f'Full month before {tag} date', lambda j, d=dname: f'=IF({mcol(j)}{CAL["end"]}<{d},1,0)')
    crow(f'df_{tag}', f'Discount factor to {tag} date', lambda j, d=dname: f'=IF({mcol(j)}{CAL["end"]}<{d},0,(1+discount_rate)^(-({mcol(j)}{CAL["end"]}-{d})/365.25))', fmt=NUM3)
crow('is_end_year', 'Terminal year flag', lambda j: f'=IF({mcol(j)}$3=end_year,1,0)')

# ================================================================== ADOPTION BUILD (pools, rates)
FUNNEL_YEARS = list(range(2027, END_YEAR + 1))
def build_funnel():
    ws = sheet('Funnel', '0070C0'); title(ws, 'Operating build: from diagnosed patients to revenue')
    put(ws, 2, 2, 'Central case, both products approved. Each step is an annual flow or year-end stock from the monthly engines; the driver beneath each step is the live input that governs it.', f_sub)
    r = 4; put(ws, r, 2, 'Year', f_hdr, fill=fill_hdr); put(ws, r, 3, 'Units', f_hdr, fill=fill_hdr)
    for i, y in enumerate(FUNNEL_YEARS): put(ws, r, 4 + i, y, f_hdr, fill=fill_hdr, align=Alignment(horizontal='right')); ws.column_dimensions[CL(4 + i)].width = 9.5
    FY = r; r += 1
    def frow(label, unit, fn, fmt=NUM0, bold=False, driver=False, indent=0):
        nonlocal r; put(ws, r, 2, ('   ' * indent) + label, f_bold if bold else f_norm); put(ws, r, 3, unit, f_sub)
        for i, y in enumerate(FUNNEL_YEARS):
            L = CL(4 + i); v = fn(L, y); c = put(ws, r, 4 + i, v, (f_link if driver else (f_bold if bold else f_norm)), fmt)
        r += 1; return r - 1
    yc = lambda L: f'{L}${FY}'
    sy = lambda sh, row: (lambda L, y: '=' + sum_year(sh, row, yc(L)))
    st = lambda sh, row: (lambda L, y: '=' + stock_year(sh, row, yc(L)))
    avail = lambda RR, sh: (lambda L, y: '=' + '+'.join(stock_year(sh, RR[(f'{o}|{s_}', 'avail')], yc(L)) for (o, s_) in CELLS))
    band(ws, r, 'Prevention (XR): from the diagnosed population to paid treatment-years', 2, 3 + len(FUNNEL_YEARS)); r += 1
    frow('US diagnosed HAE patients, year end', 'patients', st('XR Engine', RX['tot_total']))
    frow('Net population growth', '%', lambda L, y: '=pop_growth', PCT, driver=True, indent=1)
    frow('Patients still available to switch to XR, year end', 'patients', avail(RX, 'XR Engine'))
    frow('Annual switching opportunity', '%', lambda L, y: '=category_switch', PCT, driver=True, indent=1)
    frow('Oral preference among preventive patients', '%', lambda L, y: '=prevention_oral_preference', PCT, driver=True, indent=1)
    frow('PHVS share of oral choices, year average', '%', lambda L, y: f"=AVERAGEIFS({mrng('XR Engine', RX['choice'])},{mrng('XR Engine', 3)},{yc(L)})", PCT, driver=True, indent=1)
    frow('XR prescriptions written', 'prescriptions', sy('XR Engine', RX['tot_rx']), bold=True)
    frow('Paid access conversion', '%', lambda L, y: '=conversion', PCT, driver=True, indent=1)
    frow('XR paid starts', 'patients', sy('XR Engine', RX['tot_starts']), bold=True)
    frow('First-year retention', '%', lambda L, y: '=xr_first_retention', PCT, driver=True, indent=1)
    frow('Mature annual retention', '%', lambda L, y: '=xr_mature_retention', PCT, driver=True, indent=1)
    frow('XR active patients, year end', 'patients', st('XR Engine', RX['tot_active']), bold=True)
    frow('XR patient-years of exposure', 'patient-years', sy('XR Engine', RX['tot_py']))
    frow('Paid intensity (refill gaps)', '%', lambda L, y: '=xr_paid_intensity', PCT, driver=True, indent=1)
    frow('XR paid treatment-years', 'treatment-years', sy('Revenue', RV['XR_units']), bold=True)
    XRU = r - 1
    frow('XR net price per treatment-year', 'USD', lambda L, y: f"=IF({L}{XRU}=0,xr_wac_28days*365/28*xr_net_fraction,{sum_year('Revenue', RV['XR_US'], yc(L))}*1000000/{L}{XRU})", USD0, driver=True, indent=1)
    frow('XR revenue, United States', 'USD m', sy('Revenue', RV['XR_US']), NUM)
    frow('XR revenue, ex-US', 'USD m', lambda L, y: f"={sum_year('Revenue', RV['XR_direct'], yc(L))}+{sum_year('Revenue', RV['XR_roy'], yc(L))}", NUM)
    XRREV = frow('XR revenue', 'USD m', sy('Revenue', RV['XR_rec']), NUM, bold=True)
    r += 1; band(ws, r, 'On-demand (IR): from rescue users to paid doses', 2, 3 + len(FUNNEL_YEARS)); r += 1
    frow('Patients still available to adopt IR, year end', 'patients', avail(RI, 'IR Engine'))
    frow('Ekterly-to-PHVS switch transfer', 'x', lambda L, y: '=ir_oral_switch_relative', NUM2, driver=True, indent=1)
    frow('Injectable rescue opportunity multiplier', 'x', lambda L, y: '=ir_injectable_opportunity', NUM2, driver=True, indent=1)
    frow('PHVS share of oral choices, year average', '%', lambda L, y: f"=AVERAGEIFS({mrng('IR Engine', RI['choice'])},{mrng('IR Engine', 3)},{yc(L)})", PCT, driver=True, indent=1)
    frow('IR prescriptions written', 'prescriptions', sy('IR Engine', RI['tot_rx']), bold=True)
    frow('Paid access conversion', '%', lambda L, y: '=conversion', PCT, driver=True, indent=1)
    frow('IR paid starts', 'patients', sy('IR Engine', RI['tot_starts']), bold=True)
    frow('Annual membership continuation', '%', lambda L, y: '=ir_retention', PCT, driver=True, indent=1)
    frow('IR active members, year end', 'patients', st('IR Engine', RI['tot_active']), bold=True)
    frow('IR member-years', 'member-years', sy('IR Engine', RI['tot_py']))
    IRPY = r - 1
    frow('Rescue doses per year, prevention and no prevention', 'doses', lambda L, y: '=TEXT(ir_doses_ltp,"0.0")&" / "&TEXT(ir_doses_other,"0.0")', GEN, driver=True, indent=1)
    frow('Redosing quantity multiplier', 'x', lambda L, y: '=ir_quantity_multiplier', NUM3, driver=True, indent=1)
    frow('IR paid doses', 'doses', sy('Revenue', RV['IR_A_units']), bold=True)
    IRU = r - 1
    frow('Paid doses per member-year', 'doses', lambda L, y: f"=IF({L}{IRPY}=0,\"\",{L}{IRU}/{L}{IRPY})", NUM2, indent=1)
    frow('IR net price per dose', 'USD', lambda L, y: f"=IF({L}{IRU}=0,ir_wac_dose*ir_net_fraction,{sum_year('Revenue', RV['IR_A_US'], yc(L))}*1000000/{L}{IRU})", USD0, driver=True, indent=1)
    frow('IR revenue, United States', 'USD m', sy('Revenue', RV['IR_A_US']), NUM)
    frow('IR revenue, ex-US', 'USD m', lambda L, y: f"={sum_year('Revenue', RV['IR_A_direct'], yc(L))}+{sum_year('Revenue', RV['IR_A_roy'], yc(L))}", NUM)
    IRREV = frow('IR revenue', 'USD m', sy('Revenue', RV['IR_A_rec']), NUM, bold=True)
    r += 1
    frow('Total revenue, central case', 'USD m', lambda L, y: f'={L}{XRREV}+{L}{IRREV}', NUM, bold=True)
    frow('Exclusivity', '', lambda L, y: f'=IF({yc(L)}<loe_year,"protected","generic")', GEN)
    put(ws, r + 1, 2, 'Switching opportunity is applied to the patients still available in each origin pool, so starts taper as the pool is drawn down even though the PHVS share of oral choices rises to 60% by month 60. Failed authorisations retry after twelve months. Supporting builds for every driver follow below.', f_sub)
    widths(ws, {'A': 2, 'B': 52, 'C': 14}); ws.freeze_panes = 'D5'; ws.print_area = f'B1:{CL(3 + len(FUNNEL_YEARS))}{r + 1}'
    return ws
wsA = sheet('Adoption Build', '0070C0'); title(wsA, 'Patient adoption build', 'Where patients start, how fast they try PHVS, and how many stay. Rates feed the monthly XR and IR engines.')
ORIG = ['orladeyo', 'other_prevention', 'acute_only', 'diagnosed_untreated']; RESC = ['ekterly', 'injectable', 'none']
ON = {'orladeyo': 'Orladeyo', 'other_prevention': 'Other prevention', 'acute_only': 'Rescue only', 'diagnosed_untreated': 'Untreated'}; RN = {'ekterly': 'Ekterly', 'injectable': 'Injectable rescue', 'none': 'No rescue'}
r = 4; band(wsA, r, 'Ekterly active oral rescue stock at PHVS launch (projected from start forms)', 2, 12); r += 1
put(wsA, r, 2, 'Months from July 2025 to IR launch'); put(wsA, r, 5, "=YEAR(launch_IR)*12+MONTH(launch_IR)-1+launch_delay_IR-(2025*12+6)", f_norm, NUM0); name('ek_nmonths', f"'Adoption Build'!$E${r}"); r += 1
put(wsA, r, 2, 'Month', f_bold); put(wsA, r, 3, 'Index', f_bold); put(wsA, r, 4, 'Forms', f_bold); put(wsA, r, 5, 'Counted', f_bold); put(wsA, r, 6, 'Survival weight', f_bold); put(wsA, r, 7, 'Active contribution', f_bold); r += 1
EK0 = r
for i in range(24):
    m = 2025 * 12 + 6 + i; put(wsA, r, 2, dt.date(m // 12, m % 12 + 1, 1), f_norm, 'mmm-yy'); put(wsA, r, 3, i, f_norm, NUM0)
    put(wsA, r, 4, f'=IF(C{r}<8,INDEX(ekterly_forms_obs,1,C{r}+1),ekterly_future_forms)', f_norm, NUM)
    put(wsA, r, 5, f'=IF(C{r}<=ek_nmonths-2,1,0)', f_norm, NUM0)
    put(wsA, r, 6, f'=IF(E{r}=1,ekterly_continuation^((ek_nmonths-C{r}-2)/12),0)', f_norm, NUM3)
    put(wsA, r, 7, f'=D{r}*E{r}*F{r}*ekterly_paid_conversion*ekterly_unique_fraction', f_norm, NUM); r += 1
put(wsA, r, 2, 'Active Ekterly oral rescue users at launch', f_bold); put(wsA, r, 7, f'=SUM(G{EK0}:G{r-1})*ekterly_free_stock_multiplier', f_bold, NUM); name('ek_active_launch', f"'Adoption Build'!$G${r}"); r += 1
put(wsA, r, 2, 'Expressed at July 2026 (divide by population growth to launch)'); put(wsA, r, 7, '=ek_active_launch/(1+pop_growth)^((YEAR(launch_IR)*12+MONTH(launch_IR)-1+launch_delay_IR-(2026*12+6))/12)', f_norm, NUM); name('ek_oral_base', f"'Adoption Build'!$G${r}"); r += 1
put(wsA, r, 2, 'Forms are prescription attempts, not active patients. Conversion, continuation and uniqueness are assumptions; the 1,702 forms through February 2026 are the only observed anchor (KALVFORM).', f_sub); r += 2
band(wsA, r, 'Starting pools, July 2026: prevention origin crossed with primary rescue origin', 2, 12); r += 1
put(wsA, r, 2, 'Origin', f_bold); put(wsA, r, 3, 'Patients', f_bold); put(wsA, r, 4, 'No rescue', f_bold); put(wsA, r, 5, 'Rescue-eligible', f_bold); put(wsA, r, 6, 'Ekterly', f_bold); put(wsA, r, 7, 'Injectable', f_bold); put(wsA, r, 8, 'None', f_bold); r += 1
JP0 = r; POOL_CELL = {}
for k in ORIG:
    put(wsA, r, 2, ON[k]); put(wsA, r, 3, f'=population*w_{k}', f_norm, NUM)
    put(wsA, r, 4, f'=C{r}*' + ('1' if k == 'diagnosed_untreated' else 'ltp_no_rescue_fraction' if k in ('orladeyo', 'other_prevention') else '0'), f_norm, NUM)
    put(wsA, r, 5, f'=C{r}-D{r}', f_norm, NUM); r += 1
put(wsA, r, 2, 'Total', f_bold); put(wsA, r, 3, f'=SUM(C{JP0}:C{r-1})', f_bold, NUM); put(wsA, r, 4, f'=SUM(D{JP0}:D{r-1})', f_bold, NUM); put(wsA, r, 5, f'=SUM(E{JP0}:E{r-1})', f_bold, NUM); JPT = r; r += 1
for i, k in enumerate(ORIG):
    rr = JP0 + i
    put(wsA, rr, 6, f'=E{rr}/$E${JPT}*ek_oral_base', f_norm, NUM); put(wsA, rr, 7, f'=E{rr}-F{rr}', f_norm, NUM); put(wsA, rr, 8, f'=D{rr}', f_norm, NUM)
    for ci, rs in enumerate(RESC): POOL_CELL[(k, rs)] = f"'Adoption Build'!${CL(6+ci)}${rr}"
put(wsA, JPT, 6, f'=SUM(F{JP0}:F{JPT-1})', f_bold, NUM); put(wsA, JPT, 7, f'=SUM(G{JP0}:G{JPT-1})', f_bold, NUM); put(wsA, JPT, 8, f'=SUM(H{JP0}:H{JPT-1})', f_bold, NUM)
put(wsA, r, 2, 'Check: pools sum to population'); put(wsA, r, 3, f'=SUM(F{JP0}:H{JPT-1})-population', f_norm, NUM3); r += 1
put(wsA, r, 2, 'Ekterly use is allocated in proportion to each rescue-eligible origin; the true cross-tabulation is not public. Origins are fixed labels for the pre-PHVS regimen, not updated treatment states.', f_sub); r += 2
band(wsA, r, 'Annualised prescription propensities (completed-start proxies divided by the 80% normalisation)', 2, 12); r += 1
put(wsA, r, 2, 'Prevention (XR), by origin', f_bold); put(wsA, r, 5, 'Mature rate', f_bold); put(wsA, r, 8, 'How it is built', f_bold); r += 1
put(wsA, r, 2, 'Oral choice share, mature'); put(wsA, r, 5, '=oral_mature_share', f_norm, PCT); r += 1
put(wsA, r, 2, 'Injectable-origin completed starts'); put(wsA, r, 5, '=category_switch*prevention_oral_preference*oral_mature_share', f_norm, PCT); put(wsA, r, 8, 'switching 15.2% x oral preference 54% x PHVS share 60%', f_sub); INJ = r; r += 1
put(wsA, r, 2, 'Rescue-only completed starts'); put(wsA, r, 5, f'=w_other_prevention/w_acute_only*E{INJ}*entrant_rescue_to_prevention_ratio', f_norm, PCT); put(wsA, r, 8, 'scaled by pool size and the 32/52 entrant ratio', f_sub); ACU = r; r += 1
put(wsA, r, 2, 'Untreated completed starts'); put(wsA, r, 5, f'=w_acute_only/w_diagnosed_untreated*E{ACU}*entrant_naive_to_rescue_ratio*untreated_xr_reach', f_norm, PCT); put(wsA, r, 8, 'scaled by pool size, the 17/32 entrant ratio and 50% reach', f_sub); NAI = r; r += 1
XR_RATE = {}
for k, formula, note in (('orladeyo', '=category_switch*oral_destination_capture/reference_paid_conversion', 'all departures from the incumbent oral, 60% captured'), ('other_prevention', f'=E{INJ}/reference_paid_conversion', ''), ('acute_only', f'=E{ACU}/reference_paid_conversion', ''), ('diagnosed_untreated', f'=E{NAI}/reference_paid_conversion', '')):
    put(wsA, r, 2, f'XR annual prescription rate: {ON[k]}', f_bold); put(wsA, r, 5, formula, f_bold, PCT); put(wsA, r, 8, note, f_sub); XR_RATE[k] = f"'Adoption Build'!$E${r}"; r += 1
r += 1
put(wsA, r, 2, 'Rescue (IR), by rescue origin', f_bold); r += 1
put(wsA, r, 2, 'Ekterly first-8-month forms as annualised propensity'); put(wsA, r, 5, '=1-(1-ekterly_first8_forms/ekterly_launch_addressable)^(12/8)', f_norm, PCT); put(wsA, r, 8, '1,702 forms over 8 months on ~9,000 addressable, annualised', f_sub); EKA = r; r += 1
IR_RATE = {}
for k, formula, note in (('ekterly', '=category_switch*ir_oral_switch_relative*oral_destination_capture/reference_paid_conversion', 'borrows the LTP switching proxy; weakest transfer in the model'), ('injectable', f'=E{EKA}*ir_injectable_opportunity*oral_mature_share', 'first-oral forms propensity x PHVS oral choice'), ('none', '=ir_no_rescue_opportunity*rescue_oral_preference*oral_mature_share', '20% opportunity x 84% oral preference x PHVS share')):
    put(wsA, r, 2, f'IR annual prescription rate: {RN[k]}', f_bold); put(wsA, r, 5, formula, f_bold, PCT); put(wsA, r, 8, note, f_sub); IR_RATE[k] = f"'Adoption Build'!$E${r}"; r += 1
put(wsA, r, 2, 'Mature rates shown; the monthly engines scale the oral-choice terms from the launch share and apply the switching multiplier on each product launch clock. Monthly probability = 1-(1-annual)^(1/12).', f_sub); r += 2
band(wsA, r, 'Retention and paid exposure', 2, 12); r += 1
put(wsA, r, 2, 'Comparator first-year persistence (Orladeyo, 45-day gap)'); put(wsA, r, 5, 0.61, f_input, PCT); put(wsA, r, 9, 'XRPERSIST', f_sub); PERS = r; r += 1
put(wsA, r, 2, 'Share of exits that were efficacy-related (French cohort 13/30)'); put(wsA, r, 5, '=13/30', f_input, PCT); put(wsA, r, 9, 'FRENCH', f_sub); EFF = r; r += 1
put(wsA, r, 2, 'Share of efficacy exits PHVS avoids (assumption)'); put(wsA, r, 5, 0.5, f_input, PCT); AVD = r; r += 1
put(wsA, r, 2, 'Implied XR first-year retention'); put(wsA, r, 5, f'=E{PERS}+(1-E{PERS})*E{EFF}*E{AVD}', f_norm, PCT); put(wsA, r, 8, 'equals the working input 69.45%; no adverse-event benefit assumed', f_sub); r += 1
put(wsA, r, 2, 'Comparator PDC'); put(wsA, r, 5, 0.73, f_input, PCT); PDC = r; r += 1
put(wsA, r, 2, 'Continuous patient-years implied by 61% exponential survival'); put(wsA, r, 5, f'=(1-E{PERS})/(-LN(E{PERS}))', f_norm, NUM3); PY1 = r; r += 1
put(wsA, r, 2, 'Implied paid intensity (PDC / patient-years)'); put(wsA, r, 5, f'=E{PDC}/E{PY1}', f_norm, NUM3); put(wsA, r, 8, 'equals the working input; applied once after survival', f_sub); r += 2
ADOPT_ANNUAL_ROW = r   # filled later with annual summaries
widths(wsA, {'A': 2, 'B': 50, 'C': 11, 'D': 11, 'E': 12, 'F': 12, 'G': 14, 'H': 64, 'I': 11})

# ================================================================== XR / IR ENGINES
CELLS = [(o, s) for o in ORIG for s in RESC]
def engine_sheet(prod, first_name, mature_name, rate_sheet_fn, success_note):
    ws = sheet(f'{prod} Engine', '7F7F7F'); title(ws, f'{prod} monthly patient engine', f'One block per origin cell. Available -> attempts -> access -> active -> retired, conserving each pool. {success_note}')
    month_header(ws); R = {}; r = 8
    def row(key, label, fn, unit='', fmt=NUM, pad=0, note=''):
        nonlocal r; R[key] = r; mrow(ws, r, label, fn, unit, fmt, pad=pad, note=note); r += 1
    wsn = f'{prod} Engine'
    age = CAL['age_' + prod]; on = CAL['on_' + prod]
    row('age', 'Months since launch', lambda j: f"=Calendar!{mcol(j)}${age}", fmt=NUM0)
    row('on', 'Launched (and approved in this branch)', lambda j: f"=Calendar!{mcol(j)}${on}", fmt=NUM0)
    row('ramp', 'Choice ramp fraction', lambda j: f"=MIN(1,MAX(0,{mcol(j)}{R['age']})/choice_months)", fmt=NUM3)
    row('choice', 'PHVS share of oral choices', lambda j: f"=oral_launch_share+(oral_mature_share-oral_launch_share)*{mcol(j)}{R['ramp']}", fmt=PCT)
    row('factor', 'Switching opportunity factor', lambda j: f"=(1+(switch_mature_multiplier-1)*{mcol(j)}{R['ramp']})*IF({mcol(j)}{R['age']}>=60,hazard_year6_multiplier,1)", fmt=NUM3)
    row('g', 'Monthly growth', lambda j: f"=Calendar!{mcol(j)}${CAL['g']}", fmt=PCT2)
    keys = ORIG if prod == 'XR' else RESC
    for k in keys:
        row(f'rate_{k}', f'Annual rate: {(ON if prod=="XR" else RN)[k]}', lambda j, k=k: rate_sheet_fn(k, j, R), fmt=PCT)
        row(f'q_{k}', f'Monthly probability: {(ON if prod=="XR" else RN)[k]}', lambda j, k=k: f"=IF({mcol(j)}{R['on']}=1,1-(1-{mcol(j)}{R[f'rate_{k}']})^(1/12),0)", fmt=PCT2)
    # weights (in pad columns, reversed ages 11..0)
    r += 1; put(ws, r, 1, 'Cohort weights by age (months), oldest to newest', f_bold); r += 1
    R['w_open'] = r; put(ws, r, 1, 'Opening survival weight, first year'); put(ws, r, 4, 'age ->', f_sub)
    for a in range(12): put(ws, r, C0 + a, f'={first_name}^({11-a}/12)', f_norm, NUM3)
    r += 1; R['w_end'] = r; put(ws, r, 1, 'Closing survival weight, first year')
    for a in range(12): put(ws, r, C0 + a, f'={first_name}^({12-a}/12)', f_norm, NUM3)
    r += 1; R['C1'] = r; put(ws, r, 1, 'Exposure per opening member, first year (month)'); put(ws, r, 4, f'=IF({first_name}=1,1/12,(1-{first_name}^(1/12))/(-LN({first_name})))', f_norm, NUM3 + '0')
    r += 1; R['Cm'] = r; put(ws, r, 1, 'Exposure per opening member, mature (month)'); put(ws, r, 4, f'=IF({mature_name}=1,1/12,(1-{mature_name}^(1/12))/(-LN({mature_name})))', f_norm, NUM3 + '0')
    r += 1; R['sm'] = r; put(ws, r, 1, 'Monthly survival, mature'); put(ws, r, 4, f'={mature_name}^(1/12)', f_norm, NUM3 + '0')
    r += 2
    W_OPEN = f"${FIRSTPAD}${R['w_open']}:${LASTPAD}${R['w_open']}"; W_END = f"${FIRSTPAD}${R['w_end']}:${LASTPAD}${R['w_end']}"
    C1, Cm, SM = f"$D${R['C1']}", f"$D${R['Cm']}", f"$D${R['sm']}"
    for (o, s) in CELLS:
        key = f'{o}|{s}'; qk = f"q_{o if prod=='XR' else s}"
        band(ws, r, f'{ON[o]} / {RN[s]}', 1, 4); r += 1
        B = {}
        def brow(k2, label, fn, fmt=NUM, note=''):
            nonlocal r; B[k2] = r; R[(key, k2)] = r; mrow(ws, r, label, fn, '', fmt, note=note); r += 1
        brow('total', 'Pool (grows with population)', lambda j: f"={POOL_CELL[(o, s)]}" if j == 0 else f"={mcol(j-1)}{B['total']}*(1+{mcol(j)}{R['g']})")
        brow('growth_add', 'Net new patients', lambda j: "=0" if j == 0 else f"={mcol(j-1)}{B['total']}*{mcol(j)}{R['g']}")
        brow('released', 'Returning after failed access or re-entry', lambda j: f"={lag(wsn, B['rx'] if 'rx' in B else r+3, j, 'retry_months')}*(1-conversion)+{lag(wsn, r+10, j, 'reentry_delay')}*reentry_fraction", note='retry and re-entry lags')
        brow('pre', 'Available before attempts', lambda j: f"={POOL_CELL[(o, s)]}+{mcol(j)}{B['growth_add']}+{mcol(j)}{B['released']}" if j == 0 else f"={mcol(j-1)}{B['avail']}+{mcol(j)}{B['growth_add']}+{mcol(j)}{B['released']}" if False else None)
        brow('rx', 'Prescription attempts', lambda j: f"={mcol(j)}{B['pre']}*{mcol(j)}{R[qk]}")
        brow('avail', 'Available after attempts', lambda j: f"={mcol(j)}{B['pre']}-{mcol(j)}{B['rx']}")
        brow('starts', 'Paid activations (after access delay)', lambda j: f"={lag(wsn, B['rx'], j, 'delay')}*conversion")
        brow('young_open', 'First-year members, opening', lambda j: f"=SUMPRODUCT({CL(MC0+j-11)}{B['starts']}:{mcol(j)}{B['starts']},{W_OPEN})")
        brow('young_end', 'First-year members, closing', lambda j: f"=SUMPRODUCT({CL(MC0+j-11)}{B['starts']}:{mcol(j)}{B['starts']},{W_END})")
        brow('first_py', 'First-year patient-years', lambda j: f"={mcol(j)}{B['young_open']}*{C1}", fmt=NUM2)
        brow('mature_open', 'Mature members, opening', lambda j: f"={lag(wsn, B['starts'], j, '12')}*{first_name}" if j == 0 else f"={mcol(j-1)}{B['mature_open']}*{SM}+{lag(wsn, B['starts'], j, '12')}*{first_name}")
        brow('py', 'Patient-years (exposure)', lambda j: f"={mcol(j)}{B['first_py']}+{mcol(j)}{B['mature_open']}*{Cm}", fmt=NUM2)
        brow('active', 'Active members, closing', lambda j: f"={mcol(j)}{B['young_end']}+{mcol(j)}{B['mature_open']}*{SM}")
        brow('exits', 'Discontinuations', lambda j: f"={mcol(j)}{B['young_open']}+{mcol(j)}{B['mature_open']}-{mcol(j)}{B['active']}")
        brow('retired', 'Retired, cumulative', lambda j: f"={mcol(j)}{B['exits']}*(1-reentry_fraction)" if j == 0 else f"={mcol(j-1)}{B['retired']}+{mcol(j)}{B['exits']}*(1-reentry_fraction)")
        brow('check', 'Conservation check: pool less (available + pending + cooldown + active + retired)', lambda j: f"={mcol(j)}{B['total']}-({mcol(j)}{B['avail']}+{mcol(j)}{B['active']}+{mcol(j)}{B['retired']}+SUM(${mcol(0)}{B['rx']}:{mcol(j)}{B['rx']})-SUM(${mcol(0)}{B['starts']}:{mcol(j)}{B['starts']})+SUM(${mcol(0)}{B['exits']}:{mcol(j)}{B['exits']})*reentry_fraction-SUM(${mcol(0)}{B['released']}:{mcol(j)}{B['released']}))", fmt=NUM3)
        # fix 'pre' (needs avail row index) and 'released' lags now that rows are known
        for j in range(NM):
            ws.cell(B['pre'], MC0 + j, f"={POOL_CELL[(o, s)]}+{mcol(j)}{B['growth_add']}+{mcol(j)}{B['released']}" if j == 0 else f"={mcol(j-1)}{B['avail']}+{mcol(j)}{B['growth_add']}+{mcol(j)}{B['released']}").font = f_norm
            ws.cell(B['released'], MC0 + j, f"={lag(wsn, B['rx'], j, 'retry_months')}*(1-conversion)+{lag(wsn, B['exits'], j, 'reentry_delay')}*reentry_fraction").font = f_norm
        r += 1
    band(ws, r, 'Totals across origin cells', 1, 4); r += 1
    for k2, label in (('rx', 'Prescription attempts'), ('starts', 'Paid activations'), ('py', 'Patient-years'), ('first_py', 'First-year patient-years'), ('active', 'Active members, closing'), ('exits', 'Discontinuations'), ('total', 'Population')):
        row(f'tot_{k2}', label, lambda j, k2=k2: '=' + '+'.join(f"{mcol(j)}{R[(f'{o}|{s}', k2)]}" for (o, s) in CELLS), fmt=NUM2 if 'py' in k2 else NUM)
    for o in ORIG:
        row(f'rx_{o}', f'Attempts from {ON[o]}', lambda j, o=o: '=' + '+'.join(f"{mcol(j)}{R[(f'{o}|{s}', 'rx')]}" for s in RESC))
        row(f'starts_{o}', f'Activations from {ON[o]}', lambda j, o=o: '=' + '+'.join(f"{mcol(j)}{R[(f'{o}|{s}', 'starts')]}" for s in RESC))
    for s in RESC:
        row(f'rx_r_{s}', f'Attempts from {RN[s]} users', lambda j, s=s: '=' + '+'.join(f"{mcol(j)}{R[(f'{o}|{s}', 'rx')]}" for o in ORIG))
        row(f'starts_r_{s}', f'Activations from {RN[s]} users', lambda j, s=s: '=' + '+'.join(f"{mcol(j)}{R[(f'{o}|{s}', 'starts')]}" for o in ORIG))
    r += 1
    row('conserve', 'Largest conservation error across cells', lambda j: '=MAX(' + ','.join(f"ABS({mcol(j)}{R[(f'{o}|{s}', 'check')]})" for (o, s) in CELLS) + ')', fmt=NUM3)
    return ws, R

def xr_rate(k, j, R):
    base = {'orladeyo': 'category_switch*oral_destination_capture/reference_paid_conversion',
            'other_prevention': f'category_switch*prevention_oral_preference*{mcol(j)}{R["choice"]}/reference_paid_conversion',
            'acute_only': f'w_other_prevention/w_acute_only*category_switch*prevention_oral_preference*{mcol(j)}{R["choice"]}*entrant_rescue_to_prevention_ratio/reference_paid_conversion',
            'diagnosed_untreated': f'w_acute_only/w_diagnosed_untreated*w_other_prevention/w_acute_only*category_switch*prevention_oral_preference*{mcol(j)}{R["choice"]}*entrant_rescue_to_prevention_ratio*entrant_naive_to_rescue_ratio*untreated_xr_reach/reference_paid_conversion'}[k]
    return f'=({base})*{mcol(j)}{R["factor"]}'
def ir_rate(k, j, R):
    base = {'ekterly': 'category_switch*ir_oral_switch_relative*oral_destination_capture/reference_paid_conversion',
            'injectable': f'(1-(1-ekterly_first8_forms/ekterly_launch_addressable)^(12/8))*ir_injectable_opportunity*{mcol(j)}{R["choice"]}',
            'none': f'ir_no_rescue_opportunity*rescue_oral_preference*{mcol(j)}{R["choice"]}'}[k]
    return f'=({base})*{mcol(j)}{R["factor"]}'
wsX, RX = engine_sheet('XR', 'xr_first_retention', 'xr_mature_retention', xr_rate, 'Identical in the both-success and XR-only branches; zero when XR fails.')
wsR, RI = engine_sheet('IR', 'ir_retention', 'ir_retention', ir_rate, 'Population flows do not depend on XR; only dose budgets do (Revenue tab).')
wb._sheets.remove(wsA); wb._sheets.insert(wb._sheets.index(wsX), wsA)  # keep Adoption Build before engines

# ================================================================== REVENUE (monthly)
wsV = sheet('Revenue', '7F7F7F'); title(wsV, 'Monthly quantity, price and revenue', 'XR once; IR twice (with XR overlap for the both-success branch, without it for the IR-only branch). USD millions.')
month_header(wsV); RV = {}; r = 8
def vrow(key, label, fn, unit='', fmt=NUM, pad=0, note=''):
    global r; RV[key] = r; mrow(wsV, r, label, fn, unit, fmt, pad=pad, note=note); r += 1
band(wsV, r, 'Rescue dose budgets by origin cell', 1, 4); r += 1
vrow('mig', 'Non-PHVS prophylaxis migration', lambda j: f"=nonphvs_ltp_shift*MIN(1,MAX(0,Calendar!{mcol(j)}${CAL['age_IR']}+1)/120)", fmt=PCT)
for variant, tag in (('A', 'with XR overlap'), ('B', 'IR only, no XR')):
    for (o, s) in CELLS:
        key = f'{o}|{s}'
        def mk(kind, o=o, s=s, key=key, variant=variant):
            def f(j):
                OV = f"IF('XR Engine'!{mcol(j)}${RX[(key,'total')]}=0,0,MIN(1,'XR Engine'!{mcol(j)}${RX[(key,'py')]}*12/'XR Engine'!{mcol(j)}${RX[(key,'total')]}))" if variant == 'A' else '0'
                BASE = 'ir_doses_ltp' if o in ('orladeyo', 'other_prevention') else f"((1-{mcol(j)}${RV['mig']})*ir_doses_other+{mcol(j)}${RV['mig']}*ir_doses_ltp)"
                if kind == 'ov': return f"={OV}"
                ovr = f"{mcol(j)}{RV[(variant,key,'ov')]}"
                if kind == 'budget': return f"=((1-{ovr})*{BASE}+{ovr}*ir_doses_ltp*dual_use)*ir_quantity_multiplier"
                if kind == 'initial': return f"=ir_initial_doses*(1-{ovr}+{ovr}*dual_use)"
                st = f"'IR Engine'!{mcol(j)}${RI[(key,'starts')]}"; fpy = f"'IR Engine'!{mcol(j)}${RI[(key,'first_py')]}"; py = f"'IR Engine'!{mcol(j)}${RI[(key,'py')]}"
                if kind == 'paid_init': return f"={st}*{mcol(j)}{RV[(variant,key,'initial')]}*ir_initial_paid_fraction"
                if kind == 'free_init': return f"={st}*{mcol(j)}{RV[(variant,key,'initial')]}*(1-ir_initial_paid_fraction)"
                if kind == 'repl': return f"={fpy}*({mcol(j)}{RV[(variant,key,'budget')]}-{mcol(j)}{RV[(variant,key,'initial')]})+({py}-{fpy})*{mcol(j)}{RV[(variant,key,'budget')]}"
            return f
        lab = f'{ON[o]} / {RN[s]} ({tag})'
        for kind, l2, fmt in (('ov', 'XR overlap share', PCT), ('budget', 'Annual dose budget', NUM2), ('initial', 'Initial stock per start', NUM2), ('paid_init', 'Paid initial doses', NUM2), ('free_init', 'Free initial doses', NUM2), ('repl', 'Replacement doses', NUM2)):
            RV[(variant, key, kind)] = r; mrow(wsV, r, f'{lab}: {l2}', mk(kind), '', fmt); r += 1
r += 1
band(wsV, r, 'XR: paid treatment-years, price and revenue', 1, 4); r += 1
def product_rows(prod, tag, units_fn, free_fn=None):
    global r
    pre = f'{prod}{tag}'
    vrow(f'{pre}_raw', f'{prod} raw units (pre-generic)', units_fn, fmt=NUM2)
    put(wsV, r, 1, f'{prod} protected-year monthly unit anchor'); put(wsV, r, 4, f"=AVERAGEIFS({mrng('Revenue', RV[f'{pre}_raw'])},{mrng('Revenue', 3)},loe_year_{prod}-1)", f_norm, NUM2); RV[f'{pre}_anchor'] = r; r += 1
    vrow(f'{pre}_units', f'{prod} monetised units', lambda j: f"=IF({mcol(j)}$3<loe_year_{prod},{mcol(j)}{RV[f'{pre}_raw']},$D${RV[f'{pre}_anchor']}*INDEX({prod.lower()}_q_idx,MATCH({mcol(j)}$3,loe_years,0)))", fmt=NUM2)
    price = 'xr_wac_28days*365/28*xr_net_fraction*xr_price_multiplier' if prod == 'XR' else 'ir_wac_dose*ir_net_fraction*ir_price_multiplier'
    vrow(f'{pre}_price', f'{prod} net price, USD', lambda j: f"={price}*(1+price_growth)^MAX(0,{mcol(j)}$3-2027)*INDEX({prod.lower()}_p_idx,MATCH({mcol(j)}$3,loe_years,0))", fmt=USD0)
    vrow(f'{pre}_US', f'{prod} US net revenue', lambda j: f"={mcol(j)}{RV[f'{pre}_units']}*{mcol(j)}{RV[f'{pre}_price']}/1000000", fmt=NUM2)
    vrow(f'{pre}_phase', f'{prod} ex-US rollout phase', lambda j: f"=MIN(1,MAX(0,({mcol(j)}$2-Calendar!{mcol(j)}${CAL['launch_idx_'+prod]}-exus_lag+1)/exus_phase))", fmt=PCT)
    vrow(f'{pre}_dem_pre', f'{prod} ex-US demand before generics', lambda j: f"={lag('Revenue', RV[f'{pre}_US'], j, 'exus_lag')}*{prod.lower()}_exus_ratio*{mcol(j)}{RV[f'{pre}_phase']}", fmt=NUM2)
    put(wsV, r, 1, f'{prod} protected-year ex-US anchor'); put(wsV, r, 4, f"=AVERAGEIFS({mrng('Revenue', RV[f'{pre}_dem_pre'])},{mrng('Revenue', 3)},loe_year_{prod}-1)", f_norm, NUM2); RV[f'{pre}_exanchor'] = r; r += 1
    vrow(f'{pre}_dem', f'{prod} ex-US end-market demand', lambda j: f"=IF({mcol(j)}$3<loe_year_{prod},{mcol(j)}{RV[f'{pre}_dem_pre']},$D${RV[f'{pre}_exanchor']}*INDEX({prod.lower()}_r_idx,MATCH({mcol(j)}$3,loe_years,0)))", fmt=NUM2)
    vrow(f'{pre}_direct', f'{prod} ex-US direct revenue', lambda j: f"={mcol(j)}{RV[f'{pre}_dem']}*exus_direct_fraction", fmt=NUM2)
    vrow(f'{pre}_roy', f'{prod} ex-US royalty revenue', lambda j: f"={mcol(j)}{RV[f'{pre}_dem']}*(1-exus_direct_fraction)*partner_royalty", fmt=NUM2)
    vrow(f'{pre}_rec', f'{prod} recognised revenue', lambda j: f"={mcol(j)}{RV[f'{pre}_US']}+{mcol(j)}{RV[f'{pre}_direct']}+{mcol(j)}{RV[f'{pre}_roy']}", fmt=NUM2)
    if free_fn:
        vrow(f'{pre}_free_raw', f'{prod} free initial doses (pre-generic)', free_fn, fmt=NUM2)
        put(wsV, r, 1, f'{prod} protected-year free-dose anchor'); put(wsV, r, 4, f"=AVERAGEIFS({mrng('Revenue', RV[f'{pre}_free_raw'])},{mrng('Revenue', 3)},loe_year_{prod}-1)", f_norm, NUM2); RV[f'{pre}_freeanchor'] = r; r += 1
        vrow(f'{pre}_free', f'{prod} free units', lambda j: f"=IF({mcol(j)}$3<loe_year_{prod},{mcol(j)}{RV[f'{pre}_free_raw']},$D${RV[f'{pre}_freeanchor']}*INDEX(ir_q_idx,MATCH({mcol(j)}$3,loe_years,0)))", fmt=NUM2)
    r += 1
product_rows('XR', '', lambda j: f"='XR Engine'!{mcol(j)}${RX['tot_py']}*xr_paid_intensity")
for variant, tag, label in (('A', '_A', 'IR with XR overlap (both-success branch)'), ('B', '_B', 'IR without XR (IR-only branch)')):
    band(wsV, r, label, 1, 4); r += 1
    product_rows('IR', tag, lambda j, v=variant: '=' + '+'.join(f"{mcol(j)}{RV[(v,f'{o}|{s}','paid_init')]}+{mcol(j)}{RV[(v,f'{o}|{s}','repl')]}" for (o, s) in CELLS),
                 lambda j, v=variant: '=' + '+'.join(f"{mcol(j)}{RV[(v,f'{o}|{s}','free_init')]}" for (o, s) in CELLS))
band(wsV, r, 'Summary rows', 1, 4); r += 1
vrow('IR_A_paidinit', 'IR paid initial doses (both)', lambda j: '=' + '+'.join(f"{mcol(j)}{RV[('A',f'{o}|{s}','paid_init')]}" for (o, s) in CELLS), fmt=NUM2)
vrow('IR_A_repl', 'IR replacement doses (both)', lambda j: '=' + '+'.join(f"{mcol(j)}{RV[('A',f'{o}|{s}','repl')]}" for (o, s) in CELLS), fmt=NUM2)
vrow('overlap', 'Overlap member equivalents (both)', lambda j: '=' + '+'.join(f"{mcol(j)}{RV[('A',f'{o}|{s}','ov')]}*'IR Engine'!{mcol(j)}${RI[(f'{o}|{s}','py')]}*12" for (o, s) in CELLS), fmt=NUM2)

# ================================================================== FINANCIALS (monthly, four branches)
wsF = sheet('Fin Monthly', '7F7F7F'); title(wsF, 'Monthly financials by clinical branch', 'Economic cash-compensation forecast, USD millions. Four branches: both approved, IR only, XR only, neither.')
month_header(wsF); RF = {}; r = 8
STATES = [('both', 1, 1, 'XR', 'IR_A'), ('IR_only', 1, 0, None, 'IR_B'), ('XR_only', 0, 1, 'XR', None), ('neither', 0, 0, None, None)]
SNAME = {'both': 'Both approved', 'IR_only': 'IR approved, XR fails', 'XR_only': 'XR approved, IR fails', 'neither': 'Neither approved'}
for st, sIR, sXR, xrset, irset in STATES:
    band(wsF, r, SNAME[st], 1, 4); r += 1; B = {}
    def frow(key, label, fn, fmt=NUM2, note=''):
        global r; B[key] = r; RF[(st, key)] = r; mrow(wsF, r, label, fn, '', fmt, note=note); r += 1
    z = lambda j: '=0'
    frow('st_IR', 'IR costs continue', lambda j: f"=MAX({sIR},Calendar!{mcol(j)}${CAL['pre_dec_IR']})", fmt=NUM0)
    frow('st_XR', 'XR costs continue', lambda j: f"=MAX({sXR},Calendar!{mcol(j)}${CAL['pre_dec_XR']})", fmt=NUM0)
    frow('ongoing', 'Operating', (lambda j: f"=Calendar!{mcol(j)}${CAL['ongoing_fail']}") if st == 'neither' else (lambda j: '=1'), fmt=NUM0)
    frow('liq', 'Liquidation month', (lambda j: f"=Calendar!{mcol(j)}${CAL['is_shutdown']}") if st == 'neither' else z, fmt=NUM0)
    frow('nprod', 'Active products', lambda j: f"={mcol(j)}{B['st_IR']}+{mcol(j)}{B['st_XR']}", fmt=NUM0)
    frow('platform', 'Fixed platform retained', lambda j: f"=IF({mcol(j)}{B['nprod']}=2,1,IF({mcol(j)}{B['nprod']}=1,single_product_platform_fraction,winddown_platform_fraction))", fmt=PCT)
    frow('rdf', 'R&D retained', lambda j: f"=rd_IR_fraction*{mcol(j)}{B['st_IR']}+(1-rd_IR_fraction)*{mcol(j)}{B['st_XR']}", fmt=PCT)
    frow('factor', 'Cost factor (inflation x runoff x scale)', lambda j: f"=Calendar!{mcol(j)}${CAL['infl']}*Calendar!{mcol(j)}${CAL['runoff']}*cost_scale", fmt=NUM3)
    xr_rec = (lambda j: f"=Revenue!{mcol(j)}${RV['XR_rec']}") if xrset else z
    ir_rec = (lambda j: f"=Revenue!{mcol(j)}${RV[irset+'_rec']}") if irset else z
    frow('rev_IR', 'IR revenue', ir_rec); frow('rev_XR', 'XR revenue', xr_rec)
    frow('revenue', 'Revenue', lambda j: f"=({mcol(j)}{B['rev_IR']}+{mcol(j)}{B['rev_XR']})*{mcol(j)}{B['ongoing']}")
    frow('US', 'US revenue', lambda j: '=' + '+'.join([f"Revenue!{mcol(j)}${RV[s2+'_US']}" for s2 in ([xrset] if xrset else []) + ([irset] if irset else [])] or ['0']))
    frow('exdirect', 'Ex-US direct revenue', lambda j: '=' + '+'.join([f"Revenue!{mcol(j)}${RV[s2+'_direct']}" for s2 in ([xrset] if xrset else []) + ([irset] if irset else [])] or ['0']))
    frow('endmarket', 'End-market sales (royalty base)', lambda j: '=' + '+'.join([f"Revenue!{mcol(j)}${RV[s2+'_US']}+Revenue!{mcol(j)}${RV[s2+'_dem']}" for s2 in ([xrset] if xrset else []) + ([irset] if irset else [])] or ['0']))
    frow('free_cost', 'Manufacturing cost of free doses', (lambda j: f"=Revenue!{mcol(j)}${RV[irset+'_free']}*Revenue!{mcol(j)}${RV[irset+'_price']}*manufacture_rate/1000000") if irset else z, fmt=NUM3)
    frow('rd', 'Research and development', lambda j: f"=({mcol(j)}{B['ongoing']})*(INDEX(cost_rd,MATCH(Calendar!{mcol(j)}${CAL['sched_year']},cost_years,0))*{mcol(j)}{B['factor']}*{mcol(j)}{B['rdf']}/12+IF(AND({sIR}=1,{mcol(j)}$2=Calendar!{mcol(j)}${CAL['dec_IR']}),remaining_milestone_eurm*fx,0))")
    frow('commercial', 'Commercial', lambda j: f"=({mcol(j)}{B['ongoing']})*(INDEX(cost_comm,MATCH(Calendar!{mcol(j)}${CAL['sched_year']},cost_years,0))*{mcol(j)}{B['factor']}*{mcol(j)}{B['platform']}/12+{mcol(j)}{B['US']}*commercial_variable_us+{mcol(j)}{B['exdirect']}*commercial_variable_exus)")
    frow('milestone', 'BRAIN milestone at IR approval (inside R&D)', lambda j: f"=({mcol(j)}{B['ongoing']})*IF(AND({sIR}=1,{mcol(j)}$2=Calendar!{mcol(j)}${CAL['dec_IR']}),remaining_milestone_eurm*fx,0)")
    frow('comm_var', 'Variable commercial cost (inside commercial)', lambda j: f"=({mcol(j)}{B['ongoing']})*({mcol(j)}{B['US']}*commercial_variable_us+{mcol(j)}{B['exdirect']}*commercial_variable_exus)")
    frow('ga', 'General and administrative', lambda j: f"=({mcol(j)}{B['ongoing']})*INDEX(cost_ga,MATCH(Calendar!{mcol(j)}${CAL['sched_year']},cost_years,0))*{mcol(j)}{B['factor']}*{mcol(j)}{B['platform']}/12")
    frow('manufacture', 'Manufacturing and distribution', lambda j: f"=({mcol(j)}{B['ongoing']})*(({mcol(j)}{B['US']}+{mcol(j)}{B['exdirect']})*manufacture_rate+{mcol(j)}{B['free_cost']})")
    frow('license', 'BRAIN royalty', lambda j: f"=({mcol(j)}{B['ongoing']})*{mcol(j)}{B['endmarket']}*license_rate")
    frow('capex', 'Capital expenditure', lambda j: f"=({mcol(j)}{B['ongoing']})*(capex_fixed_usdm/12*Calendar!{mcol(j)}${CAL['runoff']}+capex_fraction*({mcol(j)}{B['US']}+{mcol(j)}{B['exdirect']}))", fmt=NUM3)
    frow('dep', 'Depreciation', lambda j: f"=({mcol(j)}{B['ongoing']})*MIN({('open_ppe_and_rou*fx' if j == 0 else mcol(j-1)+str(r+1))}+{mcol(j)}{B['capex']},{('open_ppe_and_rou*fx' if j == 0 else mcol(j-1)+str(r+1))}/depreciation_months+{mcol(j)}{B['capex']}/(2*depreciation_months))", fmt=NUM3)
    frow('ppe', 'PP&E, closing', lambda j: f"=IF({mcol(j)}{B['ongoing']}=0,0,IF({mcol(j)}{B['liq']}=1,0,{('open_ppe_and_rou*fx' if j == 0 else mcol(j-1)+str(r))}+{mcol(j)}{B['capex']}-{mcol(j)}{B['dep']}))", fmt=NUM3)
    frow('ebit', 'Operating income', lambda j: f"={mcol(j)}{B['revenue']}-{mcol(j)}{B['manufacture']}-{mcol(j)}{B['license']}-{mcol(j)}{B['rd']}-{mcol(j)}{B['commercial']}-{mcol(j)}{B['ga']}-{mcol(j)}{B['dep']}")
    frow('nol_used', 'NOL used', lambda j: f"=MIN({('0' if j == 0 else mcol(j-1)+str(r+1))},MAX(0,{mcol(j)}{B['ebit']})*nol_utilization_cap)")
    frow('nol', 'NOL pool, closing', lambda j: f"=IF({mcol(j)}{B['ongoing']}=0,0,{('0' if j == 0 else mcol(j-1)+str(r))}-{mcol(j)}{B['nol_used']}+MAX(0,-{mcol(j)}{B['ebit']}))")
    frow('tax', 'Tax', lambda j: f"=({mcol(j)}{B['ongoing']})*MAX(tax_floor_usdm/12*Calendar!{mcol(j)}${CAL['runoff']},MAX(0,{mcol(j)}{B['ebit']}-{mcol(j)}{B['nol_used']})*tax_rate)")
    frow('ni', 'Net income', lambda j: f"={mcol(j)}{B['ebit']}-{mcol(j)}{B['tax']}")
    frow('nwc', 'Net working capital, closing', lambda j: f"=IF({mcol(j)}{B['ongoing']}=0,0,IF({mcol(j)}{B['liq']}=1,0,open_nwc*fx+nwc_fraction*({mcol(j)}{B['US']}+{mcol(j)}{B['exdirect']})*12))")
    frow('dnwc', 'Change in NWC', lambda j: f"=({mcol(j)}{B['ongoing']})*((open_nwc*fx+nwc_fraction*({mcol(j)}{B['US']}+{mcol(j)}{B['exdirect']})*12)-{('open_nwc*fx' if j == 0 else mcol(j-1)+str(B['nwc']))})")
    frow('fcf', 'Free cash flow', lambda j: f"={mcol(j)}{B['ni']}+{mcol(j)}{B['dep']}-{mcol(j)}{B['capex']}-{mcol(j)}{B['dnwc']}")
    frow('cash_pre', 'Cash before financing', lambda j: f"={('open_cash*fx' if j == 0 else mcol(j-1)+str(r+6))}+{mcol(j)}{B['fcf']}")
    frow('fund_net', 'Net funding raised', lambda j: f"=IF(AND({mcol(j)}{B['ongoing']}=1,{mcol(j)}{B['cash_pre']}<minimum_cash_m,{mcol(j)}{B['nprod']}>0),cash_raise_target_multiple*minimum_cash_m-{mcol(j)}{B['cash_pre']},0)")
    frow('fund_gross', 'Gross funding', lambda j: f"={mcol(j)}{B['fund_net']}/(1-issue_fee)")
    frow('fund_fee', 'Issuance fees', lambda j: f"={mcol(j)}{B['fund_gross']}-{mcol(j)}{B['fund_net']}", fmt=NUM3)
    frow('new_sh', 'Shares issued', lambda j: f"={mcol(j)}{B['fund_gross']}/issue_price", fmt=NUM3)
    frow('shares', 'Shares outstanding, closing', lambda j: f"={('open_shares_m' if j == 0 else mcol(j-1)+str(r))}+{mcol(j)}{B['new_sh']}", fmt=NUM3)
    frow('cash', 'Cash, closing', lambda j: f"=IF({mcol(j)}{B['ongoing']}=0,0,IF({mcol(j)}{B['liq']}=1,0,{mcol(j)}{B['cash_pre']}+{mcol(j)}{B['fund_net']}))")
    frow('dist', 'Liquidation distribution', (lambda j: f"=IF({mcol(j)}{B['liq']}=1,MAX(0,{mcol(j)}{B['cash_pre']}+{mcol(j)}{B['fund_net']}+open_nwc*fx+nwc_fraction*({mcol(j)}{B['US']}+{mcol(j)}{B['exdirect']})*12+open_other_net_assets*fx),0)") if st == 'neither' else z)
    frow('other', 'Other net assets', lambda j: f"=IF({mcol(j)}{B['ongoing']}=0,0,IF({mcol(j)}{B['liq']}=1,0,open_other_net_assets*fx))", fmt=NUM3)
    frow('equity', 'Equity, closing', lambda j: f"=IF({mcol(j)}{B['ongoing']}=0,0,IF({mcol(j)}{B['liq']}=1,0,{('open_equity*fx' if j == 0 else mcol(j-1)+str(r))}+{mcol(j)}{B['ni']}+{mcol(j)}{B['fund_net']}))")
    frow('bal', 'Balance check', lambda j: f"={mcol(j)}{B['cash']}+{mcol(j)}{B['nwc']}+{mcol(j)}{B['ppe']}+{mcol(j)}{B['other']}-{mcol(j)}{B['equity']}", fmt=NUM3)
    r += 1

# ================================================================== VALUATION
wsVal = sheet('Valuation', 'C00000'); title(wsVal, 'Engine parity: live valuation of the central case by clinical branch')
r = 4; band(wsVal, r, 'Branch probabilities', 2, 8); r += 1
put(wsVal, r, 2, 'Both approved'); put(wsVal, r, 4, '=ir_pos*xr_pos/shared_success', f_norm, PCT); name('w_both', f'Valuation!$D${r}'); put(wsVal, r, 6, 'P(IR) x P(XR) / shared gate', f_sub); r += 1
put(wsVal, r, 2, 'IR only'); put(wsVal, r, 4, '=ir_pos-w_both', f_norm, PCT); name('w_IR_only', f'Valuation!$D${r}'); r += 1
put(wsVal, r, 2, 'XR only'); put(wsVal, r, 4, '=xr_pos-w_both', f_norm, PCT); name('w_XR_only', f'Valuation!$D${r}'); r += 1
put(wsVal, r, 2, 'Neither'); put(wsVal, r, 4, '=1-ir_pos-xr_pos+w_both', f_norm, PCT); name('w_neither', f'Valuation!$D${r}'); r += 1
put(wsVal, r, 2, 'Check: sum'); put(wsVal, r, 4, '=w_both+w_IR_only+w_XR_only+w_neither', f_norm, PCT); r += 2
VALROW = {}
for tag, dlabel in (('val', 'Valuation date'), ('tgt', 'Target date')):
    band(wsVal, r, '', 2, 8); wsVal.cell(r, 2).value = ('="Value today, "&TEXT(valuation_date,"d mmm yyyy")' if tag == 'val' else '="12-month price target, "&TEXT(target_date,"d mmm yyyy")&": the same DCF rolled forward one year; cash flows before that date accrue to cash, the rest are discounted to it"'); r += 1
    heads = ['USD millions', 'Both', 'IR only', 'XR only', 'Neither']
    for i, h in enumerate(heads): put(wsVal, r, 2 + i, h, f_hdr, fill=fill_hdr)
    r += 1; R0 = r
    lines = [('cash', 'Cash at date (partial month prorated)'), ('shares', 'Ordinary shares at date, millions'), ('pv', 'PV of free cash flow after date'), ('fees', 'PV of issuance fees'), ('tv', 'PV of declining residual after 2055'), ('liq', 'PV of liquidation distribution'), ('equity', 'Equity value'), ('strike', 'Option strike, USD'), ('basic_rsu', 'Shares plus RSUs, millions'), ('price', 'Value per share (treasury-stock)'), ('diluted', 'Diluted shares, millions'), ('funding', 'Future net funding needed after date')]
    for li, (k, lab) in enumerate(lines):
        put(wsVal, R0 + li, 2, lab, f_bold if k in ('price', 'equity') else f_norm); VALROW[(tag, k)] = R0 + li
    for ci, (st, sIR, sXR, _, _) in enumerate(STATES):
        c = 3 + ci; L = CL(c); fcf = mrng('Fin Monthly', RF[(st, 'fcf')]); frac = mrng('Calendar', CAL[f'frac_{tag}']); full = mrng('Calendar', CAL[f'full_{tag}']); df = mrng('Calendar', CAL[f'df_{tag}']); aft = mrng('Calendar', CAL[f'after_{tag}'])
        put(wsVal, VALROW[(tag, 'cash')], c, f"=open_cash*fx+SUMPRODUCT({fcf},{frac})+SUMPRODUCT({mrng('Fin Monthly', RF[(st,'fund_net')])}-{mrng('Fin Monthly', RF[(st,'dist')])},{full})", f_norm, NUM)
        put(wsVal, VALROW[(tag, 'shares')], c, f"=IF(SUM({full})=0,open_shares_m,INDEX({mrng('Fin Monthly', RF[(st,'shares')])},1,SUM({full})))", f_norm, NUM3)
        put(wsVal, VALROW[(tag, 'pv')], c, f"=SUMPRODUCT({fcf},1-{frac},{df})", f_norm, NUM)
        put(wsVal, VALROW[(tag, 'fees')], c, f"=SUMPRODUCT({mrng('Fin Monthly', RF[(st,'fund_fee')])},{aft},{df})", f_norm, NUM)
        tv = f"SUMPRODUCT({fcf},{mrng('Calendar', CAL['is_end_year'])})*(1+tail_growth)/(discount_rate-tail_growth)*(1+discount_rate)^(-(DATE(end_year,12,31)-{'valuation_date' if tag=='val' else 'target_date'})/365.25)"
        put(wsVal, VALROW[(tag, 'tv')], c, '=0' if st == 'neither' else f'={tv}', f_norm, NUM)
        put(wsVal, VALROW[(tag, 'liq')], c, f"=SUMPRODUCT({mrng('Fin Monthly', RF[(st,'dist')])},{aft},{df})" if st == 'neither' else '=0', f_norm, NUM)
        eq = f"MAX(0,{L}{VALROW[(tag,'cash')]}+{L}{VALROW[(tag,'pv')]}+{L}{VALROW[(tag,'tv')]}-{L}{VALROW[(tag,'fees')]})" if st != 'neither' else f"{L}{VALROW[(tag,'liq')]}*{L}{VALROW[(tag,'shares')]}/INDEX({mrng('Fin Monthly', RF[(st,'shares')])},1,{NM})"
        put(wsVal, VALROW[(tag, 'equity')], c, f'={eq}', f_bold, NUM)
        put(wsVal, VALROW[(tag, 'strike')], c, '=open_option_strike_eur*fx', f_norm, USD)
        put(wsVal, VALROW[(tag, 'basic_rsu')], c, f"={L}{VALROW[(tag,'shares')]}+open_rsus_m", f_norm, NUM3)
        E, S, BR = f"{L}{VALROW[(tag,'equity')]}", f"{L}{VALROW[(tag,'strike')]}", f"{L}{VALROW[(tag,'basic_rsu')]}"
        put(wsVal, VALROW[(tag, 'price')], c, f"=IF({E}/{BR}<={S},MAX(0,{E}/{BR}),({E}+open_options_m*{S})/({BR}+open_options_m))", f_bold, USD)
        put(wsVal, VALROW[(tag, 'diluted')], c, f"=IF({E}/{BR}<={S},{BR},{BR}+open_options_m*(1-{S}/{L}{VALROW[(tag,'price')]}))", f_norm, NUM3)
        put(wsVal, VALROW[(tag, 'funding')], c, f"=SUMPRODUCT({mrng('Fin Monthly', RF[(st,'fund_net')])},{aft})", f_norm, NUM)
    r = R0 + len(lines)
    put(wsVal, r, 2, 'Probability-weighted value per share', f_bold); put(wsVal, r, 3, f"=w_both*C{VALROW[(tag,'price')]}+w_IR_only*D{VALROW[(tag,'price')]}+w_XR_only*E{VALROW[(tag,'price')]}+w_neither*F{VALROW[(tag,'price')]}", f_bold, USD, fill=fill_key)
    name('value_today' if tag == 'val' else 'value_target', f'Valuation!$C${r}'); VALROW[(tag, 'expected')] = r; r += 2
name('price_both_target', f"Valuation!$C${VALROW[('tgt','price')]}"); name('price_both_today', f"Valuation!$C${VALROW[('val','price')]}")
band(wsVal, r, 'Versus the market', 2, 8); r += 1
put(wsVal, r, 2, 'Share price'); put(wsVal, r, 3, '=share_price', f_norm, USD); put(wsVal, r, 5, '=TEXT(price_date,"d mmm yyyy")', f_sub); r += 1
put(wsVal, r, 2, 'Expected value in 12 months'); put(wsVal, r, 3, '=value_target', f_norm, USD); r += 1
put(wsVal, r, 2, 'Monte Carlo expected value in 12 months (saved run)'); put(wsVal, r, 3, '=mc_mean_target', f_norm, USD); r += 1
put(wsVal, r, 2, 'Price target (Monte Carlo mean; DCF tab)', f_bold); put(wsVal, r, 3, '=mc_mean_target', f_bold, USD, fill=fill_key); name('price_target_live', f'Valuation!$C${r}'); r += 1
put(wsVal, r, 2, 'Implied 12-month return'); put(wsVal, r, 3, '=price_target/share_price-1', f_norm, PCT); r += 1
put(wsVal, r, 2, 'Return on the working point'); put(wsVal, r, 3, '=value_target/share_price-1', f_norm, PCT); r += 1
put(wsVal, r, 2, 'Return if both products are approved'); put(wsVal, r, 3, '=price_both_target/share_price-1', f_norm, PCT); r += 1
put(wsVal, r, 2, 'Shares outstanding at the price date, millions'); put(wsVal, r, 3, '=open_shares_m', f_norm, NUM3); r += 1
put(wsVal, r, 2, 'Market capitalisation'); put(wsVal, r, 3, '=share_price*open_shares_m', f_norm, NUM); r += 1
put(wsVal, r, 2, 'Future funding is a modelled requirement, not committed capital. Post-date funding is assumed at fair value, so negative cash flows carry the cost once; only the issuance fee is charged.', f_sub); r += 2
band(wsVal, r, 'Risk-adjusted DCF walk at the target date: product cash flows weighted by the branch probabilities (85.3% both, 4.7% IR only, 4.7% XR only); the 5.3% neither branch enters as liquidation value', 2, 8); r += 1
DY = list(range(2027, END_YEAR + 1)); DC0 = 3
put(wsVal, r, 2, 'Year', f_hdr, fill=fill_hdr)
for i, y in enumerate(DY):
    put(wsVal, r, DC0 + i, y, f_hdr, fill=fill_hdr, align=Alignment(horizontal='right')); wsVal.column_dimensions[CL(DC0 + i)].width = max(wsVal.column_dimensions[CL(DC0 + i)].width or 0, 9.5)
    if y >= 2032: wsVal.column_dimensions[CL(DC0 + i)].outlineLevel = 1; wsVal.column_dimensions[CL(DC0 + i)].hidden = False
DYR = r; r += 1
yrs = mrng('Fin Monthly', 3); fr = mrng('Calendar', CAL['frac_tgt']); dfr = mrng('Calendar', CAL['df_tgt'])
WN = {'both': 'w_both', 'IR_only': 'w_IR_only', 'XR_only': 'w_XR_only'}
def wsum(key, extra=''):
    return '+'.join(f"{WN[st]}*SUMPRODUCT({mrng('Fin Monthly', RF[(st, key)])}{extra})" for st in WN)
def wrow(label, key, extra, fmt=NUM, bold=False):
    global r; put(wsVal, r, 2, label, f_bold if bold else f_norm)
    for i in range(len(DY)): put(wsVal, r, DC0 + i, '=' + wsum(key, extra.replace('YY', f'{CL(DC0+i)}${DYR}')), f_bold if bold else f_norm, fmt)
    r += 1; return r - 1
W_REV = wrow('Expected revenue, USD m', 'revenue', f",--({yrs}=YY)")
W_OPX = wrow('Expected operating costs and tax, USD m', 'revenue', f",--({yrs}=YY)")
for i in range(len(DY)):
    L = CL(DC0 + i); wsVal.cell(W_OPX, DC0 + i).value = '=' + '+'.join(f"{WN[st]}*(SUMPRODUCT({mrng('Fin Monthly', RF[(st,'revenue')])},--({yrs}={L}${DYR}))-SUMPRODUCT({mrng('Fin Monthly', RF[(st,'ni')])},--({yrs}={L}${DYR})))" for st in WN)
W_FCF = wrow('Expected free cash flow after the target date, USD m', 'fcf', f",1-{fr},--({yrs}=YY)", bold=True)
W_PV = wrow('Present value at the target date, USD m', 'fcf', f",1-{fr},{dfr},--({yrs}=YY)")
put(wsVal, r, 2, 'Effective discount factor'); [put(wsVal, r, DC0 + i, f"=IF({CL(DC0+i)}{W_FCF}=0,\"\",{CL(DC0+i)}{W_PV}/{CL(DC0+i)}{W_FCF})", f_norm, NUM3) for i in range(len(DY))]; r += 1
put(wsVal, r, 2, 'Cumulative present value, USD m'); [put(wsVal, r, DC0 + i, f"=SUM($C${W_PV}:{CL(DC0+i)}${W_PV})", f_norm, NUM) for i in range(len(DY))]; r += 2
wv = lambda key: f"=w_both*C{VALROW[('tgt',key)]}+w_IR_only*D{VALROW[('tgt',key)]}+w_XR_only*E{VALROW[('tgt',key)]}"
put(wsVal, r, 2, 'Sum of present values, 2027-2055'); put(wsVal, r, 3, f"=SUM(C{W_PV}:{CL(DC0+len(DY)-1)}{W_PV})", f_norm, NUM); WS0 = r; r += 1
put(wsVal, r, 2, 'Expected residual after 2055 (declining perpetuity)'); put(wsVal, r, 3, wv('tv'), f_norm, NUM); r += 1
put(wsVal, r, 2, 'Expected value if neither is approved (liquidation proceeds to current holders)'); put(wsVal, r, 3, f"=w_neither*F{VALROW[('tgt','equity')]}", f_norm, NUM); r += 1
put(wsVal, r, 2, 'Expected cash at the target date'); put(wsVal, r, 3, wv('cash'), f_norm, NUM); r += 1
put(wsVal, r, 2, 'Less expected issuance fees'); put(wsVal, r, 3, '=-(' + wv('fees')[1:] + ')', f_norm, NUM); r += 1
put(wsVal, r, 2, 'Expected equity value', f_bold); put(wsVal, r, 3, f"=SUM(C{WS0}:C{r-1})", f_bold, NUM); WEQ = r; r += 1
put(wsVal, r, 2, 'Check: equals the probability-weighted branch equity'); put(wsVal, r, 3, f"=C{WEQ}-({wv('equity')[1:]}+w_neither*F{VALROW[('tgt','equity')]})", f_norm, NUM3); r += 1
put(wsVal, r, 2, 'Expected diluted shares, millions'); put(wsVal, r, 3, wv('diluted') + f"+w_neither*F{VALROW[('tgt','diluted')]}", f_norm, NUM3); r += 1
put(wsVal, r, 2, 'Expected equity per expected share'); put(wsVal, r, 3, f"=C{WEQ}/C{r-1}", f_norm, USD); r += 1
put(wsVal, r, 2, 'Probability-weighted value per share (headline, branch by branch)', f_bold); put(wsVal, r, 3, '=value_target', f_bold, USD, fill=fill_key); r += 1
put(wsVal, r, 2, 'Difference: option dilution and the zero floor are applied within each branch'); put(wsVal, r, 3, f"=C{r-1}-C{r-2}", f_norm, USD); r += 1
put(wsVal, r, 2, 'Expected revenue is the product cash flow weighted by the four branch probabilities; costs are weighted the same way, so a failed product carries neither its revenue nor its commercial spend. Discounting is linear, so this walk equals the weighted sum of the branch DCFs above. ', f_sub); wsVal.cell(r, 2).alignment = Alignment(wrap_text=True, vertical='top'); wsVal.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8); wsVal.row_dimensions[r].height = 40
widths(wsVal, {'A': 2, 'B': 52, 'C': 14, 'D': 14, 'E': 14, 'F': 14})

# ================================================================== analyst rationale for inputs (overrides the imported text)
RATIONALE = {
 'population': 'About 10,000 diagnosed US patients: BioCryst cites ~10,000 claims-derived diagnosed (2019) and ~9,000 company estimate; 8,500 treated plus 1,500 untreated. Not an age or genotype eligible census (POP).',
 'growth': 'Net of new diagnoses and exits. Issuer-reported treated patients grew 1.53% a year 2019-23; diagnosed 2.63%. Tested 0-3%.',
 'category_switch': 'Claims cohort: 50 of 328 commercially insured LTP initiators (15.2%) filled a different LTP within 12 months (HAEA 2025 poster). A mixed-initiator benchmark, not a stable-user switch rate; bounds 10-20%.',
 'prevention_oral_preference': 'Stated-preference study, 150 patients: 54% chose oral over injection at equal efficacy. Heterogeneity, not a national switching rate.',
 'oral_destination_capture': 'Share of departures from the incumbent oral (Orladeyo) that land on PHVS rather than injectables. Judgment; 40-80% tested.',
 'untreated_xr_reach': 'Reach discount on the historical untreated-entrant share (17% of Orladeyo triers). Judgment; 25-75% tested.',
 'oral_launch_share': 'PHVS share of patients choosing an oral at launch. Later-oral analogues (Attruby ~18-20% NBRx, Scemblix 23% first line) argue against assuming parity on day one. 25-60% tested.',
 'oral_mature_share': 'Fixed at 60% in every case: a favourable allowance for better attack control versus berotralstat, not a measured premium. No analogue identifies a universal 60%.',
 'choice_months': 'Linear ramp from launch share to 60% over five years on each product own clock. Timing is judgment (36-84 months).',
 'switch_mature_multiplier': 'Scales the switching opportunity by month 60. Working 1.0 keeps opportunity flat; 0.5 tests exhaustion of easy wins, 1.25 tests expansion. The single most influential sampled input.',
 'hazard_year6_multiplier': 'Disabled (1.0). Kept so the historical half-after-year-5 stress can be reproduced.',
 'ir_oral_switch_relative': 'Ekterly users switching to PHVS borrow the LTP switching proxy; no oral-to-oral rescue switching data exist. 0.5-1.5x tested.',
 'ir_injectable_opportunity': 'Injectable rescue users adopt at the annualised Ekterly first-8-month form propensity (1,702 forms on ~9,000) times PHVS oral choice. 0.75-1.25x tested.',
 'ir_no_rescue_opportunity': '20% of patients without rescue consider one each year, times 84% oral preference and PHVS choice. Judgment; 10-30% tested.',
 'ekterly_future_forms': 'Last observed bimonthly pace (384 forms) continued; 100-250 a month tested.',
 'ekterly_paid_conversion': 'Share of forms that become paid users. Orladeyo paid share ran 71-84%; 65-90% tested.',
 'ekterly_continuation': 'Annual survival of Ekterly users before PHVS launch. Judgment; 80-95%.',
 'ekterly_unique_fraction': 'Working assumes each form is one person; 90% tests duplicates.',
 'ekterly_free_stock_multiplier': 'Stress for unobserved free or patient-assistance oral users (1.2 = 20% more).',
 'ltp_no_rescue_fraction': 'Guidelines expect every prevention patient to hold rescue; 10% without rescue is the published stress.',
 'conversion': 'Share of prescriptions that become paid starts. ~90% proxy outside HAE less 10 points for a competitive, prior-authorised market. Orladeyo paid share 81-84% is a stock measure, not a conversion rate.',
 'delay': 'One month from prescription to first paid dispense; three in the restricted-access package.',
 'retry_months': 'Failed authorisations return to the available pool after 12 months.',
 'reentry_fraction': 'Working: discontinuers do not return. 25% returning after 24 months is the published stress.',
 'reentry_delay': 'See above.',
 'xr_first_retention': '61% comparator persistence (Zuraw 2025, 45-day gap) plus half of efficacy-related exits avoided: 0.61+0.39x(13/30)x0.5 = 69.45%. No adverse-event benefit assumed.',
 'xr_mature_retention': 'Survivors beyond year one retain better than initiators; 90% is a forecast, bounded 80-95%.',
 'ir_retention': 'Continued access and preference for PHVS rescue, not monthly dispensing. 91% repeat dispensing for Orladeyo is context only. 80-95% tested.',
 'xr_paid_intensity': 'Observed PDC 0.73 divided by the 0.789 patient-years implied by 61% survival: refill gaps reduce paid exposure by about 7.5%. Applied once.',
 'ir_doses_ltp': 'Claims mean of 8.0 doses a year in the minimal-gap LTP subgroup (11.8 across all LTP users, including non-dispensers). 6-11.8 tested.',
 'ir_doses_other': 'Broader on-demand claimant benchmark of 14 doses a year, transferred to rescue-only patients. 10-18 tested.',
 'ir_initial_doses': 'Two doses supplied at activation and counted inside the annual budget, not on top of it.',
 'ir_initial_paid_fraction': 'Working: initial stock is paid. Free initial supply is the published stress.',
 'dual_use': 'Share of a PHVS XR user rescue budget that goes to PHVS IR (same molecule). Unobserved; 0-100% tested.',
 'nonphvs_ltp_shift': 'Optional migration of rescue-only users onto other prevention over ten years; off in working.',
 'ir_quantity_multiplier': 'Thesis assumption, fixed in every case: 17% of RAPIDe-3 attacks were redosed versus 38% in KONFIDENT, so 1.17 doses an attack against 1.38. Applied once to the replacement budget; initial stock preserved.',
 'xr_wac_28days': 'Orladeyo WAC per 28 days after the January 2026 increase (BioCryst 8-K). Annualised at 365/28.',
 'xr_net_fraction': '66% net of list is an assumption, not a disclosed gross-to-net. Access packages move it to 58% or 78%.',
 'xr_price_multiplier': 'Fixed at 1.0: no efficacy price premium in any active case.',
 'ir_wac_dose': 'Ekterly WAC per dose from the KalVista approval presentation.',
 'ir_net_fraction': '84% net of list assumed; 76% or 94% in the access packages.',
 'ir_price_multiplier': 'Fixed at 1.0.',
 'ir_exus_ratio': 'Ex-US demand as a share of lagged US sales. Orladeyo history is the main anchor; Takeda global infrastructure given 20% weight for XR only.',
 'xr_exus_ratio': 'See above; 15-60% tested.',
 'exus_lag': 'Weighted country access lag of 12 months.',
 'exus_phase': 'Ex-US demand phases in over 48 months.',
 'exus_direct_fraction': 'Working sells overseas directly, carrying the commercial cost. Licensing (0%) with a 20% royalty is the published stress.',
 'partner_royalty': 'Royalty on partner sales if licensed.',
 'cost_scale': 'Scales R&D, fixed commercial and G&A together. 0.8-1.2 tested. Higher cost does not buy higher sales in the model.',
 'cost_inflation': '2% a year on fixed costs after 2031 until generic entry.',
 'commercial_variable_us': '8% of US sales: field, patient services and distribution that scale with volume.',
 'commercial_variable_exus': '18% of ex-US direct sales.',
 'manufacture_rate': '4% of direct sales for a small-molecule oral made by contract manufacturers; also charged on free doses.',
 'license_rate': 'BRAIN royalty on worldwide end-market sales, disclosed as low to mid single digits (20-F).',
 'capex_fraction': 'Variable capex on direct sales, plus the fixed USD 0.5m a year.',
 'nwc_fraction': '12% of annualised monthly sales as receivables, inventory and payables.',
 'tax_rate': '20% on positive profit after the NOL pool; no legal entity modelling.',
 'minimum_cash_m': 'When modelled cash falls below USD 75m the plan raises to USD 150m. A requirement, not committed capital.',
 'issue_price': 'Illustrative price for the share count only.',
 'issue_fee': '5% gross spread.',
 'fx': 'ECB reference rate 8 October 2026; used once to translate the June 2026 balance.',
 'ir_pos': 'Approval probability for IR (PDUFA 23 April 2027).',
 'xr_pos': 'Approval probability for XR after a planned 1H 2027 filing.',
 'shared_success': 'Shared molecule gate: P(both) = 0.9 x 0.9 / 0.95, which preserves 90% marginals and adds dependence.',
 'discount_rate': '12% nominal USD; 10% and 15% are the published stresses.',
 'loe_year': 'Licensed patents expire 23 November 2038 or later (20-F); 2039 is the first full generic year. 2043 is the later-exclusivity stress.',
 'loe_price_cap': 'Off: brand price can rise after generics as in the pooled study. On caps the price index at 1.',
}
RATIONALE.update(CRIM['rationale'])
for k, rr in INPUT_ROW.items():
    if k in RATIONALE: wsI.cell(rr, 14, RATIONALE[k]).font = f_sub

# ================================================================== helpers for annual/quarterly pulls
YEAR_ROW, Q_ROW, MONTH_ROW = 3, 4, 5
def sum_year(ws_name, r, year_cell): return f"SUMIFS({mrng(ws_name, r)},{mrng(ws_name, YEAR_ROW)},{year_cell})"
def stock_year(ws_name, r, year_cell): return f"INDEX({mrng(ws_name, r)},1,MATCH(DATE({year_cell},12,1),{mrng(ws_name, MONTH_ROW)},0))"
def sum_q(ws_name, r, q_cell): return f"SUMIFS({mrng(ws_name, r)},{mrng(ws_name, Q_ROW)},{q_cell})"
def stock_q(ws_name, r, q_cell): return f"INDEX({mrng(ws_name, r)},1,MATCH(DATE(LEFT({q_cell},4),RIGHT({q_cell},1)*3,1),{mrng(ws_name, MONTH_ROW)},0))"
FINB = lambda key: RF[('both', key)]

# ================================================================== ANNUAL FINANCIALS
wsAn = sheet('Annual', '0070C0'); title(wsAn, 'Annual operating drivers', 'Units are stated row by row. Money is in millions of the currency shown under each year: EUR as reported for 2024-2025, USD from 2026 (H1 actual translated at 1.1186 plus the H2 forecast). Forecast years assume both products are approved.')
AN_UNITS = {'pop': 'patients', 'xr_rx': 'prescriptions', 'xr_starts': 'patients', 'xr_active': 'patients', 'xr_py': 'patient-years', 'xr_units': 'treatment-years', 'xr_price': 'USD per year', 'ir_rx': 'prescriptions', 'ir_starts': 'patients', 'ir_active': 'patients', 'ir_py': 'member-years', 'ir_init': 'doses', 'ir_repl': 'doses', 'ir_units': 'doses', 'ir_price': 'USD per dose', 'growth': '%', 'margin': '%', 'new_sh': 'million shares', 'shares': 'million shares', 'bs_check': 'millions'}
YEARS = [2024, 2025] + list(range(2026, END_YEAR + 1)); YC0 = 4
put(wsAn, 5, 2, 'Fiscal year', f_hdr, fill=fill_hdr); put(wsAn, 5, 3, 'Units', f_hdr, fill=fill_hdr)
for i, y in enumerate(YEARS):
    c = YC0 + i; lab = f'{y}A' if y < 2026 else ('2026A/E' if y == 2026 else f'{y}E')
    put(wsAn, 4, c, y, f_norm, '0'); put(wsAn, 5, c, lab, f_hdr, fill=fill_hdr, align=Alignment(horizontal='right'))
    wsAn.column_dimensions[CL(c)].width = 10
    if y >= 2032: wsAn.column_dimensions[CL(c)].outlineLevel = 1; wsAn.column_dimensions[CL(c)].hidden = False
put(wsAn, 6, 2, 'Currency', f_sub); [put(wsAn, 6, YC0 + i, 'EUR' if y < 2026 else 'USD', f_sub, align=Alignment(horizontal='right')) for i, y in enumerate(YEARS)]
put(wsAn, 7, 2, 'Exclusivity', f_sub); [put(wsAn, 7, YC0 + i, f'=IF({CL(YC0+i)}$4<2027,"",IF({CL(YC0+i)}$4<loe_year,"protected","generic"))', f_sub, align=Alignment(horizontal='right')) for i, y in enumerate(YEARS)]
AN = {}; r = 8
def arow(key, label, fin=None, hist=None, kind='flow', sheet_name='Fin Monthly', fmt=NUM, bold=False, calc=None, unit='', h1=True, indent=0):
    """fin: monthly row to aggregate; hist: History key for 2024/25 (and H1 2026 when h1); calc: fn(col letter)->formula using AN rows."""
    global r; AN[key] = r; put(wsAn, r, 2, ('   ' * indent) + label, f_bold if bold else f_norm); put(wsAn, r, 3, unit or AN_UNITS.get(key, 'millions'), f_sub)
    for i, y in enumerate(YEARS):
        c = YC0 + i; L = CL(c); yc = f'{L}$4'
        if calc is not None: put(wsAn, r, c, calc(L, y), f_bold if bold else f_norm, fmt); continue
        if y < 2026:
            v = f"=History!{CL(3 + (y - 2024))}{HIST_ROW[hist]}" if hist else None
        elif fin is not None:
            core = sum_year(sheet_name, fin, yc) if kind == 'flow' else stock_year(sheet_name, fin, yc)
            v = f"={core}" + (f"+History!E{HIST_ROW[hist]}*fx" if (y == 2026 and hist and h1 and kind == 'flow') else '')
        else: v = None
        if v is not None: put(wsAn, r, c, v, f_bold if bold else f_norm, fmt)
    r += 1
band(wsAn, r, 'Operating drivers (both products approved)', 2, YC0 + len(YEARS) - 1); r += 1
arow('pop', 'US diagnosed population', fin=RX['tot_total'], sheet_name='XR Engine', kind='stock', fmt=NUM0)
arow('xr_rx', 'XR prescriptions written', fin=RX['tot_rx'], sheet_name='XR Engine', fmt=NUM0)
arow('xr_starts', 'XR paid activations', fin=RX['tot_starts'], sheet_name='XR Engine', fmt=NUM0)
arow('xr_active', 'XR active patients, year end', fin=RX['tot_active'], sheet_name='XR Engine', kind='stock', fmt=NUM0)
arow('xr_py', 'XR patient-years', fin=RX['tot_py'], sheet_name='XR Engine', fmt=NUM0)
arow('xr_units', 'XR paid treatment-years', fin=RV['XR_units'], sheet_name='Revenue', fmt=NUM0)
arow('xr_price', 'XR net price per year, USD', calc=lambda L, y: f"=IF({L}{AN['xr_units']}=0,xr_wac_28days*365/28*xr_net_fraction*xr_price_multiplier,{sum_year('Revenue', RV['XR_US'], L+'$4')}*1000000/{L}{AN['xr_units']})" if y >= 2026 else None, fmt=USD0)
arow('ir_rx', 'IR prescriptions written', fin=RI['tot_rx'], sheet_name='IR Engine', fmt=NUM0)
arow('ir_starts', 'IR paid activations', fin=RI['tot_starts'], sheet_name='IR Engine', fmt=NUM0)
arow('ir_active', 'IR active members, year end', fin=RI['tot_active'], sheet_name='IR Engine', kind='stock', fmt=NUM0)
arow('ir_py', 'IR member-years', fin=RI['tot_py'], sheet_name='IR Engine', fmt=NUM0)
arow('ir_init', 'IR paid initial doses', fin=RV['IR_A_paidinit'], sheet_name='Revenue', fmt=NUM0)
arow('ir_repl', 'IR replacement doses', fin=RV['IR_A_repl'], sheet_name='Revenue', fmt=NUM0)
arow('ir_units', 'IR paid doses', fin=RV['IR_A_units'], sheet_name='Revenue', fmt=NUM0)
arow('ir_price', 'IR net price per dose, USD', calc=lambda L, y: f"=IF({L}{AN['ir_units']}=0,ir_wac_dose*ir_net_fraction*ir_price_multiplier,{sum_year('Revenue', RV['IR_A_US'], L+'$4')}*1000000/{L}{AN['ir_units']})" if y >= 2026 else None, fmt=USD0)
r += 1
put(wsAn, r, 2, 'Patient counts are year-end stocks; prescriptions, activations, patient-years and doses are annual flows. Net prices are the realised net revenue per unit, so they fall after generic entry. The income statement, balance sheet and cash flow are on the Three Statements tab.', f_sub)
widths(wsAn, {'A': 2, 'B': 40, 'C': 8}); wsAn.freeze_panes = 'D6'


# ================================================================== THREE STATEMENTS
wsT = sheet('Three Statements', '0070C0'); title(wsT, 'Income statement, balance sheet and cash flow', 'Money in millions of the currency shown under each year (EUR as reported for 2024-2025, USD from 2026); other units stated row by row. Forecasts assume both products are approved and are cash-compensation equivalents without interest or FX.')
TS_UNITS = {'growth': '%', 'gm': '%', 'om': '%', 'em': '%', 'sh': 'million shares', 'eps': 'per share', 'etr': '%', 'nwc_pct': '%', 'rd_pct': '%', 'sm_pct': '%', 'ga_pct': '%', 'fcf_m': '%', 'runway': 'years', 'xr_u': 'treatment-years', 'xr_p': 'USD per year', 'ir_u': 'doses', 'ir_p': 'USD per dose', 'exus_pct': '%', 'sh_open': 'million shares', 'sh_new': 'million shares', 'sh_close': 'million shares', 'opt': 'million options', 'rsu': 'million RSUs', 'bs_check': 'millions', 'cf_check': 'millions'}
QYEARS = [2026, 2027, 2028, 2029, 2030, 2031]; PERIODS = []
for _y in YEARS:
    if _y in QYEARS:
        for _q in (1, 2, 3, 4):
            if _y == 2026 and _q < 3: continue
            PERIODS.append(('Q', f'{_y}Q{_q}'))
    PERIODS.append(('A', _y))
TCOLI = {per: YC0 + i for i, per in enumerate(PERIODS)}
def TC(y): return CL(TCOLI[('A', y)])
def TQ(q): return CL(TCOLI[('Q', q)])
TLAST = CL(YC0 + len(PERIODS) - 1)
fill_q = PatternFill('solid', fgColor='41658A')
put(wsT, 3, 2, 'Period', f_sub); put(wsT, 5, 2, 'Fiscal period', f_hdr, fill=fill_hdr); put(wsT, 5, 3, 'Units', f_hdr, fill=fill_hdr)
for per, c in TCOLI.items():
    if per[0] == 'A':
        y = per[1]; put(wsT, 3, c, 'A', f_sub, align=Alignment(horizontal='right')); put(wsT, 4, c, y, f_norm, '0'); put(wsT, 5, c, f'{y}A' if y < 2026 else ('2026A/E' if y == 2026 else f'{y}E'), f_hdr, fill=fill_hdr, align=Alignment(horizontal='right')); wsT.column_dimensions[CL(c)].width = 10
    else:
        put(wsT, 3, c, 'Q', f_sub, align=Alignment(horizontal='right')); put(wsT, 5, c, per[1], f_hdr, fill=fill_q, align=Alignment(horizontal='right')); wsT.column_dimensions[CL(c)].width = 9
put(wsT, 6, 2, 'Currency', f_sub); [put(wsT, 6, c, 'EUR' if (per[0] == 'A' and per[1] < 2026) else 'USD', f_sub, align=Alignment(horizontal='right')) for per, c in TCOLI.items()]
put(wsT, 7, 2, 'Exclusivity', f_sub); [put(wsT, 7, c, (f'=IF({CL(c)}$4<2027,"",IF({CL(c)}$4<loe_year,"protected","generic"))' if per[0] == 'A' else f'=IF(LEFT({CL(c)}$5,4)*1<loe_year,"protected","generic")'), f_sub, align=Alignment(horizontal='right')) for per, c in TCOLI.items()]
import re as _re2
def _qform(v, L):
    """Turn an annual formula into its quarter equivalent: SUMIFS on the quarter row, quarter-end stocks, year from the label."""
    v = v.replace("$Q$3:$NF$3," + L + "$4", "$Q$4:$NF$4," + L + "$5")
    v = v.replace(f"DATE({L}$4,12,1)", f"DATE(LEFT({L}$5,4),RIGHT({L}$5,1)*3,1)")
    v = _re2.sub(rf"(?<![A-Z$]){L}\$4(?!\d)", f"(LEFT({L}$5,4)*1)", v)
    return v
TS = {}; r = 8
def trow(key, label, fin=None, hist=None, kind='flow', fmt=NUM, bold=False, calc=None, h1=True, indent=0, histcalc=None, qcalc=None, qblank=False):
    global r; TS[key] = r; put(wsT, r, 2, ('   ' * indent) + label, f_bold if bold else f_norm); put(wsT, r, 3, TS_UNITS.get(key, 'millions'), f_sub)
    for per, c in TCOLI.items():
        L = CL(c)
        if per[0] == 'A':
            y = per[1]; H = CL(3 + (y - 2024)) if y < 2026 else None; yc = f'{L}$4'
            if y < 2026: v = histcalc(H, L) if histcalc else (f"=History!{H}{HIST_ROW[hist]}" if hist else None)
            elif calc is not None: v = calc(L, y)
            elif fin is not None: v = '=' + (sum_year('Fin Monthly', fin, yc) if kind == 'flow' else stock_year('Fin Monthly', fin, yc)) + (f"+History!E{HIST_ROW[hist]}*fx" if (y == 2026 and hist and h1 and kind == 'flow') else '')
            else: v = None
        else:
            q = per[1]; y = int(q[:4]); prevq = CL(c - 1)
            if qblank: v = None
            elif qcalc is not None: v = qcalc(L, q, prevq)
            elif fin is not None: v = '=' + (sum_q('Fin Monthly', fin, L + '$5') if kind == 'flow' else stock_q('Fin Monthly', fin, L + '$5'))
            elif calc is not None:
                v = calc(L, max(y, 2027)); v = _qform(v, L) if isinstance(v, str) else v
            else: v = None
        if v is not None: put(wsT, r, c, v, f_bold if bold else f_norm, fmt)
    r += 1
LASTC = TLAST
band(wsT, r, 'Income statement', 2, YC0 + len(PERIODS) - 1); r += 1
SY = lambda row, L: sum_year('Revenue', row, L + '$4')
trow('rev_ir_us', 'IR revenue, United States', calc=lambda L, y: f"={SY(RV['IR_A_US'], L)}", indent=2)
trow('rev_ir_ex', 'IR revenue, ex-US (direct sales and royalties)', calc=lambda L, y: f"={SY(RV['IR_A_direct'], L)}+{SY(RV['IR_A_roy'], L)}", indent=2)
trow('rev_ir', 'IR product revenue', fin=FINB('rev_IR'), indent=1)
trow('rev_xr_us', 'XR revenue, United States', calc=lambda L, y: f"={SY(RV['XR_US'], L)}", indent=2)
trow('rev_xr_ex', 'XR revenue, ex-US (direct sales and royalties)', calc=lambda L, y: f"={SY(RV['XR_direct'], L)}+{SY(RV['XR_roy'], L)}", indent=2)
trow('rev_xr', 'XR product revenue', fin=FINB('rev_XR'), indent=1)
trow('revenue', 'Revenue', fin=FINB('revenue'), hist='revenue', bold=True)
trow('growth', 'Revenue growth', qblank=True, calc=lambda L, y: f"=IF({TC(y-1)}{TS['revenue']}>0,{L}{TS['revenue']}/{TC(y-1)}{TS['revenue']}-1,\"\")" if y >= 2028 else None, fmt=PCT, indent=1)
trow('cogs', 'Cost of goods sold (manufacturing and distribution)', fin=FINB('manufacture'))
trow('royalty', 'Royalty expense (BRAIN licence)', fin=FINB('license'))
trow('gp', 'Gross profit', calc=lambda L, y: f"={L}{TS['revenue']}-{L}{TS['cogs']}-{L}{TS['royalty']}", histcalc=lambda H, L: f"=History!{H}{HIST_ROW['revenue']}", bold=True)
trow('gm', 'Gross margin', calc=lambda L, y: f"=IF({L}{TS['revenue']}>0,{L}{TS['gp']}/{L}{TS['revenue']},\"\")", fmt=PCT, indent=1)
trow('rd', 'Research and development', fin=FINB('rd'), hist='rd')
trow('milestone', 'of which BRAIN milestone at IR approval', fin=FINB('milestone'), indent=2)
trow('sm_var', 'Selling and marketing, variable (8% US, 18% ex-US direct)', fin=FINB('comm_var'), indent=2)
trow('sm', 'Selling and marketing', fin=FINB('commercial'))
trow('sm_fixed', 'of which fixed commercial platform', calc=lambda L, y: f"={L}{TS['sm']}-{L}{TS['sm_var']}", indent=2)
trow('ga', 'General and administrative', fin=FINB('ga'), hist='ga')
trow('ebitda', 'EBITDA', calc=lambda L, y: f"={L}{TS['gp']}-{L}{TS['rd']}-{L}{TS['sm']}-{L}{TS['ga']}", histcalc=lambda H, L: f"=History!{H}{HIST_ROW['revenue']}-History!{H}{HIST_ROW['opex']}+History!{H}{HIST_ROW['depreciation']}", bold=True)
trow('em', 'EBITDA margin', calc=lambda L, y: f"=IF({L}{TS['revenue']}>0,{L}{TS['ebitda']}/{L}{TS['revenue']},\"\")", fmt=PCT, indent=1)
trow('da', 'Depreciation and amortisation', fin=FINB('dep'), hist='depreciation', fmt=NUM2, h1=False)
trow('ebit', 'Operating income', calc=lambda L, y: f"={L}{TS['ebitda']}-{L}{TS['da']}", histcalc=lambda H, L: f"=History!{H}{HIST_ROW['revenue']}-History!{H}{HIST_ROW['opex']}", bold=True)
trow('om', 'Operating margin', calc=lambda L, y: f"=IF({L}{TS['revenue']}>0,{L}{TS['ebit']}/{L}{TS['revenue']},\"\")", fmt=PCT, indent=1)
trow('fin', 'Finance income (reported; not forecast)', hist='finance_income', calc=lambda L, y: f"=History!E{HIST_ROW['finance_income']}*fx" if y == 2026 else '=0', qcalc=lambda L, q, pq: '=0')
trow('ebt', 'Pre-tax income', calc=lambda L, y: f"={L}{TS['ebit']}+{L}{TS['fin']}", histcalc=lambda H, L: f"={L}{TS['ebit']}+{L}{TS['fin']}", bold=True)
trow('tax', 'Income tax', fin=FINB('tax'), hist='tax')
trow('etr', 'Effective tax rate', calc=lambda L, y: f"=IF({L}{TS['ebt']}>0,{L}{TS['tax']}/{L}{TS['ebt']},\"\")", fmt=PCT, indent=1)
trow('ni', 'Net income', calc=lambda L, y: f"={L}{TS['ebt']}-{L}{TS['tax']}", histcalc=lambda H, L: f"=History!{H}{HIST_ROW['net_income']}", bold=True)
trow('sh', 'Shares outstanding, year end', fin=FINB('shares'), hist='shares_end', kind='stock', fmt=NUM2)
trow('eps', 'Earnings per share', calc=lambda L, y: f"={L}{TS['ni']}/{L}{TS['sh']}", histcalc=lambda H, L: f"={L}{TS['ni']}/{L}{TS['sh']}", fmt=USD)
r += 1; band(wsT, r, 'Balance sheet (condensed economic balance sheet)', 2, YC0 + len(PERIODS) - 1); r += 1
trow('cash', 'Cash and equivalents', fin=FINB('cash'), hist='cash_end', kind='stock')
trow('nwc', 'Operating working capital, net', fin=FINB('nwc'), kind='stock')
trow('ppe', 'Property, equipment and right-of-use assets', fin=FINB('ppe'), kind='stock', fmt=NUM2)
trow('other', 'Other net assets', fin=FINB('other'), kind='stock', fmt=NUM2)
trow('ta', 'Total net assets', calc=lambda L, y: f"={L}{TS['cash']}+{L}{TS['nwc']}+{L}{TS['ppe']}+{L}{TS['other']}", histcalc=lambda H, L: f"=History!{H}{HIST_ROW['total_assets']}-History!{H}{HIST_ROW['liabilities']}", bold=True)
trow('debt', 'Debt', calc=lambda L, y: '=0', histcalc=lambda H, L: '=0')
trow('eq_open', 'Equity, opening', calc=lambda L, y: (f"=open_equity*fx" if y == 2026 else f"={TC(y-1)}{r+3}"), qcalc=lambda L, q, pq: '=open_equity*fx' if q == '2026Q3' else f"={pq}{r+3}", indent=1)
trow('eq_ni', 'Net income (from 30 June 2026 in the 2026 column)', calc=lambda L, y: f"={L}{TS['ni']}" + ("-History!E%d*fx" % HIST_ROW['net_income'] if y == 2026 else ''), qcalc=lambda L, q, pq: f"={L}{TS['ni']}", indent=1)
trow('eq_issue', 'Equity issued, net of fees', fin=FINB('fund_net'), indent=1)
trow('eq_close', 'Equity, closing', calc=lambda L, y: f"={L}{TS['eq_open']}+{L}{TS['eq_ni']}+{L}{TS['eq_issue']}", histcalc=lambda H, L: f"=History!{H}{HIST_ROW['equity']}", bold=True)
trow('bs_check', 'Check: net assets less debt less equity', calc=lambda L, y: f"={L}{TS['ta']}-{L}{TS['debt']}-{L}{TS['eq_close']}", histcalc=lambda H, L: f"={L}{TS['ta']}-{L}{TS['debt']}-{L}{TS['eq_close']}", fmt=NUM3)
r += 1; band(wsT, r, 'Cash flow statement', 2, YC0 + len(PERIODS) - 1); r += 1
trow('cf_ni', 'Net income', calc=lambda L, y: f"={L}{TS['ni']}", histcalc=lambda H, L: f"=History!{H}{HIST_ROW['net_income']}", indent=1)
trow('cf_da', 'Depreciation and amortisation', calc=lambda L, y: f"={L}{TS['da']}" + (f"+History!E{HIST_ROW['depreciation']}*fx" if y == 2026 else ''), qcalc=lambda L, q, pq: f"={L}{TS['da']}", histcalc=lambda H, L: f"=History!{H}{HIST_ROW['depreciation']}", fmt=NUM2, indent=1)
trow('cf_sbc', 'Share-based compensation and other non-cash (reported)', histcalc=lambda H, L: f"=History!{H}{HIST_ROW['sbc']}", calc=lambda L, y: f"=History!E{HIST_ROW['sbc']}*fx" if y == 2026 else '=0', qcalc=lambda L, q, pq: '=0', indent=1)
trow('cf_nwc', 'Change in operating working capital', fin=FINB('dnwc'), calc=lambda L, y: f"=-({sum_year('Fin Monthly', FINB('dnwc'), L+'$4')})", qcalc=lambda L, q, pq: f"=-({sum_q('Fin Monthly', FINB('dnwc'), L+'$5')})", histcalc=lambda H, L: f"=History!{H}{HIST_ROW['cfo']}-History!{H}{HIST_ROW['net_income']}-History!{H}{HIST_ROW['depreciation']}-History!{H}{HIST_ROW['sbc']}", indent=1)
trow('cfo', 'Cash from operations', calc=lambda L, y: f"={L}{TS['cf_ni']}+{L}{TS['cf_da']}+{L}{TS['cf_sbc']}+{L}{TS['cf_nwc']}" + (f"+(History!E{HIST_ROW['cfo']}*fx-History!E{HIST_ROW['net_income']}*fx-History!E{HIST_ROW['depreciation']}*fx-History!E{HIST_ROW['sbc']}*fx)" if y == 2026 else ''), histcalc=lambda H, L: f"=History!{H}{HIST_ROW['cfo']}", bold=True)
trow('capex', 'Capital expenditure', calc=lambda L, y: f"=-({sum_year('Fin Monthly', FINB('capex'), L+'$4')})" + (f"-History!E{HIST_ROW['capex']}*fx" if y == 2026 else ''), histcalc=lambda H, L: f"=-History!{H}{HIST_ROW['capex']}", fmt=NUM2, indent=1)
trow('cfi', 'Cash from investing', calc=lambda L, y: f"={L}{TS['capex']}", histcalc=lambda H, L: f"=History!{H}{HIST_ROW['cfi']}", bold=True)
trow('issue', 'Equity issued, net of fees', calc=lambda L, y: f"={sum_year('Fin Monthly', FINB('fund_net'), L+'$4')}" + (f"+History!E{HIST_ROW['cff']}*fx" if y == 2026 else ''), histcalc=lambda H, L: f"=History!{H}{HIST_ROW['cff']}", indent=1)
trow('cff', 'Cash from financing', calc=lambda L, y: f"={L}{TS['issue']}", histcalc=lambda H, L: f"=History!{H}{HIST_ROW['cff']}", bold=True)
trow('fx', 'FX on cash (reported)', histcalc=lambda H, L: f"=History!{H}{HIST_ROW['cash_fx']}", calc=lambda L, y: f"=History!E{HIST_ROW['cash_fx']}*fx" if y == 2026 else '=0', qcalc=lambda L, q, pq: '=0')
trow('net', 'Net change in cash', calc=lambda L, y: f"={L}{TS['cfo']}+{L}{TS['cfi']}+{L}{TS['cff']}+{L}{TS['fx']}", histcalc=lambda H, L: f"={L}{TS['cfo']}+{L}{TS['cfi']}+{L}{TS['cff']}+{L}{TS['fx']}", bold=True)
trow('cash_open', 'Cash, opening', calc=lambda L, y: (f"=History!E{HIST_ROW['cash_begin']}*fx" if y == 2026 else f"={TC(y-1)}{r+1}"), qcalc=lambda L, q, pq: '=open_cash*fx' if q == '2026Q3' else f"={pq}{r+1}", histcalc=lambda H, L: f"=History!{H}{HIST_ROW['cash_begin']}")
trow('cash_close', 'Cash, closing', calc=lambda L, y: f"={L}{TS['cash_open']}+{L}{TS['net']}", histcalc=lambda H, L: f"={L}{TS['cash_open']}+{L}{TS['net']}", bold=True)
trow('cf_check', 'Check: closing cash against the balance sheet', calc=lambda L, y: f"={L}{TS['cash_close']}-{L}{TS['cash']}", histcalc=lambda H, L: f"={L}{TS['cash_close']}-{L}{TS['cash']}", fmt=NUM3)
r += 1; band(wsT, r, 'Supporting schedules', 2, YC0 + len(PERIODS) - 1); r += 1
put(wsT, r, 2, 'Revenue build', f_bold); r += 1
trow('xr_u', 'XR paid treatment-years', calc=lambda L, y: f"={SY(RV['XR_units'], L)}", fmt=NUM0, indent=1)
trow('xr_p', 'XR net price per treatment-year', calc=lambda L, y: f"=IF({L}{TS['xr_u']}=0,xr_wac_28days*365/28*xr_net_fraction*(1+price_growth)^MAX(0,{L}$4-2027),{SY(RV['XR_US'], L)}*1000000/{L}{TS['xr_u']})", fmt=USD0, indent=1)
trow('ir_u', 'IR paid doses', calc=lambda L, y: f"={SY(RV['IR_A_units'], L)}", fmt=NUM0, indent=1)
trow('ir_p', 'IR net price per dose', calc=lambda L, y: f"=IF({L}{TS['ir_u']}=0,ir_wac_dose*ir_net_fraction*(1+price_growth)^MAX(0,{L}$4-2027),{SY(RV['IR_A_US'], L)}*1000000/{L}{TS['ir_u']})", fmt=USD0, indent=1)
trow('exus_pct', 'Ex-US revenue as a share of US revenue', calc=lambda L, y: f"=IF({L}{TS['rev_ir_us']}+{L}{TS['rev_xr_us']}>0,({L}{TS['rev_ir_ex']}+{L}{TS['rev_xr_ex']})/({L}{TS['rev_ir_us']}+{L}{TS['rev_xr_us']}),\"\")", fmt=PCT, indent=1)
r += 1; put(wsT, r, 2, 'Operating expense build', f_bold); r += 1
trow('rd_sched', 'R&D budget before scale, inflation and runoff', calc=lambda L, y: f"=INDEX(cost_rd,MATCH(MIN({L}$4,2031),cost_years,0))", indent=1)
trow('sm_sched', 'Fixed commercial budget before scale, inflation and runoff', calc=lambda L, y: f"=INDEX(cost_comm,MATCH(MIN({L}$4,2031),cost_years,0))", indent=1)
trow('ga_sched', 'G&A budget before scale, inflation and runoff', calc=lambda L, y: f"=INDEX(cost_ga,MATCH(MIN({L}$4,2031),cost_years,0))", indent=1)
trow('cost_factor', 'Cost factor applied (scale x inflation x post-generic runoff)', calc=lambda L, y: f"=cost_scale*(1+cost_inflation)^MAX(0,MIN({L}$4,loe_year-1)-2031)*(INDEX(ir_r_idx,MATCH(MAX(2026,{L}$4-1),loe_years,0))+INDEX(xr_r_idx,MATCH(MAX(2026,{L}$4-1),loe_years,0)))/2", fmt=NUM3, indent=1)
trow('rd_pct', 'R&D as a share of revenue', calc=lambda L, y: f"=IF({L}{TS['revenue']}>0,{L}{TS['rd']}/{L}{TS['revenue']},\"\")", fmt=PCT, indent=1)
trow('sm_pct', 'Selling and marketing as a share of revenue', calc=lambda L, y: f"=IF({L}{TS['revenue']}>0,{L}{TS['sm']}/{L}{TS['revenue']},\"\")", fmt=PCT, indent=1)
trow('ga_pct', 'G&A as a share of revenue', calc=lambda L, y: f"=IF({L}{TS['revenue']}>0,{L}{TS['ga']}/{L}{TS['revenue']},\"\")", fmt=PCT, indent=1)
r += 1; put(wsT, r, 2, 'Working capital', f_bold); r += 1
trow('nwc_open', 'Operating working capital, opening', calc=lambda L, y: f"=open_nwc*fx" if y == 2026 else f"={TC(y-1)}{TS['nwc']}", qcalc=lambda L, q, pq: '=open_nwc*fx' if q == '2026Q3' else f"={pq}{TS['nwc']}", indent=1)
trow('nwc_close', 'Operating working capital, closing', calc=lambda L, y: f"={L}{TS['nwc']}", indent=1)
trow('nwc_chg', 'Investment in working capital', calc=lambda L, y: f"={L}{TS['nwc_close']}-{L}{TS['nwc_open']}", indent=1)
trow('nwc_pct', 'Working capital as a share of direct sales', calc=lambda L, y: f"=IF({L}{TS['rev_ir_us']}+{L}{TS['rev_xr_us']}>0,({L}{TS['nwc_close']}-open_nwc*fx)/({SY(RV['IR_A_US'], L)}+{SY(RV['XR_US'], L)}+{SY(RV['IR_A_direct'], L)}+{SY(RV['XR_direct'], L)}),\"\")", fmt=PCT, indent=1)
r += 1; put(wsT, r, 2, 'Property and equipment', f_bold); r += 1
trow('ppe_open', 'Opening', calc=lambda L, y: f"=open_ppe_and_rou*fx" if y == 2026 else f"={TC(y-1)}{TS['ppe']}", qcalc=lambda L, q, pq: '=open_ppe_and_rou*fx' if q == '2026Q3' else f"={pq}{TS['ppe']}", fmt=NUM2, indent=1)
trow('ppe_capex', 'Capital expenditure (forecast months only in 2026)', calc=lambda L, y: f"={sum_year('Fin Monthly', FINB('capex'), L+'$4')}", fmt=NUM2, indent=1)
trow('ppe_dep', 'Depreciation', calc=lambda L, y: f"=-{sum_year('Fin Monthly', FINB('dep'), L+'$4')}", fmt=NUM2, indent=1)
trow('ppe_close', 'Closing', calc=lambda L, y: f"={L}{TS['ppe_open']}+{L}{TS['ppe_capex']}+{L}{TS['ppe_dep']}", fmt=NUM2, indent=1)
trow('ppe_chk', 'Check against the balance sheet', calc=lambda L, y: f"={L}{TS['ppe_close']}-{L}{TS['ppe']}", fmt=NUM3, indent=1)
r += 1; put(wsT, r, 2, 'Equity funding and share count', f_bold); r += 1
trow('fund_gross', 'Gross proceeds raised', calc=lambda L, y: f"={sum_year('Fin Monthly', FINB('fund_gross'), L+'$4')}", indent=1)
trow('fund_fee', 'Issuance fees', calc=lambda L, y: f"=-{sum_year('Fin Monthly', FINB('fund_fee'), L+'$4')}", fmt=NUM2, indent=1)
trow('fund_net', 'Net proceeds', calc=lambda L, y: f"={L}{TS['fund_gross']}+{L}{TS['fund_fee']}", indent=1)
trow('sh_open', 'Shares, opening', calc=lambda L, y: f"=open_shares_m" if y == 2026 else f"={TC(y-1)}{TS['sh']}", qcalc=lambda L, q, pq: '=open_shares_m' if q == '2026Q3' else f"={pq}{TS['sh']}", fmt=NUM2, indent=1)
trow('sh_new', 'Shares issued at the illustrative issue price', calc=lambda L, y: f"={sum_year('Fin Monthly', FINB('new_sh'), L+'$4')}", fmt=NUM2, indent=1)
trow('sh_close', 'Shares, closing', calc=lambda L, y: f"={L}{TS['sh_open']}+{L}{TS['sh_new']}", fmt=NUM2, indent=1)
trow('opt', 'Options outstanding (treasury-stock method at valuation)', calc=lambda L, y: "=open_options_m", fmt=NUM3, indent=1)
trow('rsu', 'RSUs outstanding', calc=lambda L, y: "=open_rsus_m", fmt=NUM3, indent=1)
r += 1; put(wsT, r, 2, 'Tax and loss carry-forwards', f_bold); r += 1
trow('nol_open', 'NOL pool, opening', calc=lambda L, y: "=0" if y == 2026 else f"={TC(y-1)}{r+3}", qcalc=lambda L, q, pq: '=0' if q == '2026Q3' else f"={pq}{r+3}", indent=1)
trow('nol_add', 'Losses added', calc=lambda L, y: f"={L}{r+2}-{L}{TS['nol_open']}+{L}{r+1}", indent=1)
trow('nol_used', 'NOL used against profit (80% cap)', calc=lambda L, y: f"={sum_year('Fin Monthly', FINB('nol_used'), L+'$4')}", indent=1)
trow('nol_close', 'NOL pool, closing', calc=lambda L, y: f"={stock_year('Fin Monthly', FINB('nol'), L+'$4')}", indent=1)
trow('cash_tax', 'Cash tax (20% after NOL, USD 1m floor)', calc=lambda L, y: f"={L}{TS['tax']}", indent=1)
r += 1; put(wsT, r, 2, 'Ratios', f_bold); r += 1
trow('fcf', 'Free cash flow (operations less capex)', calc=lambda L, y: f"={L}{TS['cfo']}+{L}{TS['cfi']}", histcalc=lambda H, L: f"={L}{TS['cfo']}+{L}{TS['cfi']}", indent=1)
trow('fcf_m', 'Free cash flow margin', calc=lambda L, y: f"=IF({L}{TS['revenue']}>0,{L}{TS['fcf']}/{L}{TS['revenue']},\"\")", fmt=PCT, indent=1)
trow('runway', 'Cash runway at the period\'s operating burn, years', calc=lambda L, y: f"=IF({L}{TS['cfo']}<0,{L}{TS['cash']}/-{L}{TS['cfo']},\"\")", qcalc=lambda L, q, pq: f"=IF({L}{TS['cfo']}<0,{L}{TS['cash']}/(-{L}{TS['cfo']}*4),\"\")", histcalc=lambda H, L: f"=IF({L}{TS['cfo']}<0,{L}{TS['cash']}/-{L}{TS['cfo']},\"\")", fmt=NUM2, indent=1)
r += 1
put(wsT, r, 2, 'Reported 2024-2025 follow IFRS and include finance income and non-cash share-based compensation. From H2 2026 the forecast is an economic cash-compensation view: existing awards enter dilution at valuation, future incentives sit inside the cash budgets, and there is no interest or FX. The 2026 column adds the reported H1 (EUR x 1.1186) to the H2 forecast, so its opening cash is the reported 31 December 2025 balance translated at the same rate.', f_sub); r += 1
put(wsT, r, 2, 'Working capital is carried net (12% of annualised sales plus the opening balance); liabilities beyond payables are inside other net assets. Debt is zero throughout. Explicit forecast runs to 2055; the Valuation tab adds a declining residual beyond it.', f_sub)
widths(wsT, {'A': 2, 'B': 56, 'C': 15}); wsT.freeze_panes = 'D6'; wsT.print_area = f"B3:{TLAST}{TS['runway']}"
put(wsT, 1, 2, 'Central case: every input at its most likely value (Inputs tab). The Monte Carlo on the DCF tab varies those inputs around these values.', f_sub); wsT.cell(1, 2).font = f_title
put(wsT, 2, 2, 'Central case: every input at its most likely value on the Inputs tab; the Monte Carlo and the price target vary those inputs around this path. Quarters run through 2031 and sum to the annual columns beside them.', f_sub)
wsT.cell(1, 2).value = 'Income statement, balance sheet and cash flow'

# financial lines live on Three Statements; operating drivers on Annual
FINMAP = {'ir_rev': 'rev_ir', 'xr_rev': 'rev_xr', 'revenue': 'revenue', 'ebit': 'ebit', 'ni': 'ni', 'fcf': 'fcf', 'fund': 'fund_net', 'cash': 'cash', 'shares': 'sh', 'margin': 'om'}
def AX(key):
    return ("'Three Statements'", TS[FINMAP[key]]) if key in FINMAP else ('Annual', AN[key])
def tcol(key, y): return TC(y) if key in FINMAP else CL(YC0 + YEARS.index(y))
def aref(key, y): sh, rr_ = AX(key); return f"{sh}!{tcol(key, y)}{rr_}"
def arange(key): sh, rr_ = AX(key); return f"{sh}!$D${rr_}:${TLAST if key in FINMAP else CL(YC0+len(YEARS)-1)}${rr_}"
AYEARS = f"'Three Statements'!$D$4:${TLAST}$4"; AFLAG = f"'Three Statements'!$D$3:${TLAST}$3"
def tsum(key, y0, y1): sh, rr_ = AX(key); return f"SUMIFS({sh}!$D${rr_}:${TLAST}${rr_},{AFLAG},\"A\",{AYEARS},\">={y0}\",{AYEARS},\"<={y1}\")"
name('rev_2030', f"'Three Statements'!${TC(2030)}${TS['revenue']}")
QR = {}
# ================================================================== ADOPTION BUILD: annual summary table
r = ADOPT_ANNUAL_ROW; band(wsA, r, 'Annual adoption by origin, both products approved (prescriptions and patients are counts; patient-years are exposure)', 2, 14); r += 1
AYRS = list(range(2027, END_YEAR + 1)); put(wsA, r, 2, 'Year', f_hdr, fill=fill_hdr); put(wsA, r, 3, '', f_hdr, fill=fill_hdr)
for i, y in enumerate(AYRS): put(wsA, r, 4 + i, y, f_hdr, fill=fill_hdr, align=Alignment(horizontal='right')); wsA.column_dimensions[CL(4 + i)].width = max(wsA.column_dimensions[CL(4 + i)].width or 0, 9)
r += 1
def arow2(label, ws_name, row, kind='flow', fmt=NUM0, bold=False):
    global r; put(wsA, r, 2, label, f_bold if bold else f_norm)
    for i, y in enumerate(AYRS):
        L = CL(4 + i); yc = f'{L}${r-1 if False else ADOPT_ANNUAL_ROW+1}'
        put(wsA, r, 4 + i, '=' + (sum_year(ws_name, row, yc) if kind == 'flow' else stock_year(ws_name, row, yc)), f_bold if bold else f_norm, fmt)
    r += 1
for o in ORIG: arow2(f'XR prescriptions from {ON[o]}', 'XR Engine', RX[f'rx_{o}'])
arow2('XR prescriptions, total', 'XR Engine', RX['tot_rx'], bold=True); arow2('XR paid activations', 'XR Engine', RX['tot_starts']); arow2('XR active patients, year end', 'XR Engine', RX['tot_active'], kind='stock'); arow2('XR patient-years', 'XR Engine', RX['tot_py'])
r += 1
for s_ in RESC: arow2(f'IR prescriptions from {RN[s_]} users', 'IR Engine', RI[f'rx_r_{s_}'])
arow2('IR prescriptions, total', 'IR Engine', RI['tot_rx'], bold=True); arow2('IR paid activations', 'IR Engine', RI['tot_starts']); arow2('IR active members, year end', 'IR Engine', RI['tot_active'], kind='stock'); arow2('IR member-years', 'IR Engine', RI['tot_py'])
r += 1
put(wsA, r, 2, 'Launch comparability: IR prescriptions in the first eight commercial months'); put(wsA, r, 5, f"=SUMPRODUCT({mrng('IR Engine', RI['tot_rx'])},--({mrng('Calendar', CAL['age_IR'])}>=0),--({mrng('Calendar', CAL['age_IR'])}<8))", f_norm, NUM0); put(wsA, r, 8, 'against 1,702 Ekterly start forms over its first eight months (KALVFORM). An output of the build, not an imposed discount.', f_sub); r += 1
put(wsA, r, 2, 'Ratio to Ekterly forms'); put(wsA, r, 5, f'=E{r-1}/ekterly_first8_forms', f_norm, PCT); r += 1

# ================================================================== QUANTITY & PRICE
wsP = sheet('Quantity & Price', '0070C0'); title(wsP, 'Quantity and price build', 'How patient-years and rescue memberships become paid units, and what each unit is worth.')
r = 4; band(wsP, r, 'Net price', 2, 10); r += 1
put(wsP, r, 2, 'XR comparator WAC per 28 days (Orladeyo)'); put(wsP, r, 5, '=xr_wac_28days', f_norm, USD); put(wsP, r, 7, 'XRPRICE', f_sub); r += 1
put(wsP, r, 2, 'Annualised at 365/28'); put(wsP, r, 5, f'=E{r-1}*365/28', f_norm, USD0); r += 1
put(wsP, r, 2, 'Net fraction of list'); put(wsP, r, 5, '=xr_net_fraction', f_norm, PCT); r += 1
put(wsP, r, 2, 'XR net price per treatment-year', f_bold); put(wsP, r, 5, f'=E{r-2}*E{r-1}*xr_price_multiplier', f_bold, USD0); r += 2
put(wsP, r, 2, 'IR comparator WAC per dose (Ekterly)'); put(wsP, r, 5, '=ir_wac_dose', f_norm, USD); put(wsP, r, 7, 'IRPRICE', f_sub); r += 1
put(wsP, r, 2, 'Net fraction of list'); put(wsP, r, 5, '=ir_net_fraction', f_norm, PCT); r += 1
put(wsP, r, 2, 'IR net price per dose', f_bold); put(wsP, r, 5, f'=E{r-2}*E{r-1}*ir_price_multiplier', f_bold, USD0); r += 1
put(wsP, r, 2, 'Both prices are comparator list prices with assumed net discounts, not disclosed PHVS pricing. No launch price premium and no list-price inflation are assumed.', f_sub); r += 2
band(wsP, r, 'Rescue dose budget per member-year', 2, 10); r += 1
put(wsP, r, 2, 'Origin', f_bold); put(wsP, r, 5, 'Claims budget', f_bold); put(wsP, r, 6, 'After redosing (x 1.17/1.38)', f_bold); put(wsP, r, 7, 'Initial stock', f_bold); put(wsP, r, 8, 'Replacement, year 1', f_bold); put(wsP, r, 9, 'Replacement, mature', f_bold); r += 1
for lab, base in (('On prevention (Orladeyo, other prevention)', 'ir_doses_ltp'), ('No prevention (rescue only, untreated)', 'ir_doses_other')):
    put(wsP, r, 2, lab); put(wsP, r, 5, f'={base}', f_norm, NUM2); put(wsP, r, 6, f'=E{r}*ir_quantity_multiplier', f_norm, NUM2); put(wsP, r, 7, '=ir_initial_doses', f_norm, NUM2); put(wsP, r, 8, f'=F{r}-G{r}', f_norm, NUM2); put(wsP, r, 9, f'=F{r}', f_norm, NUM2); r += 1
put(wsP, r, 2, 'A PHVS XR user who also carries PHVS rescue gets the prevention budget times the same-drug allocation (dual use). The overlap share is the XR patient-years in each origin cell divided by its population, so it rises as XR penetrates. Free initial doses earn nothing but still cost manufacturing.', f_sub); r += 2
band(wsP, r, 'Annual quantities and revenue, both products approved (units as labelled; revenue in USD millions)', 2, 16); r += 1
PYRS = list(range(2027, END_YEAR + 1)); put(wsP, r, 2, 'Year', f_hdr, fill=fill_hdr); put(wsP, r, 3, '', f_hdr, fill=fill_hdr)
for i, y in enumerate(PYRS): put(wsP, r, 4 + i, y, f_hdr, fill=fill_hdr, align=Alignment(horizontal='right')); wsP.column_dimensions[CL(4 + i)].width = 10
PYH = r; r += 1
def prow(label, ws_name, row, kind='flow', fmt=NUM0, bold=False, calc=None):
    global r; put(wsP, r, 2, label, f_bold if bold else f_norm)
    for i, y in enumerate(PYRS):
        L = CL(4 + i); yc = f'{L}${PYH}'
        put(wsP, r, 4 + i, calc(L) if calc else '=' + (sum_year(ws_name, row, yc) if kind == 'flow' else stock_year(ws_name, row, yc)), f_bold if bold else f_norm, fmt)
    r += 1
prow('XR patient-years (survival-integrated)', 'XR Engine', RX['tot_py']); prow('XR paid treatment-years (after refill intensity)', 'Revenue', RV['XR_units'], bold=True)
prow('XR net price index after generics', 'LOE', 0, calc=lambda L: f"=INDEX(xr_p_idx,MATCH({L}${PYH},loe_years,0))", fmt=NUM3); prow('XR US revenue, USD m', 'Revenue', RV['XR_US'], fmt=NUM); prow('XR ex-US recognised revenue, USD m', 'Revenue', 0, calc=lambda L: f"={sum_year('Revenue', RV['XR_direct'], L+'$'+str(PYH))}+{sum_year('Revenue', RV['XR_roy'], L+'$'+str(PYH))}", fmt=NUM)
r += 1
prow('IR member-years', 'IR Engine', RI['tot_py']); prow('IR paid initial doses', 'Revenue', RV['IR_A_paidinit']); prow('IR replacement doses', 'Revenue', RV['IR_A_repl']); prow('IR free initial doses', 'Revenue', RV['IR_A_free']); prow('IR paid doses after generics', 'Revenue', RV['IR_A_units'], bold=True)
prow('IR net price index after generics', 'LOE', 0, calc=lambda L: f"=INDEX(ir_p_idx,MATCH({L}${PYH},loe_years,0))", fmt=NUM3); prow('IR US revenue, USD m', 'Revenue', RV['IR_A_US'], fmt=NUM); prow('IR ex-US recognised revenue, USD m', 'Revenue', 0, calc=lambda L: f"={sum_year('Revenue', RV['IR_A_direct'], L+'$'+str(PYH))}+{sum_year('Revenue', RV['IR_A_roy'], L+'$'+str(PYH))}", fmt=NUM)
prow('Overlap: IR members who are also XR patients', 'Revenue', RV['overlap'], calc=lambda L: f"={sum_year('Revenue', RV['overlap'], L+'$'+str(PYH))}/12")
r += 1
put(wsP, r, 2, 'Ex-US demand equals the ex-US/US ratio times US sales twelve months earlier, phased in over four years. After the first full generic year, units freeze at the protected-year average and follow the LOE quantity index; price follows the LOE price index.', f_sub)
widths(wsP, {'A': 2, 'B': 46, 'C': 3})

# ================================================================== STREET BRIDGE (engine counterfactuals + Street data)
SPX = Path('/private/tmp/claude-501/-Users-kwu-Library-Application-Support-Claude-scratch-workspaces-665f01e2-9652-4512-9fda-0497b55d354c-06f88a99-9105-4941-a8cf-87e6b02aff65-scratch-2026-10-10-7ee708/b9401f11-15ac-4c44-8c9e-8eb278b73abd/scratchpad')
ALT_A = json.loads((SPX / 'altdata_A.json').read_text())['datasets']; ALT_B = json.loads((SPX / 'altdata_B.json').read_text())['datasets']
wsFn = build_funnel()
bridge = [(b['label'], b['both_target'], b['expected_target'], b['peak']) for b in CR['bridge']]
wsS = sheet('Street Bridge', '0070C0'); title(wsS, 'From the Street to our view', 'Where sell-side targets sit, how their revenue path compares with ours, and which assumptions carry the gap.')
st = ALT_A['street_targets']; r = 4; band(wsS, r, 'Sell-side price targets, September 2026 notes', 2, 10); r += 1
cols = ['Broker', 'Report date', 'Price target, USD', 'IR PoS', 'XR PoS', 'Discount rate', 'Method', 'Source page']
for i, h in enumerate(cols): put(wsS, r, 2 + i, h, f_hdr, fill=fill_hdr)
r += 1; ST0 = r
for row_ in st['rows']:
    for i, v in enumerate(row_[:8]): put(wsS, r, 2 + i, v, f_input if i == 2 else f_norm, USD if i == 2 else (PCT if i in (3, 4, 5) and isinstance(v, (int, float)) else GEN))
    r += 1
ST1 = r - 1
put(wsS, r, 2, 'Median target', f_bold); put(wsS, r, 4, f'=MEDIAN(D{ST0}:D{ST1})', f_bold, USD); name('street_median', f"'Street Bridge'!$D${r}"); r += 1
put(wsS, r, 2, 'Lowest / highest'); put(wsS, r, 4, f'=MIN(D{ST0}:D{ST1})', f_norm, USD); put(wsS, r, 5, f'=MAX(D{ST0}:D{ST1})', f_norm, USD); r += 1
put(wsS, r, 2, 'Share price'); put(wsS, r, 4, '=share_price', f_norm, USD); put(wsS, r, 5, '=TEXT(price_date,"d mmm yyyy")', f_sub); r += 1
put(wsS, r, 2, 'Our 12-month price target', f_bold); put(wsS, r, 4, '=price_target', f_bold, USD, fill=fill_key); r += 1
put(wsS, r, 2, 'Targets transcribed from the supplied September 2026 notes (reverse-DCF research, 7 October 2026). Leerink and Morgan Stanley later updates supersede their earlier figures; Wells Fargo states no terminal growth, which is not no terminal value.', f_sub); r += 2
gp = ALT_A['guggenheim_revenue_path']; band(wsS, r, 'Guggenheim published revenue path (as printed, USD at 1.16) against this model, both products approved. The reverse-solved ramps the share price requires are on the Reverse DCF tab.', 2, 16); r += 1
gyears = [row_[0] for row_ in gp['rows']]; gcols = gp['columns']
iIRu = [i for i, c in enumerate(gcols) if 'IR' in c and 'USD' in c][0]; iXRu = [i for i, c in enumerate(gcols) if 'XR' in c and 'USD' in c][0]
put(wsS, r, 2, 'Calendar year', f_hdr, fill=fill_hdr); put(wsS, r, 3, '', f_hdr, fill=fill_hdr); [put(wsS, r, 4 + i, y, f_hdr, fill=fill_hdr, align=Alignment(horizontal='right')) for i, y in enumerate(gyears)]; GY = r; r += 1
put(wsS, r, 2, 'Street IR revenue'); [put(wsS, r, 4 + i, row_[iIRu], f_input, NUM) for i, row_ in enumerate(gp['rows'])]; GIR = r; r += 1
put(wsS, r, 2, 'Street XR revenue'); [put(wsS, r, 4 + i, row_[iXRu], f_input, NUM) for i, row_ in enumerate(gp['rows'])]; GXR = r; r += 1
put(wsS, r, 2, 'Street total', f_bold); [put(wsS, r, 4 + i, f'={CL(4+i)}{GIR}+{CL(4+i)}{GXR}', f_bold, NUM) for i in range(len(gyears))]; GT = r; r += 1
put(wsS, r, 2, 'Model IR revenue'); [put(wsS, r, 4 + i, f"=IF({CL(4+i)}${GY}<2026,0,INDEX({arange('ir_rev')},1,MATCH({CL(4+i)}${GY},{AYEARS},0)))", f_norm, NUM) for i in range(len(gyears))]; MIR = r; r += 1
put(wsS, r, 2, 'Model XR revenue'); [put(wsS, r, 4 + i, f"=IF({CL(4+i)}${GY}<2026,0,INDEX({arange('xr_rev')},1,MATCH({CL(4+i)}${GY},{AYEARS},0)))", f_norm, NUM) for i in range(len(gyears))]; MXR = r; r += 1
put(wsS, r, 2, 'Model total', f_bold); [put(wsS, r, 4 + i, f'={CL(4+i)}{MIR}+{CL(4+i)}{MXR}', f_bold, NUM) for i in range(len(gyears))]; MT = r; r += 1
put(wsS, r, 2, 'Model as a share of Street'); [put(wsS, r, 4 + i, f'=IF({CL(4+i)}{GT}=0,"",{CL(4+i)}{MT}/{CL(4+i)}{GT})', f_norm, PCT) for i in range(len(gyears))]; r += 1
put(wsS, r, 2, 'Guggenheim 9 September 2026 note (124306319 pp1-3): 90%/90% PoS, 12.5% WACC, -5% terminal growth. Printed EUR thousands converted to millions and to USD at 1.16; whether the printed revenues already carry probability of success is not disclosed. Our path is conditional on approval and recognises direct ex-US sales.', f_sub); r += 2
from openpyxl.chart import LineChart, BarChart, Reference, Series
ch = LineChart(); ch.title = 'Revenue path, USD m'; ch.height = 8; ch.width = 18; ch.y_axis.title = 'USD m'; ch.style = 2
data = Reference(wsS, min_col=3, max_col=3 + len(gyears), min_row=GT, max_row=GT); ch.add_data(Reference(wsS, min_col=4, max_col=3 + len(gyears), min_row=GT), from_rows=True, titles_from_data=False); ch.add_data(Reference(wsS, min_col=4, max_col=3 + len(gyears), min_row=MT), from_rows=True, titles_from_data=False)
ch.set_categories(Reference(wsS, min_col=4, max_col=3 + len(gyears), min_row=GY)); ch.series[0].tx = None; from openpyxl.chart.series import SeriesLabel; ch.series[0].tx = SeriesLabel(v='Street (Guggenheim)'); ch.series[1].tx = SeriesLabel(v='This model')
CHART_S = ch
from openpyxl.worksheet.pagebreak import Break
wsS.row_breaks.append(Break(id=r - 1))
band(wsS, r, 'Value bridge: what the three thesis assumptions are worth (engine runs, 9 October 2026 inputs)', 2, 10); r += 1
for i, h in enumerate(['Step', 'Both approved, 12-month value', 'Change', 'Probability-weighted 12-month value', 'Change', 'Peak revenue, USD m']): put(wsS, r, 2 + i, h, f_hdr, fill=fill_hdr)
r += 1; BR0 = r
for i, (lab, pb, pe, pk) in enumerate(bridge):
    put(wsS, r, 2, lab); put(wsS, r, 3, pb, f_input, USD); put(wsS, r, 4, '' if i == 0 else f'=C{r}-C{r-1}', f_norm, USD); put(wsS, r, 5, pe, f_input, USD); put(wsS, r, 6, '' if i == 0 else f'=E{r}-E{r-1}', f_norm, USD); put(wsS, r, 7, pk, f_input, NUM); r += 1
put(wsS, r, 2, 'Central case, live engine (equals the last step with no overrides)', f_bold); put(wsS, r, 3, '=price_both_target', f_bold, USD); put(wsS, r, 5, '=value_target', f_bold, USD); put(wsS, r, 7, f"=MAX({arange('revenue')})", f_bold, NUM); r += 1
put(wsS, r, 2, 'Street median target for reference'); put(wsS, r, 3, '=street_median', f_norm, USD); r += 1
put(wsS, r, 2, 'The first row relaxes the redosing haircut, which the engine refuses in active cases; it is shown only to size the assumption. Each later row restores one central-case assumption. Order matters because the effects interact; the end point is the central case, before the Monte Carlo averaging that sets the price target.', f_sub); r += 1
wsS.add_chart(CHART_S, f'B{r+1}')
for rr_ in range(BR0, BR0 + len(bridge) + 3): wsS.cell(rr_, 2).alignment = Alignment(wrap_text=True, vertical='top')
widths(wsS, {'A': 2, 'B': 64, 'C': 13, 'D': 11, 'E': 13, 'F': 11, 'G': 12, 'H': 11, 'I': 11, 'J': 11, 'K': 11, 'L': 11, 'M': 11, 'N': 11, 'O': 11, 'P': 11})
wsS.print_area = f'B1:P{r+22}'


# ================================================================== REVERSE DCF
wsRD = sheet('Reverse DCF', '0070C0'); title(wsRD, 'Reverse DCF: what the share price requires', 'The revenue ramp that reprices the stock with the Street ramp shape, approval odds and discount rate held fixed. Saved solve of 7 October 2026, rescaled live to the current price.')
RESOLVE = json.loads((ROOT / 'models/excel/reverse_resolve.json').read_text()); RV_EXT = RESOLVE['variants']['extended']; RV_PLAT = RESOLVE['variants']['plateau']
rr = RV_EXT['required_ramps']; ss = RV_EXT['solve_summary']; pp = ALT_A['guggenheim_revenue_path']['subtables']['parameters']
r = 4; band(wsRD, r, 'Method', 2, 12); r += 1
for line in ('Take the Guggenheim IR and XR revenue paths (9 September 2026 note) as the shape of the ramp. Build a normalised DCF around them: 90% approval odds for each product, 12.5% discount rate, 15% tax from 2031, 50/50 pooled-study and Kuvan erosion from 2039, explicit cash flows to 2055 and a declining residual.',
             'Equity value is linear in a common scale on both ramps: value = k x (PV of IR and XR contribution) + PV of shared costs + net cash. Solve k so that value equals the target price. k above 1 means the market needs more than the Street path; below 1 means less.',
             'Two solves anchor the tab: the $31.83 spot snapshot (horizon 0) and the $44 Guggenheim target (one-year horizon). Because value is linear in k, those two points identify the line, and the live rows rescale the ramp to today\'s share price without re-running the solve. The Street path ends in 2036 and is carried to 2038 at its own 2036 growth rates before generic erosion starts in 2039.'):
    put(wsRD, r, 2, line, f_sub); wsRD.cell(r, 2).alignment = Alignment(wrap_text=True, vertical='top'); wsRD.merge_cells(start_row=r, start_column=2, end_row=r, end_column=14); wsRD.row_dimensions[r].height = 26; r += 1
r += 1; band(wsRD, r, 'Solve summary: Street path carried to 2038 at its 2036 growth rates (re-solved 10 October 2026 on the 7 October inputs)', 2, 12); r += 1
for i, h in enumerate(ss['columns']): put(wsRD, r, 2 + i, h, f_hdr, fill=fill_hdr)
r += 1; SS = {}
for row_ in ss['rows']:
    put(wsRD, r, 2, row_[0]); put(wsRD, r, 3, row_[1], f_input, NUM3 if isinstance(row_[1], float) else GEN); put(wsRD, r, 4, row_[2], f_input, NUM3 if isinstance(row_[2], float) else GEN); put(wsRD, r, 5, row_[3], f_sub); SS[row_[0]] = r; r += 1
k_spot_row = next(v for k_, v in SS.items() if 'scale' in k_.lower()); tp_row = next(v for k_, v in SS.items() if k_.startswith('Target price')); hz_row = next(v for k_, v in SS.items() if 'horizon' in k_.lower())
rdisc = next((row_[1] for row_ in pp['rows'] if 'discount' in str(row_[0]).lower() or 'wacc' in str(row_[0]).lower()), 0.125)
r += 1; band(wsRD, r, 'Live rescaling to the current share price', 2, 12); r += 1
put(wsRD, r, 2, 'Discount rate in the reverse model'); put(wsRD, r, 3, rdisc, f_input, PCT); put(wsRD, r, 5, 'Guggenheim 12.5% WACC, kept so the saved solve stays comparable', f_sub); RD_R = r; r += 1
put(wsRD, r, 2, 'Target equity PV per share, spot solve'); put(wsRD, r, 3, f'=C{tp_row}/(1+C{RD_R})^C{hz_row}', f_norm, USD); RD_P1 = r; r += 1
put(wsRD, r, 2, 'Target equity PV per share, $44 solve'); put(wsRD, r, 3, f'=D{tp_row}/(1+C{RD_R})^D{hz_row}', f_norm, USD); RD_P2 = r; r += 1
put(wsRD, r, 2, 'Scale per dollar of value (slope)'); put(wsRD, r, 3, f'=(D{k_spot_row}-C{k_spot_row})/(C{RD_P2}-C{RD_P1})', f_norm, '0.00000'); RD_A = r; r += 1
put(wsRD, r, 2, 'Intercept (shared costs and net cash, in scale units)'); put(wsRD, r, 3, f'=C{RD_A}*C{RD_P1}-C{k_spot_row}', f_norm, '0.00000'); RD_B = r; r += 1
put(wsRD, r, 2, 'Share price today'); put(wsRD, r, 3, '=share_price', f_norm, USD); put(wsRD, r, 5, '=TEXT(price_date,"d mmm yyyy")', f_sub); r += 1
put(wsRD, r, 2, 'Scale the current price requires on the Street ramp', f_bold); put(wsRD, r, 3, f'=C{RD_A}*C{r-1}-C{RD_B}', f_bold, NUM3, fill=fill_key); name('rd_k_live', f"'Reverse DCF'!$C${r}"); RD_K = r; r += 1
put(wsRD, r, 2, 'Rescaling factor versus the spot solve'); put(wsRD, r, 3, f'=C{RD_K}/C{k_spot_row}', f_norm, NUM3); RD_F = r; r += 1
put(wsRD, r, 2, 'For reference: spot scale under the 7 October convention (2036 revenue held flat to 2038)'); put(wsRD, r, 3, RV_PLAT['k_spot'], f_input, NUM3); put(wsRD, r, 5, f"Guggenheim 2036 growth of {RESOLVE['guggenheim_2036_growth']['IR']:.1%} (IR) and {RESOLVE['guggenheim_2036_growth']['XR']:.1%} (XR) carried two more years; all other inputs unchanged", f_sub); r += 1
put(wsRD, r, 2, 'The linear rescaling is exact inside the reverse model as long as the solved path stays profitable in every taxed year, which holds across the plausible range. It inherits every assumption of that model, including the imposed ramp shape.', f_sub); wsRD.cell(r, 2).alignment = Alignment(wrap_text=True, vertical='top'); wsRD.merge_cells(start_row=r, start_column=2, end_row=r, end_column=12); wsRD.row_dimensions[r].height = 28; r += 2
band(wsRD, r, 'Required success-case worldwide revenue, USD millions, against this model (both products approved)', 2, 16); r += 1
yrs_ = [row_[0] for row_ in rr['rows']]
put(wsRD, r, 2, 'Calendar year', f_hdr, fill=fill_hdr); put(wsRD, r, 3, '', f_hdr, fill=fill_hdr)
for i, y in enumerate(yrs_):
    put(wsRD, r, 4 + i, y, f_hdr, fill=fill_hdr, align=Alignment(horizontal='right')); wsRD.column_dimensions[CL(4 + i)].width = 9.5
    if y > 2036: wsRD.column_dimensions[CL(4 + i)].outlineLevel = 1; wsRD.column_dimensions[CL(4 + i)].hidden = False
RDY = r; r += 1
def rdrow(label, vals=None, fn=None, bold=False, fmt=NUM, font=None):
    global r; put(wsRD, r, 2, label, f_bold if bold else f_norm)
    for i in range(len(yrs_)):
        v = fn(i) if fn else vals[i]; put(wsRD, r, 4 + i, v, font or (f_input if vals is not None else (f_bold if bold else f_norm)), fmt)
    r += 1; return r - 1
R_SIR = rdrow('Required at $31.83 spot: IR', [row_[1] for row_ in rr['rows']])
R_SXR = rdrow('Required at $31.83 spot: XR', [row_[2] for row_ in rr['rows']])
R_ST = rdrow('Required at $31.83 spot: combined', fn=lambda i: f'={CL(4+i)}{R_SIR}+{CL(4+i)}{R_SXR}', bold=True)
R_PT = rdrow('Required at the $44 target: combined', [row_[7] for row_ in rr['rows']])
R_LIR = rdrow('Required at today\'s price: IR', fn=lambda i: f'={CL(4+i)}{R_SIR}*$C${RD_F}')
R_LXR = rdrow('Required at today\'s price: XR', fn=lambda i: f'={CL(4+i)}{R_SXR}*$C${RD_F}')
R_LT = rdrow('Required at today\'s price: combined', fn=lambda i: f'={CL(4+i)}{R_LIR}+{CL(4+i)}{R_LXR}', bold=True)
AY0, AYL = AYEARS, arange
R_MIR = rdrow('This model: IR', fn=lambda i: f"=INDEX({AYL('ir_rev')},1,MATCH({CL(4+i)}${RDY},{AY0},0))")
R_MXR = rdrow('This model: XR', fn=lambda i: f"=INDEX({AYL('xr_rev')},1,MATCH({CL(4+i)}${RDY},{AY0},0))")
R_MT = rdrow('This model: combined', fn=lambda i: f'={CL(4+i)}{R_MIR}+{CL(4+i)}{R_MXR}', bold=True)
R_PCT = rdrow('This model as a share of what today\'s price requires', fn=lambda i: f'=IF({CL(4+i)}{R_LT}=0,"",{CL(4+i)}{R_MT}/{CL(4+i)}{R_LT})', fmt=PCT)
r += 1
put(wsRD, r, 2, 'Peak required, USD m'); put(wsRD, r, 4, f'=MAX(D{R_ST}:{CL(3+len(yrs_))}{R_ST})', f_norm, NUM); put(wsRD, r, 5, 'at $31.83', f_sub); put(wsRD, r, 7, f'=MAX(D{R_LT}:{CL(3+len(yrs_))}{R_LT})', f_norm, NUM); put(wsRD, r, 8, 'at today\'s price', f_sub); put(wsRD, r, 10, f'=MAX(D{R_MT}:{CL(3+len(yrs_))}{R_MT})', f_norm, NUM); put(wsRD, r, 11, 'this model', f_sub); r += 1
put(wsRD, r, 2, 'Cumulative 2027-2036, USD m'); put(wsRD, r, 4, f'=SUM(D{R_ST}:M{R_ST})', f_norm, NUM); put(wsRD, r, 7, f'=SUM(D{R_LT}:M{R_LT})', f_norm, NUM); put(wsRD, r, 10, f'=SUM(D{R_MT}:M{R_MT})', f_norm, NUM); r += 1
put(wsRD, r, 2, 'Required ramps are success-case worldwide revenues with the Street shape imposed, carried from 2036 to 2038 at the 2036 growth rates and eroded from 2039. The reverse model has no independently fitted S-curve, so a slower early ramp with a higher peak could support the same price. Source basis of the printed Guggenheim revenues (success-case or already probability-weighted) is not disclosed; the solve treats them as success-case. Replica solve: models/reverse_dcf/resolve_extended.py.', f_sub); wsRD.cell(r, 2).alignment = Alignment(wrap_text=True, vertical='top'); wsRD.merge_cells(start_row=r, start_column=2, end_row=r, end_column=14); wsRD.row_dimensions[r].height = 42; r += 2
chr_ = LineChart(); chr_.title = 'Required revenue versus this model, USD m'; chr_.height = 8; chr_.width = 20; chr_.style = 2
for rowi, lab in ((R_ST, 'Required at $31.83'), (R_LT, "Required at today's price"), (R_MT, 'This model')):
    chr_.add_data(Reference(wsRD, min_col=4, max_col=3 + len(yrs_), min_row=rowi), from_rows=True, titles_from_data=False); chr_.series[-1].tx = SeriesLabel(v=lab)
chr_.set_categories(Reference(wsRD, min_col=4, max_col=3 + len(yrs_), min_row=RDY)); wsRD.add_chart(chr_, f'B{r}')
widths(wsRD, {'A': 2, 'B': 58, 'C': 13}); wsRD.print_area = f'B1:P{r+20}'

# ================================================================== MONTE CARLO (compact) and DCF WALK
wsK = sheet('Checks', 'C00000'); title(wsK, 'Checks')
PD = MCX['perdraw']; NDRAW = len(PD); W_ = MCX['weights']; ST4 = ['both', 'IR_only', 'XR_only', 'neither']
def qtl(vals, pp): v = sorted(vals); return v[min(len(v) - 1, int(pp * len(v)))]
et = [d['expected_target'] for d in PD]; vt = [d['value_today'] for d in PD]
pairs = sorted((d['state_target'][s_], W_[s_] / NDRAW) for d in PD for s_ in ST4)
def wq(pp):
    acc = 0.0
    for v, w in pairs:
        acc += w
        if acc >= pp: return v
    return pairs[-1][0]
edges = list(range(0, 62, 2)); hist_e = [0.0] * (len(edges) - 1); hist_o = [0.0] * (len(edges) - 1)
binidx = lambda v: min(max(int(v // 2), 0), len(edges) - 2)
for d in PD:
    hist_e[binidx(d['expected_target'])] += 1 / NDRAW
    for s_ in ST4: hist_o[binidx(d['state_target'][s_])] += W_[s_] / NDRAW
put(wsK, 4, 10, 'Monte Carlo histogram bins ($2) feeding the chart on the Monte Carlo tab', f_bold); put(wsK, 5, 10, 'Bin', f_hdr, fill=fill_hdr); put(wsK, 5, 11, 'Expected values', f_hdr, fill=fill_hdr); put(wsK, 5, 12, 'Clinical outcomes', f_hdr, fill=fill_hdr)
for i in range(len(edges) - 1): put(wsK, 6 + i, 10, edges[i], f_norm, USD0); put(wsK, 6 + i, 11, hist_e[i], f_input, PCT); put(wsK, 6 + i, 12, hist_o[i], f_input, PCT)
H0, H1 = 6, 5 + len(edges) - 1
def spearman(x, y):
    def rk(v):
        o = sorted(range(len(v)), key=lambda i: v[i]); rr_ = [0.0] * len(v); i = 0
        while i < len(o):
            j = i
            while j + 1 < len(o) and v[o[j + 1]] == v[o[i]]: j += 1
            for k_ in range(i, j + 1): rr_[o[k_]] = (i + j) / 2 + 1
            i = j + 1
        return rr_
    rx, ry = rk(x), rk(y); mx, my = sum(rx) / len(rx), sum(ry) / len(ry)
    cov = sum((a - mx) * (b - my) for a, b in zip(rx, ry)); vx = sum((a - mx) ** 2 for a in rx) ** .5; vy = sum((b - my) ** 2 for b in ry) ** .5
    return cov / (vx * vy) if vx and vy else 0.0
sp = sorted(((k_, spearman([d['inputs'][k_] for d in PD], et)) for k_ in priors['priors']), key=lambda t: -abs(t[1]))
wsM = sheet('Monte Carlo', '0070C0'); title(wsM, 'Monte Carlo')
r = 4; band(wsM, r, '4,096 draws of the 28 uncertain inputs through the patient engine, financials and four-branch DCF (seed 7202609, 11 October 2026, Crimson priors)', 2, 7); r += 1
for i, h in enumerate(['USD per share', 'Mean', 'Median', 'P10', 'P90', '']): put(wsM, r, 2 + i, h, f_hdr, fill=fill_hdr)
r += 1
put(wsM, r, 2, '12-month value: the price target is the mean', f_bold); put(wsM, r, 3, sum(et) / NDRAW, f_input, USD, fill=fill_key); name('mc_mean_target', f"'Monte Carlo'!$C${r}"); [put(wsM, r, 4 + i, qtl(et, pp), f_input, USD) for i, pp in enumerate((0.5, 0.1, 0.9))]; r += 1
put(wsM, r, 2, 'Fair value today'); put(wsM, r, 3, sum(vt) / NDRAW, f_input, USD); [put(wsM, r, 4 + i, qtl(vt, pp), f_input, USD) for i, pp in enumerate((0.5, 0.1, 0.9))]; r += 1
put(wsM, r, 2, '12-month clinical outcomes, four branches weighted once'); put(wsM, r, 3, sum(v * w for v, w in pairs), f_input, USD); [put(wsM, r, 4 + i, wq(pp), f_input, USD) for i, pp in enumerate((0.5, 0.1, 0.9))]; r += 1
put(wsM, r, 2, 'Share of draws below the share price'); put(wsM, r, 3, sum(1 for v in et if v < PRICE) / NDRAW, f_input, PCT); put(wsM, r, 4, 'a sensitivity statistic under these priors, not a probability of a profitable short', f_sub); r += 1
put(wsM, r, 2, 'Live inputs versus the simulated priors', f_bold); put(wsM, r, 3, f'=IF(SUMPRODUCT(--(Inputs!$H$9:$H${max(INPUT_ROW.values())}<>Inputs!$E$9:$E${max(INPUT_ROW.values())}))+ABS(launch_delay_IR)+ABS(launch_delay_XR)+ABS(loe_year_IR-loe_year)+ABS(loe_year_XR-loe_year)+ABS(discount_rate-{P['discount_rate']['value']})=0,"Inputs match the simulated baseline","STALE: live inputs differ from the simulated baseline")', f_flag); r += 2
band(wsM, r, 'Clinical branches', 2, 7); r += 1
for i, h in enumerate(['Branch', 'Weight', 'Median value, 12 months', 'Median funding, USD m', '', '']): put(wsM, r, 2 + i, h, f_hdr, fill=fill_hdr, align=Alignment(wrap_text=True, vertical='center'))
wsM.row_dimensions[r].height = 27
r += 1
for s_ in ST4: put(wsM, r, 2, SNAME[s_]); put(wsM, r, 3, W_[s_], f_input, PCT2); put(wsM, r, 4, qtl([d['state_target'][s_] for d in PD], 0.5), f_input, USD); put(wsM, r, 5, qtl([d['state_funding'][s_] for d in PD], 0.5), f_input, NUM); r += 1
r += 1; band(wsM, r, 'The 28 sampled inputs: triangular priors (minimum, most likely, maximum)', 2, 7); r += 1
for i, h in enumerate(['Input', 'Minimum', 'Most likely', 'Maximum', 'Rank association', 'Basis']): put(wsM, r, 2 + i, h, f_hdr, fill=fill_hdr)
r += 1; spd = dict(sp)
for k_, v in priors['priors'].items():
    put(wsM, r, 2, LABELS.get(k_, k_), f_bold if k_ in THESIS else f_norm); put(wsM, r, 3, v['min'], f_input, NUM3); put(wsM, r, 4, v['mode'], f_input, NUM3); put(wsM, r, 5, v['max'], f_input, NUM3); put(wsM, r, 6, spd[k_], f_input, NUM2); put(wsM, r, 7, ('Crimson underwriting; ' if k_ in CRIM['priors'] else '') + ', '.join(v['sources']), f_sub); r += 1
put(wsM, r, 2, 'Rank association is the Spearman correlation between the sampled input and the 12-month value across draws: prior-dependent, not causal. Fixed policies (redosing haircut, 60% mature share, price multipliers) do not vary.', f_sub); wsM.cell(r, 2).alignment = Alignment(wrap_text=True, vertical='top'); wsM.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7); wsM.row_dimensions[r].height = 28; r += 2
band(wsM, r, 'Structural stresses and access packages on the central case (engine runs)', 2, 7); r += 1
for i, h in enumerate(['Scenario', 'Value today', '12-month value', 'Change', 'Peak revenue, USD m']): put(wsM, r, 2 + i, h, f_hdr, fill=fill_hdr)
r += 1
SCN = {'hazards_halved_after_year5': 'Switching hazards halved after year 5', 'quarter_of_stoppers_retry_after24m': '25% of discontinuers return after 24 months', 'other_prevention_shift20pct_by10y': '20% of rescue-only users move to other prevention by year 10', 'generic_price_capped_flat': 'Post-generic price capped flat', 'initial_rescue_stock_free': 'Initial rescue stock supplied free', 'partner_all_overseas': 'All overseas sales licensed (20% royalty)', 'six_month_launch_delay_both': 'Both launches delayed six months', 'no_refill_intensity_adjustment': 'No refill-gap adjustment on XR', 'old_prices_only': 'Older price anchors ($14,000 and $418,000)', 'omit_license_royalty_cost_counterfactual': 'No BRAIN royalty (counterfactual)', 'redose_plus_better_IR_capture': '1.25x injectable rescue opportunity', 'LTP_ten_percent_no_rescue': '10% of preventive users hold no rescue', 'extra_Ekterly_free_stock': '20% extra Ekterly free supply', 'rebate_for_access': 'Access package: more rebate, better access', 'restricted_high_net': 'Access package: restricted access, higher net'}
for x in CR['scenarios']:
    put(wsM, r, 2, SCN.get(x['name'], x['name'])); put(wsM, r, 3, x['value_today'], f_input, USD); put(wsM, r, 4, x['expected_target'], f_input, USD); put(wsM, r, 5, x['delta_target'], f_norm, USD); put(wsM, r, 6, x['peak']['recognized_revenue_usdm'], f_input, NUM); r += 1
MC_LEFT_END = r
# right-hand column: charts, then the one-variable tornado
bc = BarChart(); bc.type = 'col'; bc.grouping = 'clustered'; bc.title = 'Distribution of the 12-month value'; bc.height = 8.5; bc.width = 17; bc.y_axis.title = 'Share of draws'; bc.x_axis.title = 'USD per share, $2 bins'; bc.style = 2; bc.gapWidth = 30
bc.add_data(Reference(wsK, min_col=11, max_col=12, min_row=H0 - 1, max_row=H1), titles_from_data=True); bc.set_categories(Reference(wsK, min_col=10, min_row=H0, max_row=H1)); wsM.add_chart(bc, 'I4')
rr0 = 4; put(wsM, rr0, 17, 'Rank association', f_hdr, fill=fill_hdr); put(wsM, rr0, 16, 'Input', f_hdr, fill=fill_hdr)
for i, (k_, v) in enumerate(sp[:14]): put(wsM, rr0 + 1 + i, 16, LABELS.get(k_, k_)); put(wsM, rr0 + 1 + i, 17, v, f_input, NUM2)
sc_ = BarChart(); sc_.type = 'bar'; sc_.title = 'Leading drivers under these priors'; sc_.height = 8.5; sc_.width = 17; sc_.style = 2; sc_.legend = None
sc_.add_data(Reference(wsM, min_col=17, min_row=rr0 + 1, max_row=rr0 + 14)); sc_.set_categories(Reference(wsM, min_col=16, min_row=rr0 + 1, max_row=rr0 + 14)); sc_.x_axis.scaling.orientation = 'maxMin'; wsM.add_chart(sc_, 'I23')
r = 42; band(wsM, r, 'One-variable sensitivities on the central case, ranked by span (engine runs)', 9, 15); r += 1
for i, h in enumerate(['Input', 'Low', 'High', 'Value at low', 'Value at high', 'Span']): put(wsM, r, 9 + i, h, f_hdr, fill=fill_hdr)
r += 1
for x in CR['sensitivities']:
    k_ = x['parameter']; put(wsM, r, 9, LABELS.get(k_, k_), f_bold if k_ in THESIS else f_norm); put(wsM, r, 10, x['low'], f_norm, NUM3); put(wsM, r, 11, x['high'], f_norm, NUM3); put(wsM, r, 12, x['target_low'], f_input, USD); put(wsM, r, 13, x['target_high'], f_input, USD); put(wsM, r, 14, abs(x['target_high'] - x['target_low']), f_norm, USD); r += 1
put(wsM, r, 9, 'Spans are chosen bounds, not statistical importance; the central case is the point at every most-likely value.', f_sub)
widths(wsM, {'A': 2, 'B': 50, 'C': 13, 'D': 14, 'E': 14, 'F': 17, 'G': 44, 'H': 3, 'I': 44, 'J': 10, 'K': 10, 'L': 12, 'M': 12, 'N': 10, 'O': 3, 'P': 36, 'Q': 17})
wsM.print_area = f'B1:Q{max(MC_LEFT_END, r) + 1}'

# ---- DCF walk on the expected cash flows
wsW = sheet('DCF Walk', '0070C0'); title(wsW, 'Discounted cash flow')
DLC = 3 + (END_YEAR - 2027)
EXA = MCX['expected']['annual']; MS = MCX['states']; W3 = {s_: W_[s_] for s_ in ('both', 'IR_only', 'XR_only')}
def ex3(y, key): return sum(W3[s_] * MS[s_]['annual'][str(y)][key] for s_ in W3)
def ex4(y, key): return sum(W_[s_] * MS[s_]['annual'][str(y)][key] for s_ in ST4)
def exv3(key): return sum(W3[s_] * MS[s_]['target'][key] for s_ in W3)
r = 4; band(wsW, r, 'Price target', 2, DLC); r += 1
put(wsW, r, 2, '12-month price target: mean value across the 4,096 Monte Carlo draws', f_bold); put(wsW, r, 3, '=mc_mean_target', f_bold, USD, fill=fill_key); name('price_target', f"'DCF Walk'!$C${r}"); r += 1
put(wsW, r, 2, 'Share price'); put(wsW, r, 3, '=share_price', f_norm, USD); put(wsW, r, 4, '=TEXT(price_date,"d mmm yyyy")', f_sub); r += 1
put(wsW, r, 2, 'Implied 12-month return', f_bold); put(wsW, r, 3, '=price_target/share_price-1', f_bold, PCT); name('implied_return', f"'DCF Walk'!$C${r}"); r += 1
put(wsW, r, 2, 'Each draw samples the 28 uncertain inputs, runs the patient engine, the financials and a dated DCF for the four clinical branches, and returns one value. Discounting is linear, so the mean of those values is the DCF of the mean cash flows shown below; only the per-share step (option dilution and the zero floor inside each branch) is applied draw by draw.', f_sub); wsW.cell(r, 2).alignment = Alignment(wrap_text=True, vertical='top'); wsW.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10); wsW.row_dimensions[r].height = 40; r += 2
band(wsW, r, 'Clinical branches (probabilities applied once, inside the expected cash flows)', 2, DLC); r += 1
for s_, lab in (('both', 'Both approved'), ('IR_only', 'IR approved, XR fails'), ('XR_only', 'XR approved, IR fails'), ('neither', 'Neither approved')):
    put(wsW, r, 2, lab); put(wsW, r, 3, f'=w_{s_}', f_norm, PCT); r += 1
put(wsW, r, 2, 'P(both) = P(IR) x P(XR) / shared molecule gate; the marginals stay at 90% each.', f_sub); r += 2
band(wsW, r, 'Expected cash flows: probability-weighted across branches and averaged over draws, USD millions', 2, DLC); r += 1
DY = list(range(2027, END_YEAR + 1)); DC0 = 3
put(wsW, r, 2, 'Year', f_hdr, fill=fill_hdr)
for i, y in enumerate(DY): put(wsW, r, DC0 + i, y, f_hdr, fill=fill_hdr, align=Alignment(horizontal='right')); wsW.column_dimensions[CL(DC0 + i)].width = 9.5
DYR = r; r += 1
def wrow(label, vals, fmt=NUM, bold=False, font=None):
    global r; put(wsW, r, 2, label, f_bold if bold else f_norm)
    for i, v in enumerate(vals): put(wsW, r, DC0 + i, v, font or (f_input if not (isinstance(v, str) and v.startswith('=')) else (f_bold if bold else f_norm)), fmt)
    r += 1; return r - 1
W_REV = wrow('Expected revenue', [ex4(y, 'revenue') for y in DY])
W_CEN = wrow('   Central case revenue, both approved (Three Statements, live)', ['=' + aref('revenue', y) for y in DY])
W_COST = wrow('Expected operating costs and tax', [ex4(y, 'revenue') - ex4(y, 'net_income') for y in DY])
W_NI = wrow('Expected net income', [ex4(y, 'net_income') for y in DY])
W_FCF = wrow('Expected free cash flow after the target date (success branches)', [ex3(y, 'fcf_after_target') for y in DY], bold=True)
W_PV = wrow('Present value at 9 October 2027', [ex3(y, 'pv_target') for y in DY])
wrow('Effective discount factor', [f'=IF({CL(DC0+i)}{W_FCF}=0,"",{CL(DC0+i)}{W_PV}/{CL(DC0+i)}{W_FCF})' for i in range(len(DY))], fmt=NUM3)
wrow('Cumulative present value', [f'=SUM($C${W_PV}:{CL(DC0+i)}${W_PV})' for i in range(len(DY))])
r += 1
put(wsW, r, 2, 'Sum of present values, 2027-2055'); put(wsW, r, 3, f'=SUM(C{W_PV}:{CL(DC0+len(DY)-1)}{W_PV})', f_norm, NUM); WS0 = r; r += 1
put(wsW, r, 2, 'Expected residual after 2055 (declining perpetuity)'); put(wsW, r, 3, exv3('terminal_pv_usdm'), f_input, NUM); r += 1
put(wsW, r, 2, 'Expected cash at the target date'); put(wsW, r, 3, exv3('cash_usdm'), f_input, NUM); r += 1
put(wsW, r, 2, 'Less expected issuance fees'); put(wsW, r, 3, -exv3('financing_fee_pv_usdm'), f_input, NUM); r += 1
put(wsW, r, 2, 'Expected value if neither product is approved (liquidation)'); put(wsW, r, 3, W_['neither'] * MS['neither']['target']['equity_usdm'], f_input, NUM); r += 1
put(wsW, r, 2, 'Expected equity value', f_bold); put(wsW, r, 3, f'=SUM(C{WS0}:C{r-1})', f_bold, NUM); WEQ = r; r += 1
put(wsW, r, 2, 'Expected diluted shares, millions'); put(wsW, r, 3, sum(W_[s_] * MS[s_]['target']['diluted_shares_m'] for s_ in ST4), f_input, NUM3); r += 1
put(wsW, r, 2, 'Expected equity per expected share'); put(wsW, r, 3, f'=C{WEQ}/C{r-1}', f_norm, USD); r += 1
put(wsW, r, 2, 'Price target: mean of the per-draw values', f_bold); put(wsW, r, 3, '=price_target', f_bold, USD, fill=fill_key); r += 1
put(wsW, r, 2, 'Difference: option dilution and the zero floor applied within each draw and branch'); put(wsW, r, 3, f'=C{r-1}-C{r-2}', f_norm, USD); r += 1
put(wsW, r, 2, 'Expected revenue is below the central case for two reasons: it carries the 15% probability that one or both products fail, and the priors are lopsided (several inputs can only disappoint from their most-likely value). The central case on Three Statements is the live engine; this walk is the Monte Carlo mean of that engine.', f_sub); wsW.cell(r, 2).alignment = Alignment(wrap_text=True, vertical='top'); wsW.merge_cells(start_row=r, start_column=2, end_row=r, end_column=12); wsW.row_dimensions[r].height = 40; r += 2
SENSW = {x['parameter']: x for x in CR['sensitivities']}
band(wsW, r, 'Discount rate: WACC build from public data (Damodaran, January 2026; US Treasury par yields, October 2026)', 2, DLC); r += 1
def wl(label, val, fmt, note='', nm=None, bold=False, fill=None):
    global r; put(wsW, r, 2, label, f_bold if bold else f_norm)
    put(wsW, r, 3, val, (f_bold if bold else f_norm) if (isinstance(val, str) and val.startswith('=')) else f_input, fmt, fill=fill)
    if note: put(wsW, r, 5, note, f_sub)
    if nm: name(nm, f"'DCF Walk'!$C${r}")
    r += 1; return r - 1
R_RF = wl('Risk-free rate: 10-year Treasury par yield, 9 October 2026', 0.0524, PCT2, 'US Treasury daily par yield curve, last close before the 10 October price date (a Saturday); the ten-year tenor matches the cash-flow horizon.', 'rf_rate')
R_ERP = wl('Equity risk premium: implied, S&P 500, January 2026', 0.0423, PCT2, 'Damodaran implied premium from the index level and expected cash flows, on a 4.18% T-bond; equals his mature-market premium after the 0.23% US default spread.', 'erp')
R_BU = wl('Unlevered beta corrected for cash: Drugs (Biotechnology), 496 US firms', 1.08, NUM2, 'Damodaran Betas by Sector, January 2026: levered 1.14, D/E 13.0%, unlevered 1.03, cash 4.2% of firm value. Cash-corrected because cash is added separately above.', 'beta_unlevered')
R_D = wl('Debt, USD m', 0, NUM, 'June 2026 interim: no term debt; leases are immaterial.', 'debt_usdm')
R_E = wl('Market capitalisation at the price date, USD m', '=share_price*open_shares_m', NUM, 'share price times basic shares; options enter through the treasury-stock dilution in the per-share step.')
R_DE = wl('Debt to equity', f'=IF(C{R_E}=0,0,C{R_D}/C{R_E})', PCT2)
R_T = wl('Marginal tax rate', '=tax_rate', PCT, 'Inputs; applied once the NOL pool is used.')
R_BL = wl('Relevered beta', f'=C{R_BU}*(1+(1-C{R_T})*C{R_DE})', NUM2, 'levered beta = unlevered beta x (1 + (1 - t) x D/E); equals the unlevered beta with no debt. The traded five-year beta is negative on a short history and is not usable.', 'beta')
R_KE = wl('Cost of equity (CAPM)', '=rf_rate+beta*erp', PCT2, 'risk-free rate + beta x equity risk premium', 'cost_of_equity', bold=True)
R_SPR = wl('Default spread if Pharvaris borrowed: synthetic rating on negative EBIT', 0.19, PCT2, 'Damodaran ratings table, January 2026: interest cover below 0.2 maps to D with a 19.00% spread; no investment-grade debt capacity, so the model funds with equity.')
R_KD = wl('Pre-tax cost of debt', f'=rf_rate+C{R_SPR}', PCT2)
R_KDA = wl('After-tax cost of debt', f'=C{R_KD}*(1-C{R_T})', PCT2)
R_WE = wl('Equity weight', f'=IF(C{R_D}+C{R_E}=0,1,C{R_E}/(C{R_D}+C{R_E}))', PCT)
R_WD = wl('Debt weight', f'=1-C{R_WE}', PCT)
R_W = wl('WACC', f'=C{R_WE}*C{R_KE}+C{R_WD}*C{R_KDA}', PCT2, 'all-equity capital structure, so WACC equals the cost of equity', 'wacc', bold=True)
R_PR = wl('Premium for single-asset execution and funding risk', 0.022, PCT2, 'analyst choice; approval risk sits in the branch weights, so this covers launch execution, the equity raises the model expects and thin trading history.', 'single_asset_premium')
R_SUM = wl('WACC plus premium', f'=C{R_W}+C{R_PR}', PCT2)
wl('Discount rate used: rounded to 0.1% and written to Inputs', '=discount_rate', PCT, 'Inputs takes ROUND(WACC + premium, 0.1%); it sets the Calendar discount factors for the central case and was the rate inside every Monte Carlo draw.', bold=True, fill=fill_key)
wl('Street discount rates, September 2026 notes: lowest / median / highest', f"=TEXT(MIN('Street Bridge'!G{ST0}:G{ST1}),\"0.0%\")&\" / \"&TEXT(MEDIAN('Street Bridge'!G{ST0}:G{ST1}),\"0.0%\")&\" / \"&TEXT(MAX('Street Bridge'!G{ST0}:G{ST1}),\"0.0%\")", GEN, 'from the broker table on the Reverse DCF tab; Guggenheim uses a 12.5% WACC')
if 'discount_rate' in SENSW: wl('Central case 12-month value at a 10% and a 15% discount rate (engine runs)', f"${SENSW['discount_rate']['target_low']:.2f} / ${SENSW['discount_rate']['target_high']:.2f}", GEN, 'the saved one-variable sensitivity on the Monte Carlo tab; the price target itself is fixed at the 12% used in the draws')
widths(wsW, {'A': 2, 'B': 66, 'C': 12}); wsW.print_area = f'B1:AE{r}'

# ---- Checks: control totals from the Crimson central case
title(wsK, 'Checks')
r = 4; band(wsK, r, 'Control totals: live workbook against the saved engine run of the central case', 2, 8); r += 1
for i, h in enumerate(['Metric', 'Engine (saved)', 'Workbook (live)', 'Difference', 'Tolerance', 'Status']): put(wsK, r, 2 + i, h, f_hdr, fill=fill_hdr)
r += 1; K0 = r
w = CR['base']; ann = {a['year']: a for a in w['annual_both_success']}
def krow(label, engine_val, live_formula, tol, fmt=USD):
    global r; put(wsK, r, 2, label); put(wsK, r, 3, engine_val, f_input, fmt); put(wsK, r, 4, live_formula, f_norm, fmt); put(wsK, r, 5, f'=D{r}-C{r}', f_norm, fmt); put(wsK, r, 6, tol, f_sub, fmt); put(wsK, r, 7, f'=IF(ABS(E{r})<=F{r},"OK","CHECK")', f_bold); r += 1
krow('Value today, probability-weighted', w['risk_adjusted_value_today'], '=value_today', 0.01)
krow('12-month value, probability-weighted', w['expected_12month_value'], '=value_target', 0.01)
for s_ in ('both', 'IR_only', 'XR_only', 'neither'):
    ci_ = 3 + ['both', 'IR_only', 'XR_only', 'neither'].index(s_)
    krow(f'{SNAME[s_]}: 12-month value', w['conditional_valuations'][s_]['target']['price'], f"=Valuation!{CL(ci_)}{VALROW[('tgt','price')]}", 0.01)
    krow(f'{SNAME[s_]}: value today', w['conditional_valuations'][s_]['valuation']['price'], f"=Valuation!{CL(ci_)}{VALROW[('val','price')]}", 0.01)
krow('Peak revenue, USD m', w['conditional_peak']['recognized_revenue_usdm'], f"=MAX({arange('revenue')})", 0.01, NUM2)
for y in (2027, 2028, 2029, 2030, 2031, 2038):
    L = CL(YC0 + YEARS.index(y)); krow(f'{y} revenue, USD m', ann[y]['revenue'], '=' + aref('revenue', y), 0.01, NUM2)
for y in (2030, 2038):
    L = CL(YC0 + YEARS.index(y)); krow(f'{y} IR revenue', ann[y]['IR_recognized'], '=' + aref('ir_rev', y), 0.01, NUM2); krow(f'{y} XR revenue', ann[y]['XR_recognized'], '=' + aref('xr_rev', y), 0.01, NUM2)
for y in (2027, 2028, 2031):
    L = CL(YC0 + YEARS.index(y)); krow(f'{y} free cash flow', ann[y]['fcf'], '=' + aref('fcf', y), 0.01, NUM2); krow(f'{y} year-end cash', ann[y]['cash_end'], '=' + aref('cash', y), 0.01, NUM2); krow(f'{y} net funding', ann[y]['funding_net'], '=' + aref('fund', y), 0.01, NUM2); krow(f'{y} shares, millions', ann[y]['shares_end'], '=' + aref('shares', y), 0.001, NUM3)
krow('2030 XR patient-years', ann[2030]['xr_patient_years'], f'=Annual!{CL(YC0+YEARS.index(2030))}{AN["xr_py"]}', 0.01, NUM2); krow('2030 IR paid doses', ann[2030]['IR_units'], f'=Annual!{CL(YC0+YEARS.index(2030))}{AN["ir_units"]}', 0.01, NUM2)
K1 = r - 1
put(wsK, r, 2, 'All control totals pass', f_bold); put(wsK, r, 7, f'=IF(COUNTIF(G{K0}:G{K1},"CHECK")=0,"OK","CHECK")', f_flag); name('checks_controls', f'Checks!$G${r}'); r += 2
band(wsK, r, 'Identities (should be zero or true in every case)', 2, 8); r += 1
def irow(label, formula, tol=1e-6, fmt=NUM3):
    global r; put(wsK, r, 2, label); put(wsK, r, 4, formula, f_norm, fmt); put(wsK, r, 6, tol, f_sub, fmt); put(wsK, r, 7, f'=IF(ABS(D{r})<=F{r},"OK","CHECK")', f_bold); r += 1
I0 = r
irow('Origin weights sum to 100% (difference)', '=w_orladeyo+w_other_prevention+w_acute_only+w_diagnosed_untreated-1')
irow('Joint pools sum to population (difference)', f"='Adoption Build'!C{JPT+1}")
irow('Clinical branch weights sum to 100% (difference)', '=w_both+w_IR_only+w_XR_only+w_neither-1')
for s_ in ('both', 'IR_only', 'XR_only', 'neither'): irow(f'{SNAME[s_]}: largest monthly balance-sheet error', f"=SUMPRODUCT(MAX(ABS({mrng('Fin Monthly', RF[(s_,'bal')])})))")
irow('XR engine: largest population conservation error', f"=MAX({mrng('XR Engine', RX['conserve'])})")
irow('IR engine: largest population conservation error', f"=MAX({mrng('IR Engine', RI['conserve'])})")
irow('XR engine: most negative available pool', '=MIN(' + ','.join(f"MIN({mrng('XR Engine', RX[(f'{o}|{s}', 'avail')])})" for (o, s) in CELLS) + ')', 1e-8)
irow('IR engine: most negative available pool', '=MIN(' + ','.join(f"MIN({mrng('IR Engine', RI[(f'{o}|{s}', 'avail')])})" for (o, s) in CELLS) + ')', 1e-8)
irow('Monthly revenue equals annual revenue (difference)', f"=SUM({mrng('Fin Monthly', FINB('revenue'))})-{tsum('revenue', 2026, END_YEAR)}")
irow('Monthly revenue equals quarterly revenue through 2031 (difference)', f"=SUMIFS({mrng('Fin Monthly', FINB('revenue'))},{mrng('Fin Monthly', YEAR_ROW)},\"<=2031\")-SUMIFS('Three Statements'!$D${TS['revenue']}:${TLAST}${TS['revenue']},{AFLAG},\"Q\")")
irow('Revenue before IR launch (should be zero)', f"=SUMPRODUCT({mrng('Fin Monthly', FINB('revenue'))},--({mrng('Calendar', CAL['age_IR'])}<0))")
irow('Price x units equals US revenue, XR (largest difference)', f"=SUMPRODUCT(MAX(ABS({mrng('Revenue', RV['XR_units'])}*{mrng('Revenue', RV['XR_price'])}/1000000-{mrng('Revenue', RV['XR_US'])})))")
irow('Equity value equals price x diluted shares, both approved (difference)', f"=Valuation!C{VALROW[('tgt','price')]}*Valuation!C{VALROW[('tgt','diluted')]}-Valuation!C{VALROW[('tgt','equity')]}")
irow('Three Statements: largest balance-sheet check', f"=SUMPRODUCT(MAX(ABS('Three Statements'!$D${TS['bs_check']}:${TLAST}${TS['bs_check']})))")
irow('Three Statements: largest cash-flow check', f"=SUMPRODUCT(MAX(ABS('Three Statements'!$D${TS['cf_check']}:${TLAST}${TS['cf_check']})))")
irow('Opening balance: assets equal equity (EUR m)', f"=History!C{OPEN['equity_eurm']+1}")
irow('Reported cash roll-forward, FY2025', f"=History!D{HIST_ROW['shares_end']+1}")
I1 = r - 1
put(wsK, r, 2, 'All identities pass', f_bold); put(wsK, r, 7, f'=IF(COUNTIF(G{I0}:G{I1},"CHECK")=0,"OK","CHECK")', f_flag); name('checks_identities', f'Checks!$G${r}'); r += 2
put(wsK, r, 2, 'Passing checks show the workbook reproduces the engine arithmetic. They do not validate the commercial assumptions; the Inputs tab and the research archive carry that argument.', f_sub)
widths(wsK, {'A': 2, 'B': 56, 'C': 14, 'D': 14, 'E': 12, 'F': 10, 'G': 9, 'H': 3, 'I': 3, 'J': 10, 'K': 14, 'L': 14})

# ================================================================== SOURCES
wsSo = sheet('Sources', 'FFC000'); title(wsSo, 'Sources')
r = 4
for c_, h in ((2, 'Source'), (3, 'Title'), (4, 'Date'), (12, 'Where in the source'), (13, 'Link'), (14, 'What it establishes, and its limitation')): put(wsSo, r, c_, h, f_hdr, fill=fill_hdr)
for c_ in range(2, 15): wsSo.cell(r, c_).fill = fill_hdr
r += 1
for k, v in src.items():
    put(wsSo, r, 2, k, f_bold); put(wsSo, r, 3, v.get('title', '')); put(wsSo, r, 4, v.get('date', '') or ''); put(wsSo, r, 12, fix_spacing(v.get('locator', '') or ''), f_sub)
    u = v.get('url'); c = put(wsSo, r, 13, 'open' if u else '', f_link)
    if u: c.hyperlink = u
    put(wsSo, r, 14, fix_spacing((v.get('facts', '') or '') + ((' Limitation: ' + v['limitation']) if v.get('limitation') else '')), f_sub)
    for cc in (3, 12, 14): wsSo.cell(r, cc).alignment = Alignment(wrap_text=True, vertical='top')
    r += 1
put(wsSo, r + 1, 2, 'Public filings, publications and company materials only; expert-call material is not an input.', f_sub)
widths(wsSo, {'A': 2, 'B': 14, 'C': 40, 'D': 12, 'E': 3, 'F': 3, 'G': 3, 'H': 3, 'I': 3, 'J': 3, 'K': 3, 'L': 34, 'M': 8, 'N': 90}); wsSo.freeze_panes = 'C5'
# ================================================================== ALT DATA TABS
def write_dataset(ws, r, ds, level=0):
    """Generic writer for the agent-compiled datasets."""
    t = ds.get('title', ''); u = ds.get('unit', '')
    if t: band(ws, r, t, 2, 12); r += 1
    if u: put(ws, r, 2, u, f_sub); r += 1
    if ds.get('columns') and ds.get('rows') is not None:
        for i, h in enumerate(ds['columns']): put(ws, r, 2 + i, h, f_hdr, fill=fill_hdr, align=Alignment(wrap_text=True))
        r += 1
        for row_ in ds['rows']:
            for i, v in enumerate(row_):
                if isinstance(v, (list, dict)): v = json.dumps(v)
                c = put(ws, r, 2 + i, v, f_input if isinstance(v, (int, float)) and not isinstance(v, bool) else f_norm, NUM2 if isinstance(v, float) else GEN)
                c.alignment = Alignment(wrap_text=True, vertical='top')
            r += 1
        r += 1
    for key in ('tables', 'subtables'):
        for sk, sd in (ds.get(key) or {}).items():
            if not sd.get('title'): sd = {**sd, 'title': sk.replace('_', ' ').capitalize()}
            r = write_dataset(ws, r, sd, level + 1)
    for n in ds.get('notes', []) or []: put(ws, r, 2, n, f_sub); r += 1
    srcs = ds.get('sources') or []
    if srcs:
        put(ws, r, 2, 'Sources', f_bold); r += 1
        for s_ in srcs:
            nm_ = s_.get('name') or s_.get('source') or ''; d_ = s_.get('date') or ''; u_ = s_.get('url') or ''; p_ = s_.get('path') or ', '.join(s_.get('repo_files', []) or [])
            put(ws, r, 2, f'{nm_} ({d_})' if d_ else nm_, f_sub); c = put(ws, r, 5, u_ or p_, f_link if u_ else f_sub)
            if u_: c.hyperlink = u_
            r += 1
    return r + 1

ALT_TABS = [('Street Targets', [ALT_A['street_targets']]), ('Street Revenue Path', [ALT_A['guggenheim_revenue_path']]), ('Orladeyo Launch', [ALT_B['orladeyo_launch']]), ('Ekterly Launch', [ALT_B['ekterly_launch']]), ('HAE Injectables', [ALT_A['hae_injectable_revenues']]), ('Generic Erosion', [ALT_A['generic_erosion_panel']]), ('Clinical Cross-Trial', [ALT_A['clinical_cross_trial']]), ('Second-Oral Analogues', [ALT_A['second_oral_analogues']]), ('Persistence Benchmarks', [ALT_B['persistence_switching_benchmarks']]), ('Payer Policies', [ALT_B['payer_policies']]), ('HAE Population', [ALT_B['hae_population']]), ('Company Facts', [ALT_B['phvs_company_facts']])]
for nm_, dss in ALT_TABS:
    ws_ = sheet(nm_, '7F7F7F'); title(ws_, nm_); r = 4
    for ds in dss: r = write_dataset(ws_, r, ds)
    widths(ws_, {'A': 2, 'B': 44, 'C': 18, 'D': 14, 'E': 16, 'F': 22, 'G': 22, 'H': 22, 'I': 22, 'J': 22, 'K': 22, 'L': 22})
    ws_.page_setup.fitToWidth = 1

wsDiv = sheet('SCRAP & DATA ------>', '0070C0'); put(wsDiv, 2, 2, 'Reference data', f_title); put(wsDiv, 3, 2, 'The primary sources and evidence registers behind the Inputs tab. Values only; no forecast cell depends on these tabs.', f_norm)
# ================================================================== SUMMARY
wsSu = sheet('Summary', '002060'); title(wsSu, 'Pharvaris (NASDAQ: PHVS)')
EXA = MCX['expected']['annual']; XPEAK_Y = max(EXA, key=lambda y: EXA[y]['revenue']); XPEAK = EXA[XPEAK_Y]['revenue']
r = 4; put(wsSu, r, 2, 'Recommendation', f_bold); put(wsSu, r, 4, '=IF(implied_return<0,"SHORT","LONG")', f_title); r += 1
put(wsSu, r, 2, 'Share price'); put(wsSu, r, 4, '=share_price', f_norm, USD); put(wsSu, r, 5, '=TEXT(price_date,"d mmm yyyy")', f_sub); r += 1
put(wsSu, r, 2, '12-month price target', f_bold); put(wsSu, r, 4, '=price_target', f_bold, USD, fill=fill_key); put(wsSu, r, 5, 'mean of 4,096 Monte Carlo draws of the patient model, financials and DCF', f_sub); r += 1
put(wsSu, r, 2, 'Implied return', f_bold); put(wsSu, r, 4, '=implied_return', f_bold, PCT); r += 1
put(wsSu, r, 2, 'Street median target'); put(wsSu, r, 4, '=street_median', f_norm, USD); r += 1
put(wsSu, r, 2, 'Model checks'); put(wsSu, r, 4, '=IF(AND(checks_controls="OK",checks_identities="OK"),"All checks pass",IF(checks_identities="OK","Identities pass; an override is active so control totals differ from the saved engine","CHECK"))', f_norm); r += 2
band(wsSu, r, 'How the target is built (DCF tab)', 2, 7); r += 1
for lab, ref_ in (('Present value of expected free cash flow, 2027-2055', f"='DCF Walk'!C{WS0}"), ('Expected residual after 2055', f"='DCF Walk'!C{WS0+1}"), ('Expected cash at the target date', f"='DCF Walk'!C{WS0+2}"), ('Less expected issuance fees', f"='DCF Walk'!C{WS0+3}"), ('Expected value if neither product is approved', f"='DCF Walk'!C{WS0+4}"), ('Expected equity value', f"='DCF Walk'!C{WEQ}"), ('Expected diluted shares, millions', f"='DCF Walk'!C{WEQ+1}"), ('Price target, mean of the per-draw values', '=price_target')):
    put(wsSu, r, 2, lab, f_bold if 'Expected equity' in lab or 'Price target' in lab else f_norm); put(wsSu, r, 4, ref_, f_bold if 'Expected equity' in lab or 'Price target' in lab else f_norm, USD if 'Price target' in lab else (NUM3 if 'shares' in lab else NUM)); r += 1
r += 1; wsSu.row_breaks.append(Break(id=r - 1)); band(wsSu, r, 'Expected financials: probability-weighted mean of the 4,096 draws (USD millions unless the row says otherwise)', 2, 11); r += 1
HY = [2027, 2028, 2029, 2030, 2031, 2033, 2035, 2038, 2040]
put(wsSu, r, 2, 'Fiscal year', f_hdr, fill=fill_hdr); [put(wsSu, r, 3 + i, f'{y}E', f_hdr, fill=fill_hdr, align=Alignment(horizontal='right')) for i, y in enumerate(HY)]; HYR = r; r += 1
for key, lab, fmt in (('IR_recognized', 'IR revenue', NUM), ('XR_recognized', 'XR revenue', NUM), ('revenue', 'Total revenue', NUM), ('__central', 'Central case revenue (Three Statements, live)', NUM), ('operating_income', 'Operating income', NUM), ('net_income', 'Net income', NUM), ('fcf', 'Free cash flow', NUM), ('funding_net', 'Equity funding required', NUM), ('cash_end', 'Cash, year end', NUM), ('shares_end', 'Shares outstanding, million shares', NUM2), ('xr_active', 'XR active patients, year end (count)', NUM0), ('ir_access_active', 'IR active members, year end (count)', NUM0)):
    put(wsSu, r, 2, lab, f_bold if key in ('revenue', 'fcf') else f_norm)
    for i, y in enumerate(HY):
        v = ('=' + aref('revenue', y)) if key == '__central' else EXA[str(y)][key]
        put(wsSu, r, 3 + i, v, f_bold if key in ('revenue', 'fcf') else (f_norm if key == '__central' else f_input), fmt)
    r += 1
put(wsSu, r, 2, 'Expected revenue carries the 15% chance that one or both products fail and the lopsided input ranges; the central case is the live engine at the most-likely inputs.', f_sub); r += 2
band(wsSu, r, 'The assumptions that carry the short', 2, 11); r += 1
for c_, h in ((2, 'Assumption'), (3, 'Central'), (4, 'Range tested'), (6, 'Value across the range'), (8, 'Why it is set here')): put(wsSu, r, c_, h, f_hdr, fill=fill_hdr)
for c_ in range(2, 12): wsSu.cell(r, c_).fill = fill_hdr; wsSu.cell(r, c_).font = f_hdr
wsSu.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5); wsSu.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7); wsSu.merge_cells(start_row=r, start_column=8, end_row=r, end_column=11)
r += 1; SENS = {x['parameter']: x for x in CR['sensitivities']}
for k in ['ir_quantity_multiplier', 'category_switch', 'switch_mature_multiplier', 'oral_launch_share', 'ir_retention']:
    put(wsSu, r, 2, LABELS[k], f_bold, fill=fill_key); put(wsSu, r, 3, f"={k}", f_norm, NUM3)
    if k in SENS: x = SENS[k]; put(wsSu, r, 4, f"{x['low']} to {x['high']}", f_norm); put(wsSu, r, 6, f"${x['target_low']:.2f} to ${x['target_high']:.2f}", f_norm)
    else: put(wsSu, r, 4, 'fixed in every case', f_sub); put(wsSu, r, 6, 'relaxing it adds about $%.2f' % (bridge[0][2] - bridge[1][2]), f_sub)
    wsSu.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5); wsSu.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)
    put(wsSu, r, 8, RATIONALE[k], f_sub); wsSu.cell(r, 8).alignment = Alignment(wrap_text=True, vertical='top')
    for c_ in range(2, 8): wsSu.cell(r, c_).alignment = Alignment(vertical='top', horizontal=wsSu.cell(r, c_).alignment.horizontal)
    wsSu.merge_cells(start_row=r, start_column=8, end_row=r, end_column=11); wsSu.row_dimensions[r].height = max(26, 13 * (len(RATIONALE[k]) // 62 + 1)); r += 1
r += 1; band(wsSu, r, 'Key Drivers (most likely inputs for Monte Carlo)', 2, 6); band(wsSu, r, 'KPIs: key outputs (expected path unless stated)', 7, 11); r += 1
KPI_L = [('US diagnosed HAE patients, July 2026', '=population', NUM0), ('Net population growth', '=pop_growth', PCT), ('Oral preference among preventive patients', '=prevention_oral_preference', PCT), ('Annual category switching opportunity', '=category_switch', PCT), ('PHVS share of oral choices, launch to mature', '=TEXT(oral_launch_share,"0%")&" to "&TEXT(oral_mature_share,"0%")', GEN), ('PHVS capture of departing oral users', '=oral_destination_capture', PCT), ('Paid access conversion', '=conversion', PCT), ('XR retention, first year and mature', '=TEXT(xr_first_retention,"0.0%")&" / "&TEXT(xr_mature_retention,"0%")', GEN), ('IR annual membership continuation', '=ir_retention', PCT), ('Redosing quantity multiplier (1.17/1.38)', '=ir_quantity_multiplier', NUM3), ('Rescue doses per year, prevention and no prevention', '=TEXT(ir_doses_ltp,"0.0")&" / "&TEXT(ir_doses_other,"0.0")', GEN), ('XR WAC per 28 days / gross-to-net discount', '=TEXT(xr_wac_28days,"$#,##0")&" / "&TEXT(1-xr_net_fraction,"0%")', GEN), ('IR WAC per dose / gross-to-net discount', '=TEXT(ir_wac_dose,"$#,##0")&" / "&TEXT(1-ir_net_fraction,"0%")', GEN), ('XR net price per treatment-year', '=xr_wac_28days*365/28*xr_net_fraction', USD0), ('IR net price per dose', '=ir_wac_dose*ir_net_fraction', USD0), ('Approval odds, IR and XR', '=TEXT(ir_pos,"0%")&" / "&TEXT(xr_pos,"0%")', GEN), ('First full generic year', '=loe_year', '0'), ('Brand revenue drop after generic entry, year 1 / year 4', '=TEXT(1-INDEX(xr_r_idx,MATCH(loe_year,loe_years,0)),"0%")&" / "&TEXT(1-INDEX(xr_r_idx,MATCH(loe_year+3,loe_years,0)),"0%")', GEN), ('Discount rate used / WACC build', '=TEXT(discount_rate,"0.0%")&" / "&TEXT(wacc,"0.0%")', GEN)]
cum = sum(EXA[str(y)]['revenue'] for y in range(2027, 2037)); first_pos = next((y for y in range(2027, END_YEAR + 1) if EXA[str(y)]['operating_income'] > 0), None)
KPI_R = [('Expected peak revenue, USD m', XPEAK, NUM), ('Expected peak year', int(XPEAK_Y), '0'), ('Expected revenue 2030, USD m', EXA['2030']['revenue'], NUM), ('Expected revenue 2035, USD m', EXA['2035']['revenue'], NUM), ('Expected cumulative revenue 2027-2036, USD m', cum, NUM), ('XR share of expected peak-year revenue', EXA[XPEAK_Y]['XR_recognized'] / XPEAK, PCT), ('Peak XR penetration of diagnosed patients (central, live)', f'=SUMPRODUCT(MAX(({arange("pop")}>0)*{arange("xr_active")}/({arange("pop")}+({arange("pop")}=0))))', PCT), ('Peak IR penetration of diagnosed patients (central, live)', f'=SUMPRODUCT(MAX(({arange("pop")}>0)*{arange("ir_active")}/({arange("pop")}+({arange("pop")}=0))))', PCT), ('Expected XR active patients, 2031 / 2038', f"{EXA['2031']['xr_active']:,.0f} / {EXA['2038']['xr_active']:,.0f}", GEN), ('Expected IR active members, 2031 / 2038', f"{EXA['2031']['ir_access_active']:,.0f} / {EXA['2038']['ir_access_active']:,.0f}", GEN), ('First year of positive expected operating income', first_pos, '0'), ('Expected equity funding after today, USD m', MCX['expected']['valuation']['future_funding_net_usdm'], NUM), ('Expected shares after modelled funding, m', EXA['2031']['shares_end'], NUM2), ('Central case peak revenue, USD m (live)', f'=MAX({arange("revenue")})', NUM), ('Price target and implied return', '=TEXT(price_target,"$0.00")&" / "&TEXT(implied_return,"0.0%")', GEN), ("Required scale on the Street ramp, today's price", '=rd_k_live', NUM3), ('Monte Carlo 12-month value: median / P10 / P90', f"=TEXT('Monte Carlo'!D6,\"$0.00\")&\" / \"&TEXT('Monte Carlo'!E6,\"$0.00\")&\" / \"&TEXT('Monte Carlo'!F6,\"$0.00\")", GEN), ('Share of draws below the share price', f"='Monte Carlo'!C9", PCT)]
for i in range(max(len(KPI_L), len(KPI_R))):
    if i < len(KPI_L): lab, f_, fm = KPI_L[i]; put(wsSu, r + i, 2, lab); put(wsSu, r + i, 5, f_, f_norm, fm, align=Alignment(horizontal='right')); wsSu.merge_cells(start_row=r + i, start_column=5, end_row=r + i, end_column=6)
    if i < len(KPI_R): lab, f_, fm = KPI_R[i]; put(wsSu, r + i, 7, lab); wsSu.merge_cells(start_row=r + i, start_column=7, end_row=r + i, end_column=9); put(wsSu, r + i, 10, f_, (f_norm if isinstance(f_, str) and f_.startswith('=') else f_input), fm, align=Alignment(horizontal='right')); wsSu.merge_cells(start_row=r + i, start_column=10, end_row=r + i, end_column=11)
r += max(len(KPI_L), len(KPI_R)) + 2
rc = BarChart(); rc.type = 'col'; rc.grouping = 'stacked'; rc.overlap = 100; rc.title = 'Expected revenue by product, USD m'; rc.height = 7.5; rc.width = 16; rc.style = 2
rc.add_data(Reference(wsSu, min_col=2, max_col=2 + len(HY), min_row=HYR + 1, max_row=HYR + 2), from_rows=True, titles_from_data=True); rc.set_categories(Reference(wsSu, min_col=3, max_col=2 + len(HY), min_row=HYR)); wsSu.add_chart(rc, f'B{r}'); wsSu.print_area = f'B1:K{r+18}'
widths(wsSu, {'A': 2, 'B': 44, 'C': 11, 'D': 11, 'E': 11, 'F': 11, 'G': 16, 'H': 16, 'I': 16, 'J': 12, 'K': 12})

# ================================================================== MODEL MAP
def draw_model_map(path):
    import matplotlib; matplotlib.use('Agg'); import matplotlib.pyplot as plt
    from matplotlib.patches import FancyBboxPatch
    plt.rcParams['font.family'] = 'Arial'
    NAVY, GREY, LINE = '#002060', '#5A6B8C', '#9AA7BF'
    layers = [
     ('Primary sources', '33 public documents', GREY, [
        'Pharvaris 20-F FY2025', 'Pharvaris H1 2026 interim', 'RAPIDe-3, CHAPTER-3 results\n(CTIS, FDA, publications)', 'Ekterly label; KalVista\nlaunch updates, 14D-9', 'BioCryst 10-K, 8-K, investor decks', 'HAEA 2025 LTP claims poster', 'Zuraw 2025 persistence study', 'French berotralstat cohort', 'HAE preference survey, n=150', 'Takeda FY2025 results', 'Health Affairs\ngeneric-entry appendix', 'BioMarin Kuvan filings', 'CMS Part D spending', '12 sell-side notes, Sep 2026', 'Payer policies:\nUHC, Cigna, BCBS', 'ECB FX; prevalence studies']),
     ('Evidence registers', '12 reference tabs', GREY, [
        'Orladeyo and Ekterly\nlaunch histories', 'Persistence, switching\nand dosing benchmarks', 'Clinical cross-trial comparison', 'Second-oral analogues', 'HAE injectable revenues', 'Generic erosion panel', 'Street targets and\nGuggenheim path', 'Payer policies and\nHAE population']),
     ('Inputs', '75 assumptions, 3 cases', NAVY, [
        'Population and net growth', 'Switching, preference,\noral choice ramp', 'Access, delay, retry', 'Retention and paid intensity', 'Rescue dose budgets\nand redosing', 'Net prices and geography', 'Cost schedule, tax, funding', 'Approval odds, discount\nrate, exclusivity']),
     ('Patient engine', '12 origin cells x 354 months', NAVY, [
        'Joint pools and\nEkterly launch stock', 'Prescription propensities\nby origin', 'XR ledger: attempts, access,\nactive, retired', 'IR ledger: attempts, access,\nactive, retired', 'Cohort survival\nand exposure']),
     ('Economics', '4 clinical branches', NAVY, [
        'XR paid treatment-years\nx net price', 'IR doses, with and\nwithout XR overlap', 'Ex-US lag and rollout', 'Post-generic price and volume', 'Costs, tax, working\ncapital, funding']),
     ('Statements and valuation', 'dated DCF, 9 Oct 2026 and 2027', NAVY, [
        'Three Statements\nand schedules', 'Annual and Quarterly', 'DCF per branch,\noption dilution', 'Reverse DCF:\nwhat the price requires', 'Monte Carlo: 4,096 draws', 'Sensitivities and\nStreet bridge']),
     ('Outputs', 'Summary and Cover', NAVY, [
        'Price target and implied return', 'KPIs and thesis assumptions', 'Equity bridge and\nfunding need', 'Checks: engine parity,\nidentities']),
    ]
    fig, ax = plt.subplots(figsize=(19, 8.8), dpi=120); ax.set_xlim(0, 19); ax.set_ylim(0, 8.8); ax.axis('off')
    W, H0, H1 = 2.35, 7.6, 0.55; xs = [0.25 + i * 2.68 for i in range(len(layers))]
    centers = []
    for li, (head, sub, col, nodes) in enumerate(layers):
        x = xs[li]; n = len(nodes); gap = (H0 - H1) / n; h = min(0.72, gap - 0.07); cs = []
        ax.text(x + W / 2, 8.45, head, ha='center', va='center', fontsize=11.5, fontweight='bold', color=NAVY)
        ax.text(x + W / 2, 8.12, sub, ha='center', va='center', fontsize=8.5, color='#404040')
        for k, lab in enumerate(nodes):
            yc = H0 - gap * (k + 0.5); cs.append(yc)
            ax.add_patch(FancyBboxPatch((x, yc - h / 2), W, h, boxstyle='round,pad=0.015,rounding_size=0.06', fc=col, ec=col))
            ax.text(x + W / 2, yc, lab, ha='center', va='center', fontsize=8, color='white', linespacing=1.15)
        centers.append(cs)
    for li in range(len(layers) - 1):
        x1 = xs[li] + W; x2 = xs[li + 1]
        for y1 in centers[li]:
            for y2 in centers[li + 1]:
                ax.plot([x1, x2], [y1, y2], color='#7F8FB0', lw=0.7, alpha=0.75, zorder=0)
    ax.text(0.25, 0.22, 'Every node in the navy layers is live in this workbook; grey layers are the research archive and reference tabs it draws on. Lines show that each layer feeds the whole of the next, as the monthly engine recomputes from every input.', fontsize=8.4, color='#404040', va='center')
    fig.savefig(path, bbox_inches='tight', facecolor='white'); plt.close(fig)
MAP_PNG = ROOT / 'models/excel/model_map.png'; draw_model_map(MAP_PNG)
from openpyxl.drawing.image import Image as XLImage


# ================================================================== COVER
wsCo = sheet('Cover', '002060'); 
put(wsCo, 3, 2, 'Pharvaris N.V. (NASDAQ: PHVS)', Font(name='Arial', size=20, bold=True, color=NAVY))
put(wsCo, 4, 2, 'Point72 Academy Case Competition, October 2026', Font(name='Arial', size=12, color='595959'))
r = 6
for lab, v, fmt in (('Team', TEAM, GEN), ('Ticker', 'NASDAQ: PHVS', GEN), ('Recommendation', '=IF(implied_return<0,"SHORT","LONG")', GEN), ('Share price', '=share_price', USD), ('Price date', '=price_date', 'd mmm yyyy'), ('12-month price target', '=price_target', USD), ('Implied return', '=implied_return', PCT), ('Valuation date', '=valuation_date', 'd mmm yyyy'), ('Time Horizon', '12 Months', GEN)):
    put(wsCo, r, 2, lab, f_bold); put(wsCo, r, 4, v, f_input if lab in ('Team', 'Ticker') else (f_sub if lab == 'Price target basis' else f_bold), fmt, align=Alignment(horizontal='right') if lab == 'Time Horizon' else None); r += 1
put(wsCo, 6, 10, 'READ ME !!!', Font(name='Arial', size=26, bold=True), align=Alignment(horizontal='left')); wsCo.row_dimensions[6].height = 33
put(wsCo, 7, 10, 'Model Map', Font(name='Arial', size=11, bold=True, color='FFFFFF'), fill=PatternFill('solid', fgColor='001F60'))
_img2 = XLImage(str(MAP_PNG)); _img2.width, _img2.height = 1616, 695; wsCo.add_image(_img2, 'I8')
put(wsCo, 12, 5, 'Mean 12-month value across Monte Carlo draws', f_sub)
r += 2; band(wsCo, r, 'Implied Headline KPIs', 2, 8); r += 1
for lab, f_, fm in (('Expected peak revenue and year', f'${XPEAK:,.0f}m in {XPEAK_Y}', GEN), ('Peak XR penetration, diagnosed US patients', f'=SUMPRODUCT(MAX(({arange("pop")}>0)*{arange("xr_active")}/({arange("pop")}+({arange("pop")}=0))))', PCT), ('Peak IR penetration, diagnosed US patients', f'=SUMPRODUCT(MAX(({arange("pop")}>0)*{arange("ir_active")}/({arange("pop")}+({arange("pop")}=0))))', PCT), ('XR WAC per 28 days / gross-to-net discount', '=TEXT(xr_wac_28days,"$#,##0")&" / "&TEXT(1-xr_net_fraction,"0%")', GEN), ('IR WAC per dose / gross-to-net discount', '=TEXT(ir_wac_dose,"$#,##0")&" / "&TEXT(1-ir_net_fraction,"0%")', GEN), ('Brand revenue drop post-LOE, year 1 / year 4', '=TEXT(1-INDEX(xr_r_idx,MATCH(loe_year,loe_years,0)),"0%")&" / "&TEXT(1-INDEX(xr_r_idx,MATCH(loe_year+3,loe_years,0)),"0%")&" from "&TEXT(loe_year,"0")', GEN), ('Oral preference among preventive patients', '=prevention_oral_preference', PCT), ('PHVS share of oral choices, launch to mature', '=TEXT(oral_launch_share,"0%")&" to "&TEXT(oral_mature_share,"0%")', GEN), ('Paid access conversion', '=conversion', PCT), ('XR retention, first year / mature', '=TEXT(xr_first_retention,"0.0%")&" / "&TEXT(xr_mature_retention,"0%")', GEN), ('IR annual membership continuation', '=ir_retention', PCT), ('Redosing quantity multiplier (1.17/1.38)', '=ir_quantity_multiplier', NUM3), ('Approval odds, IR and XR', '=TEXT(ir_pos,"0%")&" / "&TEXT(xr_pos,"0%")', GEN), ('Discount rate used; WACC build', '=TEXT(discount_rate,"0.0%")&"; "&TEXT(wacc,"0.0%")', GEN), ('Expected equity funding after today, USD m', MCX['expected']['valuation']['future_funding_net_usdm'], NUM)):
    put(wsCo, r, 2, lab, f_bold); put(wsCo, r, 4, f_, f_norm if isinstance(f_, str) and f_.startswith('=') else f_input, fm, align=Alignment(horizontal='right')); r += 1
r += 1; band(wsCo, r, 'Contents', 2, 8); r += 1
TOC = [('Summary', 'Recommendation, how the target is built, expected financials, thesis assumptions, KPIs'), ('DCF', 'Price target, branch probabilities, expected cash flow walk, cost of equity'), ('Three Statements', 'Central case: income statement, balance sheet, cash flow and schedules, quarterly through 2031 inside the annual columns'), ('Reverse DCF', 'Sell-side targets, Guggenheim path, the ramp the share price requires, value of the thesis assumptions'), ('Operating Build', 'Funnel from diagnosed patients to revenue with its drivers, then the supporting builds'), ('Monte Carlo', '4,096 draws: distribution, branches, priors, drivers, stresses and one-variable sensitivities'), ('INPUTS ------>', 'Every assumption with override, bounds, source and rationale; reported financials; generic-erosion indices'), ('Patient Engine', 'Monthly calendar and the XR and IR patient ledgers by origin cell'), ('Financial Engine', 'Monthly units, prices, revenue and financials for the four clinical branches'), ('SCRAP & DATA ------>', 'Sources register, then one tab per evidence dataset: launches, benchmarks, clinical, analogues, injectables, erosion, Street, payers, population, company facts'), ('Checks', 'Live workbook against the saved engine; accounting and population identities')]
for nm_, desc in TOC:
    c = put(wsCo, r, 2, nm_, f_bold); sh = nm_.split(' / ')[0]
    from openpyxl.worksheet.hyperlink import Hyperlink as _HL
    c.hyperlink = _HL(ref=c.coordinate, location=f"'{sh}'!A1"); c.font = Font(name='Arial', size=10, bold=True, color='0563C1', underline='single')
    put(wsCo, r, 4, desc); wsCo.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8); r += 1
r += 1; band(wsCo, r, 'How to read the workbook', 2, 8); r += 1
for lab, desc, font, fill in (('Blue text', 'Hardcoded input or saved result', f_input, None), ('Black text', 'Formula', f_norm, None), ('Green text', 'Link to another sheet', f_link, None), ('Yellow cell', 'Override column on Inputs', f_norm, fill_in), ('Orange label', 'Assumption that carries the thesis', f_bold, fill_key)):
    c = put(wsCo, r, 2, lab, font); c.fill = fill or PatternFill(); put(wsCo, r, 4, desc); r += 1
r += 1
widths(wsCo, {'A': 2, 'B': 44, 'C': 2, 'D': 28, 'E': 16, 'F': 16, 'G': 16, 'H': 16}); wsCo.print_area = f'B1:AF{r}'; wsCo.page_setup.fitToHeight = 1

# ================================================================== order, print setup, save
def _chart_bottom(ws):
    from openpyxl.utils.cell import coordinate_from_string
    b = 0
    for ch in ws._charts:
        a = ch.anchor
        try: row = int(coordinate_from_string(a)[1]) if isinstance(a, str) else a._from.row + 1
        except Exception: continue
        b = max(b, row + int((ch.height or 7.5) / 0.5) + 2)
    return b
GROUPS = [('Cover', ['Cover']), ('Summary', ['Summary']), ('DCF', ['DCF Walk']), ('Three Statements', ['Three Statements']), ('Reverse DCF', ['Street Bridge', 'Reverse DCF']),
          ('Operating Build', ['Funnel', 'Adoption Build', 'Quantity & Price', 'Annual']), ('Monte Carlo', ['Monte Carlo']), ('INPUTS ------>', ['Inputs', 'History', 'LOE']),
          ('Patient Engine', ['Calendar', 'XR Engine', 'IR Engine']), ('Financial Engine', ['Revenue', 'Fin Monthly']), ('SCRAP & DATA ------>', ['SCRAP & DATA ------>']), ('Sources', ['Sources'])] + [(t, [t]) for t, _ in ALT_TABS] + [('Checks', ['Checks', 'Valuation'])]
GROUP_WIDTHS = {'Three Statements': {'B': 56, 'C': 15}, 'Reverse DCF': {'B': 64, 'C': 13}, 'Operating Build': {'B': 52, 'C': 14}, 'INPUTS ------>': None, 'Patient Engine': None, 'Financial Engine': None, 'Checks': {'B': 56, 'C': 14, 'D': 14, 'E': 12, 'F': 10, 'G': 9}}
import re as _re
from copy import copy as _copy
from openpyxl.worksheet.pagebreak import Break as _Break
OFFS = {}
def _chart_extra(ws):
    extra = 0
    for ch in ws._charts:
        anc = ch.anchor if isinstance(ch.anchor, str) else 'B2'; row_ = int(_re.match(r'[A-Z]+(\d+)', anc).group(1))
        extra = max(extra, row_ + int((ch.height or 7.5) * 2.1) + 1 - ws.max_row)
    return max(0, extra)
for gname, parts in GROUPS:
    off = 0
    for pn in parts: OFFS[pn] = (gname, off); off += wb[pn].max_row + 2 + _chart_extra(wb[pn])
_ANY = _re.compile(r"(?:(?:'([^']+)'|(?<![A-Za-z0-9_])([A-Za-z_][A-Za-z0-9_]*))!)?(?<![A-Za-z_'$\d.])(\$?[A-Z]{1,3})(\$?)(\d+)(?::(\$?[A-Z]{1,3})(\$?)(\d+))?(?![\d(A-Za-z_])")
def _rw(text, own_sheet):
    """Rewrite every cell reference in a formula or defined name for the consolidated layout, each exactly once."""
    own_off = OFFS.get(own_sheet, (own_sheet, 0))[1]
    parts = text.split('"'); out = []
    for i, seg in enumerate(parts):
        if i % 2 == 1: out.append(seg); continue
        def f_ref(m):
            sh = m.group(1) or m.group(2)
            if sh:
                g, off = OFFS.get(sh, (sh, 0)); prefix = (f"'{g}'" if ' ' in g else g) + '!'
            else:
                off = own_off; prefix = ''
            a_ = f"{m.group(3)}{m.group(4)}{int(m.group(5)) + off}"
            b_ = f":{m.group(6)}{m.group(7)}{int(m.group(8)) + off}" if m.group(6) else ''
            return prefix + a_ + b_
        out.append(_ANY.sub(f_ref, seg))
    return '"'.join(out)
for gname, parts in GROUPS:
    tmp = wb.create_sheet(gname + '__new'); first = wb[parts[0]]
    tmp.sheet_properties.tabColor = first.sheet_properties.tabColor
    tmp.sheet_view.showGridLines = False; tmp.sheet_properties.tabColor = first.sheet_properties.tabColor
    tmp.page_setup.orientation = 'landscape'; tmp.page_setup.fitToWidth = 1; tmp.page_setup.fitToHeight = 0; tmp.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    tmp.oddFooter.left.text = TEAM; tmp.oddFooter.center.text = gname; tmp.oddFooter.right.text = 'Page &P of &N'; tmp.oddHeader.left.text = 'Pharvaris (PHVS)'; tmp.oddHeader.right.text = '&D'
    maxcol = 1
    for pn in parts:
        ws = wb[pn]; off = OFFS[pn][1]
        for row in ws.iter_rows():
            for c in row:
                if c.value is None and not c.has_style: continue
                n = tmp.cell(c.row + off, c.column)
                v = c.value
                if isinstance(v, str) and v.startswith('='): v = _rw(v, pn)
                n.value = v
                if c.has_style: n.font = _copy(c.font); n.fill = _copy(c.fill); n.number_format = c.number_format; n.alignment = _copy(c.alignment); n.border = _copy(c.border)
                if c.hyperlink:
                    from openpyxl.worksheet.hyperlink import Hyperlink as _HL2
                    n.hyperlink = _HL2(ref=n.coordinate, location=c.hyperlink.location) if c.hyperlink.location else c.hyperlink.target
                maxcol = max(maxcol, c.column)
        for rr, d in ws.row_dimensions.items():
            if d.height: tmp.row_dimensions[rr + off].height = d.height
        for mr in ws.merged_cells.ranges: tmp.merge_cells(start_row=mr.min_row + off, start_column=mr.min_col, end_row=mr.max_row + off, end_column=mr.max_col)
        for ch in ws._charts:
            for ser in ch.series:
                if ser.val is not None and ser.val.numRef is not None: ser.val.numRef.f = _rw(ser.val.numRef.f, pn)
                if ser.cat is not None:
                    if ser.cat.numRef is not None: ser.cat.numRef.f = _rw(ser.cat.numRef.f, pn)
                    if ser.cat.strRef is not None: ser.cat.strRef.f = _rw(ser.cat.strRef.f, pn)
                if ser.tx is not None and ser.tx.strRef is not None: ser.tx.strRef.f = _rw(ser.tx.strRef.f, pn)
            anc = ch.anchor if isinstance(ch.anchor, str) else 'B2'
            m = _re.match(r'([A-Z]+)(\d+)', anc); tmp.add_chart(ch, f"{m.group(1)}{int(m.group(2)) + off}")
        for im in ws._images:
            anc = im.anchor if isinstance(im.anchor, str) else 'B2'; m = _re.match(r'([A-Z]+)(\d+)', anc); im.anchor = f"{m.group(1)}{int(m.group(2)) + off}"; tmp.add_image(im)
        for br in ws.row_breaks.brk: tmp.row_breaks.append(_Break(id=br.id + off))
        if off: tmp.row_breaks.append(_Break(id=off))
    # widths: explicit group widths, else the first part's, else the max across parts
    gw = GROUP_WIDTHS.get(gname, {})
    for col in range(1, maxcol + 1):
        L = CL(col)
        if gw and L in gw: tmp.column_dimensions[L].width = gw[L]; continue
        ws_w = [wb[pn].column_dimensions[L].width for pn in parts if L in wb[pn].column_dimensions and wb[pn].column_dimensions[L].width]
        if gw is not None and gw and ws_w: tmp.column_dimensions[L].width = (10 if col > 3 else max(ws_w))
        elif ws_w: tmp.column_dimensions[L].width = (first.column_dimensions[L].width if L in first.column_dimensions and first.column_dimensions[L].width else max(ws_w))
        if L in first.column_dimensions and first.column_dimensions[L].outlineLevel: tmp.column_dimensions[L].outlineLevel = first.column_dimensions[L].outlineLevel
    tmp.freeze_panes = first.freeze_panes
    if first.print_title_rows: tmp.print_title_rows = first.print_title_rows
    if gname in ('Patient Engine', 'Financial Engine'): tmp.page_setup.fitToWidth = 1; tmp.print_area = f'A1:{mcol(23)}{tmp.max_row}'
    else: tmp.print_area = f"A1:{CL(maxcol)}{max(tmp.max_row, _chart_bottom(tmp))}"
    for pn in parts: wb.remove(wb[pn])
    tmp.title = gname
wb['Cover'].print_area = f"A1:AF{wb['Cover'].max_row}"; wb['Cover'].page_setup.fitToHeight = 1
for nm_ in list(wb.defined_names):
    dn = wb.defined_names[nm_]; dn.attr_text = _rw(dn.attr_text, '')
wb._sheets = [wb[g] for g, _ in GROUPS]
wb.active = 0
wb.save(OUT); print('saved', OUT)
_maps = {'RX': {str(k): v for k, v in RX.items()}, 'RI': {str(k): v for k, v in RI.items()}, 'RV': {str(k): v for k, v in RV.items()}, 'RF': {str(k): v for k, v in RF.items()}, 'CAL': CAL, 'VALROW': {str(k): v for k, v in VALROW.items()}, 'MC0': MC0, 'NM': NM, 'AN': AN, 'TS': TS, 'OFFS': OFFS, 'INPUT_ROW': INPUT_ROW}
(ROOT / 'models/excel/rowmap.json').write_text(json.dumps(_maps))
