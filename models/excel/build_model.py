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
    if sub: ws.cell(2, 2, sub).font = f_sub

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
    for k in range(PAD): ws.column_dimensions[CL(C0 + k)].width = 6
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
put(wsI, 4, 2, 'Case selector', f_bold); put(wsI, 4, 4, 'Working', f_input, fill=fill_in); name('case_selector', "Inputs!$D$4")
put(wsI, 4, 6, 'Type Working, Cautious or Upside. Overrides in column H beat the case.', f_sub)
put(wsI, 5, 2, 'Case index', f_bold); put(wsI, 5, 4, '=MATCH(case_selector,$E$7:$G$7,0)', f_norm, '0'); name('case_index', "Inputs!$D$5")
put(wsI, 6, 2, 'Price target basis', f_bold); put(wsI, 6, 4, 'Monte Carlo', f_input, fill=fill_in); name('pt_basis', "Inputs!$D$6"); put(wsI, 6, 6, 'Monte Carlo = expected value across the 4,096 saved draws; Working = the deterministic point case.', f_sub)
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
THESIS = {'ir_quantity_multiplier', 'category_switch', 'switch_mature_multiplier', 'ir_retention'}
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
 'ir_pos': 'IR approval probability', 'xr_pos': 'XR approval probability', 'shared_success': 'Shared molecule gate', 'discount_rate': 'Discount rate', 'loe_year': 'First full generic year', 'loe_price_cap': 'Cap post-generic price index at 1'}
INPUT_ROW = {}
r = 8
for gname, keys in groups:
    band(wsI, r, gname, 2, 14); r += 1
    for k in keys:
        v = P[k]; val = v['value']; cases = {c: cfg['cases'][c].get(k, val) for c in ('working', 'cautious', 'upside_challenge')}
        put(wsI, r, 2, LABELS.get(k, k), f_bold if k in THESIS else f_norm, fill=fill_key if k in THESIS else None)
        put(wsI, r, 3, k, f_sub); put(wsI, r, 4, v.get('units', ''), f_sub)
        fmt = PCT if ('fraction' in v.get('units', '') or 'probability' in v.get('units', '') or 'survival' in v.get('units', '') or k in ('growth', 'cost_inflation', 'discount_rate', 'tax_rate', 'dual_use', 'oral_launch_share', 'oral_mature_share')) and not k.endswith('multiplier') else (NUM0 if isinstance(val, int) and not isinstance(val, bool) else NUM3)
        if k in ('xr_wac_28days', 'ir_wac_dose', 'issue_price'): fmt = USD
        if k in ('fx',): fmt = NUM3
        for i, c in enumerate(('working', 'cautious', 'upside_challenge')):
            cv = cases[c]; put(wsI, r, 5 + i, (int(cv) if isinstance(cv, bool) else cv), f_input, fmt)
        put(wsI, r, 8, f'=IF($I{r}<>"",$I{r},INDEX($E{r}:$G{r},1,case_index))', f_bold, fmt)
        wsI.cell(r, 9).fill = fill_in; wsI.cell(r, 9).font = f_input; wsI.cell(r, 9).number_format = fmt
        b = v.get('bounds') or ['', '']
        put(wsI, r, 10, b[0], f_sub, fmt); put(wsI, r, 11, b[1], f_sub, fmt)
        put(wsI, r, 12, v.get('status', ''), f_sub); put(wsI, r, 13, ', '.join(v.get('sources', [])), f_sub)
        put(wsI, r, 14, fix_spacing(v.get('translation', '')), f_sub)
        name('pop_growth' if k == 'growth' else k, f"Inputs!$H${r}"); INPUT_ROW[k] = r; r += 1
    r += 1
# structural inputs
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
widths(wsI, {'A': 2, 'B': 44, 'C': 24, 'D': 16, 'E': 11, 'F': 11, 'G': 11, 'H': 11, 'I': 10, 'J': 9, 'K': 9, 'L': 24, 'M': 18, 'N': 80})
for rr_ in range(8, wsI.max_row + 1):
    if wsI.cell(rr_, 14).value: wsI.cell(rr_, 14).alignment = Alignment(wrap_text=True, vertical='top')
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
wsA = sheet('Adoption Build', '0070C0'); title(wsA, 'Patient adoption build', 'Where patients start, how fast they try PHVS, and how many stay. Rates feed the monthly XR and IR engines.')
ORIG = ['orladeyo', 'other_prevention', 'acute_only', 'diagnosed_untreated']; RESC = ['ekterly', 'injectable', 'none']
ON = {'orladeyo': 'Orladeyo', 'other_prevention': 'Other prevention', 'acute_only': 'Rescue only', 'diagnosed_untreated': 'Untreated'}; RN = {'ekterly': 'Ekterly', 'injectable': 'Injectable rescue', 'none': 'No rescue'}
r = 4; band(wsA, r, 'Ekterly active oral rescue stock at PHVS launch (projected from start forms)', 2, 12); r += 1
put(wsA, r, 2, 'Months from July 2025 to IR launch'); put(wsA, r, 5, "=YEAR(launch_IR)*12+MONTH(launch_IR)-1+launch_delay_IR-(2025*12+6)", f_norm, NUM0); name('ek_nmonths', f"'Adoption Build'!$E${r}"); r += 1
put(wsA, r, 2, 'Month', f_bold); put(wsA, r, 3, 'Index', f_bold); put(wsA, r, 4, 'Forms', f_bold); put(wsA, r, 5, 'Counted', f_bold); put(wsA, r, 6, 'Survival weight', f_bold); put(wsA, r, 7, 'Active contribution', f_bold); r += 1
EK0 = r
for i in range(48):
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
widths(wsA, {'A': 2, 'B': 58, 'C': 12, 'D': 12, 'E': 14, 'F': 14, 'G': 16, 'H': 70, 'I': 12})

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
    vrow(f'{pre}_price', f'{prod} net price, USD', lambda j: f"={price}*INDEX({prod.lower()}_p_idx,MATCH({mcol(j)}$3,loe_years,0))", fmt=USD0)
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
wsVal = sheet('Valuation', '0070C0'); title(wsVal, 'Valuation', 'Dated DCF at the valuation and target dates for each clinical branch, weighted once by branch probability. USD millions except per share.')
r = 4; band(wsVal, r, 'Branch probabilities', 2, 8); r += 1
put(wsVal, r, 2, 'Both approved'); put(wsVal, r, 4, '=ir_pos*xr_pos/shared_success', f_norm, PCT); name('w_both', f'Valuation!$D${r}'); put(wsVal, r, 6, 'P(IR) x P(XR) / shared gate', f_sub); r += 1
put(wsVal, r, 2, 'IR only'); put(wsVal, r, 4, '=ir_pos-w_both', f_norm, PCT); name('w_IR_only', f'Valuation!$D${r}'); r += 1
put(wsVal, r, 2, 'XR only'); put(wsVal, r, 4, '=xr_pos-w_both', f_norm, PCT); name('w_XR_only', f'Valuation!$D${r}'); r += 1
put(wsVal, r, 2, 'Neither'); put(wsVal, r, 4, '=1-ir_pos-xr_pos+w_both', f_norm, PCT); name('w_neither', f'Valuation!$D${r}'); r += 1
put(wsVal, r, 2, 'Check: sum'); put(wsVal, r, 4, '=w_both+w_IR_only+w_XR_only+w_neither', f_norm, PCT); r += 2
VALROW = {}
for tag, dlabel in (('val', 'Valuation date'), ('tgt', 'Target date')):
    band(wsVal, r, f'{dlabel}: =TEXT({"valuation_date" if tag=="val" else "target_date"},"d mmm yyyy")', 2, 8); wsVal.cell(r, 2).value = f'="{dlabel}: "&TEXT({"valuation_date" if tag=="val" else "target_date"},"d mmm yyyy")'; r += 1
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
put(wsVal, r, 2, 'Price target', f_bold); put(wsVal, r, 3, '=IF(pt_basis="Monte Carlo",mc_mean_target,value_target)', f_bold, USD, fill=fill_key); put(wsVal, r, 5, '="basis: "&pt_basis&" (Inputs tab)"', f_sub); name('price_target', f'Valuation!$C${r}'); r += 1
put(wsVal, r, 2, 'Implied 12-month return', f_bold); put(wsVal, r, 3, '=price_target/share_price-1', f_bold, PCT); name('implied_return', f'Valuation!$C${r}'); r += 1
put(wsVal, r, 2, 'Return on the working point'); put(wsVal, r, 3, '=value_target/share_price-1', f_norm, PCT); r += 1
put(wsVal, r, 2, 'Return if both products are approved'); put(wsVal, r, 3, '=price_both_target/share_price-1', f_norm, PCT); r += 1
put(wsVal, r, 2, 'Shares outstanding at the price date, millions'); put(wsVal, r, 3, '=open_shares_m', f_norm, NUM3); r += 1
put(wsVal, r, 2, 'Market capitalisation'); put(wsVal, r, 3, '=share_price*open_shares_m', f_norm, NUM); r += 1
put(wsVal, r, 2, 'Future funding is a modelled requirement, not committed capital. Post-date funding is assumed at fair value, so negative cash flows carry the cost once; only the issuance fee is charged.', f_sub)
widths(wsVal, {'A': 2, 'B': 48, 'C': 14, 'D': 14, 'E': 14, 'F': 14, 'G': 4, 'H': 30})

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
wsAn = sheet('Annual', '0070C0'); title(wsAn, 'Annual financials and operating drivers', 'EUR millions as reported for 2024-2025; USD millions from 2026. 2026 combines the H1 actual (translated at 1.1186) with the H2 forecast. Forecast years assume both products are approved.')
YEARS = [2024, 2025] + list(range(2026, END_YEAR + 1)); YC0 = 4
put(wsAn, 5, 2, 'Fiscal year', f_bold); put(wsAn, 5, 3, 'Units', f_bold)
for i, y in enumerate(YEARS):
    c = YC0 + i; lab = f'{y}A' if y < 2026 else ('2026A/E' if y == 2026 else f'{y}E')
    put(wsAn, 4, c, y, f_norm, '0'); put(wsAn, 5, c, lab, f_hdr, fill=fill_hdr, align=Alignment(horizontal='right'))
    wsAn.column_dimensions[CL(c)].width = 10
    if y >= 2032: wsAn.column_dimensions[CL(c)].outlineLevel = 1; wsAn.column_dimensions[CL(c)].hidden = True
put(wsAn, 6, 2, 'Currency', f_sub); [put(wsAn, 6, YC0 + i, 'EUR' if y < 2026 else 'USD', f_sub, align=Alignment(horizontal='right')) for i, y in enumerate(YEARS)]
AN = {}; r = 8
def arow(key, label, fin=None, hist=None, kind='flow', sheet_name='Fin Monthly', fmt=NUM, bold=False, calc=None, unit='', h1=True, indent=0):
    """fin: monthly row to aggregate; hist: History key for 2024/25 (and H1 2026 when h1); calc: fn(col letter)->formula using AN rows."""
    global r; AN[key] = r; put(wsAn, r, 2, ('   ' * indent) + label, f_bold if bold else f_norm); put(wsAn, r, 3, unit, f_sub)
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
r += 1; band(wsAn, r, 'Income statement', 2, YC0 + len(YEARS) - 1); r += 1
arow('ir_rev_us', 'IR US revenue', fin=RV['IR_A_US'], sheet_name='Revenue', indent=1)
arow('ir_rev_ex', 'IR ex-US revenue', calc=lambda L, y: f"={sum_year('Revenue', RV['IR_A_direct'], L+'$4')}+{sum_year('Revenue', RV['IR_A_roy'], L+'$4')}" if y >= 2026 else None, indent=1)
arow('ir_rev', 'IR revenue', fin=RV['IR_A_rec'], sheet_name='Revenue', bold=True)
arow('xr_rev_us', 'XR US revenue', fin=RV['XR_US'], sheet_name='Revenue', indent=1)
arow('xr_rev_ex', 'XR ex-US revenue', calc=lambda L, y: f"={sum_year('Revenue', RV['XR_direct'], L+'$4')}+{sum_year('Revenue', RV['XR_roy'], L+'$4')}" if y >= 2026 else None, indent=1)
arow('xr_rev', 'XR revenue', fin=RV['XR_rec'], sheet_name='Revenue', bold=True)
arow('revenue', 'Total revenue', fin=FINB('revenue'), hist='revenue', bold=True)
arow('growth', 'Growth', calc=lambda L, y: f"=IF(OR({CL(YC0+YEARS.index(y)-1)}{AN['revenue']}=0,{CL(YC0+YEARS.index(y)-1)}{AN['revenue']}=\"\"),\"\",{L}{AN['revenue']}/{CL(YC0+YEARS.index(y)-1)}{AN['revenue']}-1)" if y >= 2028 else None, fmt=PCT, indent=1)
arow('manuf', 'Manufacturing and distribution', fin=FINB('manufacture'))
arow('license', 'BRAIN royalty', fin=FINB('license'))
arow('gross', 'Gross profit', calc=lambda L, y: f"={L}{AN['revenue']}-{L}{AN['manuf']}-{L}{AN['license']}" if y >= 2026 else None, bold=True)
arow('rd', 'Research and development', fin=FINB('rd'), hist='rd')
arow('comm', 'Commercial', fin=FINB('commercial'))
arow('ga', 'General and administrative', fin=FINB('ga'), hist='ga')
arow('dep', 'Depreciation', fin=FINB('dep'), hist='depreciation', fmt=NUM2, h1=False)
arow('ebit', 'Operating income', calc=lambda L, y: f"={L}{AN['gross']}-{L}{AN['rd']}-{L}{AN['comm']}-{L}{AN['ga']}-{L}{AN['dep']}" if y >= 2026 else f"=History!{CL(3+(y-2024))}{HIST_ROW['revenue']}-History!{CL(3+(y-2024))}{HIST_ROW['opex']}", bold=True)
arow('margin', 'Operating margin', calc=lambda L, y: f"=IF({L}{AN['revenue']}>0,{L}{AN['ebit']}/{L}{AN['revenue']},\"\")", fmt=PCT, indent=1)
arow('finance', 'Finance income (reported, not forecast)', hist='finance_income')
arow('tax', 'Tax', fin=FINB('tax'), hist='tax')
arow('ni', 'Net income', fin=FINB('ni'), hist='net_income', bold=True)
r += 1; band(wsAn, r, 'Cash flow and funding', 2, YC0 + len(YEARS) - 1); r += 1
arow('cf_ni', 'Net income', fin=FINB('ni'), hist='net_income')
arow('cf_dep', 'Depreciation', fin=FINB('dep'), hist='depreciation', fmt=NUM2)
arow('cf_capex', 'Capital expenditure', fin=FINB('capex'), hist='capex', fmt=NUM2)
arow('cf_nwc', 'Increase in net working capital', fin=FINB('dnwc'))
arow('fcf', 'Free cash flow', calc=lambda L, y: f"={L}{AN['cf_ni']}+{L}{AN['cf_dep']}-{L}{AN['cf_capex']}-{L}{AN['cf_nwc']}" if y >= 2026 else f"=History!{CL(3+(y-2024))}{HIST_ROW['cfo']}+History!{CL(3+(y-2024))}{HIST_ROW['cfi']}", bold=True)
arow('fund', 'Net equity funding required', fin=FINB('fund_net'), hist='cff')
arow('new_sh', 'Shares issued, millions', fin=FINB('new_sh'), fmt=NUM2)
arow('cash', 'Cash, year end', fin=FINB('cash'), hist='cash_end', kind='stock', bold=True)
arow('shares', 'Shares outstanding, millions', fin=FINB('shares'), hist='shares_end', kind='stock', fmt=NUM2)
r += 1; band(wsAn, r, 'Economic balance sheet', 2, YC0 + len(YEARS) - 1); r += 1
arow('bs_cash', 'Cash', fin=FINB('cash'), hist='cash_end', kind='stock')
arow('bs_nwc', 'Net working capital', fin=FINB('nwc'), kind='stock')
arow('bs_ppe', 'PP&E', fin=FINB('ppe'), kind='stock', fmt=NUM2)
arow('bs_other', 'Other net assets', fin=FINB('other'), kind='stock', fmt=NUM2)
arow('bs_equity', 'Equity', fin=FINB('equity'), hist='equity', kind='stock', bold=True)
arow('bs_check', 'Check: assets less equity', calc=lambda L, y: f"={L}{AN['bs_cash']}+{L}{AN['bs_nwc']}+{L}{AN['bs_ppe']}+{L}{AN['bs_other']}-{L}{AN['bs_equity']}" if y >= 2026 else None, fmt=NUM3)
arow('nol', 'NOL pool, year end', fin=FINB('nol'), kind='stock')
r += 1
put(wsAn, r, 2, 'Reported 2024-2025 are IFRS EUR figures and include finance income and non-cash share-based compensation; forecasts are economic cash-compensation equivalents with no interest or FX. 2026 net income adds the reported H1 loss (EUR x 1.1186) to the H2 forecast.', f_sub); r += 1
put(wsAn, r, 2, 'Years after 2031 are grouped; expand the outline to see the explicit forecast through 2055 and the generic erosion from 2039.', f_sub)
widths(wsAn, {'A': 2, 'B': 40, 'C': 8}); wsAn.freeze_panes = 'D6'
name('rev_2030', f"Annual!${CL(YC0+YEARS.index(2030))}${AN['revenue']}")


# ================================================================== THREE STATEMENTS
wsT = sheet('Three Statements', '0070C0'); title(wsT, 'Income statement, balance sheet and cash flow', 'Reported EUR millions for 2024-2025; USD millions from 2026 (H1 actual at 1.1186 plus H2 forecast). Forecasts assume both products are approved and are cash-compensation equivalents without interest or FX.')
put(wsT, 5, 2, 'Fiscal year', f_bold); put(wsT, 5, 3, 'Units', f_bold)
for i, y in enumerate(YEARS):
    c = YC0 + i; put(wsT, 4, c, y, f_norm, '0'); put(wsT, 5, c, f'{y}A' if y < 2026 else ('2026A/E' if y == 2026 else f'{y}E'), f_hdr, fill=fill_hdr, align=Alignment(horizontal='right')); wsT.column_dimensions[CL(c)].width = 10
    if y >= 2032: wsT.column_dimensions[CL(c)].outlineLevel = 1; wsT.column_dimensions[CL(c)].hidden = True
put(wsT, 6, 2, 'Currency', f_sub); [put(wsT, 6, YC0 + i, 'EUR' if y < 2026 else 'USD', f_sub, align=Alignment(horizontal='right')) for i, y in enumerate(YEARS)]
TS = {}; r = 8
def trow(key, label, fin=None, hist=None, kind='flow', fmt=NUM, bold=False, calc=None, h1=True, indent=0, histcalc=None):
    global r; TS[key] = r; put(wsT, r, 2, ('   ' * indent) + label, f_bold if bold else f_norm)
    for i, y in enumerate(YEARS):
        c = YC0 + i; L = CL(c); yc = f'{L}$4'; H = CL(3 + (y - 2024)) if y < 2026 else None
        if y < 2026:
            v = histcalc(H, L) if histcalc else (f"=History!{H}{HIST_ROW[hist]}" if hist else None)
        elif calc is not None: v = calc(L, y)
        elif fin is not None:
            v = '=' + (sum_year('Fin Monthly', fin, yc) if kind == 'flow' else stock_year('Fin Monthly', fin, yc)) + (f"+History!E{HIST_ROW[hist]}*fx" if (y == 2026 and hist and h1 and kind == 'flow') else '')
        else: v = None
        if v is not None: put(wsT, r, c, v, f_bold if bold else f_norm, fmt)
    r += 1
LASTC = CL(YC0 + len(YEARS) - 1)
band(wsT, r, 'Income statement', 2, YC0 + len(YEARS) - 1); r += 1
trow('rev_ir', 'IR product revenue', fin=FINB('rev_IR'), indent=1); trow('rev_xr', 'XR product revenue', fin=FINB('rev_XR'), indent=1)
trow('revenue', 'Revenue', fin=FINB('revenue'), hist='revenue', bold=True)
trow('cogs', 'Cost of goods sold (manufacturing and distribution)', fin=FINB('manufacture'))
trow('royalty', 'Royalty expense (BRAIN licence)', fin=FINB('license'))
trow('gp', 'Gross profit', calc=lambda L, y: f"={L}{TS['revenue']}-{L}{TS['cogs']}-{L}{TS['royalty']}", histcalc=lambda H, L: f"=History!{H}{HIST_ROW['revenue']}", bold=True)
trow('gm', 'Gross margin', calc=lambda L, y: f"=IF({L}{TS['revenue']}>0,{L}{TS['gp']}/{L}{TS['revenue']},\"\")", fmt=PCT, indent=1)
trow('rd', 'Research and development', fin=FINB('rd'), hist='rd')
trow('sm', 'Selling and marketing', fin=FINB('commercial'))
trow('ga', 'General and administrative', fin=FINB('ga'), hist='ga')
trow('da', 'Depreciation and amortisation', fin=FINB('dep'), hist='depreciation', fmt=NUM2, h1=False)
trow('ebit', 'Operating income', calc=lambda L, y: f"={L}{TS['gp']}-{L}{TS['rd']}-{L}{TS['sm']}-{L}{TS['ga']}-{L}{TS['da']}", histcalc=lambda H, L: f"=History!{H}{HIST_ROW['revenue']}-History!{H}{HIST_ROW['opex']}", bold=True)
trow('om', 'Operating margin', calc=lambda L, y: f"=IF({L}{TS['revenue']}>0,{L}{TS['ebit']}/{L}{TS['revenue']},\"\")", fmt=PCT, indent=1)
trow('fin', 'Finance income (reported; not forecast)', hist='finance_income', calc=lambda L, y: f"=History!E{HIST_ROW['finance_income']}*fx" if y == 2026 else '=0')
trow('ebt', 'Pre-tax income', calc=lambda L, y: f"={L}{TS['ebit']}+{L}{TS['fin']}", histcalc=lambda H, L: f"={L}{TS['ebit']}+{L}{TS['fin']}", bold=True)
trow('tax', 'Income tax', fin=FINB('tax'), hist='tax')
trow('ni', 'Net income', calc=lambda L, y: f"={L}{TS['ebt']}-{L}{TS['tax']}", histcalc=lambda H, L: f"=History!{H}{HIST_ROW['net_income']}", bold=True)
trow('sh', 'Shares outstanding, year end, millions', fin=FINB('shares'), hist='shares_end', kind='stock', fmt=NUM2)
trow('eps', 'Earnings per share', calc=lambda L, y: f"={L}{TS['ni']}/{L}{TS['sh']}", histcalc=lambda H, L: f"={L}{TS['ni']}/{L}{TS['sh']}", fmt=USD)
r += 1; band(wsT, r, 'Balance sheet (condensed economic balance sheet)', 2, YC0 + len(YEARS) - 1); r += 1
trow('cash', 'Cash and equivalents', fin=FINB('cash'), hist='cash_end', kind='stock')
trow('nwc', 'Operating working capital, net', fin=FINB('nwc'), kind='stock')
trow('ppe', 'Property, equipment and right-of-use assets', fin=FINB('ppe'), kind='stock', fmt=NUM2)
trow('other', 'Other net assets', fin=FINB('other'), kind='stock', fmt=NUM2)
trow('ta', 'Total net assets', calc=lambda L, y: f"={L}{TS['cash']}+{L}{TS['nwc']}+{L}{TS['ppe']}+{L}{TS['other']}", histcalc=lambda H, L: f"=History!{H}{HIST_ROW['total_assets']}-History!{H}{HIST_ROW['liabilities']}", bold=True)
trow('debt', 'Debt', calc=lambda L, y: '=0', histcalc=lambda H, L: '=0')
trow('eq_open', 'Equity, opening', calc=lambda L, y: (f"=open_equity*fx" if y == 2026 else f"={CL(YC0+YEARS.index(y)-1)}{r+3}"), indent=1)
trow('eq_ni', 'Net income (from 30 June 2026 in the 2026 column)', calc=lambda L, y: f"={L}{TS['ni']}" + ("-History!E%d*fx" % HIST_ROW['net_income'] if y == 2026 else ''), indent=1)
trow('eq_issue', 'Equity issued, net of fees', fin=FINB('fund_net'), indent=1)
trow('eq_close', 'Equity, closing', calc=lambda L, y: f"={L}{TS['eq_open']}+{L}{TS['eq_ni']}+{L}{TS['eq_issue']}", histcalc=lambda H, L: f"=History!{H}{HIST_ROW['equity']}", bold=True)
trow('bs_check', 'Check: net assets less debt less equity', calc=lambda L, y: f"={L}{TS['ta']}-{L}{TS['debt']}-{L}{TS['eq_close']}", histcalc=lambda H, L: f"={L}{TS['ta']}-{L}{TS['debt']}-{L}{TS['eq_close']}", fmt=NUM3)
r += 1; band(wsT, r, 'Cash flow statement', 2, YC0 + len(YEARS) - 1); r += 1
trow('cf_ni', 'Net income', calc=lambda L, y: f"={L}{TS['ni']}", histcalc=lambda H, L: f"=History!{H}{HIST_ROW['net_income']}", indent=1)
trow('cf_da', 'Depreciation and amortisation', calc=lambda L, y: f"={L}{TS['da']}" + (f"+History!E{HIST_ROW['depreciation']}*fx" if y == 2026 else ''), histcalc=lambda H, L: f"=History!{H}{HIST_ROW['depreciation']}", fmt=NUM2, indent=1)
trow('cf_sbc', 'Share-based compensation and other non-cash (reported)', histcalc=lambda H, L: f"=History!{H}{HIST_ROW['sbc']}", calc=lambda L, y: f"=History!E{HIST_ROW['sbc']}*fx" if y == 2026 else '=0', indent=1)
trow('cf_nwc', 'Change in operating working capital', fin=FINB('dnwc'), calc=lambda L, y: f"=-({sum_year('Fin Monthly', FINB('dnwc'), L+'$4')})", histcalc=lambda H, L: f"=History!{H}{HIST_ROW['cfo']}-History!{H}{HIST_ROW['net_income']}-History!{H}{HIST_ROW['depreciation']}-History!{H}{HIST_ROW['sbc']}", indent=1)
trow('cfo', 'Cash from operations', calc=lambda L, y: f"={L}{TS['cf_ni']}+{L}{TS['cf_da']}+{L}{TS['cf_sbc']}+{L}{TS['cf_nwc']}" + (f"+(History!E{HIST_ROW['cfo']}*fx-History!E{HIST_ROW['net_income']}*fx-History!E{HIST_ROW['depreciation']}*fx-History!E{HIST_ROW['sbc']}*fx)" if y == 2026 else ''), histcalc=lambda H, L: f"=History!{H}{HIST_ROW['cfo']}", bold=True)
trow('capex', 'Capital expenditure', calc=lambda L, y: f"=-({sum_year('Fin Monthly', FINB('capex'), L+'$4')})" + (f"-History!E{HIST_ROW['capex']}*fx" if y == 2026 else ''), histcalc=lambda H, L: f"=-History!{H}{HIST_ROW['capex']}", fmt=NUM2, indent=1)
trow('cfi', 'Cash from investing', calc=lambda L, y: f"={L}{TS['capex']}", histcalc=lambda H, L: f"=History!{H}{HIST_ROW['cfi']}", bold=True)
trow('issue', 'Equity issued, net of fees', calc=lambda L, y: f"={sum_year('Fin Monthly', FINB('fund_net'), L+'$4')}" + (f"+History!E{HIST_ROW['cff']}*fx" if y == 2026 else ''), histcalc=lambda H, L: f"=History!{H}{HIST_ROW['cff']}", indent=1)
trow('cff', 'Cash from financing', calc=lambda L, y: f"={L}{TS['issue']}", histcalc=lambda H, L: f"=History!{H}{HIST_ROW['cff']}", bold=True)
trow('fx', 'FX on cash (reported)', histcalc=lambda H, L: f"=History!{H}{HIST_ROW['cash_fx']}", calc=lambda L, y: f"=History!E{HIST_ROW['cash_fx']}*fx" if y == 2026 else '=0')
trow('net', 'Net change in cash', calc=lambda L, y: f"={L}{TS['cfo']}+{L}{TS['cfi']}+{L}{TS['cff']}+{L}{TS['fx']}", histcalc=lambda H, L: f"={L}{TS['cfo']}+{L}{TS['cfi']}+{L}{TS['cff']}+{L}{TS['fx']}", bold=True)
trow('cash_open', 'Cash, opening', calc=lambda L, y: (f"=History!E{HIST_ROW['cash_begin']}*fx" if y == 2026 else f"={CL(YC0+YEARS.index(y)-1)}{r+1}"), histcalc=lambda H, L: f"=History!{H}{HIST_ROW['cash_begin']}")
trow('cash_close', 'Cash, closing', calc=lambda L, y: f"={L}{TS['cash_open']}+{L}{TS['net']}", histcalc=lambda H, L: f"={L}{TS['cash_open']}+{L}{TS['net']}", bold=True)
trow('cf_check', 'Check: closing cash against the balance sheet', calc=lambda L, y: f"={L}{TS['cash_close']}-{L}{TS['cash']}", histcalc=lambda H, L: f"={L}{TS['cash_close']}-{L}{TS['cash']}", fmt=NUM3)
r += 1
put(wsT, r, 2, 'Reported 2024-2025 follow IFRS and include finance income and non-cash share-based compensation. From H2 2026 the forecast is an economic cash-compensation view: existing awards enter dilution at valuation, future incentives sit inside the cash budgets, and there is no interest or FX. The 2026 column adds the reported H1 (EUR x 1.1186) to the H2 forecast, so its opening cash is the reported 31 December 2025 balance translated at the same rate.', f_sub); r += 1
put(wsT, r, 2, 'Working capital is carried net (12% of annualised sales plus the opening balance); liabilities beyond payables are inside other net assets. Debt is zero throughout. Years after 2031 are grouped.', f_sub)
widths(wsT, {'A': 2, 'B': 48, 'C': 8}); wsT.freeze_panes = 'D6'; wsT.print_area = f"B4:{CL(YC0+YEARS.index(2031))}{TS['cf_check']}"

# ================================================================== QUARTERLY
wsQ = sheet('Quarterly', '0070C0'); title(wsQ, 'Quarterly financials, 2026 H2 to 2029', 'USD millions, both products approved. Flows sum the monthly engine; balances take the quarter-end month.')
QS = [f'{y}Q{q}' for y in (2026, 2027, 2028, 2029) for q in (1, 2, 3, 4) if not (y == 2026 and q < 3)]
put(wsQ, 5, 2, 'Quarter', f_bold)
for i, q in enumerate(QS): put(wsQ, 5, 4 + i, q, f_hdr, fill=fill_hdr, align=Alignment(horizontal='right')); wsQ.column_dimensions[CL(4 + i)].width = 10
QR = {}; r = 7
def qrow(key, label, fin=None, kind='flow', sheet_name='Fin Monthly', fmt=NUM, bold=False, calc=None, indent=0):
    global r; QR[key] = r; put(wsQ, r, 2, ('   ' * indent) + label, f_bold if bold else f_norm)
    for i, q in enumerate(QS):
        c = 4 + i; L = CL(c); qc = f'{L}$5'
        v = calc(L) if calc else ('=' + (sum_q(sheet_name, fin, qc) if kind == 'flow' else stock_q(sheet_name, fin, qc)))
        put(wsQ, r, c, v, f_bold if bold else f_norm, fmt)
    r += 1
band(wsQ, r, 'Patients', 2, 3 + len(QS)); r += 1
qrow('xr_starts', 'XR paid activations', RX['tot_starts'], sheet_name='XR Engine', fmt=NUM0); qrow('xr_active', 'XR active patients, quarter end', RX['tot_active'], kind='stock', sheet_name='XR Engine', fmt=NUM0)
qrow('ir_starts', 'IR paid activations', RI['tot_starts'], sheet_name='IR Engine', fmt=NUM0); qrow('ir_active', 'IR active members, quarter end', RI['tot_active'], kind='stock', sheet_name='IR Engine', fmt=NUM0)
qrow('ir_units', 'IR paid doses', RV['IR_A_units'], sheet_name='Revenue', fmt=NUM0); qrow('xr_units', 'XR paid treatment-years', RV['XR_units'], sheet_name='Revenue', fmt=NUM0)
r += 1; band(wsQ, r, 'Income statement', 2, 3 + len(QS)); r += 1
qrow('ir_rev', 'IR revenue', RV['IR_A_rec'], sheet_name='Revenue'); qrow('xr_rev', 'XR revenue', RV['XR_rec'], sheet_name='Revenue')
qrow('revenue', 'Total revenue', FINB('revenue'), bold=True)
qrow('cogs', 'Manufacturing and royalty', calc=lambda L: f"={sum_q('Fin Monthly', FINB('manufacture'), L+'$5')}+{sum_q('Fin Monthly', FINB('license'), L+'$5')}")
qrow('rd', 'Research and development', FINB('rd')); qrow('comm', 'Commercial', FINB('commercial')); qrow('ga', 'General and administrative', FINB('ga')); qrow('dep', 'Depreciation', FINB('dep'), fmt=NUM2)
qrow('ebit', 'Operating income', calc=lambda L: f"={L}{QR['revenue']}-{L}{QR['cogs']}-{L}{QR['rd']}-{L}{QR['comm']}-{L}{QR['ga']}-{L}{QR['dep']}", bold=True)
qrow('tax', 'Tax', FINB('tax')); qrow('ni', 'Net income', FINB('ni'), bold=True)
r += 1; band(wsQ, r, 'Cash flow and balance', 2, 3 + len(QS)); r += 1
qrow('fcf', 'Free cash flow', FINB('fcf'), bold=True); qrow('fund', 'Net equity funding required', FINB('fund_net')); qrow('cash', 'Cash, quarter end', FINB('cash'), kind='stock', bold=True); qrow('shares', 'Shares outstanding, millions', FINB('shares'), kind='stock', fmt=NUM2)
r += 1; put(wsQ, r, 2, 'Commercial launch of IR is assumed in May 2027 after the 23 April 2027 PDUFA date; XR in June 2028. First paid activations follow the one-month access delay.', f_sub)
widths(wsQ, {'A': 2, 'B': 40, 'C': 4}); wsQ.freeze_panes = 'D6'

# ================================================================== ADOPTION BUILD: annual summary table
r = ADOPT_ANNUAL_ROW; band(wsA, r, 'Annual adoption by origin (both products approved)', 2, 14); r += 1
AYRS = list(range(2027, 2038)); put(wsA, r, 2, 'Year', f_bold)
for i, y in enumerate(AYRS): put(wsA, r, 4 + i, y, f_hdr, fill=fill_hdr, align=Alignment(horizontal='right'))
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
band(wsP, r, 'Annual quantities and revenue (both products approved)', 2, 16); r += 1
PYRS = list(range(2027, 2041)); put(wsP, r, 2, 'Year', f_bold)
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
widths(wsP, {'A': 2, 'B': 52, 'C': 4})

# ================================================================== STREET BRIDGE (engine counterfactuals + Street data)
SPX = Path('/private/tmp/claude-501/-Users-kwu-Library-Application-Support-Claude-scratch-workspaces-665f01e2-9652-4512-9fda-0497b55d354c-06f88a99-9105-4941-a8cf-87e6b02aff65-scratch-2026-10-10-7ee708/b9401f11-15ac-4c44-8c9e-8eb278b73abd/scratchpad')
ALT_A = json.loads((SPX / 'altdata_A.json').read_text())['datasets']; ALT_B = json.loads((SPX / 'altdata_B.json').read_text())['datasets']
sys.path.insert(0, str(FIN)); import engine as E, inspect
_src = inspect.getsource(E.resolve).replace("    assert abs(p['ir_quantity_multiplier'] - 1.17 / 1.38) < 1e-12, 'All active forecasts require the selected redosing discount'\n", '')
_ns = {}; exec(_src, E.__dict__, _ns); resolve_relaxed = _ns['resolve']
def run_counterfactual(overrides):
    saved = E.resolve; E.resolve = resolve_relaxed
    try: out = E.evaluate(cfg, 'working', overrides)
    finally: E.resolve = saved
    return out
BRIDGE_STEPS = [
 ('Our engine with the three thesis assumptions at their favourable bounds: no redosing haircut, switching at the top of its range, 95% rescue continuation', {'ir_quantity_multiplier': 1.0, 'category_switch': 0.2, 'switch_mature_multiplier': 1.25, 'ir_retention': 0.95}),
 ('Apply the redosing quantity haircut (1.17/1.38 doses per attack)', {'category_switch': 0.2, 'switch_mature_multiplier': 1.25, 'ir_retention': 0.95}),
 ('Switching opportunity at the claims benchmark (15.2%, flat)', {'ir_retention': 0.95}),
 ('Rescue continuation at 90%', {}),
]
bridge = []
for lab, ov in BRIDGE_STEPS:
    o = run_counterfactual(ov); bridge.append((lab, o['conditional_valuations']['both']['target']['price'], o['expected_12month_value'], o['conditional_peak']['recognized_revenue_usdm']))
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
put(wsS, r, 2, 'Our expected value in 12 months', f_bold); put(wsS, r, 4, '=value_target', f_bold, USD, fill=fill_key); r += 1
put(wsS, r, 2, 'Our value if both products are approved'); put(wsS, r, 4, '=price_both_target', f_norm, USD); r += 1
put(wsS, r, 2, 'Targets transcribed from the supplied September 2026 notes (reverse-DCF research, 7 October 2026). Leerink and Morgan Stanley later updates supersede their earlier figures; Wells Fargo states no terminal growth, which is not no terminal value.', f_sub); r += 2
gp = ALT_A['guggenheim_revenue_path']; band(wsS, r, 'Revenue path: Guggenheim (USD at 1.16) against this model, both products approved', 2, 16); r += 1
gyears = [row_[0] for row_ in gp['rows']]; gcols = gp['columns']
iIRu = [i for i, c in enumerate(gcols) if 'IR' in c and 'USD' in c][0]; iXRu = [i for i, c in enumerate(gcols) if 'XR' in c and 'USD' in c][0]
put(wsS, r, 2, 'Calendar year', f_bold); [put(wsS, r, 4 + i, y, f_hdr, fill=fill_hdr, align=Alignment(horizontal='right')) for i, y in enumerate(gyears)]; GY = r; r += 1
put(wsS, r, 2, 'Street IR revenue'); [put(wsS, r, 4 + i, row_[iIRu], f_input, NUM) for i, row_ in enumerate(gp['rows'])]; GIR = r; r += 1
put(wsS, r, 2, 'Street XR revenue'); [put(wsS, r, 4 + i, row_[iXRu], f_input, NUM) for i, row_ in enumerate(gp['rows'])]; GXR = r; r += 1
put(wsS, r, 2, 'Street total', f_bold); [put(wsS, r, 4 + i, f'={CL(4+i)}{GIR}+{CL(4+i)}{GXR}', f_bold, NUM) for i in range(len(gyears))]; GT = r; r += 1
put(wsS, r, 2, 'Model IR revenue'); [put(wsS, r, 4 + i, f"=IF({CL(4+i)}${GY}<2026,0,INDEX(Annual!$D${AN['ir_rev']}:${CL(YC0+len(YEARS)-1)}${AN['ir_rev']},1,MATCH({CL(4+i)}${GY},Annual!$D$4:${CL(YC0+len(YEARS)-1)}$4,0)))", f_norm, NUM) for i in range(len(gyears))]; MIR = r; r += 1
put(wsS, r, 2, 'Model XR revenue'); [put(wsS, r, 4 + i, f"=IF({CL(4+i)}${GY}<2026,0,INDEX(Annual!$D${AN['xr_rev']}:${CL(YC0+len(YEARS)-1)}${AN['xr_rev']},1,MATCH({CL(4+i)}${GY},Annual!$D$4:${CL(YC0+len(YEARS)-1)}$4,0)))", f_norm, NUM) for i in range(len(gyears))]; MXR = r; r += 1
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
put(wsS, r, 2, 'Live model (should equal the last step when the Working case is selected)', f_bold); put(wsS, r, 3, '=price_both_target', f_bold, USD); put(wsS, r, 5, '=value_target', f_bold, USD); put(wsS, r, 7, f"=MAX(Annual!$D${AN['revenue']}:${CL(YC0+len(YEARS)-1)}${AN['revenue']})", f_bold, NUM); r += 1
put(wsS, r, 2, 'Street median target for reference'); put(wsS, r, 3, '=street_median', f_norm, USD); r += 1
put(wsS, r, 2, 'The first row relaxes the redosing haircut, which the published engine refuses in active cases; it is shown only to size the assumption. Each later row restores one working assumption. Order matters because the effects interact; the end point is the Working case.', f_sub); r += 1
wsS.add_chart(CHART_S, f'B{r+1}')
for rr_ in range(BR0, BR0 + len(bridge) + 3): wsS.cell(rr_, 2).alignment = Alignment(wrap_text=True, vertical='top')
widths(wsS, {'A': 2, 'B': 64, 'C': 13, 'D': 11, 'E': 13, 'F': 11, 'G': 12, 'H': 11, 'I': 11, 'J': 11, 'K': 11, 'L': 11, 'M': 11, 'N': 11, 'O': 11, 'P': 11})
wsS.print_area = f'B1:P{r+22}'

# ================================================================== SENSITIVITIES
wsSe = sheet('Sensitivities', '0070C0'); title(wsSe, 'Sensitivities and scenarios', 'One-variable ranges and structural stresses from the Python engine on the 9 October 2026 inputs. The live model reproduces any row when the same input is changed on the Inputs tab.')
sm = {x['name']: x for x in sens['results']}; base_t = sens['base_target']
r = 4; band(wsSe, r, 'Assumptions that carry the thesis', 2, 10); r += 1
for i, h in enumerate(['Assumption', 'Working', 'Low', 'High', 'Value at low', 'Value at high', 'Span', 'Live value']): put(wsSe, r, 2 + i, h, f_hdr, fill=fill_hdr)
r += 1
for k in ['ir_quantity_multiplier', 'category_switch', 'switch_mature_multiplier', 'ir_retention']:
    put(wsSe, r, 2, LABELS[k], f_bold, fill=fill_key); put(wsSe, r, 3, f'=INDEX(Inputs!$E${INPUT_ROW[k]},1)', f_norm, NUM3); put(wsSe, r, 9, f"={'pop_growth' if k == 'growth' else k}", f_norm, NUM3)
    tests = [x for n, x in sm.items() if n.startswith(k + '=')]
    if tests:
        lo, hi = sorted(tests, key=lambda x: list(x['overrides'].values())[0])
        put(wsSe, r, 4, list(lo['overrides'].values())[0], f_norm, NUM3); put(wsSe, r, 5, list(hi['overrides'].values())[0], f_norm, NUM3); put(wsSe, r, 6, lo['expected_target'], f_input, USD); put(wsSe, r, 7, hi['expected_target'], f_input, USD); put(wsSe, r, 8, f'=G{r}-F{r}', f_norm, USD)
    else:
        put(wsSe, r, 4, 'fixed', f_sub); put(wsSe, r, 6, 'see Street Bridge', f_sub)
    r += 1
put(wsSe, r, 2, f'Working expected 12-month value in the engine: ${base_t:.2f}. Spans are chosen bounds, not statistical importance.', f_sub); r += 2
band(wsSe, r, 'All one-variable ranges, ranked by span', 2, 10); r += 1
for i, h in enumerate(['Input', 'Low', 'High', 'Value at low', 'Value at high', 'Span']): put(wsSe, r, 2 + i, h, f_hdr, fill=fill_hdr)
r += 1
for x in rank:
    k = x['parameter']; tests = sorted([t for n, t in sm.items() if n.startswith(k + '=')], key=lambda t: list(t['overrides'].values())[0])
    put(wsSe, r, 2, LABELS.get(k, k), f_bold if k in THESIS else f_norm); put(wsSe, r, 3, list(tests[0]['overrides'].values())[0], f_norm, NUM3); put(wsSe, r, 4, list(tests[-1]['overrides'].values())[0], f_norm, NUM3)
    put(wsSe, r, 5, tests[0]['expected_target'], f_input, USD); put(wsSe, r, 6, tests[-1]['expected_target'], f_input, USD); put(wsSe, r, 7, x['span'], f_norm, USD); r += 1
r += 1; band(wsSe, r, 'Structural stresses and access packages', 2, 10); r += 1
for i, h in enumerate(['Scenario', 'Value today', '12-month value', 'Change', 'Peak revenue, USD m', 'Overrides']): put(wsSe, r, 2 + i, h, f_hdr, fill=fill_hdr)
r += 1
SCN = {'hazards_halved_after_year5': 'Switching hazards halved after year 5', 'quarter_of_stoppers_retry_after24m': '25% of discontinuers return after 24 months', 'other_prevention_shift20pct_by10y': '20% of rescue-only users move to other prevention by year 10', 'generic_price_capped_flat': 'Post-generic price capped flat', 'initial_rescue_stock_free': 'Initial rescue stock supplied free', 'partner_all_overseas': 'All overseas sales licensed (20% royalty)', 'six_month_launch_delay_both': 'Both launches delayed six months', 'no_refill_intensity_adjustment': 'No refill-gap adjustment on XR', 'old_prices_only': 'Older price anchors ($14,000 and $418,000)', 'omit_license_royalty_cost_counterfactual': 'No BRAIN royalty (counterfactual)', 'redose_plus_better_IR_capture': 'Redosing haircut plus 1.25x injectable opportunity', 'LTP_ten_percent_no_rescue': '10% of preventive users hold no rescue', 'extra_Ekterly_free_stock': '20% extra Ekterly free supply', 'rebate_for_access': 'Access package: more rebate, better access (90% conversion, 58%/76% net)', 'restricted_high_net': 'Access package: restricted access, higher net (65% conversion, 3-month delay, 78%/94% net)'}
for n, lab in SCN.items():
    x = sm[n]; put(wsSe, r, 2, lab); put(wsSe, r, 3, x['value_today'], f_input, USD); put(wsSe, r, 4, x['expected_target'], f_input, USD); put(wsSe, r, 5, x['delta_target'], f_norm, USD); put(wsSe, r, 6, x['peak']['recognized_revenue_usdm'], f_input, NUM); put(wsSe, r, 7, ', '.join(f'{k}={v}' for k, v in x['overrides'].items()), f_sub); r += 1
r += 1; band(wsSe, r, 'Commercial cases', 2, 10); r += 1
for i, h in enumerate(['Case', 'Value today (engine)', '12-month value (engine)', 'Peak revenue, USD m', 'Peak year', 'Live model when selected']): put(wsSe, r, 2 + i, h, f_hdr, fill=fill_hdr)
r += 1
for c, lab in (('working', 'Working'), ('cautious', 'Cautious'), ('upside_challenge', 'Upside')):
    x = results[c]; put(wsSe, r, 2, lab); put(wsSe, r, 3, x['risk_adjusted_value_today'], f_input, USD); put(wsSe, r, 4, x['expected_12month_value'], f_input, USD); put(wsSe, r, 5, x['conditional_peak']['recognized_revenue_usdm'], f_input, NUM); put(wsSe, r, 6, x['conditional_peak']['year'], f_norm, '0')
    put(wsSe, r, 7, f'=IF(case_selector="{lab}",value_target,"")', f_norm, USD); r += 1
put(wsSe, r, 2, 'Cautious and Upside are bundles of overrides listed on the Inputs tab, not Monte Carlo percentiles. The cautious bundle needs about USD 1bn of funding before cash flow turns, which is why its equity value rounds to zero.', f_sub)
widths(wsSe, {'A': 2, 'B': 60, 'C': 14, 'D': 14, 'E': 14, 'F': 14, 'G': 14, 'H': 60, 'I': 12})

# ================================================================== MONTE CARLO
wsM = sheet('Monte Carlo', '0070C0'); title(wsM, 'Monte Carlo: saved simulation, 9 October 2026 assumptions', 'Published run on the Python engine. These results do not change when workbook inputs change; the stale flag below says whether the live inputs still match the run.')
pr = mc['primary']; r = 4; band(wsM, r, 'Run identity', 2, 8); r += 1
for lab, v in (('Source commit', mc['source_commit']), ('Seed', pr['seed']), ('Primary draws', pr['draws']), ('Access package', pr['access']), ('Structure', pr['structural']), ('Dependence', pr['dependence']), ('Python', pr['python_version']), ('Engine hash (engine.py)', pr['hashes']['models/finalized/engine.py'][:16] + '...'), ('Inputs hash (inputs.json)', pr['hashes']['models/finalized/inputs.json'][:16] + '...')):
    put(wsM, r, 2, lab); put(wsM, r, 4, v, f_input); r += 1
put(wsM, r, 2, 'Live inputs versus the saved run', f_bold)
put(wsM, r, 4, f'=IF(SUMPRODUCT(--(Inputs!$H$9:$H${max(INPUT_ROW.values())}<>Inputs!$E$9:$E${max(INPUT_ROW.values())}))+ABS(launch_delay_IR)+ABS(launch_delay_XR)+ABS(loe_year_IR-loe_year)+ABS(loe_year_XR-loe_year)=0,"Inputs match the saved run","STALE: live inputs differ from the simulated baseline")', f_flag); r += 2
band(wsM, r, 'Distribution of the 12-month value, USD per share', 2, 8); r += 1
for i, h in enumerate(['', 'Mean', 'Median', 'P10', 'P90']): put(wsM, r, 2 + i, h, f_hdr, fill=fill_hdr)
r += 1
for lab, key in (('Risk-adjusted expected value across commercial draws', 'expected_target'), ('Conditional clinical outcomes, weighted once', 'clinical_outcomes_target')):
    x = pr[key]; put(wsM, r, 2, lab); [put(wsM, r, 3 + i, x[k], f_input, USD) for i, k in enumerate(('mean', 'median', 'p10', 'p90'))]
    if key == 'expected_target': name('mc_mean_target', f"'Monte Carlo'!$C${r}")
    r += 1
put(wsM, r, 2, 'Value today, risk-adjusted'); x = pr['value_today']; [put(wsM, r, 3 + i, x[k], f_input, USD) for i, k in enumerate(('mean', 'median', 'p10', 'p90'))]; r += 1
put(wsM, r, 2, 'Working case (zero-volatility point)'); put(wsM, r, 3, results['working']['expected_12month_value'], f_input, USD); put(wsM, r, 5, 'the mode of a triangular prior is not its mean, so the point case need not equal the simulation mean', f_sub); r += 1
put(wsM, r, 2, 'The middle 80% is an analyst-prior sensitivity range, not an empirical confidence interval or a stock-price forecast. No probability of a profitable short is derived from it.', f_sub); r += 2
# histogram data from compact draws
with gzip.open(ROOT / 'models/monte_carlo/results/seed_7202609/working__primary__primitive_independent/draws_compact.json.gz', 'rt') as fh: dc = json.load(fh)
ci = {c: i for i, c in enumerate(dc['columns'])}; rows_ = dc['rows']; N = len(rows_); W_ = dc['clinical_weights']
edges = list(range(0, 72, 2)); hist_e = [0.0] * (len(edges) - 1); hist_o = [0.0] * (len(edges) - 1)
def binidx(v): return min(max(int(v // 2), 0), len(edges) - 2)
for row_ in rows_:
    hist_e[binidx(row_[ci['expected_target']])] += 1 / N
    for s_ in ('both', 'IR_only', 'XR_only', 'neither'): hist_o[binidx(row_[ci[f'{s_}.target.price']])] += W_[s_] / N
band(wsM, r, 'Histogram, $2 bins (share of draws)', 2, 8); r += 1
put(wsM, r, 2, 'Bin lower edge', f_hdr, fill=fill_hdr); put(wsM, r, 3, 'Risk-adjusted expected values', f_hdr, fill=fill_hdr); put(wsM, r, 4, 'Clinical outcomes weighted once', f_hdr, fill=fill_hdr); r += 1; H0 = r
for i in range(len(edges) - 1): put(wsM, r, 2, edges[i], f_norm, USD0); put(wsM, r, 3, hist_e[i], f_input, PCT); put(wsM, r, 4, hist_o[i], f_input, PCT); r += 1
H1 = r - 1
bc = BarChart(); bc.type = 'col'; bc.grouping = 'clustered'; bc.title = 'How the 12-month value moves across commercial assumptions'; bc.height = 9; bc.width = 20; bc.y_axis.title = 'Share of draws'; bc.x_axis.title = 'USD per share, $2 bins'; bc.style = 2; bc.gapWidth = 30
bc.add_data(Reference(wsM, min_col=3, max_col=4, min_row=H0 - 1, max_row=H1), titles_from_data=True); bc.set_categories(Reference(wsM, min_col=2, min_row=H0, max_row=H1)); wsM.add_chart(bc, 'G4')
r += 1; band(wsM, r, 'Which uncertain inputs move the value (Spearman rank association, not dollars or causation)', 2, 8); r += 1
put(wsM, r, 2, 'Input', f_hdr, fill=fill_hdr); put(wsM, r, 3, 'Spearman', f_hdr, fill=fill_hdr); r += 1; S0 = r
for x in pr['rank_associations_prior_dependent_not_causal'][:14]: put(wsM, r, 2, LABELS.get(x['parameter'], x['parameter'])); put(wsM, r, 3, x['spearman'], f_input, NUM2); r += 1
S1 = r - 1
sc_ = BarChart(); sc_.type = 'bar'; sc_.title = 'Leading drivers under these priors'; sc_.height = 9; sc_.width = 16; sc_.style = 2; sc_.legend = None
sc_.add_data(Reference(wsM, min_col=3, min_row=S0, max_row=S1)); sc_.set_categories(Reference(wsM, min_col=2, min_row=S0, max_row=S1)); sc_.x_axis.scaling.orientation = 'maxMin'; wsM.add_chart(sc_, 'G24')
put(wsM, r, 2, 'Fixed assumptions (redosing haircut, 60% mature share, price multipliers) do not vary and so cannot appear here.', f_sub); r += 2
band(wsM, r, 'Clinical branches', 2, 8); r += 1
for i, h in enumerate(['Branch', 'Weight', 'Median 12-month value', 'Median future net funding, USD m']): put(wsM, r, 2 + i, h, f_hdr, fill=fill_hdr)
r += 1
for s_ in ('both', 'IR_only', 'XR_only', 'neither'): put(wsM, r, 2, SNAME[s_]); put(wsM, r, 3, W_[s_], f_input, PCT2); put(wsM, r, 4, pr['states'][s_]['target_price']['median'], f_input, USD); put(wsM, r, 5, pr['states'][s_]['future_net_funding_today_usdm']['median'], f_input, NUM); r += 1
r += 1; band(wsM, r, 'Paired scenarios (same draws, one change each)', 2, 8); r += 1
for i, h in enumerate(['Scenario', 'Draws', 'Mean', 'Median', 'P10', 'P90', 'Change vs matched primary']): put(wsM, r, 2 + i, h, f_hdr, fill=fill_hdr)
r += 1
SCLAB = {'rebate_for_access__primary': 'More rebates for better access', 'restricted_high_net__primary': 'Restricted access, higher net price', 'working__extra_Ekterly_free_supply': '20% extra Ekterly free supply', 'working__higher_discount_rate': '15% discount rate', 'working__initial_stock_free': 'Free initial rescue stock', 'working__later_LOE': 'Loss of exclusivity in 2043', 'working__licensed_overseas': 'Overseas licensing', 'working__lower_discount_rate': '10% discount rate', 'working__no_rescue_in_10pct_LTP': '10% of preventive users lack rescue', 'working__primary__commercial_rank_linked': 'Commercial variables rank-linked', 'working__six_month_delay': 'Both launches delayed six months', 'working__some_discontinuers_return': '25% of discontinuers can return'}
for x in mc['paired_scenarios']:
    key = x['scenario']; lab = next((v for k, v in SCLAB.items() if key.startswith(k)), key)
    put(wsM, r, 2, lab); put(wsM, r, 3, x['draws'], f_norm, NUM0); e = x['expected_target']; [put(wsM, r, 4 + i, e[k], f_input, USD) for i, k in enumerate(('mean', 'median', 'p10', 'p90'))]; put(wsM, r, 8, x['paired_expected_target_change']['mean'] if isinstance(x['paired_expected_target_change'], dict) else x['paired_expected_target_change'], f_input, USD); r += 1
r += 1; band(wsM, r, 'The 28 sampled inputs (triangular priors)', 2, 8); r += 1
for i, h in enumerate(['Input', 'Minimum', 'Mode', 'Maximum', 'Sources', 'Why this range']): put(wsM, r, 2 + i, h, f_hdr, fill=fill_hdr)
r += 1
for k, v in priors['priors'].items():
    put(wsM, r, 2, LABELS.get(k, k)); put(wsM, r, 3, v['min'], f_input, NUM3); put(wsM, r, 4, v['mode'], f_input, NUM3); put(wsM, r, 5, v['max'], f_input, NUM3); put(wsM, r, 6, ', '.join(v['sources']), f_sub); put(wsM, r, 7, RATIONALE.get(k, fix_spacing(v['rationale'])), f_sub); wsM.cell(r, 7).alignment = Alignment(wrap_text=True, vertical='top'); r += 1
r += 1; put(wsM, r, 2, 'Refresh: python3 models/monte_carlo/run.py --draws 4096 --seed 7202609 --workers 4, then package_results.py and verify_artifacts.py; rebuild this workbook with models/excel/build_model.py. Rerun only after the engine inputs are deliberately changed and re-verified.', f_sub)
widths(wsM, {'A': 2, 'B': 48, 'C': 16, 'D': 18, 'E': 18, 'F': 16, 'G': 16, 'H': 70})

# ================================================================== CHECKS
wsK = sheet('Checks', 'C00000'); title(wsK, 'Checks', 'Live workbook against the saved Python engine (Working case), plus accounting and population identities. Differences should be zero when the Working case is selected with no overrides.')
r = 4; band(wsK, r, 'Control totals, Working case', 2, 8); r += 1
for i, h in enumerate(['Metric', 'Engine (saved)', 'Workbook (live)', 'Difference', 'Tolerance', 'Status']): put(wsK, r, 2 + i, h, f_hdr, fill=fill_hdr)
r += 1; K0 = r
w = results['working']; ann = {a['year']: a for a in w['annual_both_success']}
def krow(label, engine_val, live_formula, tol, fmt=USD):
    global r; put(wsK, r, 2, label); put(wsK, r, 3, engine_val, f_input, fmt); put(wsK, r, 4, live_formula, f_norm, fmt); put(wsK, r, 5, f'=D{r}-C{r}', f_norm, fmt); put(wsK, r, 6, tol, f_sub, fmt); put(wsK, r, 7, f'=IF(ABS(E{r})<=F{r},"OK","CHECK")', f_bold); r += 1
krow('Value today, probability-weighted', w['risk_adjusted_value_today'], '=value_today', 0.01)
krow('12-month value, probability-weighted', w['expected_12month_value'], '=value_target', 0.01)
for s_ in ('both', 'IR_only', 'XR_only', 'neither'):
    ci_ = 3 + ['both', 'IR_only', 'XR_only', 'neither'].index(s_)
    krow(f'{SNAME[s_]}: 12-month value', w['conditional_valuations'][s_]['target']['price'], f"=Valuation!{CL(ci_)}{VALROW[('tgt','price')]}", 0.01)
    krow(f'{SNAME[s_]}: value today', w['conditional_valuations'][s_]['valuation']['price'], f"=Valuation!{CL(ci_)}{VALROW[('val','price')]}", 0.01)
krow('Peak revenue, USD m', w['conditional_peak']['recognized_revenue_usdm'], f"=MAX(Annual!$D${AN['revenue']}:${CL(YC0+len(YEARS)-1)}${AN['revenue']})", 0.01, NUM2)
for y in (2027, 2028, 2029, 2030, 2031, 2038):
    L = CL(YC0 + YEARS.index(y)); krow(f'{y} revenue, USD m', ann[y]['revenue'], f'=Annual!{L}{AN["revenue"]}', 0.01, NUM2)
for y in (2030, 2038):
    L = CL(YC0 + YEARS.index(y)); krow(f'{y} IR revenue', ann[y]['IR_recognized'], f'=Annual!{L}{AN["ir_rev"]}', 0.01, NUM2); krow(f'{y} XR revenue', ann[y]['XR_recognized'], f'=Annual!{L}{AN["xr_rev"]}', 0.01, NUM2)
for y in (2027, 2028, 2031):
    L = CL(YC0 + YEARS.index(y)); krow(f'{y} free cash flow', ann[y]['fcf'], f'=Annual!{L}{AN["fcf"]}', 0.01, NUM2); krow(f'{y} year-end cash', ann[y]['cash_end'], f'=Annual!{L}{AN["cash"]}', 0.01, NUM2); krow(f'{y} net funding', ann[y]['funding_net'], f'=Annual!{L}{AN["fund"]}', 0.01, NUM2); krow(f'{y} shares, millions', ann[y]['shares_end'], f'=Annual!{L}{AN["shares"]}', 0.001, NUM3)
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
irow('Monthly revenue equals annual revenue (difference)', f"=SUM({mrng('Fin Monthly', FINB('revenue'))})-SUM(Annual!$F${AN['revenue']}:${CL(YC0+len(YEARS)-1)}${AN['revenue']})")
irow('Monthly revenue equals quarterly revenue through 2029 (difference)', f"=SUMIFS({mrng('Fin Monthly', FINB('revenue'))},{mrng('Fin Monthly', YEAR_ROW)},\"<=2029\")-SUM(Quarterly!D{QR['revenue']}:{CL(3+len(QS))}{QR['revenue']})")
irow('Revenue before IR launch (should be zero)', f"=SUMPRODUCT({mrng('Fin Monthly', FINB('revenue'))},--({mrng('Calendar', CAL['age_IR'])}<0))")
irow('Price x units equals US revenue, XR (largest difference)', f"=SUMPRODUCT(MAX(ABS({mrng('Revenue', RV['XR_units'])}*{mrng('Revenue', RV['XR_price'])}/1000000-{mrng('Revenue', RV['XR_US'])})))")
irow('Equity value equals price x diluted shares, both approved (difference)', f"=Valuation!C{VALROW[('tgt','price')]}*Valuation!C{VALROW[('tgt','diluted')]}-Valuation!C{VALROW[('tgt','equity')]}")
irow('Three Statements: largest balance-sheet check', f"=SUMPRODUCT(MAX(ABS('Three Statements'!$D${TS['bs_check']}:${CL(YC0+len(YEARS)-1)}${TS['bs_check']})))")
irow('Three Statements: largest cash-flow check', f"=SUMPRODUCT(MAX(ABS('Three Statements'!$D${TS['cf_check']}:${CL(YC0+len(YEARS)-1)}${TS['cf_check']})))")
irow('Opening balance: assets equal equity (EUR m)', f"=History!C{OPEN['equity_eurm']+1}")
irow('Reported cash roll-forward, FY2025', f"=History!D{HIST_ROW['shares_end']+1}")
I1 = r - 1
put(wsK, r, 2, 'All identities pass', f_bold); put(wsK, r, 7, f'=IF(COUNTIF(G{I0}:G{I1},"CHECK")=0,"OK","CHECK")', f_flag); name('checks_identities', f'Checks!$G${r}'); r += 2
put(wsK, r, 2, 'Passing checks show the workbook reproduces the engine arithmetic. They do not validate the commercial assumptions; the Inputs tab and the research archive carry that argument.', f_sub)
widths(wsK, {'A': 2, 'B': 62, 'C': 16, 'D': 16, 'E': 14, 'F': 12, 'G': 10})

# ================================================================== SOURCES
wsSo = sheet('Sources', 'FFC000'); title(wsSo, 'Sources', 'Source IDs used on the Inputs tab. Public filings, publications and company materials only; expert-call material is not an input.')
r = 4
for i, h in enumerate(['ID', 'Title', 'Date', 'Where in the source', 'What it establishes', 'Limitation', 'URL']): put(wsSo, r, 2 + i, h, f_hdr, fill=fill_hdr)
r += 1
for k, v in src.items():
    put(wsSo, r, 2, k, f_bold); put(wsSo, r, 3, v.get('title', '')); put(wsSo, r, 4, v.get('date', '') or ''); put(wsSo, r, 5, fix_spacing(v.get('locator', '') or ''), f_sub); put(wsSo, r, 6, fix_spacing(v.get('facts', '') or ''), f_sub); put(wsSo, r, 7, fix_spacing(v.get('limitation', '') or ''), f_sub)
    u = v.get('url'); c = put(wsSo, r, 8, u or '', f_link)
    if u: c.hyperlink = u
    for cc in range(3, 9): wsSo.cell(r, cc).alignment = Alignment(wrap_text=True, vertical='top')
    r += 1
widths(wsSo, {'A': 2, 'B': 12, 'C': 36, 'D': 12, 'E': 30, 'F': 50, 'G': 44, 'H': 40}); wsSo.freeze_panes = 'C5'

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
wsD = sheet('Alt Data ->', '0070C0'); put(wsD, 2, 2, 'Reference datasets', f_title); put(wsD, 3, 2, 'Quantified evidence gathered during the research: launch analogues, benchmarks, Street figures and clinical comparisons. Values only; no forecast cell depends on these tabs.', f_sub)
ALT_TABS = [('Street Targets', [ALT_A['street_targets']]), ('Street Revenue Path', [ALT_A['guggenheim_revenue_path']]), ('Orladeyo Launch', [ALT_B['orladeyo_launch']]), ('Ekterly Launch', [ALT_B['ekterly_launch']]), ('HAE Injectables', [ALT_A['hae_injectable_revenues']]), ('Generic Erosion', [ALT_A['generic_erosion_panel']]), ('Clinical Cross-Trial', [ALT_A['clinical_cross_trial']]), ('Second-Oral Analogues', [ALT_A['second_oral_analogues']]), ('Persistence Benchmarks', [ALT_B['persistence_switching_benchmarks']]), ('Payer Policies', [ALT_B['payer_policies']]), ('HAE Population', [ALT_B['hae_population']]), ('Company Facts', [ALT_B['phvs_company_facts']])]
for nm_, dss in ALT_TABS:
    ws_ = sheet(nm_, '7F7F7F'); title(ws_, nm_); r = 4
    for ds in dss: r = write_dataset(ws_, r, ds)
    widths(ws_, {'A': 2, 'B': 44, 'C': 18, 'D': 14, 'E': 16, 'F': 22, 'G': 22, 'H': 22, 'I': 22, 'J': 22, 'K': 22, 'L': 22})
    ws_.page_setup.fitToWidth = 1

# ================================================================== SUMMARY
wsSu = sheet('Summary', '002060'); title(wsSu, 'Pharvaris (NASDAQ: PHVS)', 'USD millions except per share. Standalone valuation of deucrictibant: IR on-demand (PDUFA 23 April 2027) and XR prophylaxis (approval assumed May 2028).')
r = 4; put(wsSu, r, 2, 'Recommendation', f_bold); put(wsSu, r, 4, '=IF(implied_return<0,"SHORT","LONG")', f_title); r += 1
put(wsSu, r, 2, 'Share price'); put(wsSu, r, 4, '=share_price', f_norm, USD); put(wsSu, r, 5, '=TEXT(price_date,"d mmm yyyy")', f_sub); r += 1
put(wsSu, r, 2, '12-month price target', f_bold); put(wsSu, r, 4, '=price_target', f_bold, USD, fill=fill_key); put(wsSu, r, 5, '="basis: "&pt_basis', f_sub); r += 1
put(wsSu, r, 2, 'Implied return', f_bold); put(wsSu, r, 4, '=implied_return', f_bold, PCT); r += 1
put(wsSu, r, 2, 'Working point: 12-month value, probability-weighted'); put(wsSu, r, 4, '=value_target', f_norm, USD); r += 1
put(wsSu, r, 2, 'Monte Carlo: expected 12-month value, mean of draws'); put(wsSu, r, 4, '=mc_mean_target', f_norm, USD); r += 1
put(wsSu, r, 2, 'Value today (working point)'); put(wsSu, r, 4, '=value_today', f_norm, USD); r += 1
put(wsSu, r, 2, 'Value in 12 months if both products are approved'); put(wsSu, r, 4, '=price_both_target', f_norm, USD); r += 1
put(wsSu, r, 2, 'Street median target'); put(wsSu, r, 4, '=street_median', f_norm, USD); r += 1
put(wsSu, r, 2, 'Case selected'); put(wsSu, r, 4, '=case_selector', f_norm); put(wsSu, r, 5, 'change on the Inputs tab', f_sub); r += 1
put(wsSu, r, 2, 'Model checks'); put(wsSu, r, 4, '=IF(AND(checks_controls="OK",checks_identities="OK"),"All checks pass",IF(checks_identities="OK","Identities pass; control totals differ from the saved engine (expected unless Working is selected)","CHECK"))', f_norm); r += 2
band(wsSu, r, 'Clinical branches and value', 2, 7); r += 1
for i, h in enumerate(['Branch', 'Probability', 'Value today', '12-month value', 'Funding needed after today, USD m', 'Weighted 12-month value']): put(wsSu, r, 2 + i, h, f_hdr, fill=fill_hdr)
r += 1; SB0 = r
for ci_, s_ in enumerate(('both', 'IR_only', 'XR_only', 'neither')):
    L = CL(3 + ci_); put(wsSu, r, 2, SNAME[s_]); put(wsSu, r, 3, f'=w_{s_}', f_norm, PCT); put(wsSu, r, 4, f"=Valuation!{L}{VALROW[('val','price')]}", f_norm, USD); put(wsSu, r, 5, f"=Valuation!{L}{VALROW[('tgt','price')]}", f_norm, USD); put(wsSu, r, 6, f"=Valuation!{L}{VALROW[('val','funding')]}", f_norm, NUM); put(wsSu, r, 7, f'=C{r}*E{r}', f_norm, USD); r += 1
put(wsSu, r, 2, 'Probability-weighted', f_bold); put(wsSu, r, 3, f'=SUM(C{SB0}:C{r-1})', f_bold, PCT); put(wsSu, r, 4, '=value_today', f_bold, USD); put(wsSu, r, 5, '=value_target', f_bold, USD); put(wsSu, r, 7, f'=SUM(G{SB0}:G{r-1})', f_bold, USD); r += 2
band(wsSu, r, 'Equity bridge, both products approved, at the target date', 2, 7); r += 1
for k, lab in (('cash', 'Cash at the target date'), ('pv', 'PV of free cash flow, 2027-2055'), ('tv', 'PV of the declining residual after 2055'), ('fees', 'Less PV of issuance fees'), ('equity', 'Equity value'), ('diluted', 'Diluted shares, millions'), ('price', 'Value per share')):
    put(wsSu, r, 2, lab, f_bold if k in ('equity', 'price') else f_norm); put(wsSu, r, 4, f"=Valuation!C{VALROW[('tgt', k)]}" + ('*-1' if k == 'fees' else ''), f_bold if k in ('equity', 'price') else f_norm, USD if k == 'price' else (NUM3 if k == 'diluted' else NUM)); r += 1
r += 1; band(wsSu, r, 'Headline financials, both products approved (USD m)', 2, 11); r += 1
HY = [2026, 2027, 2028, 2029, 2030, 2031, 2033, 2035, 2038]
put(wsSu, r, 2, 'Fiscal year', f_bold); [put(wsSu, r, 3 + i, f'{y}E' if y > 2026 else '2026A/E', f_hdr, fill=fill_hdr, align=Alignment(horizontal='right')) for i, y in enumerate(HY)]; HYR = r; r += 1
for key, lab, fmt in (('ir_rev', 'IR revenue', NUM), ('xr_rev', 'XR revenue', NUM), ('revenue', 'Total revenue', NUM), ('ebit', 'Operating income', NUM), ('ni', 'Net income', NUM), ('fcf', 'Free cash flow', NUM), ('fund', 'Net equity funding required', NUM), ('cash', 'Cash, year end', NUM), ('shares', 'Shares outstanding, millions', NUM2), ('xr_active', 'XR active patients, year end', NUM0), ('ir_active', 'IR active members, year end', NUM0)):
    put(wsSu, r, 2, lab, f_bold if key in ('revenue', 'fcf') else f_norm)
    for i, y in enumerate(HY): put(wsSu, r, 3 + i, f"=Annual!{CL(YC0+YEARS.index(y))}{AN[key]}", f_bold if key in ('revenue', 'fcf') else f_norm, fmt)
    r += 1
r += 1; band(wsSu, r, 'The three assumptions that carry the short', 2, 11); r += 1
for i, h in enumerate(['Assumption', 'Working', 'Range tested', '12-month value across the range', 'Why it is set here']): put(wsSu, r, 2 + i, h, f_hdr, fill=fill_hdr)
for c_ in range(7, 12): wsSu.cell(r, c_).fill = fill_hdr
r += 1
for k, rng_ in (('ir_quantity_multiplier', 'fixed in every case; relaxing it adds about $%.2f (Street Bridge)' % (bridge[0][2] - bridge[1][2])), ('category_switch', None), ('switch_mature_multiplier', None), ('ir_retention', None)):
    put(wsSu, r, 2, LABELS[k], f_bold, fill=fill_key); put(wsSu, r, 3, f"={'pop_growth' if k == 'growth' else k}", f_norm, NUM3)
    tests = sorted([t for n, t in sm.items() if n.startswith(k + '=')], key=lambda t: list(t['overrides'].values())[0])
    if tests: put(wsSu, r, 4, f"{list(tests[0]['overrides'].values())[0]} to {list(tests[-1]['overrides'].values())[0]}", f_norm); put(wsSu, r, 5, f"${tests[0]['expected_target']:.2f} to ${tests[-1]['expected_target']:.2f}", f_norm)
    else: put(wsSu, r, 4, rng_, f_sub)
    put(wsSu, r, 6, RATIONALE[k], f_sub); wsSu.cell(r, 6).alignment = Alignment(wrap_text=True, vertical='top'); wsSu.merge_cells(start_row=r, start_column=6, end_row=r, end_column=11); wsSu.row_dimensions[r].height = 44; r += 1
r += 1; band(wsSu, r, 'Monte Carlo on the commercial inputs (saved run, 4,096 draws)', 2, 11); r += 1
put(wsSu, r, 2, '12-month value: mean / median / P10 / P90'); put(wsSu, r, 3, pr['expected_target']['mean'], f_input, USD); put(wsSu, r, 4, pr['expected_target']['median'], f_input, USD); put(wsSu, r, 5, pr['expected_target']['p10'], f_input, USD); put(wsSu, r, 6, pr['expected_target']['p90'], f_input, USD); put(wsSu, r, 7, 'analyst-prior range, not a confidence interval', f_sub); r += 1
for lab, key in (('Peak revenue, USD m: mean / median / P10 / P90', 'both_success_peak_usdm'), ('2030 IR revenue, USD m', 'both_success_2030_IR_usdm'), ('2030 XR revenue, USD m', 'both_success_2030_XR_usdm')):
    x = pr[key]; put(wsSu, r, 2, lab); [put(wsSu, r, 3 + i, x[k], f_input, NUM) for i, k in enumerate(('mean', 'median', 'p10', 'p90'))]; r += 1
put(wsSu, r, 2, 'Share of draws below the share price'); put(wsSu, r, 3, sum(1 for row_ in rows_ if row_[ci['expected_target']] < PRICE) / N, f_input, PCT); put(wsSu, r, 4, f'at ${PRICE:.2f}; a sensitivity statistic under these priors, not a probability the short works', f_sub); r += 2
wsSu.print_area = f'B1:K{r}'
rc = BarChart(); rc.type = 'col'; rc.grouping = 'stacked'; rc.overlap = 100; rc.title = 'Revenue by product, USD m'; rc.height = 7.5; rc.width = 16; rc.style = 2
rc.add_data(Reference(wsSu, min_col=2, max_col=2 + len(HY), min_row=HYR + 1, max_row=HYR + 2), from_rows=True, titles_from_data=True); rc.set_categories(Reference(wsSu, min_col=3, max_col=2 + len(HY), min_row=HYR)); wsSu.add_chart(rc, f'B{r}'); wsSu.print_area = f'B1:K{r+18}'
widths(wsSu, {'A': 2, 'B': 46, 'C': 11, 'D': 11, 'E': 11, 'F': 11, 'G': 11, 'H': 11, 'I': 11, 'J': 11, 'K': 11})

# ================================================================== COVER
wsCo = sheet('Cover', '002060'); 
put(wsCo, 3, 2, 'Pharvaris N.V. (NASDAQ: PHVS)', Font(name='Arial', size=20, bold=True, color=NAVY))
put(wsCo, 4, 2, 'Point72 Academy Case Competition, October 2026', Font(name='Arial', size=12, color='595959'))
r = 6
for lab, v, fmt in (('Team', TEAM, GEN), ('Team members', 'Kendall Wu and team (fill in)', GEN), ('Ticker', 'NASDAQ: PHVS', GEN), ('Recommendation', '=IF(implied_return<0,"SHORT","LONG")', GEN), ('Share price', '=share_price', USD), ('Price date', '=price_date', 'd mmm yyyy'), ('12-month price target', '=price_target', USD), ('Implied return', '=implied_return', PCT), ('Valuation date', '=valuation_date', 'd mmm yyyy'), ('Target date', '=target_date', 'd mmm yyyy'), ('Case selected', '=case_selector', GEN)):
    put(wsCo, r, 2, lab, f_bold); put(wsCo, r, 4, v, f_input if lab in ('Team', 'Team members', 'Ticker') else f_bold, fmt); r += 1
wsCo.cell(7, 4).fill = fill_in
r += 1; band(wsCo, r, 'Contents', 2, 8); r += 1
TOC = [('Summary', 'Recommendation, branch values, equity bridge, headline financials, thesis assumptions'), ('Valuation', 'Dated DCF by clinical branch, option dilution, probability weighting'), ('Street Bridge', 'Sell-side targets, Street revenue path against ours, value of the three thesis assumptions'), ('Annual', 'Operating drivers and the annual P&L, cash flow and balances through 2055'), ('Three Statements', 'Income statement, balance sheet and cash flow, reported and forecast'), ('Quarterly', '2026 H2 to 2029 by quarter'), ('Adoption Build', 'Patient pools, Ekterly stock, prescription propensities, retention, annual adoption by origin'), ('Quantity & Price', 'Dose budgets, net prices, ex-US, annual units'), ('Sensitivities', 'One-variable ranges, structural stresses, access packages, cases'), ('Monte Carlo', 'Saved 4,096-draw simulation: distribution, drivers, branches, priors'), ('Checks', 'Live workbook against the saved engine; accounting and population identities'), ('Inputs', 'Every assumption with units, case values, override, bounds, source and rationale'), ('Sources', 'Source register with links'), ('History', 'Reported EUR financials and the opening balance'), ('LOE', 'Post-generic price and quantity indices'), ('Calendar', 'Monthly timing flags'), ('XR Engine / IR Engine', 'Monthly patient ledgers by origin cell'), ('Revenue', 'Monthly units, prices and revenue, with and without XR overlap'), ('Fin Monthly', 'Monthly financials for the four clinical branches'), ('Alt Data ->', 'Reference datasets: launches, benchmarks, Street, clinical, erosion, payers, population')]
for nm_, desc in TOC:
    c = put(wsCo, r, 2, nm_, f_bold); sh = nm_.split(' / ')[0]
    if sh in wb.sheetnames: c.hyperlink = f"#'{sh}'!A1"; c.font = Font(name='Arial', size=10, bold=True, color='0563C1', underline='single')
    put(wsCo, r, 4, desc); r += 1
r += 1; band(wsCo, r, 'How to read the workbook', 2, 8); r += 1
for lab, desc, font, fill in (('Blue text', 'Hardcoded input or saved result', f_input, None), ('Black text', 'Formula', f_norm, None), ('Green text', 'Link to another sheet', f_link, None), ('Yellow cell', 'Case selector and override column on Inputs', f_norm, fill_in), ('Orange label', 'Assumption that carries the thesis', f_bold, fill_key)):
    c = put(wsCo, r, 2, lab, font); c.fill = fill or PatternFill(); put(wsCo, r, 4, desc); r += 1
r += 1
for line in ('Pick Working, Cautious or Upside on the Inputs tab, or type an override next to any assumption; every forecast, valuation and check recalculates.',
             'Forecast statements show the both-approved branch. The Valuation tab runs all four clinical branches and weights them once; cash is never multiplied by a probability.',
             'Monte Carlo, sensitivity and Street-bridge tables are saved results from the Python engine in the research repository (commit e7594c4, 9 October 2026); the Monte Carlo tab flags when live inputs no longer match that run.',
             'Public information only. Comparator list prices and claims benchmarks are anchors; net prices, switching and continuation are underwriting assumptions documented on Inputs.'):
    put(wsCo, r, 2, line, f_sub); r += 1
widths(wsCo, {'A': 2, 'B': 24, 'C': 2, 'D': 90})

# ================================================================== order, print setup, save
ORDER = ['Cover', 'Summary', 'Valuation', 'Street Bridge', 'Annual', 'Three Statements', 'Quarterly', 'Adoption Build', 'Quantity & Price', 'Sensitivities', 'Monte Carlo', 'Checks', 'Inputs', 'Sources', 'History', 'LOE', 'Calendar', 'XR Engine', 'IR Engine', 'Revenue', 'Fin Monthly', 'Alt Data ->'] + [t for t, _ in ALT_TABS]
def unify_bands(ws):
    navy = lambda c: c.fill is not None and c.fill.fill_type == 'solid' and str(c.fill.fgColor.rgb).endswith(NAVY)
    rows = {}
    for rr_ in range(1, ws.max_row + 1):
        if navy(ws.cell(rr_, 2)): rows[rr_] = max(c_ for c_ in range(2, ws.max_column + 1) if navy(ws.cell(rr_, c_)))
    for rr_, ext in list(rows.items()):
        if rr_ + 1 in rows:
            ext2 = max(ext, rows[rr_ + 1])
            for rr2 in (rr_, rr_ + 1):
                for c_ in range(2, ext2 + 1):
                    cell = ws.cell(rr2, c_); cell.fill = fill_hdr; cell.font = f_hdr
for ws_ in wb.worksheets: unify_bands(ws_)
wb._sheets = [wb[n] for n in ORDER]
for n in ('Calendar', 'XR Engine', 'IR Engine', 'Revenue', 'Fin Monthly'):
    ws_ = wb[n]; ws_.page_setup.fitToWidth = 1; ws_.page_setup.fitToHeight = 0; ws_.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True); ws_.print_title_rows = '1:6'; ws_.print_area = f'A1:{mcol(23)}{ws_.max_row}'
wb['Annual'].print_area = f"B4:{CL(YC0+YEARS.index(2031))}{AN['nol']}"
wb.active = 0
wb.save(OUT); print('saved', OUT)
_maps = {'RX': {str(k): v for k, v in RX.items()}, 'RI': {str(k): v for k, v in RI.items()}, 'RV': {str(k): v for k, v in RV.items()}, 'RF': {str(k): v for k, v in RF.items()}, 'CAL': CAL, 'VALROW': {str(k): v for k, v in VALROW.items()}, 'MC0': MC0, 'NM': NM, 'AN': AN, 'QR': QR, 'TS': TS}
(ROOT / 'models/excel/rowmap.json').write_text(json.dumps(_maps))
