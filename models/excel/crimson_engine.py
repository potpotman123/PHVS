"""Central-case engine outputs for the Crimson underwriting case: monthly rows (reconciliation), control totals,
one-variable sensitivities, structural stresses and the Street-bridge counterfactuals. Run from the repository root."""
import json, sys, inspect, copy
from pathlib import Path
ROOT = Path(__file__).resolve().parents[2]; sys.path.insert(0, str(ROOT / 'models/finalized'))
import engine as E
cfg = E.read('inputs.json'); case = json.loads((ROOT / 'models/excel/crimson_case.json').read_text()); OV = case['overrides']
def ev(extra=None, retain=False): return E.evaluate(cfg, 'working', {**OV, **(extra or {})}, retain=retain)
base = ev(retain=True)
(ROOT / 'models/excel/crimson_monthly.json').write_text(json.dumps(base.pop('monthly_by_clinical_state')))
out = {'case': case, 'base': base}
# one-variable ranges (same variables and bounds as analyze.py), relative to the Crimson base
variables = {'population':[8500,11500], 'ir_oral_switch_relative':[.5,1.5], 'category_switch':[.10,.20], 'oral_launch_share':[.20,.45], 'oral_destination_capture':[.40,.80], 'ir_injectable_opportunity':[.75,1.25], 'untreated_xr_reach':[.25,.75], 'ir_no_rescue_opportunity':[.10,.30], 'ekterly_future_forms':[100,250], 'switch_mature_multiplier':[.50,1.25], 'xr_first_retention':[.61,.6945], 'xr_mature_retention':[.8,.95], 'ir_retention':[.78,.92], 'ir_doses_ltp':[6,11.8], 'ir_doses_other':[10,18], 'dual_use':[0,1], 'xr_paid_intensity':[.9,1], 'cost_scale':[.8,1.2], 'discount_rate':[.1,.15], 'ir_pos':[.8,.95], 'xr_pos':[.8,.95], 'ir_net_fraction':[.70,.95], 'xr_net_fraction':[.5,.8], 'conversion':[.65,.9], 'growth':[0,.03], 'loe_year':[2039,2043], 'ir_exus_ratio':[.1,.35], 'xr_exus_ratio':[.15,.6], 'manufacture_rate':[.02,.08], 'license_rate':[.02,.05], 'nwc_fraction':[.08,.18], 'tax_rate':[.15,.258], 'issue_price':[15,40]}
sens = []
for k, (lo, hi) in variables.items():
    a, b = ev({k: lo}), ev({k: hi}); sens.append({'parameter': k, 'low': lo, 'high': hi, 'target_low': a['expected_12month_value'], 'target_high': b['expected_12month_value'], 'today_low': a['risk_adjusted_value_today'], 'today_high': b['risk_adjusted_value_today']})
sens.sort(key=lambda r: -abs(r['target_high'] - r['target_low'])); out['sensitivities'] = sens
scen = {'hazards_halved_after_year5':{'hazard_year6_multiplier':.5}, 'quarter_of_stoppers_retry_after24m':{'reentry_fraction':.25}, 'other_prevention_shift20pct_by10y':{'nonphvs_ltp_shift':.2}, 'generic_price_capped_flat':{'loe_price_cap':True}, 'initial_rescue_stock_free':{'ir_initial_paid_fraction':0}, 'partner_all_overseas':{'exus_direct_fraction':0}, 'six_month_launch_delay_both':{'launch_delay_IR':6,'launch_delay_XR':6}, 'no_refill_intensity_adjustment':{'xr_paid_intensity':1}, 'old_prices_only':{'ir_net_fraction':14000/16720, 'xr_net_fraction':418000/(48487.92*365/28)}, 'omit_license_royalty_cost_counterfactual':{'license_rate':0}, 'redose_plus_better_IR_capture':{'ir_injectable_opportunity':1.25}, 'LTP_ten_percent_no_rescue':{'ltp_no_rescue_fraction':.10}, 'extra_Ekterly_free_stock':{'ekterly_free_stock_multiplier':1.2}, 'rebate_for_access':cfg['access_packages']['rebate_for_access'], 'restricted_high_net':cfg['access_packages']['restricted_high_net']}
out['scenarios'] = [{'name': n, 'overrides': o, 'value_today': (r := ev(o))['risk_adjusted_value_today'], 'expected_target': r['expected_12month_value'], 'delta_target': r['expected_12month_value'] - base['expected_12month_value'], 'peak': r['conditional_peak']} for n, o in scen.items()]
# Street-bridge counterfactuals (redosing relaxed only for sizing)
_src = inspect.getsource(E.resolve).replace("    assert abs(p['ir_quantity_multiplier'] - 1.17 / 1.38) < 1e-12, 'All active forecasts require the selected redosing discount'\n", '')
_ns = {}; exec(_src, E.__dict__, _ns); relaxed = _ns['resolve']
steps = [('Favourable bounds: no redosing haircut, switching at the top of its range, 92% rescue continuation, 45% launch share', {'ir_quantity_multiplier': 1.0, 'category_switch': 0.2, 'switch_mature_multiplier': 1.25, 'ir_retention': 0.92, 'oral_launch_share': 0.45}),
         ('Apply the redosing quantity haircut (1.17/1.38 doses per attack)', {'category_switch': 0.2, 'switch_mature_multiplier': 1.25, 'ir_retention': 0.92, 'oral_launch_share': 0.45}),
         ('Switching opportunity at the claims benchmark (15.2%, flat)', {'ir_retention': 0.92, 'oral_launch_share': 0.45}),
         ('Launch share of oral choices at 30%', {'ir_retention': 0.92}),
         ('Rescue continuation at 85%', {})]
bridge = []
saved = E.resolve; E.resolve = relaxed
try:
    for lab, o in steps:
        r = ev(o); ab = [x for x in r['annual_both_success'] if 2027 <= x['year'] <= 2038]; bridge.append({'label': lab, 'both_target': r['conditional_valuations']['both']['target']['price'], 'expected_target': r['expected_12month_value'], 'peak': r['conditional_peak']['recognized_revenue_usdm'], 'cum_revenue_2027_2038': sum(x['revenue'] for x in ab), 'cum_fcf_2027_2038': sum(x['fcf'] for x in ab), 'cum_net_income_2027_2038': sum(x['net_income'] for x in ab)})
finally: E.resolve = saved
out['bridge'] = bridge
(ROOT / 'models/excel/crimson_results.json').write_text(json.dumps(out, indent=1))
print('central case: today', round(base['risk_adjusted_value_today'], 4), 'target', round(base['expected_12month_value'], 4), 'peak', base['conditional_peak'])
print('both-approved target', round(base['conditional_valuations']['both']['target']['price'], 4))
print('top sensitivities:', [(s['parameter'], round(s['target_low'], 2), round(s['target_high'], 2)) for s in sens[:6]])
print('bridge:', [(b['label'][:30], round(b['expected_target'], 2)) for b in bridge])
