"""Deterministic sensitivities and portable audit outputs, not Monte Carlo draws."""
import copy
import json
from engine import BASE, read, write, evaluate, resolve, aggregate, month_index, clinical_weights

cfg = read('inputs.json')
base = read('results.json')['working']
variables = {
    'population':[8500,11500], 'ir_oral_switch_relative':[.5,1.5], 'category_switch':[.10,.20],
    'oral_launch_share':[.25,.60],
    'oral_destination_capture':[.40,.80], 'ir_injectable_opportunity':[.75,1.25],
    'untreated_xr_reach':[.25,.75], 'ir_no_rescue_opportunity':[.10,.30],
    'ekterly_future_forms':[100,250], 'switch_mature_multiplier':[.50,1.25],
    'xr_first_retention':[.61,.6945], 'xr_mature_retention':[.8,.95],
    'ir_retention':[.8,.95], 'ir_doses_ltp':[6,11.8], 'ir_doses_other':[10,18],
    'dual_use':[0,1], 'xr_paid_intensity':[.9,1], 'cost_scale':[.8,1.2],
    'discount_rate':[.1,.15], 'ir_pos':[.8,.95], 'xr_pos':[.8,.95],
    'ir_net_fraction':[.70,.95], 'xr_net_fraction':[.5,.8],
    'conversion':[.65,.9], 'growth':[0,.03], 'loe_year':[2039,2043],
    'ir_exus_ratio':[.1,.35], 'xr_exus_ratio':[.15,.6],
    'manufacture_rate':[.02,.08], 'license_rate':[.02,.05],
    'nwc_fraction':[.08,.18], 'tax_rate':[.15,.258], 'issue_price':[15,40]
}
scenarios = {k+'='+str(v): {k:v} for k, values in variables.items() for v in values}
scenarios.update({
    'hazards_halved_after_year5':{'hazard_year6_multiplier':.5},
    'quarter_of_stoppers_retry_after24m':{'reentry_fraction':.25},
    'other_prevention_shift20pct_by10y':{'nonphvs_ltp_shift':.2},
    'generic_price_capped_flat':{'loe_price_cap':True},
    'initial_rescue_stock_free':{'ir_initial_paid_fraction':0},
    'partner_all_overseas':{'exus_direct_fraction':0},
    'six_month_launch_delay_both':{'launch_delay_IR':6,'launch_delay_XR':6},
    'no_refill_intensity_adjustment':{'xr_paid_intensity':1},
    'old_prices_only':{'ir_net_fraction':14000/16720, 'xr_net_fraction':418000/(48487.92*365/28)},
    'omit_license_royalty_cost_counterfactual':{'license_rate':0},
    'redose_plus_better_IR_capture':{'ir_quantity_multiplier':1.17/1.38,'ir_injectable_opportunity':1.25},
    'LTP_ten_percent_no_rescue':{'ltp_no_rescue_fraction':.10},
    'extra_Ekterly_free_stock':{'ekterly_free_stock_multiplier':1.2},
    'rebate_for_access':cfg['access_packages']['rebate_for_access'],
    'restricted_high_net':cfg['access_packages']['restricted_high_net'],
})
results = []
for name, overrides in scenarios.items():
    r = evaluate(cfg, overrides=overrides)
    results.append(dict(name=name, overrides=overrides, value_today=r['risk_adjusted_value_today'],
                        expected_target=r['expected_12month_value'], delta_target=r['expected_12month_value']-base['expected_12month_value'], peak=r['conditional_peak']))
write('sensitivities.json', {'base_target':base['expected_12month_value'],'range_type':'Underwriting and structural stresses; not confidence intervals. Price/access one-variable runs isolate arithmetic; use paired packages for coherent cases.','results':results})
rank = []
for k in variables:
    tests = [r for r in results if r['name'].startswith(k+'=')]
    rank.append({'parameter':k,'target_low':min(r['expected_target'] for r in tests),'target_high':max(r['expected_target'] for r in tests),'span':max(r['expected_target'] for r in tests)-min(r['expected_target'] for r in tests)})
rank.sort(key=lambda r:-r['span'])
write('driver_ranking.json',rank)
# Alternative dependence structures retain the same clinical marginal PoS.
weights={m:clinical_weights(resolve(cfg),m) for m in ['shared_gate','independent','perfect_positive']}
write('clinical_dependence.json',{m:{'weights':w,'target':sum(w[s]*base['conditional_valuations'][s]['target']['price'] for s in w)} for m,w in weights.items()})
# Calendar tables: native history separate; 2026 H1 actual translated only for a
# clearly flagged planning bridge. No claim these are reported USD financials.
all_results=read('results.json'); quarters=read('working_quarterly.json')
history=read('history.json'); fx=resolve(cfg)['fx']; reports=[]
for year in [2024,2025]:
    reports.append({'period':str(year),'status':'reported','currency':'EURm','source':'F25',**history['periods'][str(year)]})
h1=history['periods']['H1_2026']; h2=copy.deepcopy(base['annual_both_success'][0]); f2026={'period':'2026','status':'H1 actual plus H2 forecast, constant-FX USD planning view','currency':'USDm'}
for k in ['revenue','rd','ga','net_income','capex']:
    f2026[k]=h1[k]*fx+h2[k]
f2026.update(commercial=h2['commercial'],cash_end=h2['cash_end'],shares_end=h2['shares_end'],cfo=h1['cfo']*fx+h2['net_income']+h2['depreciation']-h2['delta_nwc'],funding_net=h2['funding_net'],warning='Historical finance income/FX and noncash SBC remain in H1 actual; H2 economic-compensation convention differs. See history.json for full cash bridge.')
reports.append(f2026)
for r in base['annual_both_success'][1:]:
    reports.append({'period':str(r['year']),'status':'conditional on both approvals','currency':'USDm',**r})
write('financial_tables.json',{'history_and_forecast':reports,'quarterly_2026H2_to_2028':[r for r in quarters if str(r['quarter'])[:4] in ['2026','2027','2028']],'forecast_warning':'Working operating statements conditional on both products succeeding; valuation separately branches four clinical outcomes.'})
monthly=read('working_monthly.json')['both']; launch=month_index(cfg['launches']['IR']['commercial_month'])
window=[r for r in monthly if launch<=month_index(r['month'])<launch+8]
launch_checks={'ir_first8months':{'forms':sum(r['ir_rx'] for r in window),'paid_access_starts':sum(r['ir_access_starts'] for r in window),'US_revenue_m':sum(r['IR_US'] for r in window),'Ekterly_forms':1702,'forms_ratio':sum(r['ir_rx'] for r in window)/1702},'comparability':'Segmented origins with projected Ekterly launch stock and rising PHVS oral choice; ratio is an output, not an imposed first-mover discount. Forms include attempts, not proven unique people. Paid starts lag. No per-patient revenue inferred from cumulativeforms.'}
launch_checks['IR_full_launch_years']=[]
for i in range(4):
    window=[r for r in monthly if launch+i*12<=month_index(r['month'])<launch+(i+1)*12]
    launch_checks['IR_full_launch_years'].append({'launch_year':i+1,'US_revenue_m':sum(r['IR_US'] for r in window),'forms':sum(r['ir_rx'] for r in window),'paid_doses':sum(r['IR_units'] for r in window)})
write('launch_checks.json',launch_checks)
lines=['PHVS INDEPENDENT UNDERWRITING — USDm unless per-share; 2026-10-09','Conditional revenue; risk-adjusted standalone valuation; no takeover premium','Case | Today/share | Expected12m/share | Conditional peak | Peak year']
for k,r in all_results.items():
    lines.append(f"{k} | {r['risk_adjusted_value_today']:.2f} | {r['expected_12month_value']:.2f} | {r['conditional_peak']['recognized_revenue_usdm']:.1f} | {r['conditional_peak']['year']}")
lines+=['','Working conditional calendar financials','Year | IR | XR | Revenue | R&D | Commercial | G&A | FCF | Net new funding']
for r in base['annual_both_success']:
    if r['year']<=2038:
        lines.append(f"{r['year']} | {r['IR_recognized']:.1f} | {r['XR_recognized']:.1f} | {r['revenue']:.1f} | {r['rd']:.1f} | {r['commercial']:.1f} | {r['ga']:.1f} | {r['fcf']:.1f} | {r['funding_net']:.1f}")
lines+=['2026 line above is H2 only; full-year planning bridge in financial_tables.json.','','Largest one-variable target ranges (chosen bounds, not statistical importance):']
for r in rank[:10]:lines.append(f"{r['parameter']}: ${r['target_low']:.2f}–${r['target_high']:.2f}")
(BASE/'summary.txt').write_text('\n'.join(lines)+'\n')
print('\n'.join(lines[:6]+lines[-11:]))
