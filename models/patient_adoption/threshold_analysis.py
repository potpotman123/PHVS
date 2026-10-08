"""Invert the saved annual revenue hurdle into hypothetical XR cohorts and IR units.

No valuation re-solve and no inferred population estimates. All scenarios are
conditional; pool exhaustion is flagged instead of silently creating patients.
"""
import hashlib
import json
import math
from copy import deepcopy
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
HERE = Path(__file__).resolve().parent


def integrated_survival(age, p):
    """Exact expected treated months over [age,age+1] for monthly piecewise hazards."""
    r = p['first_year_retention']
    w = p['first_three_months_log_hazard_share']
    h1 = -math.log(r) * w / 3
    h2 = -math.log(r) * (1-w) / 9
    h3 = -math.log(p['mature_annual_retention']) / 12
    cumulative = h1*min(age, 3) + h2*min(max(age-3, 0), 9) + h3*max(age-12, 0)
    hazard = h1 if age < 3 else h2 if age < 12 else h3
    return math.exp(-cumulative) * (-math.expm1(-hazard)/hazard if hazard else 1)


def survival(age, p):
    if age <= 3:
        return p['first_year_retention'] ** (p['first_three_months_log_hazard_share'] * age/3)
    if age <= 12:
        return p['first_year_retention'] ** (p['first_three_months_log_hazard_share'] +
                  (1-p['first_three_months_log_hazard_share'])*(age-3)/9)
    return p['first_year_retention'] * p['mature_annual_retention']**((age-12)/12)


def exposure(cohorts, first_month, last_month, p):
    return sum(size * integrated_survival(t-start, p)/12
               for start, size in cohorts
               for t in range(max(start, first_month), last_month+1))


def solve_xr(p, targets):
    import numpy as np
    from scipy.optimize import linprog, minimize
    assert 0 < p['prescription_to_paid_probability'] <= 1
    assert 0 < p['first_year_retention'] <= 1
    assert 0 < p['mature_annual_retention'] <= 1
    assert p['paid_delay_months'] >= 0
    assert 0 <= p['first_three_months_log_hazard_share'] <= 1
    targets=[(y,v) for y,v in targets if y >= p['xr_launch_year']]
    final_month=12*(targets[-1][0]-p['xr_launch_year'])+11
    months=list(range(p['xr_first_prescription_month']-1, final_month+1))
    A=np.array([[exposure([(t+p['paid_delay_months'],p['prescription_to_paid_probability'])],
                         12*(year-p['xr_launch_year']),12*(year-p['xr_launch_year'])+11,p)
                 for t in months] for year,_ in targets])
    b=np.array([v*1e6*p['us_revenue_fraction']/p['xr_annual_net_price_usd'] for _,v in targets])
    # Mathematical minimum first prescriptions subject to every annual revenue target.
    low=linprog(np.ones(len(months)), A_eq=A, b_eq=b, bounds=(0,None), method='highs')
    if not low.success: raise ValueError('Annual revenue path has no nonnegative cohort allocation: '+low.message)
    # Smooth monthly allocation, not a statistically estimated adoption curve.
    # Minimize squared adjacent-month changes; no artificial annual reset.
    scale=max(float(low.x.mean()),1)
    def roughness_gradient(z):
        d=np.diff(z);g=np.zeros_like(z);g[:-1]-=d;g[1:]+=d
        return g
    opt=minimize(lambda z: .5*np.sum(np.diff(z)**2), low.x/scale,
                 jac=roughness_gradient,
                 constraints=[dict(type='eq',fun=lambda z:(A*z).sum(axis=1)-b/scale,jac=lambda z:A)],
                 bounds=[(0,None)]*len(months),method='SLSQP',
                 options={'ftol':1e-10,'maxiter':500})
    if not opt.success: raise ValueError('Smooth allocation failed: '+opt.message)
    x=np.maximum(opt.x*scale,0)
    if max(abs((A*x).sum(axis=1)-b)) > 1e-6: raise ValueError('Allocation did not reconcile annual patient-years')
    cohorts=[(t+p['paid_delay_months'],float(n)*p['prescription_to_paid_probability']) for t,n in zip(months,x)]
    rows=[];cumulative_existing=cumulative_rx=0;first_exhaustion=None
    for year,worldwide in targets:
        first_month=12*(year-p['xr_launch_year']);last_month=first_month+11
        year_attempts=[float(n) for t,n in zip(months,x) if first_month<=t<=last_month]
        total_rx=sum(year_attempts)
        new_dx_rx=sum(min(n,p['annual_new_diagnoses']*p['new_diagnosis_prescription_capture']/12) for n in year_attempts)
        existing_rx=total_rx-new_dx_rx
        available_existing_start=p['existing_eligible_pool']-cumulative_existing
        cumulative_existing+=existing_rx;cumulative_rx+=total_rx
        if first_exhaustion is None and cumulative_existing>p['existing_eligible_pool']+1e-8:first_exhaustion=year
        target_revenue=worldwide*1e6*p['us_revenue_fraction']
        realized_py=exposure(cohorts,first_month,last_month,p)
        end_active=sum(size*survival(last_month+1-start,p) for start,size in cohorts if start<=last_month)
        rows.append(dict(year=year,worldwide_target_USDm=worldwide,us_target_USDm=target_revenue/1e6,
                         paid_patient_years_required=target_revenue/p['xr_annual_net_price_usd'],modeled_paid_patient_years=realized_py,
                         revenue_gap_USD=realized_py*p['xr_annual_net_price_usd']-target_revenue,
                         new_unique_prescriptions=total_rx,new_diagnosis_prescriptions=new_dx_rx,
                         existing_patient_prescriptions=existing_rx,cumulative_existing_prescriptions=cumulative_existing,
                         cumulative_all_prescriptions=cumulative_rx,remaining_existing_never_attempted=p['existing_eligible_pool']-cumulative_existing,
                         annual_existing_attempt_fraction=existing_rx/available_existing_start if available_existing_start>0 else None,
                         year_end_active_paid=end_active))
    return dict(inputs=p,annual=rows,monthly_prescriptions=[dict(month=t,first_prescriptions=float(n)) for t,n in zip(months,x)],
                first_pool_exhaustion_year=first_exhaustion,required_initial_existing_pool=cumulative_existing,
                required_share_of_assumed_existing_pool=cumulative_existing/p['existing_eligible_pool'],
                all_annual_targets_matched=all(abs(z['revenue_gap_USD'])<1 for z in rows),
                absolute_minimum_all_prescriptions=float(low.fun),smooth_all_prescriptions=float(x.sum()),
                allocation_method='Minimum squared month-to-month prescription changes subject to nonnegative cohorts and exact annual revenue. Also report linear-program minimum cumulative all-source prescriptions.',
                note='Unconstrained recruitment requirement. Negative remaining pool means infeasible under unique-attempt/no-return assumptions, not actual negative patients.')


def main():
    config=json.loads((HERE/'threshold_inputs.json').read_text())
    sourcepath=ROOT/config['valuation_source']
    val=json.loads(sourcepath.read_text())['base']
    targets=[(config['first_source_year']+i,s[1]) for i,s in enumerate(val['sales'])
             if config['first_source_year']+i <= config['last_preplateau_year']]
    ir_end=val['sales'][config['last_preplateau_year']-config['first_source_year']][0]
    base=config['illustrative_reference']; grid=config['sensitivity_values']
    ref=solve_xr(deepcopy(base),targets)
    retention_access=[]
    for retain in grid['first_year_retention']:
        for convert in grid['prescription_to_paid_probability']:
            p={**base,'first_year_retention':retain,'prescription_to_paid_probability':convert}
            retention_access.append(solve_xr(p,targets))
    price_geography=[]
    for price in grid['xr_annual_net_price_usd']:
        for us in grid['us_revenue_fraction']:
            price_geography.append(solve_xr({**base,'xr_annual_net_price_usd':price,'us_revenue_fraction':us},targets))
    variants=[]
    for key in ['mature_annual_retention','paid_delay_months','xr_first_prescription_month','existing_eligible_pool']:
        for value in grid[key]:
            if value != base[key]:
                out=solve_xr({**base,key:value},targets);out['changed_parameter']=key;variants.append(out)
    # Front-loaded versus constant year1 hazard, same survival checkpoints.
    out=solve_xr({**base,'first_three_months_log_hazard_share':.25},targets)
    out['changed_parameter']='constant_year1_hazard';variants.append(out)
    # Source-flow sensitivity: remove all captured new diagnoses.
    out=solve_xr({**base,'new_diagnosis_prescription_capture':0},targets)
    out['changed_parameter']='no_new_diagnosis_capture';variants.append(out)
    joint=[]
    for name,changes in [
        ('stacked_stress_not_probability_weighted',dict(us_revenue_fraction=.9,xr_annual_net_price_usd=300000,
          prescription_to_paid_probability=.6,first_year_retention=.61,mature_annual_retention=.9,existing_eligible_pool=4000)),
        ('discount_with_better_access',dict(xr_annual_net_price_usd=300000,prescription_to_paid_probability=.95))]:
        out=solve_xr({**base,**changes},targets);out['scenario_name']=name;joint.append(out)
    ir=[]
    for us in grid['us_revenue_fraction']:
        for spend in grid['ir_annual_net_revenue_per_paid_user_usd']:
            ir.append(dict(us_fraction=us,annual_net_spend_USD=spend,required_annual_paid_users=ir_end*1e6*us/spend))
    ir_units=[]
    for price in grid['ir_net_revenue_per_paid_pack_usd']:
        for packs in grid['ir_paid_packs_per_user_year']:
            ir_units.append(dict(net_pack_USD=price,paid_packs_per_user_year=packs,
                                 required_paid_packs=ir_end*1e6*base['us_revenue_fraction']/price,
                                 required_annual_paid_users=ir_end*1e6*base['us_revenue_fraction']/price/packs))
    results=dict(status=config['status'],valuation_snapshot=config['valuation_snapshot'],
                 source_file_sha256=hashlib.sha256(sourcepath.read_bytes()).hexdigest(),
                 worldwide_2036_USDm={'XR':targets[-1][1],'IR':ir_end},
                 reference=ref,retention_access=retention_access,price_geography=price_geography,
                 one_at_a_time=variants,joint_scenarios=joint,IR_spend_thresholds=ir,IR_pack_thresholds=ir_units,
                 IR_spend_required_by_user_count=[dict(paid_users=n,annual_net_spend_USD=ir_end*1e6*base['us_revenue_fraction']/n)
                                                for n in grid['ir_paid_user_counts']],
                 interpretation='Matching saved revenue is not a new DCF solve. No replacement ramp, national TAM, scenario probability, post2036 patient path or live valuation is estimated.')
    (HERE/'threshold_results.json').write_text(json.dumps(results,indent=2)+'\n')
    report(results)
    charts(results)
    print(json.dumps({'required_existing_pool':ref['required_initial_existing_pool'],
                      'peak_paid_patient_years':ref['annual'][-1]['paid_patient_years_required'],
                      'cumulative_prescriptions':ref['annual'][-1]['cumulative_all_prescriptions'],
                      'exhaustion':ref['first_pool_exhaustion_year']}))


def report(res):
    ref=res['reference'];p=ref['inputs'];a=ref['annual'][-1]
    lines=['PHVS conditional adoption thresholds — 8 October 2026',
    'Hypothetical scenarios; no calibrated central forecast or recommended position.','',
    'HURDLE AND METHOD',
    f"Saved $31.83 scenario:2036 worldwide success-case XR ${res['worldwide_2036_USDm']['XR']:.1f}m; IR ${res['worldwide_2036_USDm']['IR']:.1f}m.",
    'This is the latest50/50 post-generic case, not the superseded Kuvan-only scenario.',
    '90% approval probabilities are already in valuation. Do not probability-adjust patient requirements again.',
    'No current stock-price refresh. The inherited shape and product mix remain assumptions.',
    '', 'ILLUSTRATIVE REFERENCE — NOT A BASE CASE',
    '80% US revenue; XR $418k net/paid patient-year (broker assumption);80% conversion;',
    '80% year1 retention;95% annual retention thereafter;1-month access delay;',
    'January2028 first prescriptions;200 new diagnoses/year with50% prescription capture;',
    '6000 initially eligible existing patients. These are scenarios, not observed PHVS inputs.',
    'Monthly first-attempt prescriptions minimize adjacent-month changes subject to exact annual sales.',
    'This is a transparent allocation convention, not an empirically fitted adoption curve.',
    'Paid starts occur at month-start after the delay; exact integrated survival yields patient-years.',
    '60% of first-year log hazard falls in months1–3; this shape is assumed, not calibrated.',
    'Converted users have survival measured from first paid supply. No retries, returns,',
    'subsequent free-to-paid recycling or reauthorization interruptions are modeled.',
    '', 'REFERENCE OUTPUT',
    f"2036 US XR revenue ${a['us_target_USDm']:.1f}m needs {a['paid_patient_years_required']:,.0f} paid patient-years.",
    f"Smooth allocation2028–36 requires {a['cumulative_all_prescriptions']:,.0f} unique first prescriptions, including",
    f"{ref['required_initial_existing_pool']:,.0f} existing-patient attempts ({ref['required_share_of_assumed_existing_pool']:.1%} of assumed6000).",
    f"Absolute mathematical minimum all-source first prescriptions: {ref['absolute_minimum_all_prescriptions']:,.0f}; this permits lumpy recruitment.",
    'Existing-patient attempts include previously diagnosed rescue-only/untreated entrants to prevention,',
    'not just switching from incumbent prophylaxis. No current-regimen allocation is inferred.',
    '', 'YEAR-BY-YEAR XR REQUIREMENT',
    'Year | US sales $m | Paid patient-years | All first Rx | Existing-patient Rx | Remaining existing pool']
    for x in ref['annual']:
        lines.append(f"{x['year']} | {x['us_target_USDm']:.1f} | {x['paid_patient_years_required']:.0f} | {x['new_unique_prescriptions']:.0f} | {x['existing_patient_prescriptions']:.0f} | {x['remaining_existing_never_attempted']:.0f}")
    lines += ['', 'RETENTION x ACCESS: REQUIRED EXISTING-PATIENT FIRST ATTEMPTS THROUGH2036',
              'Year1 retention | Eventual paid conversion | Existing attempts | % of6000 | First exhaustion year']
    for s in res['retention_access']:
        i=s['inputs']; lines.append(f"{i['first_year_retention']:.0%} | {i['prescription_to_paid_probability']:.0%} | {s['required_initial_existing_pool']:.0f} | {s['required_share_of_assumed_existing_pool']:.1%} | {s['first_pool_exhaustion_year'] or 'none'}")
    lines += ['', 'NET PRICE x US MIX:2036 PAID PATIENT-YEARS / CUMULATIVE EXISTING ATTEMPTS']
    for s in res['price_geography']:
        i=s['inputs'];lines.append(f"${i['xr_annual_net_price_usd']:,.0f}; US{i['us_revenue_fraction']:.0%}: {s['annual'][-1]['paid_patient_years_required']:.0f} / {s['required_initial_existing_pool']:.0f}")
    lines += ['', 'ONE-AT-A-TIME SENSITIVITIES: CUMULATIVE EXISTING ATTEMPTS / FIRST EXHAUSTION']
    for s in res['one_at_a_time']:
        key=s['changed_parameter'];value=s['inputs'].get(key,'see inputs')
        lines.append(f"{key}={value}: {s['required_initial_existing_pool']:.0f} / {s['first_pool_exhaustion_year'] or 'none'}")
    lines += ['', 'JOINT SCENARIOS — NOT PROBABILITY ESTIMATES',
              'Stress:90%US, $300knet,60%conversion,61%year1,90%mature,4000existing eligibles.',
              'Discount/access: $300knet with95%conversion; other reference assumptions unchanged.']
    for s in res['joint_scenarios']:
        lines.append(f"{s['scenario_name']}: {s['required_initial_existing_pool']:.0f} existing attempts; {s['required_share_of_assumed_existing_pool']:.1%} of assumed pool; first exhaustion {s['first_pool_exhaustion_year'] or 'none'}")
    lines += ['', 'IR: ANNUAL PAID USERS, NOT RESCUE OWNERS',
              'US share | Net annual sales per paid user |2036 annual paid users required']
    for x in res['IR_spend_thresholds']:
        lines.append(f"{x['us_fraction']:.0%} | ${x['annual_net_spend_USD']:,.0f} | {x['required_annual_paid_users']:,.0f}")
    lines += ['', 'At80%US mix, required net annual revenue per paid user:']
    for x in res['IR_spend_required_by_user_count']:
        lines.append(f"{x['paid_users']:,} users -> ${x['annual_net_spend_USD']:,.0f}")
    lines += ['', 'Equivalent pack grid is saved in threshold_results.json. Pack definition, price and',
    'utilization are arbitrary arithmetic scenarios, not PHVS dosing assumptions.',
    'A paid user here has annual paid dispensing; a person carrying old/free stock may generate',
    'no sales this year. IR users can overlap XR patients; do not sum them as unique HAE patients.',
    'Run inventory/replenishment modeling before interpreting pack counts as attacks.',
    '', 'WHAT THE THRESHOLDS DO AND DO NOT ESTABLISH',
    'A required pool above the assumed eligible stock is infeasible under these no-return rules.',
    'A required pool below it is only arithmetically possible, not evidence of achievable capture.',
    'Compare requirements to a future validated eligible stock, not all diagnosed patients blindly.',
    'The same annual sales can result from different starts, persistence and prices. Revenue alone',
    'cannot identify each. Low conversion can increase cumulative unique recruitment dramatically.',
    'Later launch/delay can concentrate initial prescriptions. Smoothness and recruitment timing',
    'affect cumulative requirements; the linear-program minimum all-source count is a lower bound.',
    'Discounts may improve access; these grids isolate effects, not independent random distributions.',
    'First-year survival61% is an observed analogue in a different population;95% is a management',
    'illustration. Applying either to post-conversion PHVS cohorts is hypothetical.',
    'IR/XR mix is inherited, not independently identified by the stock price: a different mix can',
    'support the same valuation. This pass does not solve that product-allocation frontier.',
    'The prescribed ramp is held fixed through2036. No faster/slower ramp is revalued; plateau and',
    'generic-era patient dynamics remain outside this bridge. Ex-US economics also remain assumed.',
    'No annual spend, paid conversion, pool size or retention is selected as the most likely case.',
    '', 'REPRODUCIBILITY',
    'Inputs: models/patient_adoption/threshold_inputs.json',
    'Source: models/reverse_dcf/verification.json, base.sales (2027 onward)',
    'Engine: models/patient_adoption/threshold_analysis.py',
    'Results: models/patient_adoption/threshold_results.json',
    'Checks: models/patient_adoption/threshold_validation.json',
    'Run threshold_analysis.py, then validate_thresholds.py. Existing DCF files are read only.']
    (ROOT/'output/PHVS_conditional_thresholds_2026-10-08.txt').write_text('\n'.join(lines)+'\n')


def charts(res):
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    fig,ax=plt.subplots(1,2,figsize=(13,5.2),layout='constrained')
    matrix=[[s['required_initial_existing_pool'] for s in res['retention_access'][i:i+3]] for i in [0,3,6]]
    ax[0].imshow(matrix,cmap='YlOrRd',vmin=0,vmax=max(max(row) for row in matrix))
    ax[0].set(xticks=range(3),xticklabels=['60%','80%','95%'],yticks=range(3),yticklabels=['61%','80%','95%'],
              xlabel='Prescription-to-paid conversion',ylabel='First-year retention',title='XR: existing-patient first attempts\nCumulative 2028–2036')
    for i,row in enumerate(matrix):
        for j,n in enumerate(row):ax[0].text(j,i,f'{n:,.0f}\n{n/6000:.0%} of 6,000',ha='center',va='center',fontsize=11,color='white' if n>2200 else 'black')
    xs=[50000,100000,150000,200000]
    for share in [.6,.8,.9]:
        vals=[x['required_annual_paid_users'] for x in res['IR_spend_thresholds'] if x['us_fraction']==share]
        ax[1].plot([x/1000 for x in xs],vals,marker='o',label=f'{share:.0%} US sales')
    ax[1].set(xlabel='Net annual revenue per paid IR user ($000)',ylabel='Annual paid users required',title='IR: paid-user hurdle in 2036')
    ax[1].legend();ax[1].grid(alpha=.2)
    fig.suptitle('Conditional requirements — assumptions, not a calibrated forecast',fontsize=14)
    fig.supxlabel('XR holds US share80%, net price$418k, mature retention95%, delay1month; IR net spend is hypothetical.',fontsize=9)
    fig.savefig(ROOT/'output/PHVS_conditional_thresholds_2026-10-08.png',dpi=160)
    plt.close(fig)


if __name__ == '__main__':
    main()
