"""Independent accounting checks for conditional thresholds, not market validation."""
import json
import math
from pathlib import Path
from threshold_analysis import integrated_survival, survival, exposure, solve_xr

HERE=Path(__file__).resolve().parent
config=json.loads((HERE/'threshold_inputs.json').read_text())
results=json.loads((HERE/'threshold_results.json').read_text())
p=config['illustrative_reference']
checks=[]

# A twelve-month endpoint must equal the stated survival, regardless of hazard shape.
assert abs(survival(12,p)-p['first_year_retention'])<1e-12
assert abs(survival(24,p)-p['first_year_retention']*p['mature_annual_retention'])<1e-12
# Independent numerical integral using 2000 subintervals per month.
for age in [0,2,3,8,11,12,23,60]:
    numeric=sum(survival(age+(j+.5)/2000,p) for j in range(2000))/2000
    assert abs(numeric-integrated_survival(age,p))<1e-9
checks.append('Monthly treated-time integration agrees with independent fine-grid integration')

# Closed-form no-attrition benchmark: one paid start each month has6.5 patient-years.
no_attrition={**p,'first_year_retention':1.,'mature_annual_retention':1.}
assert abs(exposure([(m,1.) for m in range(12)],0,11,no_attrition)-6.5)<1e-12
checks.append('Zero-attrition monthly cohort result equals closed-form6.5 patient-years')

all_scenarios=[results['reference']]+results['retention_access']+results['price_geography']+results['one_at_a_time']
if 'joint_scenarios' in results:all_scenarios+=results['joint_scenarios']
max_gap=0
for s in all_scenarios:
    assert s['all_annual_targets_matched']
    assert all(math.isfinite(x['first_prescriptions']) and x['first_prescriptions']>=0 for x in s['monthly_prescriptions'])
    total_existing=sum(a['existing_patient_prescriptions'] for a in s['annual'])
    total_new=sum(a['new_diagnosis_prescriptions'] for a in s['annual'])
    total_rx=sum(x['first_prescriptions'] for x in s['monthly_prescriptions'])
    assert abs(total_rx-total_existing-total_new)<1e-6
    assert abs(s['required_initial_existing_pool']-total_existing)<1e-6
    assert s['absolute_minimum_all_prescriptions']<=total_rx+1e-6
    assert abs(s['annual'][-1]['remaining_existing_never_attempted']+total_existing-s['inputs']['existing_eligible_pool'])<1e-6
    max_gap=max(max_gap,max(abs(a['revenue_gap_USD']) for a in s['annual']))
checks+=['All scenario years reconcile to saved revenue hurdle within$1',
         'All monthly prescriptions are finite/nonnegative',
         'Existing plus newly diagnosed prescriptions conserve total unique attempts',
         'Smooth allocations never use fewer prescriptions than the linear-program minimum']

# IR unit identities and price monotonicity without the cohort engine.
target=results['worldwide_2036_USDm']['IR']*1e6
for row in results['IR_pack_thresholds']:
    assert abs(row['required_annual_paid_users']*row['paid_packs_per_user_year']*row['net_pack_USD']-target*.8)<1e-5
for row in results['IR_spend_thresholds']:
    assert abs(row['required_annual_paid_users']*row['annual_net_spend_USD']-target*row['us_fraction'])<1e-5
checks.append('IR users x packs x net price and users x annual spend reconcile')

# A deliberately undersized pool must flag infeasibility, not silently scale sales down.
targets=[(a['year'],a['worldwide_target_USDm']) for a in results['reference']['annual']]
tiny=solve_xr({**p,'existing_eligible_pool':1},targets)
assert tiny['first_pool_exhaustion_year']==p['xr_launch_year']
assert tiny['annual'][-1]['remaining_existing_never_attempted']<0
checks.append('Undersized eligible pool correctly flags recruitment infeasibility')

out={'status':'passed','checks':checks,'scenario_count':len(all_scenarios),
     'max_absolute_revenue_reconciliation_error_USD':max_gap,
     'market_validation':False,'DCF_repriced':False,
     'correction':'Initial year-flat prescription inversion caused alternating cohorts and a missed revenue target. Replaced before delivery with nonnegative monthly allocation and exact annual constraints.'}
(HERE/'threshold_validation.json').write_text(json.dumps(out,indent=2)+'\n')
print(json.dumps(out,indent=2))
