"""Illustrative closed-cohort sensitivity, not a PHVS adoption forecast.
Annual transitions: incumbent -> attempted PHVS -> initially paid or failed/free.
No reattempts; no new diagnoses; discontinued PHVS users leave the active pool.
Retention after year1 is deliberately explicit and assumed, not observed.
"""
import json
from pathlib import Path

def simulate(n=1000,attempt=.10,paid=.80,first_retention=.61,later_retention=.90,years=5):
    remaining=float(n)
    cohorts=[]
    out=[]
    for year in range(1,years+1):
        cohorts=[x*later_retention for x in cohorts]
        attempts=remaining*attempt
        remaining-=attempts
        starts=attempts*paid
        cohorts.append(starts*first_retention)
        out.append(dict(year=year,attempts=attempts,initial_paid_starts=starts,year_end_active_paid=sum(cohorts),remaining_never_attempted=remaining))
    assert 0 <= sum(cohorts) <= n
    return out

base=Path(__file__).parent
assumptions=json.loads((base/'working_assumptions.json').read_text())
xr=assumptions['xr']
retentions={name: xr['historical_first_year_retention'] +
    (1-xr['historical_first_year_retention'])*xr['clinical_exit_share_proxy']*effect
    for name,effect in xr['fraction_of_clinical_exits_avoided'].items()}
rows=[]
for name,retention in retentions.items():
    for attempt in [.05,.10,.20,.30]:
        for paid in [.60,.80,1.0]:
            result=simulate(attempt=attempt,paid=paid,first_retention=retention,
                            later_retention=xr['mature_annual_retention'])
            rows.append(dict(retention_case=name,annual_attempt_probability=attempt,
                initial_paid_probability=paid,first_year_retention=retention,
                subsequent_retention=xr['mature_annual_retention'],result=result))
ir=assumptions['ir'];phvs=1+ir['phvs_pivotal_repeat_within12h']
comparisons={name:dict(comparator_doses_per_100_treated_attacks=100*(1+rate),
    phvs_doses_per_100_treated_attacks=100*phvs,
    phvs_conditional_volume_difference=phvs/(1+rate)-1,
    extra_treated_attacks_or_net_price_per_dose_to_offset=(1+rate)/phvs-1)
    for name,rate in ir['ekterly_comparators'].items()}
output={'status':assumptions['status'],'population':'1000 hypothetical eligible incumbents',
 'timing':'Annual cohorts enter at year start. Year-end active patients are NOT annual patient-years.',
 'retention_cases':retentions,'retention_note':xr['working_choice_rationale'],
 'access_note':'Attempt/access grid unmeasured. Not stock paid share or survey intention.',
 'rows':rows,'IR_comparisons':comparisons,
 'IR_note':ir['status'],
 'LTP_extension_oral_first_fraction':ir['oral_first_attacks']/ir['all_attacks'],
 'LTP_extension_any_treated_fraction':ir['any_treated_attacks']/ir['all_attacks']}
(base/'illustrative_sensitivity.json').write_text(json.dumps(output,indent=2)+'\n')
for row in rows:
    if row['annual_attempt_probability']==.10 and row['initial_paid_probability']==.80:
        print(row['retention_case'],row['first_year_retention'],round(row['result'][-1]['year_end_active_paid'],1))
