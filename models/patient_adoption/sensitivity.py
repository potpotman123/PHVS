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

rows=[]
for attempt in [.05,.10,.20,.30]:
    for paid in [.60,.80,1.0]:
        result=simulate(attempt=attempt,paid=paid)
        rows.append(dict(annual_attempt_probability=attempt,initial_paid_probability=paid,first_year_retention=.61,subsequent_retention=.90,result=result))
base=Path(__file__).parent
(base/'illustrative_sensitivity.json').write_text(json.dumps({'status':'Hypothetical mechanics only. NOT fitted PHVS forecast or scenario probabilities.','population':'1000 initially eligible incumbent patients; no geography implied','timing':'Each annual cohort is treated as entering at the start of its year; year-end active counts are NOT annual patient-years.','retention_note':'61% is an Orladeyo initiation-cohort benchmark, not PHVS retention or paid-only estimate. Applying it to paid starts is a scenario assumption. Subsequent90% is hypothetical.','access_note':'60/80/100% are hypothetical initial-paid probabilities, NOT Orladeyo observed stock paid share. No free-to-paid lag modeled.','rows':rows},indent=2)+'\n')
for x in rows:
    if x['initial_paid_probability']==.8:print(x['annual_attempt_probability'],round(x['result'][-1]['year_end_active_paid'],1))
