"""Conditional monthly stock reconciliation; not estimated patient-level transitions."""
import json
import math
from itertools import product
from pathlib import Path
BASE = Path(__file__).resolve().parent
ROOT = BASE.parents[1]

def simulate(opening, opening_paid_share, annual_entries, first_retention,
             mature_retention, entry_paid_share=1., monthly_conversion=0., years=1):
    # Entries at month-start. New users receive first-year hazard for 12 months;
    # opening users receive assumed prevalent-cohort hazard, since ages are unknown.
    cohorts = [dict(p=opening*opening_paid_share, f=opening*(1-opening_paid_share), age=12)]
    rows=[]; total_entries=direct_paid=transfers=paid_exits=free_exits=paid_years=all_years=0.
    for month in range(1,12*years+1):
        e=annual_entries/12
        cohorts.append(dict(p=e*entry_paid_share, f=e*(1-entry_paid_share), age=0))
        total_entries+=e; direct_paid+=e*entry_paid_share
        # Conversion at month-start, then attrition. No reverse transfer or reactivation.
        for c in cohorts:
            t=c['f']*monthly_conversion
            c['f']-=t; c['p']+=t; transfers+=t
            r=(first_retention if c['age']<12 else mature_retention)**(1/12)
            paid_exits+=c['p']*(1-r); free_exits+=c['f']*(1-r)
            c['p']*=r; c['f']*=r; c['age']+=1
        p1=sum(c['p'] for c in cohorts); f1=sum(c['f'] for c in cohorts)
        # Exposure integrates exponential survival exactly after month-start conversion.
        for c in cohorts:
            r=(first_retention if c['age']<=12 else mature_retention)**(1/12)
            factor=1 if abs(r-1)<1e-12 else (1-r)/(-math.log(r))
            paid_years+=(c['p']/r)*factor/12
            all_years+=((c['p']+c['f'])/r)*factor/12
        assert abs(opening+total_entries-paid_exits-free_exits-p1-f1)<1e-7
        assert abs(opening*opening_paid_share+direct_paid+transfers-paid_exits-p1)<1e-7
        assert abs(opening*(1-opening_paid_share)+(total_entries-direct_paid)-transfers-free_exits-f1)<1e-7
        rows.append(dict(month=month,active=p1+f1,paid=p1,free=f1,
            cumulative_entries=total_entries,cumulative_direct_paid_entries=direct_paid,
            cumulative_free_to_paid=transfers,cumulative_paid_exits=paid_exits,
            cumulative_free_exits=free_exits,cumulative_paid_patient_years=paid_years,
            cumulative_all_patient_years=all_years))
    return rows

def solve_entries(opening, target, first, mature):
    survivors=simulate(opening,1,0,first,mature)[-1]['active']
    retained_per_entry=simulate(0,1,1,first,mature)[-1]['active']
    entries=(target-survivors)/retained_per_entry
    if entries<0: raise ValueError('Target inconsistent with nonnegative entries')
    return entries

def main():
    cfg=json.loads((BASE/'reconciliation_inputs.json').read_text())
    opening=1104/1.30; target=1104; target_paid=1104*.715
    history=[]; payment=[]
    for first,mature in product(cfg['historical_first_year_retention'],cfg['prevalent_annual_retention']):
        entries=solve_entries(opening,target,first,mature)
        result=simulate(opening,1,entries,first,mature)[-1]
        assert abs(result['active']-target)<1e-7
        history.append(dict(first_year_retention=first,prevalent_retention=mature,
          inferred_annual_entries_paid_plus_PAP=entries,inferred_exits=result['cumulative_paid_exits'],
          net_growth=target-opening,retained_opening=opening*mature))
        for initial,direct in product(cfg['opening_paid_share'],cfg['entry_paid_share']):
            args=(opening,initial,entries,first,mature,direct)
            lo=simulate(*args,monthly_conversion=0)[-1]['paid']
            hi=simulate(*args,monthly_conversion=1)[-1]['paid']
            record=dict(first=first,mature=mature,opening_paid_share=initial,entry_paid_share=direct)
            if not lo<=target_paid<=hi:
                record.update(feasible=False,reason='Cannot match paid target with nonnegative one-way free-to-paid conversion; requires different assumptions or reverse transfers.')
            else:
                a,b=0.,1.
                for _ in range(45):
                    mid=(a+b)/2
                    if simulate(*args,monthly_conversion=mid)[-1]['paid']<target_paid:a=mid
                    else:b=mid
                conversion=(a+b)/2; run=simulate(*args,monthly_conversion=conversion)
                assert abs(run[-1]['paid']-target_paid)<1e-7
                record.update(feasible=True,monthly_free_to_paid_assumed_model_fit=conversion,result=run[-1])
            payment.append(record)
    # This is a reference activity level conditional on61%first-year and90%prevalent retention.
    reference_entries=solve_entries(opening,target,.61,.9)
    forecasts=[]
    for retention,acq in product(cfg['phvs_first_year_retention'],cfg['acquisition_multipliers']):
        run=simulate(0,1,reference_entries*acq,retention,.9,entry_paid_share=1,years=5)
        forecasts.append(dict(first_year_retention=retention,entry_activity_vs_reference=acq,
            status='Conditional active-treatment comparison; paid status deliberately unspecified',
            annual_state_entries=reference_entries*acq,
            annual_snapshots=[{k:v for k,v in row.items() if k in ['month','active','cumulative_entries','cumulative_all_patient_years']} for row in run if row['month']%12==0]))
    # Population scenarios combine dated proxy anchors explicitly; no current count is observed.
    populations=[]
    for diagnosed,treated,ltp,oral in product(cfg['diagnosed_scenarios'],cfg['treated_scenarios'],cfg['ltp_share_of_treated'],cfg['oral_prevention_scenarios']):
        if treated>diagnosed or oral>treated*ltp:continue
        row=dict(diagnosed=diagnosed,oral_prevention=oral,other_prevention=treated*ltp-oral,
                 acute_only_proxy=treated*(1-ltp),diagnosed_untreated=diagnosed-treated)
        assert abs(sum(row[k] for k in ['oral_prevention','other_prevention','acute_only_proxy','diagnosed_untreated'])-diagnosed)<1e-7
        populations.append(row)
    result=dict(status=cfg['status'],reference_opening_paid_plus_PAP=opening,
                reference_annual_entries=reference_entries,historical_reconciliations=history,
                payment_reconciliations=payment,phvs_activity_scenarios=forecasts,population_scenarios=populations,
                limitations=cfg['limitations'])
    (BASE/'reconciliation_results.json').write_text(json.dumps(result,indent=2)+'\n')
    print('Historical conditional entries',round(min(x['inferred_annual_entries_paid_plus_PAP'] for x in history)),round(max(x['inferred_annual_entries_paid_plus_PAP'] for x in history)))
    print('Reference entries',round(reference_entries,1),'feasible payment cases',sum(x['feasible'] for x in payment),'of',len(payment))
    for x in forecasts:
        if x['entry_activity_vs_reference'] in [.75,1]:print('PHVS',x['first_year_retention'],x['entry_activity_vs_reference'],round(x['annual_snapshots'][-1]['active']),round(x['annual_snapshots'][-1]['cumulative_all_patient_years']))
    # Meaningful limiting cases independent of fitted data.
    zero=simulate(100, .5, 0, 1, 1, monthly_conversion=0)[-1]
    assert zero['active']==100 and abs(zero['cumulative_all_patient_years']-100)<1e-7
    noattrition=simulate(0,1,120,1,1)[-1]
    assert noattrition['active']==120 and abs(noattrition['cumulative_all_patient_years']-65)<1e-7
    none=simulate(0,1,0,.61,.9)[-1];assert none['active']==0
    print('Checks passed: monthly stock/payment identities, fit residuals, population accounting, zero flow and no attrition exposure.')

if __name__=='__main__': main()
