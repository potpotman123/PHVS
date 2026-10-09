"""First-year prescriptions -> first paid supply -> retained paid exposure.
No probability fitting; inputs distinguish evidence, proxies and underwriting choices.
"""
import json
import math
from pathlib import Path
BASE=Path(__file__).resolve().parent

def cohort(prescriptions, conversion, paid_time, retention, horizon=1.):
    eventual=prescriptions*conversion
    if paid_time>=horizon:
        return dict(eventual_paid=eventual,first_paid=0.,pending_paid=eventual,active=0.,exits=0.,patient_years=0.)
    exposure=horizon-paid_time
    hazard=-math.log(retention)
    active=eventual*math.exp(-hazard*exposure)
    py=eventual*exposure if hazard==0 else (eventual-active)/hazard
    return dict(eventual_paid=eventual,first_paid=eventual,pending_paid=0.,active=active,exits=eventual-active,patient_years=py)

def run(cfg,idx,conversion,delay_days,retention,growth):
    rows=[]
    for origin,pool in cfg['origin_pools'].items():
        rate=cfg['annual_prescription_probability'][origin][idx]
        sums={k:0. for k in ['prescriptions','eventual_paid','first_paid','pending_paid','active','exits','patient_years']}
        for m in range(365):
            time=(m+.5)/365
            # One year's composite unique-prescription fraction allocated across daily midpoint cohorts.
            # Growth changes the denominator modestly; not a repeated monthly annual hazard.
            rx=pool*(1+growth)**time*rate/365
            c=cohort(rx,conversion,time+delay_days/365,retention)
            sums['prescriptions']+=rx
            for k,v in c.items(): sums[k]+=v
        sums['nonconverters']=sums['prescriptions']-sums['eventual_paid']
        assert abs(sums['prescriptions']-sums['nonconverters']-sums['first_paid']-sums['pending_paid'])<1e-7
        assert abs(sums['first_paid']-sums['active']-sums['exits'])<1e-7
        assert 0<=sums['patient_years']<=sums['first_paid']
        assert sums['prescriptions']<=pool*(1+growth)
        rows.append(dict(origin=origin,initial_pool=pool,annual_composite_prescription_probability=rate,**sums))
    totals={k:sum(x[k] for x in rows) for k in sums}
    return dict(rate_index=idx,conversion=conversion,delay_days=delay_days,retention=retention,growth=growth,by_origin=rows,totals=totals)

def main():
    cfg=json.loads((BASE/'xr_entry_inputs.json').read_text())
    w=json.loads((BASE/'working_assumptions.json').read_text());x=w['xr']
    ret={k:x['historical_first_year_retention']+(1-x['historical_first_year_retention'])*x['clinical_exit_share_proxy']*v for k,v in x['fraction_of_clinical_exits_avoided'].items()}
    g=w['population_growth']['working_annual_rate']
    cases={name:run(cfg,c['rate_index'],c['first_paid_probability'],c['days_to_paid_conditional'],ret['working'],g) for name,c in cfg['cases'].items()}
    # Separate one-variable sweeps avoid calling correlated bundled cases probability bounds.
    sweep=[]
    for axis,values in [('rate_index',[0,1,2]),('conversion',[.65,.8,.9]),('delay_days',[21,30,90]),('retention',list(ret.values()))]:
        for value in values:
            kwargs=dict(idx=1,conversion=.8,delay_days=30,retention=ret['working'],growth=g)
            kwargs['idx' if axis=='rate_index' else axis]=value
            sweep.append(dict(axis=axis,value=value,totals=run(cfg,**kwargs)['totals']))
    out={'status':cfg['status'],'cases':cases,'one_variable_sensitivities':sweep,
         'population_growth_note':cfg['growth_source'],'limitations':cfg['limitations'],
         'definitions':{'first_paid':'Actual first-paid entrants within first calendar launch year','eventual_paid':'Successful prescriptions, including conversions after year-end','active':'Paid survivors at year-end','patient_years':'Integrated paid exposure; multiply net annual realized revenue per paid-treatment year only when estimated'}}
    (BASE/'xr_entry_results.json').write_text(json.dumps(out,indent=2)+'\n')
    # Limiting-case tests are independent of case arithmetic.
    assert cohort(100,0,.1,.74)['active']==0
    assert cohort(100,1,1.1,.74)['first_paid']==0
    c=cohort(100,1,0,1);assert c['patient_years']==100 and c['active']==100
    c=cohort(100,1,0,.74);assert abs(c['active']-74)<1e-9
    base=cases['working']['totals']
    assert run(cfg,1,.9,30,ret['working'],g)['totals']['patient_years']>base['patient_years']
    assert run(cfg,1,.8,90,ret['working'],g)['totals']['patient_years']<base['patient_years']
    for name,result in cases.items():print(name,{k:round(v,1) for k,v in result['totals'].items()})
    print('Passed: prescription/paid stock identities, exposure bounds, delay spillover, no-conversion/no-attrition cases and monotonicity.')

if __name__=='__main__':main()
