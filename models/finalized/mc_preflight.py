"""Bounded interface/edge tests, not a production investment-outcome simulation."""
import copy
import hashlib
import json
import time
from engine import BASE, read, write, evaluate, resolve, commercial, month_index
from segmented_adoption import joint_pools, rates
from mc_prepare import overrides_at, draws, clinical_outputs, triangular_quantile


def run():
    cfg=read('inputs.json');spec=read('monte_carlo_priors.json');checks=[]
    def check(name, condition):
        assert condition,name
        checks.append(name)
    check('No unknown sampler inputs',set(spec['priors']) <= set(cfg['parameters']))
    check('No overlapping sampled and hard-fixed parameters',not(set(spec['priors']) & set(spec['hard_fixed'])))
    for k,v in spec['priors'].items():
        check(k+' prior within recorded underwriting bounds',cfg['parameters'][k]['bounds'][0]<=v['min']<=v['mode']<=v['max']<=cfg['parameters'][k]['bounds'][1])
        check(k+' endpoints',abs(triangular_quantile(0,v['min'],v['mode'],v['max'])-v['min'])<1e-10 and abs(triangular_quantile(1,v['min'],v['mode'],v['max'])-v['max'])<1e-10)
    q={k:.5 for k in spec['priors']}
    z=overrides_at(q,spec,zero_volatility=True)
    check('Zero volatility resolves to deterministic working inputs',resolve(cfg,overrides=z)==resolve(cfg))
    started=time.perf_counter();zero=evaluate(cfg,overrides=z);elapsed=time.perf_counter()-started
    check('Zero volatility reproduces current working target',abs(zero['expected_12month_value']-read('results.json')['working']['expected_12month_value'])<1e-8)
    check('Seeded draws reproducible',list(draws(8))==list(draws(8)))
    short=copy.deepcopy(cfg);short['end_year']=2029
    corner_checks=[]
    for side in [0.,1.]:
        for access in spec['access_strata']:
            ov=overrides_at({k:side for k in spec['priors']},spec,access=access)
            p=resolve(cfg,overrides=ov);rs,err=commercial(short,p)
            check(f'Boundary {side}/{access}: conservation and finite revenue',err<1e-6 and all(0<=r['revenue']<float('inf') for r in rs))
            corner_checks.append({'quantile_endpoint':side,'access':access,'max_population_error':err})
    start=month_index(cfg['start_month']);launch=month_index(cfg['launches']['IR']['commercial_month'])
    for i,q in draws(128):
        for dep in spec['dependence_modes']:
            p=resolve(cfg,overrides=overrides_at(q,spec,dependence=dep))
            pools=joint_pools(cfg,p,start,launch)
            for age in [0,60,120]:
                xr,ir=rates(cfg,p,pools,age,age)
                assert all(0<=v<1 for v in list(xr.values())+list(ir.values()))
    check('128 seeded draws x2dependence modes preserve joint-stock capacity and probabilities',True)
    pilot=[]
    for i,q in draws(4):
        ov=overrides_at(q,spec);t=time.perf_counter();r=evaluate(cfg,overrides=ov)
        out=clinical_outputs(r)
        check(f'Financial smoke draw{i}: state weights and finite equity',abs(sum(s['probability'] for s in out['clinical_outcomes'])-1)<1e-12 and 0<=out['expected_target']<float('inf'))
        pilot.append({'draw_id':i,'overrides':ov,'elapsed_seconds':time.perf_counter()-t,**out})
    result={'as_of':cfg['as_of'],'status':'Passed bounded preflight only; no investment distribution estimated.','checks_passed':len(checks),'checks':checks,'boundary_cases':corner_checks,'zero_volatility_target':zero['expected_12month_value'],'zero_volatility_runtime_seconds':elapsed,'financial_smoke_draws':pilot,'sha256':{n:hashlib.sha256((BASE/n).read_bytes()).hexdigest() for n in ['inputs.json','engine.py','segmented_adoption.py','monte_carlo_priors.json','mc_prepare.py']}}
    write('mc_preflight_results.json',result)
    print(json.dumps({'checks_passed':len(checks),'zero_target':zero['expected_12month_value'],'seconds_per_expected_value_draw':sum(r['elapsed_seconds'] for r in pilot)/len(pilot)}))

if __name__=='__main__':run()
