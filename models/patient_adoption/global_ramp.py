"""Evidence-anchored geographic overlay; success-case product sales, not a DCF."""
import json
from pathlib import Path
from revenue_ramp import simulate
P=Path(__file__).resolve().parent

def globalize(rows,cfg,geo):
    out=[]
    for t,row in enumerate(rows):
        r={'year':t+1,'US':row['combined_revenue']}
        for product in ['IR','XR']:
            age=t+1-cfg[product+'_launch_month']/12
            lag=geo['lag_years']; phase=max(0,min(1,(age-lag)/geo['phase_years']))
            j=t-lag
            ref=rows[j][product+'_revenue'] if j>=0 else 0
            r[product+'_exUS']=(geo[product+'_mature_multiplier']-1)*ref*phase
            r[product+'_global']=row[product+'_revenue']+r[product+'_exUS']
        r['global']=r['IR_global']+r['XR_global'];out.append(r)
    return out

def main():
    cfg=json.loads((P/'revenue_ramp_inputs.json').read_text());g=json.loads((P/'global_ramp_inputs.json').read_text())
    us=simulate(cfg,'working')
    result={'status':'Conditional product sales; direct-commercialization equivalent, not PHVS accounting revenue if partnered; no PoS, expenses or DCF', 'geographic_scenarios':{k:globalize(us,cfg,v) for k,v in g['cases'].items()}}
    base=result['geographic_scenarios']['working'][-1]['global'];sens=[]
    inputs=json.loads((P/'revenue_ramp_results.json').read_text())
    for k,v in inputs['one_variable_sensitivities'].items():
        r=globalize(v,cfg,g['cases']['working'])[-1]
        sens.append({'input':k,'year10_global':r['global'],'change_vs_working':r['global']-base})
    result['one_variable_sensitivities']=sorted(sens,key=lambda r:abs(r['change_vs_working']),reverse=True)
    old=globalize(simulate(cfg,'working',{'XR_first_retention':.74}),cfg,g['cases']['working'])
    result['prior_retention_same_geography']=old
    assert all(r['global']>=r['US'] for rows in result['geographic_scenarios'].values() for r in rows)
    zero=globalize(us,cfg,dict(g['cases']['working'],IR_mature_multiplier=1,XR_mature_multiplier=1))
    assert all(r['global']==r['US'] for r in zero)
    assert result['geographic_scenarios']['working'][0]['IR_exUS']==0
    assert result['geographic_scenarios']['working'][1]['XR_exUS']==0
    # Constant mature US sales converge to the specified multiplier, independently of growth.
    flat=[{'IR_revenue':100.,'XR_revenue':100.,'combined_revenue':200.} for _ in range(20)]
    f=globalize(flat,cfg,g['cases']['working'])[-1]
    assert abs(f['IR_global']-120)<1e-8 and abs(f['XR_global']-125)<1e-8
    (P/'global_ramp_results.json').write_text(json.dumps(result,indent=2)+'\n')
    for r in result['geographic_scenarios']['working']:
        print(r['year'],*[round(r[k]/1e6,1) for k in ['US','IR_exUS','XR_exUS','global']])
    print('Top sensitivities:',[(x['input'],round(x['change_vs_working']/1e6,1)) for x in result['one_variable_sensitivities'][:8]])
if __name__=='__main__':main()
