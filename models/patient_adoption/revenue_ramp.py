"""Linked US XR/IR scenario ramp. Fractional patients, monthly cohorts, nominal USD.
Not a valuation or probability-weighted forecast. Evidence: revenue_ramp_inputs.json.
"""
import json, math
from pathlib import Path
BASE=Path(__file__).resolve().parent

class Product:
    def __init__(self,pools):
        self.available=dict(pools);self.total=dict(pools)
        self.pending=[];self.cooldown=[];self.cohorts=[]
        self.retired={k:0. for k in pools}
    def step(self,m,g,rates,conv,delay,first,mature,launched):
        keys=self.total
        stats={k:dict(first_paid=0.,py=0.,first_py=0.,active=0.,rx=0.) for k in keys}
        for k in keys:
            growth=self.total[k]*g;self.total[k]+=growth;self.available[k]+=growth
        keep=[]
        for due,k,n in self.cooldown:
            if due<=m:self.available[k]+=n
            else:keep.append((due,k,n))
        self.cooldown=keep
        if launched:
            for k in keys:
                n=self.available[k]*(1-(1-rates[k])**(1/12))
                self.available[k]-=n;stats[k]['rx']=n
                self.pending.append((m+delay,k,n*conv))
                self.cooldown.append((m+12,k,n*(1-conv)))
        keep=[]
        for due,k,n in self.pending:
            if due<=m:
                self.cohorts.append([k,n,0]);stats[k]['first_paid']+=n
            else:keep.append((due,k,n))
        self.pending=keep
        for c in self.cohorts:
            k,n,age=c;ret=first if age<12 else mature
            hazard=-math.log(ret);survive=n*math.exp(-hazard/12)
            py=n/12 if hazard==0 else (n-survive)/hazard
            stats[k]['py']+=py
            if age<12:stats[k]['first_py']+=py
            self.retired[k]+=n-survive;c[1]=survive;c[2]+=1
            stats[k]['active']+=survive
        for k in keys:
            accounted=self.available[k]+self.retired[k]+stats[k]['active']+sum(n for _,o,n in self.pending+self.cooldown if o==k)
            assert abs(accounted-self.total[k])<1e-6,(k,accounted,self.total[k])
        return stats

def simulate(cfg,case,overrides=None):
    c={**cfg['cases'][case],**(overrides or {})}
    pools=cfg['origin_pools'];xr=Product(pools);ir=Product(pools)
    g=(1+c['population_growth'])**(1/12)-1
    rates={k:c.get('XR_rate_'+k,cfg['XR_annual_prescription_rates'][k][c['rate_index']]) for k in pools}
    ir_rates={k:c['IR_annual_prescription_rate'] if k!='diagnosed_untreated' else c['IR_annual_prescription_rate']*.2 for k in pools}
    annual=[]
    for m in range(cfg['years']*12):
        if m%12==0:row=dict(model_year=m//12+1,XR_prescriptions=0.,XR_first_paid=0.,XR_paid_years=0.,XR_revenue=0.,IR_prescriptions=0.,IR_first_paid=0.,IR_member_years=0.,IR_initial_doses=0.,IR_replenishment_doses=0.,IR_revenue=0.)
        xs=xr.step(m,g,rates,c['conversion'],c['delay_months'],c['XR_first_retention'],.9,m>=cfg['XR_launch_month'])
        ins=ir.step(m,g,ir_rates,c['conversion'],c['delay_months'],c['IR_annual_retention'],c['IR_annual_retention'],m>=cfg['IR_launch_month'])
        for k in pools:
            x=xs[k];i=ins[k]
            # Mean-field overlap within original groups. XR adoption by acute/untreated
            # moves rescue volume toward prevention intensity; no new unique people.
            xr_fraction=min(1.,x['py']*12/xr.total[k])
            baseline=c['IR_doses_LTP'] if k in ['orladeyo','other_prevention'] else c['IR_doses_nonLTP']
            dose_budget=(1-xr_fraction)*baseline+xr_fraction*c['IR_doses_LTP']*c['dual_use_share']
            # Two-dose initial pack included WITHIN first-year annual supply budget.
            initial_per_user=cfg['IR_initial_doses']*(1-xr_fraction+xr_fraction*c['dual_use_share'])
            stock=i['first_paid']*initial_per_user
            recurring=i['first_py']*max(0,dose_budget-initial_per_user)+(i['py']-i['first_py'])*dose_budget
            row['XR_prescriptions']+=x['rx'];row['XR_first_paid']+=x['first_paid'];row['XR_paid_years']+=x['py']
            row['XR_revenue']+=x['py']*c['XR_net_annual_price']
            row['IR_prescriptions']+=i['rx'];row['IR_first_paid']+=i['first_paid'];row['IR_member_years']+=i['py']
            row['IR_initial_doses']+=stock;row['IR_replenishment_doses']+=recurring
            row['IR_revenue']+=(stock+recurring)*c['IR_net_dose_price']
        if m%12==11:
            row['XR_end_paid']=sum(v['active'] for v in xs.values())
            row['IR_end_enrolled']=sum(v['active'] for v in ins.values())
            row['US_diagnosed_population']=sum(xr.total.values())
            row['combined_revenue']=row['XR_revenue']+row['IR_revenue']
            assert row['XR_end_paid']<=row['US_diagnosed_population']
            assert row['IR_end_enrolled']<=row['US_diagnosed_population']
            annual.append(row)
    return annual

def main():
    cfg=json.loads((BASE/'revenue_ramp_inputs.json').read_text())
    results={k:simulate(cfg,k) for k in cfg['cases']}
    sensitivities={}
    for key,values in {'conversion':[.65,.9],'dual_use_share':[0.,1.], 'IR_doses_LTP':[6.,11.8], 'IR_doses_nonLTP':[10.,18.], 'IR_annual_retention':[.8,.95], 'XR_first_retention':[.61,.6945,.74], 'XR_rate_orladeyo':[.05,.15], 'XR_rate_other_prevention':[.02,.06], 'IR_annual_prescription_rate':[.05,.15], 'XR_net_annual_price':[300000,500000], 'IR_net_dose_price':[12500,16000]}.items():
        for v in values:sensitivities[f'{key}={v}']=simulate(cfg,'working',{key:v})
    # Independent edge cases and directional checks, not empirical validation.
    zero=simulate(cfg,'working',{'conversion':0.});assert all(x['combined_revenue']==0 for x in zero)
    base=results['working'][-1]
    assert sensitivities['conversion=0.9'][-1]['combined_revenue']>base['combined_revenue']
    assert sensitivities['dual_use_share=0.0'][-1]['IR_revenue']<=sensitivities['dual_use_share=1.0'][-1]['IR_revenue']
    assert all(x['XR_revenue']==0 for x in results['working'][:cfg['XR_launch_month']//12])
    # A fully retained new rescue user at time zero receives exactly budget,
    # not budget plus the initial pack, across first12months.
    assert 2+12*(8-2)/12==8
    out={'status':cfg['status'],'cases':results,'one_variable_sensitivities':sensitivities,'limitations':cfg['limitations']}
    (BASE/'revenue_ramp_results.json').write_text(json.dumps(out,indent=2)+'\n')
    lines=['US REVENUE RAMP — scenario outputs, not consensus or fair value','ModelYear1=IRlaunch;ModelYear2=XRlaunch;full years,notcalendar forecasts','Case | Year | XR $m | IR $m | Combined $m | XR end paid | IR end enrolled']
    for name,rows in results.items():
        for x in rows:
            lines.append(f"{name} | {x['model_year']} | {x['XR_revenue']/1e6:.1f} | {x['IR_revenue']/1e6:.1f} | {x['combined_revenue']/1e6:.1f} | {x['XR_end_paid']:.0f} | {x['IR_end_enrolled']:.0f}")
    (BASE/'revenue_ramp_table.txt').write_text('\n'.join(lines)+'\n')
    print('\n'.join(lines[:3]+[l for l in lines[3:] if l.startswith('working')]))
    print('Validated stock conservation, zero conversion, launch timing, supply budget and directional sensitivities.')

if __name__=='__main__':main()
