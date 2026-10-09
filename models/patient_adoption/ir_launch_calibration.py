"""Compare same-window PHVS prescription scenarios with Ekterly start forms.
No interpretation of forms as paid/unique retained patients; no fitted capture factor.
"""
import json
from pathlib import Path
from revenue_ramp import Product,simulate
from global_ramp import globalize
P=Path(__file__).resolve().parent
cfg=json.loads((P/'revenue_ramp_inputs.json').read_text())
geo=json.loads((P/'global_ramp_inputs.json').read_text())['cases']['working']
observed={2:460,4:937,6:1318,8:1702}
out={'window_caveat':'July7 launch to monthends approximately2/4/6/8months;modelusesfullmonths. Dates8months2026Feb vs no paid-stock match.','observed_forms':observed,'cases':{}}
for name in cfg['cases']:
 c=cfg['cases'][name];prod=Product(cfg['origin_pools']);g=(1+c['population_growth'])**(1/12)-1
 rates={k:c['IR_annual_prescription_rate']*(.2 if k=='diagnosed_untreated' else 1) for k in cfg['origin_pools']}
 forms=paid=0;rows=[]
 for m in range(12):
  s=prod.step(m,g,rates,c['conversion'],c['delay_months'],c['IR_annual_retention'],c['IR_annual_retention'],True)
  forms+=sum(x['rx'] for x in s.values());paid+=sum(x['first_paid'] for x in s.values())
  if m+1 in observed:
   rows.append({'month':m+1,'PHVS_Rx':forms,'PHVS_paid_starts':paid,'PHVS_active':sum(x['active'] for x in s.values()),'Ekterly_forms':observed[m+1],'PHVS_Rx_over_Ekterly_forms':forms/observed[m+1]})
 # first12months have no retry; independent analytic no-growth check
 q=Product(cfg['origin_pools']);check=0
 for m in range(12):
  z=q.step(m,0,rates,c['conversion'],c['delay_months'],c['IR_annual_retention'],c['IR_annual_retention'],True)
  check+=sum(x['rx'] for x in z.values())
 assert abs(check-sum(cfg['origin_pools'][k]*rates[k] for k in rates))<1e-6
 out['cases'][name]={'windows':rows,'first12m_Rx':forms}
out['retention_annual_scenarios']={}
for retention in [.8,.9,.95]:
 r=simulate(cfg,'working',{'IR_annual_retention':retention});w=globalize(r,cfg,geo)[-1]
 out['retention_annual_scenarios'][str(retention)]={'year5_cohort_survival':retention**5,'year10_IR_US':r[-1]['IR_revenue'],'year10_global_total':w['global']}
out['rules']=['No annualized20%penetration treated as measuredannualhazard.','No>50%refillrevenue interpreted as>50%patientretention.','90%membershipretention implies59%cohortsurvivalafter5years,not90%paidrefillprobability.','NoactualformsorclaimsidentifyPHVSsecondoralcapture;rangeisunderwritingassumption.','Revenue/forms cannot separate netprice, conversion anddoseintensity.','10%IRworkingratealready implies a material same-window discounttofirstoralforms;do not add genericsecondoralhaircut again.']
(P/'ir_launch_calibration_results.json').write_text(json.dumps(out,indent=2)+'\n')
for k,v in out['cases'].items():print(k, v['windows'][-1], 'year1Rx',round(v['first12m_Rx']))
print(out['retention_annual_scenarios'])
