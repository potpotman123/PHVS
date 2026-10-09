"""Benchmark-linked challenge scenarios, NOT estimates from cross-disease analogues."""
import json
from pathlib import Path
from revenue_ramp import Product,simulate
from global_ramp import globalize
P=Path(__file__).resolve().parent
cfg=json.loads((P/'revenue_ramp_inputs.json').read_text())
geo=json.loads((P/'global_ramp_inputs.json').read_text())['cases']['working']
def first_eight(rate):
 c=cfg['cases']['working'];prod=Product(cfg['origin_pools']);g=(1+c['population_growth'])**(1/12)-1
 rates={k:rate*(.2 if k=='diagnosed_untreated' else 1) for k in cfg['origin_pools']};total=0
 for m in range(8):
  s=prod.step(m,g,rates,c['conversion'],c['delay_months'],.9,.9,True)
  total+=sum(x['rx'] for x in s.values())
 return total
def rate_for(target):
 lo,hi=0.,.99
 for _ in range(50):
  mid=(lo+hi)/2
  if first_eight(mid)<target:lo=mid
  else:hi=mid
 return (lo+hi)/2
rows=[]
for label,rate in [('legacy_low',.05),('working',.1),('legacy_high',.15)]+[(f'{int(f*100)}pct_Ekterly_forms',rate_for(1702*f)) for f in [.5,.75,1.]]:
 us=simulate(cfg,'working',{'IR_annual_prescription_rate':rate});world=globalize(us,cfg,geo)
 rows.append({'scenario':label,'annual_IR_Rx_probability':rate,'first8m_Rx':first_eight(rate),'first8m_Ekterly_form_ratio':first_eight(rate)/1702,'year10_IR_US':us[-1]['IR_revenue'],'year10_global_total':world[-1]['global']})
for r in rows:
 assert abs(r['first8m_Rx']-1702*r['first8m_Ekterly_form_ratio'])<1e-6
out={'status':'Illustrative challenge scenarios;50/75/100%ofEkterlysame-windowforms are judgments,not analog-derived probabilities. Workingcaseunchanged.','method':'Solve annualconstantprobability for first8mforms, thenmaintainsameprobability10years. Calibrationof8monthsdoesnotvalidatepersistentlongrunhazard. Allotherworkingparametersheldfixed.','rows':rows,'limitations':['Noempiricalreasonsecondoralshouldcapannualacquisitionat15%.','Purposivelyselectedsuccessfulanaloguesfalsifyblanketpenaltybutcannotestablishmeanentrantperformance.','Historicaleligibilityandmarketdenominatorsdiffer;formsnotpaidpatients.','ComparativePHVSsuperioritynotprovenheadtohead;clinicaldifferentiationmayaffectprice/access/retentiontoo.','NoDCFvaluationrecalculation.']}
(P/'follower_challenge_results.json').write_text(json.dumps(out,indent=2)+'\n')
for r in rows:print(r['scenario'],round(r['annual_IR_Rx_probability']*100,2),round(r['first8m_Rx']),round(r['year10_global_total']/1e6))
