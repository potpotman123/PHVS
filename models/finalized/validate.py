"""Accounting/edge-case tests. Passing does not validate the market assumptions."""
import copy
import datetime as dt
import json
import math
import hashlib
from engine import *

cfg = read('inputs.json'); p = resolve(cfg); tests=[]
def check(name, predicate, detail=''):
    assert predicate, (name,detail)
    tests.append({'test':name,'passed':True,'detail':detail})

check('Origin weights conserve disease population',abs(sum(cfg['pool_weights'].values())-1)<1e-12)
# Independent cohort calculation; one initial person, no acquisition or growth.
one=Population({'k':1.});one.available['k']=0.;one.pending=[(0,'k',1.)]
pp={**p,'conversion':1.,'reentry_fraction':0.,'delay':0}
first_py=0.; end=0.
for m in range(24):
    a=one.step(m,0,{'k':0},pp,.6945,.9,False)['k']
    if m<12:first_py+=a['py']
    if m==11:end=a['active']
check('Twelve-month survival matches chosen retention',abs(end-.6945)<1e-12)
check('Mature retention applies after month12',abs(a['active']-.6945*.9)<1e-12)
check('Patient-year integration matches analytic survival',abs(first_py-(1-.6945)/(-math.log(.6945)))<1e-12)
one=Population({'k':1.});one.available['k']=0.;one.pending=[(0,'k',1.)]
stock=replen=0.
for m in range(12):
    a=one.step(m,0,{'k':0},pp,1.,1.,False)['k']
    stock+=a['starts']*2;replen+=a['first_py']*(8-2)+(a['py']-a['first_py'])*8
check('Actual cohort initial stock is inside first-year dose budget',abs(stock+replen-8)<1e-12,{'initial':stock,'replacement':replen})
check('Refill proxy exactly reconciles benchmark PDC',abs(p['xr_paid_intensity']*(1-.61)/(-math.log(.61))-.73)<1e-12,'Calibration identity, not validation of PHVS paid exposure.')
short=copy.deepcopy(cfg);short['end_year']=2029
zero,err=commercial(short,{**p,'conversion':0},(True,True))
check('Zero paid conversion gives zero revenue',all(r['revenue']==0 for r in zero))
check('Zero conversion and failed-access retries conserve patients',err<1e-6)
rev,err=commercial(short,p,(True,True))
free,_=commercial(short,{**p,'ir_initial_paid_fraction':0},(True,True))
check('Free initial supply never earns revenue',all(r['ir_initial_doses']==0 for r in free) and sum(r['ir_free_initial_doses'] for r in free)>0)
check('Free initial supply still reduces replacement budget',all(abs(a['ir_replacement_doses']-b['ir_replacement_doses'])<1e-9 for a,b in zip(rev,free)))
check('Free-stock revenue reduction equals withheld paid initial units',all(abs((a['IR_US']-b['IR_US'])-a['ir_initial_doses']*a['IR_price_usd']/1e6)<1e-8 for a,b in zip(rev,free)))
for prod in ['IR','XR']:
    launch=month_index(cfg['launches'][prod]['commercial_month'])
    check(prod+' no revenue before paid-access date',all(r[prod+'_US']==0 for r in rev if month_index(r['month'])<launch+int(p['delay'])))
    check(prod+' net price units reconcile',all(abs(r[prod+'_US']-r[prod+'_units']*r[prod+'_price_usd']/1e6)<1e-9 for r in rev))
check('Overlap does not create more unique users than disease patients',all(r['xr_patient_years']*12+r['ir_member_years']*12-r['overlap_members']<=r['population']+1e-7 for r in rev))
partner,_=commercial(short,{**p,'exus_direct_fraction':0},(True,True))
check('Partner route recognizes royalties only, no gross overseas double count',all(abs(r['revenue']-(r['IR_US']+r['XR_US']+(r['IR_exUS_endmarket']+r['XR_exUS_endmarket'])*p['partner_royalty']))<1e-9 for r in partner))
for mode in ['shared_gate','independent','perfect_positive']:
    w=clinical_weights(p,mode)
    check(mode+' clinical weights/marginals',abs(sum(w.values())-1)<1e-12 and min(w.values())>=-1e-12 and abs(w['both']+w['IR_only']-.9)<1e-12 and abs(w['both']+w['XR_only']-.9)<1e-12)
for y in range(2038,2056):
    a,b,c=erosion(cfg,p,y,'IR')
    check(f'LOE price times volume {y}',abs(a*b-c)<1e-12)
check('LOE first-year blend retains original framework',abs(erosion(cfg,p,2039,'IR')[2]-(.5*(88/64*.24)+.5*285.776/457.7))<1e-12)
check('Flat-price stress cannot exceed original erosion revenue',all(erosion(cfg,{**p,'loe_price_cap':True},y,'IR')[2]<=erosion(cfg,p,y,'IR')[2] for y in range(2039,2056)))
h=read('history.json')
for k,r in h['periods'].items():
    check(k+' reported cash rollforward',abs(r['cash_begin']+r['cfo']+r['cfi']+r['cff']+r['cash_fx']-r['cash_end'])<1e-6)
    check(k+' reported balance sheet',abs(r['total_assets']-r['liabilities']-r['equity'])<1e-6)
    check(k+' historical P&L',abs(r['revenue']-r['opex']+r['finance_income']-r['tax']-r['net_income'])<1e-6)
for case in cfg['cases']:
    rows=read(case+'_monthly.json')
    for state,rs in rows.items():
        check(case+'/'+state+' full population/balance checks',read('results.json')[case]['checks'][state]['max_population_error']<1e-6 and max(abs(r['balance_error']) for r in rs)<1e-6)
        check(case+'/'+state+' operating cash identity',all(abs(r['fcf']-(r['net_income']+r['depreciation']-r['capex']-r['delta_nwc']))<1e-7 for r in rs))
        check(case+'/'+state+' cash financing identity',all(abs(r['cash_end']-(r['cash_begin']+r['fcf']+r['funding_net']))<1e-6 for r in rs if not r['shutdown']))
        check(case+'/'+state+' funding/share units',all(abs(r['financing_shares']*resolve(cfg,case)['issue_price']-r['funding_gross'])<1e-6 for r in rs))
        years=aggregate(rs,'year');quarters=aggregate(rs,'quarter')
        check(case+'/'+state+' monthly quarterly annual revenue reconcile',abs(sum(r['revenue'] for r in rs)-sum(r['revenue'] for r in years))<1e-6 and abs(sum(r['revenue'] for r in rs)-sum(r['revenue'] for r in quarters))<1e-6)
        # Target is a separately dated valuation, not an extra 0.9 revenue haircut.
        val=read('results.json')[case]['conditional_valuations'][state]['target']
        check(case+'/'+state+' equity/shares reconcile',abs(val['price']*val['diluted_shares_m']-val['equity_usdm'])<1e-6)
# Deterministic monotonicity, not independent evidence for the parameter bounds.
s=read('sensitivities.json')['results']; sm={r['name']:r for r in s}
check('Dose-price increase increases value',sm['ir_net_fraction=0.95']['expected_target']>sm['ir_net_fraction=0.7']['expected_target'])
check('Faster IR acquisition increases value',sm['ir_injectable_opportunity=1.25']['expected_target']>sm['ir_injectable_opportunity=0.75']['expected_target'])
check('Higher costs reduce value',sm['cost_scale=1.2']['expected_target']<sm['cost_scale=0.8']['expected_target'])
# All source IDs must resolve; no private source is a pivotal model dependency.
sources=read('sources.json')['sources']
check('Every parameter source resolves',all(s in sources for a in cfg['parameters'].values() for s in a['sources']))
check('Every financial-policy source resolves',all(s in sources for a in cfg['financial_policy'].values() if isinstance(a,dict) for s in a['source']))
check('No canonical input cites licensed files',all('alphasense' not in str(v).lower() and 'broker' not in ' '.join(v['sources']).lower() for v in cfg['parameters'].values()))
check('Working applies pivotal redosing stress once',abs(p['ir_quantity_multiplier']-1.17/1.38)<1e-12)
check('No efficacy price premium in any configured case',all(resolve(cfg,k)['ir_price_multiplier']==resolve(cfg,k)['xr_price_multiplier']==1 for k in cfg['cases']))
# Integrated source-stock and timing tests, independently compared with calibration.
from segmented_adoption import joint_pools, oral_share, rates as source_rates
start=month_index(cfg['start_month']); launch=month_index(cfg['launches']['IR']['commercial_month'])
joint=joint_pools(cfg,p,start,launch)
check('Joint baseline stock conserves population',abs(sum(joint.values())-p['population'])<1e-8)
for origin,w in cfg['pool_weights'].items():
    check('Joint stock row '+origin,abs(sum(n for k,n in joint.items() if k.split('|')[0]==origin)-p['population']*w)<1e-8)
old_stock=read('adoption_calibration/final_mechanics_results.json')['stocks']['working']['launch_mix']['rescue_pools']['Ekterly']
check('Projected Ekterly launch stock matches separate calibration',abs(sum(n for k,n in joint.items() if k.endswith('|ekterly'))*(1+p['growth'])**((launch-start)/12)-old_stock)<1e-8)
for case in cfg['cases']:
    cp=resolve(cfg,case)
    check(case+' oral choice reaches60% bymonth60',abs(oral_share(cp,60)-.6)<1e-12 and abs(oral_share(cp,120)-.6)<1e-12)
check('Working launch oral choice40%',abs(oral_share(p,0)-.4)<1e-12)
rx,ri=source_rates(cfg,p,joint,60,60)
check('XR incumbent oral is all-destination capture',abs(rx['orladeyo|ekterly']-(50/328)*.6/.8)<1e-12)
check('XR injectable formula matches independent arithmetic',abs(rx['other_prevention|injectable']-(50/328)*.54*.6/.8)<1e-12)
check('IR injectable uses forms normalization without second access adjustment',abs(ri['other_prevention|injectable']-(1-(1-1702/9000)**1.5)*.6)<1e-12)
check('Source rates independent of actual paid conversion',source_rates(cfg,{**p,'conversion':.65},joint,60,60)==(rx,ri))
check('Source Rx and paid starts reconcile to totals',all(abs(sum(v for k,v in r.items() if k.startswith(prefix))-r[total])<1e-7 for r in rev for prefix,total in [('ir_rx_from_','ir_rx'),('xr_rx_from_','xr_rx'),('ir_starts_from_','ir_access_starts'),('xr_starts_from_','xr_starts')]))
check('Population grows once from its dated anchor',all(abs(r['population']-p['population']*(1+p['growth'])**((month_index(r['month'])-start)/12))<1e-6 for r in rev))
check('Every active case has the same redosing discount',all(abs(resolve(cfg,k)['ir_quantity_multiplier']-1.17/1.38)<1e-12 for k in cfg['cases']))
check('Every active sensitivity retains the redosing discount',all(abs(resolve(cfg,overrides=x['overrides'])['ir_quantity_multiplier']-1.17/1.38)<1e-12 for x in sm.values()))
rejected=False
try: resolve(cfg,overrides={'ir_quantity_multiplier':1.})
except AssertionError: rejected=True
check('Engine rejects forecasts without the selected discount',rejected)

# Regulatory delay must move liquidation after the latest decision, not just XR's original date.
delayed={**p,'launch_delay_IR':24,'launch_delay_XR':6}
cr,_=commercial(short,delayed,(False,False));fr=financial(short,delayed,cr,(False,False))
expected=max(month_index(v['approval_assumption'])+int(delayed.get('launch_delay_'+k,0)) for k,v in cfg['launches'].items())+cfg['financial_policy']['winddown_months']['value']
check('Both-fail shutdown follows latest delayed decision',next(month_index(r['month']) for r in fr if r['shutdown'])==expected)
rx2,ri2=source_rates(cfg,{**p,'ir_oral_switch_relative':.5},joint,60,60)
check('IR transfer uncertainty leaves XR unchanged',rx2==rx)
check('IR transfer uncertainty isolates incumbent oral origin',all(abs(ri2[k]-ri[k]*(.5 if k.endswith('|ekterly') else 1))<1e-12 for k in ri))
rejected=False
try: resolve(cfg,overrides={'ir_price_multiplier':1.1})
except AssertionError: rejected=True
check('Engine rejects efficacy price-premium forecasts',rejected)
write('validation.json',{'as_of':cfg['as_of'],'passed':len(tests),'failed':0,'sha256':{n:hashlib.sha256((BASE/n).read_bytes()).hexdigest() for n in ['inputs.json','history.json','sources.json','engine.py','segmented_adoption.py','analyze.py','validate.py']},'tests':tests,'meaning':'Mechanical verification only. Historical calibration and public-source checking are distinct; no independent commercial validation or statistical CI implied.'})
print(f'{len(tests)} validation checks passed.')
