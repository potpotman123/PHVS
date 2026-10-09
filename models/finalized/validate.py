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
check('Acquisition replaces standalone value',acquisition_outcome(20,40,'2027-10-09','2027-10-09')==40)
check('Failed completion uses standalone outcome',acquisition_outcome(20,40,'2027-10-09','2027-10-09',completed=False)==20)
check('Later deal discounted to target',acquisition_outcome(20,40,'2028-10-09','2027-10-09')<40)
check('Earlier cash does not automatically earn equity WACC',acquisition_outcome(20,40,'2027-04-09','2027-10-09')==40)
b=buyer_value_interface(1000,200,100,50,75,.5)
check('Buyer value differs from negotiated consideration',b['buyer_economic_value']==1175 and b['negotiated_price_scenario']==1087.5)
check('Negative net synergy does not support an acceptable reservation-price bid',not buyer_value_interface(1000,0,0,50,75,.5)['economically_acceptable_at_reservation'])
# Deterministic monotonicity, not independent evidence for the parameter bounds.
s=read('sensitivities.json')['results']; sm={r['name']:r for r in s}
check('Dose-price increase increases value',sm['ir_net_fraction=0.95']['expected_target']>sm['ir_net_fraction=0.7']['expected_target'])
check('Faster IR acquisition increases value',sm['ir_rx=0.274463']['expected_target']>sm['ir_rx=0.05']['expected_target'])
check('Higher costs reduce value',sm['cost_scale=1.2']['expected_target']<sm['cost_scale=0.8']['expected_target'])
# All source IDs must resolve; no private source is a pivotal model dependency.
sources=read('sources.json')['sources']
check('Every parameter source resolves',all(s in sources for a in cfg['parameters'].values() for s in a['sources']))
check('Every financial-policy source resolves',all(s in sources for a in cfg['financial_policy'].values() if isinstance(a,dict) for s in a['source']))
check('No canonical input cites licensed files',all('alphasense' not in str(v).lower() and 'broker' not in ' '.join(v['sources']).lower() for v in cfg['parameters'].values()))
write('validation.json',{'as_of':cfg['as_of'],'passed':len(tests),'failed':0,'sha256':{n:hashlib.sha256((BASE/n).read_bytes()).hexdigest() for n in ['inputs.json','history.json','sources.json','engine.py','analyze.py']},'tests':tests,'meaning':'Mechanical verification only. Historical calibration and public-source checking are distinct; no independent commercial validation or statistical CI implied.'})
print(f'{len(tests)} validation checks passed.')
