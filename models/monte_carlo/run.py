"""Reproducible standalone PHVS Monte Carlo using the unmodified finalized engine.

Standard library only. Execute from any directory; see README.md for interpretation.
An execution samples analyst uncertainty, not empirical probabilities of stock returns.
"""
from __future__ import annotations
import argparse
from concurrent.futures import ProcessPoolExecutor
import gzip
import hashlib
import json
import math
import multiprocessing
from pathlib import Path
import platform
import statistics
import sys
import time
import traceback

HERE = Path(__file__).resolve().parent
MODEL = HERE.parent / 'finalized'
sys.path.insert(0, str(MODEL))
import engine
from mc_prepare import draws, overrides_at

CORE = ['engine.py', 'segmented_adoption.py', 'mc_prepare.py', 'inputs.json',
        'history.json', 'sources.json', 'monte_carlo_priors.json']
CFG = engine.read('inputs.json')
SPEC = engine.read('monte_carlo_priors.json')
TOL = 1e-6

def hashes():
    return {str(p.relative_to(HERE.parent.parent)): hashlib.sha256(p.read_bytes()).hexdigest()
            for p in [*(MODEL / n for n in CORE), HERE / 'run.py']}

def dump(path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, indent=2, allow_nan=False) + '\n')

def require(ok, message):
    if not ok:
        raise ValueError(message)

def quantile(values, q):
    """Empirical inverse CDF, without interpolating discrete outcome masses."""
    values = sorted(values)
    require(bool(values) and 0 <= q <= 1, 'Invalid quantile')
    return values[max(0, math.ceil(q * len(values)) - 1)]

def weighted_quantile(pairs, q):
    pairs = sorted((x, w) for x, w in pairs if w > 0)
    require(bool(pairs) and 0 <= q <= 1, 'Invalid weighted quantile')
    threshold = q * math.fsum(w for _, w in pairs)
    cumulative = 0.
    for x, w in pairs:
        cumulative += w
        if cumulative >= threshold - 1e-12:
            return x
    return pairs[-1][0]

def describe(values):
    return {'mean': statistics.fmean(values), 'median': quantile(values, .5),
            'p10': quantile(values, .1), 'p90': quantile(values, .9),
            'min': min(values), 'max': max(values),
            'mean_mc_se': statistics.stdev(values)/math.sqrt(len(values)) if len(values)>1 else None}

def ranks(values):
    order = sorted(range(len(values)), key=values.__getitem__)
    result = [0.] * len(values)
    start = 0
    while start < len(order):
        end = start + 1
        while end < len(order) and values[order[end]] == values[order[start]]:
            end += 1
        rank = (start + 1 + end) / 2
        for i in order[start:end]:
            result[i] = rank
        start = end
    return result

def correlation(x, y):
    mx, my = statistics.fmean(x), statistics.fmean(y)
    numerator = math.fsum((a-mx)*(b-my) for a,b in zip(x,y))
    denominator = math.sqrt(math.fsum((a-mx)**2 for a in x)*math.fsum((b-my)**2 for b in y))
    return numerator / denominator if denominator else None

def assert_result(result, p):
    weights = result['clinical_weights']
    require(abs(math.fsum(weights.values())-1) < 1e-12 and min(weights.values()) >= 0, 'Clinical probability bounds')
    require(abs(weights['both']+weights['IR_only']-p['ir_pos'])<1e-12, 'IR marginal probability')
    require(abs(weights['both']+weights['XR_only']-p['xr_pos'])<1e-12, 'XR marginal probability')
    max_stock = max_balance = max_cash = max_source = 0.
    for state, rows in result['monthly_by_clinical_state'].items():
        max_stock = max(max_stock, result['checks'][state]['max_population_error'])
        for r in rows:
            require(all(math.isfinite(x) for x in r.values() if isinstance(x,(float,int))), 'Nonfinite monthly value')
            max_balance = max(max_balance, abs(r['balance_error']))
            require(min(r[k] for k in ['revenue','IR_units','XR_units','ir_initial_doses','ir_free_initial_doses','ir_replacement_doses','xr_active','ir_access_active']) >= -1e-9, 'Negative quantity or revenue')
            require(abs(r['revenue']-r['IR_recognized']-r['XR_recognized'])<TOL, 'Product revenue reconciliation')
            require(abs(r['fcf']-r['net_income']-r['depreciation']+r['capex']+r['delta_nwc'])<TOL, 'FCF identity')
            require(abs(r['financing_shares']*p['issue_price']-r['funding_gross'])<TOL, 'Funding/share units')
            if not r['shutdown']:
                max_cash = max(max_cash, abs(r['cash_end']-r['cash_begin']-r['fcf']-r['funding_net']))
            require(r['xr_patient_years']*12+r['ir_member_years']*12-r['overlap_members'] <= r['population']+TOL, 'Shared-population overlap')
            for prefix, total in [('ir_rx_from_','ir_rx'),('xr_rx_from_','xr_rx'),('ir_starts_from_','ir_access_starts'),('xr_starts_from_','xr_starts')]:
                max_source = max(max_source,abs(math.fsum(v for k,v in r.items() if k.startswith(prefix))-r[total]))
            for prod in ['IR','XR']:
                require(abs(r[prod+'_US']-r[prod+'_units']*r[prod+'_price_usd']/1e6)<TOL, 'Price times units')
                if not engine.STATES[state][0 if prod=='IR' else 1]:
                    require(abs(r[prod+'_recognized'])<1e-12, 'Failed-product revenue')
        for date in ['valuation','target']:
            v = result['conditional_valuations'][state][date]
            require(math.isfinite(v['price']) and v['price'] >= 0, 'Invalid state equity value')
            require(abs(v['price']*v['diluted_shares_m']-v['equity_usdm'])<TOL, 'Equity-to-share price')
    require(max(max_stock,max_balance,max_cash,max_source)<TOL,'Conservation / accounting tolerance')
    target = math.fsum(weights[s]*result['conditional_valuations'][s]['target']['price'] for s in weights)
    require(abs(target-result['expected_12month_value'])<1e-9,'Double or missing clinical weighting')
    return {'population':max_stock,'balance':max_balance,'cash':max_cash,'source_totals':max_source}

def evaluate_task(task):
    i,q,access,structural,dependence = task
    try:
        override = overrides_at(q,SPEC,access=access,structural=structural,dependence=dependence)
        p=engine.resolve(CFG,overrides=override)
        for key,value in SPEC['hard_fixed'].items():
            require(p[key]==value,'Changed fixed policy: '+key)
        for key,prior in SPEC['priors'].items():
            require(prior['min']<=p[key]<=prior['max'],'Outside prior: '+key)
        result=engine.evaluate(CFG,overrides=override,retain=True)
        checks=assert_result(result,p)
        state_values={s:{**result['conditional_valuations'][s],
            'weight':result['clinical_weights'][s],
            'total_net_funding_usdm':math.fsum(r['funding_net'] for r in rs),
            'total_gross_funding_usdm':math.fsum(r['funding_gross'] for r in rs)}
            for s,rs in result['monthly_by_clinical_state'].items()}
        annual=result['annual_both_success']
        selected={str(r['year']):{k:r[k] for k in ['IR_recognized','XR_recognized','revenue','xr_active','ir_access_active','funding_net']}
                  for r in annual if r['year'] in [2028,2030,2035,2038]}
        return {'draw_id':i,'inputs':override,'expected_target':result['expected_12month_value'],
                'value_today':result['risk_adjusted_value_today'],'states':state_values,
                'both_success_peak':result['conditional_peak'],'annual_both_success':selected,
                'max_errors':checks}
    except Exception as exc:
        return {'draw_id':i,'error':str(exc),'traceback':traceback.format_exc(),
                'quantiles':q,'access':access,'structural':structural,'dependence':dependence}

def scenario_name(access,structural,dependence):
    return access+'__'+structural+'__'+dependence

def batch(n,seed,access,structural,dependence,workers,output):
    name=scenario_name(access,structural,dependence)
    folder=output/f'seed_{seed}'/name
    folder.mkdir(parents=True,exist_ok=True)
    identity={'hashes':hashes(),'seed':seed,'access':access,'structural':structural,'dependence':dependence}
    identity_path=folder/'identity.json'
    if identity_path.exists():
        require(json.loads(identity_path.read_text())==identity,'Cached results do not match source / runner hashes')
    else:dump(identity_path,identity)
    records_path=folder/'draws.jsonl'
    existing=[json.loads(l) for l in records_path.read_text().splitlines()] if records_path.exists() else []
    require([r['draw_id'] for r in existing]==list(range(len(existing))),'Checkpoint is not a contiguous seed prefix')
    require(len(existing)<=n,'Refusing to truncate completed draws')
    tasks=[(i,q,access,structural,dependence) for i,q in draws(n,seed) if i>=len(existing)]
    started=time.perf_counter()
    with records_path.open('a') as fh:
        with ProcessPoolExecutor(max_workers=workers,mp_context=multiprocessing.get_context('spawn')) as pool:
            for result in pool.map(evaluate_task,tasks,chunksize=4):
                if 'error' in result:
                    dump(folder/'failure.json',result)
                    raise RuntimeError(f"Invalid draw {result['draw_id']}: {result['error']}; no redraw performed")
                fh.write(json.dumps(result,separators=(',',':'),allow_nan=False)+'\n');fh.flush()
                existing.append(result)
                if len(existing)%64==0:
                    print(f'{name}: {len(existing)}/{n}',flush=True)
    elapsed=time.perf_counter()-started
    summary=summarize(existing)
    summary.update(identity)
    summary.update({'draws':n,'new_draws_this_execution':len(tasks),'execution_seconds':elapsed,
                    'python_version':platform.python_version(),'workers':workers,
                    'valuation_date':CFG['valuation_date'],'target_date':CFG['target_date'],
                    'scope':'Standalone; conditional on stated analyst priors and fixed policies. No current-market-return claim.'})
    dump(folder/'summary.json',summary)
    compressed=gzip.compress(records_path.read_bytes(),mtime=0)
    (folder/'draws.jsonl.gz').write_bytes(compressed)
    print(json.dumps({'scenario':name,'draws':n,'seconds':elapsed,'expected_target':summary['expected_target'],'outcomes':summary['clinical_outcomes_target']}),flush=True)
    return summary

def outcome_summary(records,mode='shared_gate',date='target'):
    p=engine.resolve(CFG)
    weights=engine.clinical_weights(p,mode)
    pairs=[(r['states'][s][date]['price'],weights[s]/len(records)) for r in records for s in weights]
    return {'mean':math.fsum(x*w for x,w in pairs),'median':weighted_quantile(pairs,.5),
            'p10':weighted_quantile(pairs,.1),'p90':weighted_quantile(pairs,.9),
            'probability_zero_equity':math.fsum(w for x,w in pairs if x==0),
            'clinical_weights':weights}

def summarize(records):
    expected=[r['expected_target'] for r in records]
    ranked_y=ranks(expected)
    associations=[]
    for k in SPEC['priors']:
        rho=correlation(ranks([r['inputs'][k] for r in records]),ranked_y)
        associations.append({'parameter':k,'spearman':rho})
    associations.sort(key=lambda x:abs(x['spearman'] or 0),reverse=True)
    outcome=outcome_summary(records)
    require(abs(outcome['mean']-statistics.fmean(expected))<1e-9,'Expected and outcome means diverge')
    states={s:{'target_price':describe([r['states'][s]['target']['price'] for r in records]),
               'future_net_funding_today_usdm':describe([r['states'][s]['valuation']['future_funding_net_usdm'] for r in records])}
            for s in engine.STATES}
    convergence=[]
    for n in [64,128,256,512,1024,2048,4096,8192]:
        if n<=len(records):
            convergence.append({'draws':n,'expected_target':describe(expected[:n]),'clinical_outcomes_target':outcome_summary(records[:n])})
    # Batch-to-batch variation uses independent commercial draws as clusters;
    # never treats the four clinical states as four independent simulations.
    b=min(8,len(records)//16)
    batch_se={}
    if b>=2:
        blocks=[records[i::b] for i in range(b)]
        for metric,q in [('p10',.1),('median',.5),('p90',.9)]:
            x=[quantile([r['expected_target'] for r in block],q) for block in blocks]
            y=[outcome_summary(block)[metric] for block in blocks]
            batch_se[metric]={'expected_target':statistics.stdev(x)/math.sqrt(b),'clinical_outcomes_target':statistics.stdev(y)/math.sqrt(b)}
    stability=None
    if len(convergence)>1:
        a,z=convergence[-2:]
        comparisons=[]
        for family in ['expected_target','clinical_outcomes_target']:
            for metric in ['mean','p10','median','p90']:
                change=abs(z[family][metric]-a[family][metric]);limit=max(.5,.02*abs(z[family][metric]))
                comparisons.append({'family':family,'metric':metric,'absolute_change':change,'tolerance':limit,'within_tolerance':change<=limit})
        stability={'prefixes':[a['draws'],z['draws']],'rule':'Absolute movement <= max($0.50, 2% of later estimate); a diagnostic, not proof of convergence.','comparisons':comparisons,'all_within_tolerance':all(x['within_tolerance'] for x in comparisons)}
    return {'expected_target':describe(expected),'value_today':describe([r['value_today'] for r in records]),
            'clinical_outcomes_target':outcome,'clinical_outcomes_today':outcome_summary(records,date='valuation'),
            'both_success_peak_usdm':describe([r['both_success_peak']['recognized_revenue_usdm'] for r in records]),
            'both_success_2030_IR_usdm':describe([r['annual_both_success']['2030']['IR_recognized'] for r in records]),
            'both_success_2030_XR_usdm':describe([r['annual_both_success']['2030']['XR_recognized'] for r in records]),
            'states':states,'rank_associations_prior_dependent_not_causal':associations,
            'clinical_dependence_reweighting':{m:outcome_summary(records,m) for m in ['independent','perfect_positive']},
            'prefix_convergence':convergence,'prefix_stability':stability,'batch_quantile_mc_se_approximate':batch_se,
            'max_errors':{k:max(r['max_errors'][k] for r in records) for k in records[0]['max_errors']},
            'invalid_draws':0,'failed_draw_policy':'Fail immediately, log inputs; never clip or replace an invalid draw.'}

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--draws',type=int,default=512)
    parser.add_argument('--seed',type=int,default=SPEC['seed'])
    parser.add_argument('--workers',type=int,default=4)
    parser.add_argument('--access',choices=SPEC['access_strata'],default='working')
    parser.add_argument('--structural',choices=SPEC['structural_strata'],default='primary')
    parser.add_argument('--dependence',choices=SPEC['dependence_modes'],default='primitive_independent')
    parser.add_argument('--output',type=Path,default=HERE/'results')
    args=parser.parse_args()
    require(args.draws>0 and args.workers>0,'Positive draw count and workers required')
    before=hashes()
    batch(args.draws,args.seed,args.access,args.structural,args.dependence,args.workers,args.output)
    require(before==hashes(),'Source files changed during calculation')

if __name__=='__main__':
    main()
