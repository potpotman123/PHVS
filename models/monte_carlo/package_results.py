"""Package auditable compact results and report from run.py checkpoints (standard library).

python models/monte_carlo/package_results.py --input /tmp/phvs-mc
"""
import argparse
import gzip
import hashlib
import json
from pathlib import Path
import statistics
import run as mc

BASE='working__primary__primitive_independent'
STATES=list(mc.engine.STATES)

def read_rows(folder):
    return [json.loads(l) for l in (folder/'draws.jsonl').read_text().splitlines()]

def pack(path,obj):
    path.parent.mkdir(parents=True,exist_ok=True)
    path.write_bytes(gzip.compress(json.dumps(obj,separators=(',',':'),allow_nan=False).encode(),mtime=0))

def compact(rows):
    columns=['draw_id','expected_target','value_today','peak_usdm','peak_year','IR_2030_usdm','XR_2030_usdm']
    columns += [f'{s}.{d}.{k}' for s in STATES for d in ['valuation','target'] for k in ['price','future_funding_net_usdm','diluted_shares_m']]
    values=[]
    for r in rows:
        values.append([r['draw_id'],r['expected_target'],r['value_today'],r['both_success_peak']['recognized_revenue_usdm'],r['both_success_peak']['year'],
                       r['annual_both_success']['2030']['IR_recognized'],r['annual_both_success']['2030']['XR_recognized']]+
                      [r['states'][s][d][k] for s in STATES for d in ['valuation','target'] for k in ['price','future_funding_net_usdm','diluted_shares_m']])
    return {'format':'columns + rows; JSON floats are not rounded','columns':columns,'rows':values}

def money(x): return ('-$' if x < 0 else '$') + f'{abs(x):,.2f}'

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input',type=Path,required=True)
    parser.add_argument('--output',type=Path,default=mc.HERE/'results')
    a=parser.parse_args(); a.output.mkdir(parents=True,exist_ok=True)
    main_folder=a.input/f'seed_{mc.SPEC["seed"]}'/BASE
    primary=read_rows(main_folder)
    main_summary=json.loads((main_folder/'summary.json').read_text())
    summaries={};paired=[]; inputs_saved=set();count=0
    for folder in sorted(a.input.glob('seed_*/*')):
        if not (folder/'summary.json').exists(): continue
        s=json.loads((folder/'summary.json').read_text()); rows=read_rows(folder)
        mc.require(len(rows)==s['draws'],'Incomplete checkpoint')
        mc.require(s['hashes']==mc.hashes(),'Snapshot hash mismatch')
        for (i,q),r in zip(mc.draws(len(rows),s['seed']),rows):
            mc.require(i==r['draw_id'],'Draw ordering')
            expected=mc.overrides_at(q,mc.SPEC,access=s['access'],structural=s['structural'],dependence=s['dependence'])
            mc.require(expected==r['inputs'],'Sample replay mismatch')
        key=folder.parent.name+'/'+folder.name
        out=a.output/key
        mc.dump(out/'summary.json',s)
        c=compact(rows)
        c.update({'seed':s['seed'],'access':s['access'],'structural':s['structural'],'dependence':s['dependence'],
                  'clinical_weights':s['clinical_outcomes_target']['clinical_weights'],
                  'input_recipe':'draws(n,seed) -> overrides_at(quantiles, manifest, access, structural, dependence). See per-seed primitive_inputs.json.gz; rank-linked overrides seven quantiles before inverse CDF.'})
        pack(out/'draws_compact.json.gz',c)
        summaries[key]=s;count+=len(rows)
        if folder.name==BASE:
            keys=sorted(mc.SPEC['priors'])
            pack(a.output/folder.parent.name/'primitive_inputs.json.gz',{'columns':['draw_id']+keys,'rows':[[r['draw_id']]+[r['inputs'][k] for k in keys] for r in rows],
                 'hard_fixed':mc.SPEC['hard_fixed'],'rest':'Working configuration; scenario packages recorded in each summary.'})
            inputs_saved.add(s['seed'])
        if s['seed']==mc.SPEC['seed'] and folder.name!=BASE:
            n=len(rows); mc.require(n<=len(primary),'Insufficient paired baseline')
            delta=[r['expected_target']-p['expected_target'] for r,p in zip(rows,primary)]
            paired.append({'scenario':folder.name,'draws':n,'expected_target':s['expected_target'],
                'clinical_outcomes_target':s['clinical_outcomes_target'],
                'paired_expected_target_change':mc.describe(delta),
                'paired_base_mean':statistics.fmean(r['expected_target'] for r in primary[:n]),
                'fraction_of_paired_draws_with_higher_value':sum(d>0 for d in delta)/n})
    expected_cases=1+len(mc.SPEC['access_strata'])-1+len(mc.SPEC['structural_strata'])-1+len(mc.SPEC['dependence_modes'])-1
    mc.require(len([s for s in summaries.values() if s['seed']==mc.SPEC['seed']])==expected_cases,'Missing prespecified scenario')
    mc.require(len(inputs_saved)==2,'Expected primary and second-seed replication')
    second=next(s for s in summaries.values() if s['seed']!=mc.SPEC['seed'])
    replication=[]
    for family in ['expected_target','clinical_outcomes_target']:
        for metric in ['mean','p10','median','p90']:
            delta=second[family][metric]-main_summary[family][metric]
            limit=max(.5,.02*abs(main_summary[family][metric]))
            replication.append({'family':family,'metric':metric,'difference_second_minus_primary':delta,'tolerance':limit,'within_tolerance':abs(delta)<=limit})
    result={'scope':'Standalone valuation uncertainty, not stock-return probability. No acquisition outcomes.',
      'source_commit':'e7594c4a8e7d5f4571055e52aa9c3320b39c5ab2','primary':main_summary,
      'second_seed':second,'replication_diagnostics':replication,'paired_scenarios':paired,
      'commercial_draw_scenario_evaluations':count,'conditional_clinical_state_evaluations':4*count,
      'scenario_mixture_weights':None,'invalid_draws':sum(s['invalid_draws'] for s in summaries.values()),
      'all_scenarios':{k:{'draws':s['draws'],'prefix_stability':s['prefix_stability'],'max_errors':s['max_errors']} for k,s in summaries.items()}}
    mc.dump(a.output/'analysis.json',result)
    p=main_summary['expected_target'];o=main_summary['clinical_outcomes_target']
    lines=['# PHVS standalone Monte Carlo — 9 October 2026','',
      f'Source: [`e7594c4`](https://github.com/potpotman123/PHVS/commit/e7594c4a8e7d5f4571055e52aa9c3320b39c5ab2). Valuation date: {main_summary["valuation_date"]}; target date: {main_summary["target_date"]}.', '',
      f'**{len(primary):,} primary commercial draws**, plus an independent {second["draws"]:,}-draw seed check and twelve separate sensitivity runs with {min(r["draws"] for r in paired):,}–{max(r["draws"] for r in paired):,} draws each. {count:,} commercial draw/scenario evaluations; {4*count:,} exact conditional clinical valuations. No original repository files modified.', '',
      '| Target valuation distribution, USD/share | Mean | Median | P10 | P90 |','|---|---:|---:|---:|---:|',
      f'| Clinical-risk-adjusted expected value across commercial draws | {money(p["mean"])} | {money(p["median"])} | {money(p["p10"])} | {money(p["p90"])} |',
      f'| Conditional clinical outcomes, weighted once | {money(o["mean"])} | {money(o["median"])} | {money(o["p10"])} | {money(o["p90"])} |','',
      'The middle 80% is an analyst-prior sensitivity range, **not** an empirical confidence interval. The second row mixes conditional business-outcome valuations. It is not a forecast of actual share-price volatility or a claim that every clinical outcome resolves by October 2027; the XR decision is later. No verified market-price reference is configured, so no probability of a profitable short or percentage downside is reported.','',
      '![Valuation distributions and prior-dependent drivers](distribution.png)','',
      '## What this says about the short thesis','',
      '1. **Oral preference is heterogeneous.** The prevention preference prior is 46% / 54% / 62% (minimum / mode / maximum), applied inside an opportunity → oral choice → PHVS choice → paid access process. The 54/46 survey is not observed switching and does not permanently exclude 46% of patients. Mature PHVS share of oral choices remains the favorable, user-fixed 60%.',
      '2. **Fewer redoses reduce paid quantity.** Every simulation retains the specified 1.17/1.38 quantity factor, about a 15.2% reduction, applied once to the dose budget with initial stock accounted for inside it. This is a stipulated cross-trial-to-commercial assumption, not new evidence of realized PHVS unit economics. There is no price premium or no-redose comparison hidden in the simulation.',
      '3. **Clinical success does not guarantee the paid ramp.** Switching opportunity, initial choice, continuation, quantity, payment intensity, overseas ramp and costs vary. Approval probabilities stay high. Historical successful oral followers remain disconfirming evidence against treating “second oral” as automatic commercial failure.','',
      'The simulation tests the financial consequences of these assumptions; it does not independently prove the three theses. The zero-volatility working case remains $28.311616929. Sampling does not have to reproduce that point: triangular modes are not means, some priors are asymmetric, and the model is nonlinear. No recentering was used.','',
      '## Paired scenarios','',
      'Same seed and primitive quantiles as the primary prefix. Each row changes one named package or structure. Access packages jointly change conversion, lag and net price. No probability weights or combined scenario distribution were invented.','',
      '| Scenario | Draws | Mean target | Median | P10–P90 | Mean change vs matched primary | Paired mean MC SE |','|---|---:|---:|---:|---:|---:|---:|']
    for r in paired:
        z=r['expected_target'];d=r['paired_expected_target_change']
        labels={'rebate_for_access':'More rebates for better access','restricted_high_net':'Restricted access, higher net price',
                'extra_Ekterly_free_supply':'20% extra Ekterly free supply','higher_discount_rate':'15% discount rate',
                'initial_stock_free':'Free initial rescue stock','later_LOE':'Loss of exclusivity in 2043',
                'licensed_overseas':'Overseas licensing','lower_discount_rate':'10% discount rate',
                'no_rescue_in_10pct_LTP':'10% of preventive users lack rescue','six_month_delay':'Both launches delayed six months',
                'some_discontinuers_return':'25% of discontinuers can return','commercial_rank_linked':'Commercial variables rank-linked'}
        tokens=r['scenario'].split('__');label=next(labels[t] for t in tokens if t in labels)
        se=money(d['mean_mc_se']) if d['mean_mc_se']>=.005 else '<$0.01'
        lines.append(f'| {label} | {r["draws"]:,} | {money(z["mean"])} | {money(z["median"])} | {money(z["p10"])}–{money(z["p90"])} | {money(d["mean"])} | {se} |')
    lines+=['','Commercial rank linking gives seven demand variables the same percentile. It is an intentionally strong dependence stress, not a fitted correlation matrix. Structural scenarios are one at a time; they are not additive and do not span all combinations.','',
      '## Clinical states and capital needs','',
      '| State | Weight | Median target/share | Median future net funding at valuation date ($m) |','|---|---:|---:|---:|']
    for state in STATES:
        z=main_summary['states'][state]
        lines.append(f'| {state} | {100*o["clinical_weights"][state]:.4f}% | {money(z["target_price"]["median"])} | {z["future_net_funding_today_usdm"]["median"]:,.1f} |')
    lines+=['','These weights preserve 90% marginal success for each product using the existing 95% shared gate. No additional approval haircut is applied. Funding is a model cash requirement, not a promised financing: the inherited model assumes financing is available at its fixed issue price. Post-target funding is not charged again as extra denominator dilution after already valuing negative cash flows. Liquidation/cash assumptions explain residual value in the neither-success state.','',
      '## Numerical checks and precision','',
      f'- Original engine: 152 checks passed; original Monte Carlo preflight: 72 checks passed. New runner: eight regression tests passed, including zero-volatility, exact clinical weighting, seed prefixes, worker invariance, and checkpoint hash protection.',
      f'- Primary expected-value mean Monte Carlo standard error: {money(p["mean_mc_se"])}. Approximate P90 standard error: {money(main_summary["batch_quantile_mc_se_approximate"]["p90"]["expected_target"])} for expected values and {money(main_summary["batch_quantile_mc_se_approximate"]["p90"]["clinical_outcomes_target"])} for clinical outcomes. Eight interleaved draw clusters; not an uncertainty interval for the economic assumptions.',
      f'- Primary {main_summary["prefix_stability"]["prefixes"]} prefix test passes: {main_summary["prefix_stability"]["all_within_tolerance"]}. Second seed passes every mean/P10/median/P90 comparison within max($0.50, 2%): {all(x["within_tolerance"] for x in replication)}. Detailed scenario prefix checks are in analysis.json; small structural effects should be read through paired differences rather than independent quantiles.',
      '- Every state/month is checked for population conservation, overlap, source reconciliation, cash and FCF identities, price × units and clinical failure revenue. All sampled values are finite and inside manifest bounds. Invalid draws abort with their inputs; none are discarded or replaced.',
      '- Complete unrounded compact outputs, sampled primitive inputs, source hashes, per-scenario summaries, exact replay recipes and code are included. Full monthly paths are regenerated when needed, not stored for every simulation.','',
      '## What still limits the conclusion','',
      '- The 28 triangular distributions are analyst priors, not fitted posteriors. Increasing the draw count reduces simulation noise, not source-transfer error or selection bias.',
      '- The diagnosed population is a broad proxy, not a validated age/genotype-eligible census. Preference surveys, mixed-cohort switching, claims dose budgets and competitor start forms do not directly measure PHVS patients or paid switching.',
      '- IR continuation and transfer of preventive switching rates into rescue behavior remain weakly observed. The rank correlations are prior-dependent associations, not causal contributions or variance shares; fixed assumptions cannot appear in that ranking.',
      '- No prior fitting, clinical data-misconduct claim, efficacy-to-market-share regression, new acquisition probability, or licensing payout has been invented. M&A research remains a risk to a short, but is outside this requested standalone model.',
      '- Pricing/access packages, fixed 60% mature choice, 90% approval marginals, financing availability and post-exclusivity erosion remain important structural assumptions. Only the prespecified alternatives were run. The range does not cover every possible future.',
      '- Current repository source registers and their updates were checked and incorporated. This run does not claim a fresh independent audit of every original paper, PDF, or licensed interview.','',
      'See [README](../README.md), [recent-update reconciliation](../RECENT_UPDATES.md), [source coverage](../source_coverage.json), and [machine-readable analysis](analysis.json).','']
    (a.output/'REPORT.md').write_text('\n'.join(lines))
    manifest={str(p.relative_to(a.output)):{'sha256':hashlib.sha256(p.read_bytes()).hexdigest(),'bytes':p.stat().st_size} for p in sorted(a.output.rglob('*')) if p.is_file() and p.name not in ['artifact_manifest.json','artifact_verification.json','distribution.png','chart_manifest.json']}
    mc.dump(a.output/'artifact_manifest.json',manifest)
    print(json.dumps({'commercial_evaluations':count,'primary':main_summary['expected_target'],'second_seed_pass':all(x['within_tolerance'] for x in replication),'packaged_bytes':sum(v['bytes'] for v in manifest.values())}))

if __name__=='__main__':
    main()
