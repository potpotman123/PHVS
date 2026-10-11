"""Reproduce the primary Monte Carlo draws and accumulate the expected (mean) annual cash-flow path.
Discounting is linear, so a DCF of these expected cash flows equals the mean of the per-draw DCF values
(except for the per-share step). Output: models/excel/mc_expected_path.json. Run from the repository root.
"""
import json, sys, time, datetime as dt
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path
ROOT = Path(__file__).resolve().parents[2]; sys.path.insert(0, str(ROOT / 'models/finalized'))
import engine as E
from mc_prepare import draws, overrides_at
CFG = E.read('inputs.json'); SPEC = E.read('monte_carlo_priors.json')
import os, copy
CASE = {'overrides': {}, 'priors': {}, 'label': 'research'}
if os.environ.get('MC_CASE'):
    CASE = json.loads(Path(os.environ['MC_CASE']).read_text()); SPEC = copy.deepcopy(SPEC)
    for k, v in CASE.get('priors', {}).items(): SPEC['priors'][k].update(v)
YEARS = list(range(2026, CFG['end_year'] + 1)); STATES = ['both', 'IR_only', 'XR_only', 'neither']
FLOW = ['revenue', 'IR_recognized', 'XR_recognized', 'IR_US', 'XR_US', 'manufacture', 'license', 'rd', 'commercial', 'ga', 'depreciation', 'operating_income', 'tax', 'net_income', 'capex', 'delta_nwc', 'fcf', 'funding_net', 'funding_fee', 'distribution', 'IR_units', 'XR_units', 'xr_starts', 'ir_access_starts', 'xr_patient_years', 'ir_member_years', 'xr_rx', 'ir_rx']
STOCK = ['cash_end', 'shares_end', 'xr_active', 'ir_access_active', 'nwc_end', 'ppe_end', 'equity_end']
VAL = ['price', 'equity_usdm', 'operating_pv_usdm', 'cash_usdm', 'terminal_pv_usdm', 'financing_fee_pv_usdm', 'basic_shares_m', 'diluted_shares_m', 'future_funding_net_usdm']

def task(args):
    i, q = args
    ov = {**CASE.get('overrides', {}), **overrides_at(q, SPEC)}; res = E.evaluate(CFG, overrides=ov, retain=True); p = E.resolve(CFG, overrides=ov)
    tgt = dt.date.fromisoformat(CFG['target_date']); peak = res['conditional_peak']; a30 = next(r for r in res['annual_both_success'] if r['year'] == 2030)
    out = {'i': i, 'expected_target': res['expected_12month_value'], 'value_today': res['risk_adjusted_value_today'], 'weights': res['clinical_weights'], 'states': {}, 'inputs': {k: ov[k] for k in SPEC['priors']}, 'peak': peak['recognized_revenue_usdm'], 'peak_year': peak['year'], 'ir2030': a30['IR_recognized'], 'xr2030': a30['XR_recognized'], 'state_target': {s_: res['conditional_valuations'][s_]['target']['price'] for s_ in STATES}, 'state_funding': {s_: res['conditional_valuations'][s_]['valuation']['future_funding_net_usdm'] for s_ in STATES}}
    for s in STATES:
        rows = res['monthly_by_clinical_state'][s]; acc = {y: {k: 0.0 for k in FLOW + ['fcf_after_target', 'pv_target']} for y in YEARS}; stock = {y: {} for y in YEARS}
        for r in rows:
            y = r['year']; a = acc[y]
            for k in FLOW: a[k] += r.get(k, 0.0)
            for k in STOCK: stock[y][k] = r.get(k, 0.0)
            frac = E.elapsed_fraction(r['month'], tgt); end = E.dates(E.month_index(r['month']))[1]
            if end >= tgt:
                t = (end - tgt).days / 365.25; a['fcf_after_target'] += r['fcf'] * (1 - frac); a['pv_target'] += r['fcf'] * (1 - frac) / (1 + p['discount_rate']) ** t
        out['states'][s] = {'annual': {y: {**acc[y], **stock[y]} for y in YEARS}, 'target': {k: res['conditional_valuations'][s]['target'][k] for k in VAL}, 'valuation': {k: res['conditional_valuations'][s]['valuation'][k] for k in VAL}}
    return out

if __name__ == '__main__':
    n = int(sys.argv[1]) if len(sys.argv) > 1 else 4096; workers = int(sys.argv[2]) if len(sys.argv) > 2 else 4
    t0 = time.time(); sums = None; cnt = 0; et = []; vt = []; perdraw = []
    with ProcessPoolExecutor(max_workers=workers) as ex:
        for out in ex.map(task, list(draws(n, SPEC['seed'])), chunksize=8):
            cnt += 1; et.append(out['expected_target']); vt.append(out['value_today']); perdraw.append({k: out[k] for k in ('i', 'expected_target', 'value_today', 'inputs', 'peak', 'peak_year', 'ir2030', 'xr2030', 'state_target', 'state_funding')})
            if sums is None:
                sums = {'weights': dict(out['weights']), 'states': {s: {'annual': {y: dict(out['states'][s]['annual'][y]) for y in YEARS}, 'target': dict(out['states'][s]['target']), 'valuation': dict(out['states'][s]['valuation'])} for s in STATES}}
            else:
                for s in STATES:
                    for y in YEARS:
                        for k, v in out['states'][s]['annual'][y].items(): sums['states'][s]['annual'][y][k] += v
                    for d in ('target', 'valuation'):
                        for k, v in out['states'][s][d].items(): sums['states'][s][d][k] += v
            if cnt % 128 == 0:
                (ROOT / 'models/excel/mc_expected_path.progress').write_text(f'{cnt}/{n} draws, {time.time()-t0:.0f}s, running mean target {sum(et)/cnt:.6f}\n')
    mean = {'draws': cnt, 'seed': SPEC['seed'], 'case': CASE.get('label', 'research'), 'case_overrides': CASE.get('overrides', {}), 'case_priors': CASE.get('priors', {}), 'perdraw': perdraw, 'seconds': time.time() - t0, 'mean_expected_target': sum(et) / cnt, 'mean_value_today': sum(vt) / cnt, 'weights': sums['weights'], 'years': YEARS, 'states': {}}
    for s in STATES:
        mean['states'][s] = {'annual': {y: {k: v / cnt for k, v in sums['states'][s]['annual'][y].items()} for y in YEARS}, 'target': {k: v / cnt for k, v in sums['states'][s]['target'].items()}, 'valuation': {k: v / cnt for k, v in sums['states'][s]['valuation'].items()}}
    w = mean['weights']
    mean['expected'] = {'annual': {y: {k: sum(w[s] * mean['states'][s]['annual'][y][k] for s in STATES) for k in mean['states']['both']['annual'][y]} for y in YEARS},
                        'target': {k: sum(w[s] * mean['states'][s]['target'][k] for s in STATES) for k in VAL}, 'valuation': {k: sum(w[s] * mean['states'][s]['valuation'][k] for s in STATES) for k in VAL}}
    et.sort(); mean['check'] = {'p10': et[int(0.1 * cnt)], 'median': et[cnt // 2], 'p90': et[int(0.9 * cnt)], 'published_mean': 26.378087943701694}
    outp = Path(os.environ.get('MC_OUT', str(ROOT / 'models/excel/mc_expected_path.json'))); outp.write_text(json.dumps(mean))
    print(json.dumps({k: mean[k] for k in ('draws', 'seconds', 'mean_expected_target', 'mean_value_today', 'check')}, indent=1))
