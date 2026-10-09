"""Sampler and preflight interfaces only; production Monte Carlo belongs to next agent."""
import math
import random
from engine import read, resolve


def triangular_quantile(u, lo, mode, hi):
    assert 0 <= u <= 1 and lo <= mode <= hi
    if lo == hi:
        return lo
    f=(mode-lo)/(hi-lo)
    if u < f:
        return lo+math.sqrt(u*(hi-lo)*(mode-lo))
    return hi-math.sqrt((1-u)*(hi-lo)*(hi-mode))


def overrides_at(quantiles, manifest, access='working', structural='primary', dependence='primitive_independent', zero_volatility=False):
    assert access in manifest['access_strata'] and structural in manifest['structural_strata']
    assert dependence in manifest['dependence_modes']
    out=dict(manifest['hard_fixed'])
    for k,v in manifest['priors'].items():
        u=quantiles[k]
        if dependence=='commercial_rank_linked' and k in manifest['commercial_rank_linked']:
            u=quantiles['category_switch']
        x=v['mode'] if zero_volatility else triangular_quantile(u,v['min'],v['mode'],v['max'])
        out[k]=int(round(x)) if v['round_integer'] else x
    out.update(manifest['access_strata'][access])
    out.update(manifest['structural_strata'][structural])
    for k,v in manifest['hard_fixed'].items():
        assert out[k]==v, 'A scenario changed a mandatory fixed assumption'
    return out


def draws(n, seed=None):
    manifest=read('monte_carlo_priors.json')
    rng=random.Random(manifest['seed'] if seed is None else seed)
    for i in range(n):
        yield i,{k:rng.random() for k in sorted(manifest['priors'])}


def clinical_outputs(result):
    """Keep expectation and categorical outcomes distinct; no second PoS haircut."""
    states=[{'clinical_state':k,'probability':w,'target_price':result['conditional_valuations'][k]['target']['price']} for k,w in result['clinical_weights'].items()]
    expected=sum(s['probability']*s['target_price'] for s in states)
    assert abs(expected-result['expected_12month_value'])<1e-9
    return {'expected_target':expected,'clinical_outcomes':states}
