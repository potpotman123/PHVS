"""Source-specific stock and propensity translation. Authoritative values in inputs.json."""
def oral_share(p, age):
    f=min(1.,max(0.,age)/p['choice_months'])
    return p['oral_launch_share']+(p['oral_mature_share']-p['oral_launch_share'])*f


def joint_pools(cfg,p,start,launch):
    c=cfg['segmented_constants']
    nmonths=launch-(2025*12+6)
    observed=c['ekterly_monthly_forms_Jul2025_Feb2026']
    forms=(observed+[p['ekterly_future_forms']]*max(0,nmonths-len(observed)))[:nmonths]
    active=0.
    for m in range(nmonths):
        active*=p['ekterly_continuation']**(1/12)
        if m:active+=forms[m-1]*p['ekterly_paid_conversion']*p['ekterly_unique_fraction']
    active*=p['ekterly_free_stock_multiplier']
    # Project absolute Ekterly stock to launch, then express fixed origin shares
    # at baseline. Normal growth restores launch counts; no extra diagnostic inflow.
    factor=(1+p['growth'])**((launch-start)/12)
    prev={k:p['population']*w for k,w in cfg['pool_weights'].items()}
    no={k:n*(1 if k=='diagnosed_untreated' else p['ltp_no_rescue_fraction'] if k in ('orladeyo','other_prevention') else 0) for k,n in prev.items()}
    eligible=sum(prev.values())-sum(no.values())
    oral=active/factor
    assert 0<=oral<=eligible, 'Projected Ekterly stock exceeds available rescue population'
    pools={}
    for k,n in prev.items():
        e=(n-no[k])/eligible*oral
        for r,v in [('ekterly',e),('injectable',n-no[k]-e),('none',no[k])]:
            if v>0:pools[k+'|'+r]=v
    assert abs(sum(pools.values())-p['population'])<1e-7
    return pools


def rates(cfg,p,pools,age_ir,age_xr):
    c=cfg['segmented_constants']; norm=c['reference_paid_conversion']
    xrchoice=oral_share(p,age_xr); irchoice=oral_share(p,age_ir)
    opportunity=p['category_switch']
    inj=opportunity*p['prevention_oral_preference']*xrchoice
    weights=cfg['pool_weights']
    acute=weights['other_prevention']/weights['acute_only']*inj*c['entrant_rescue_to_prevention_ratio']
    naive=weights['acute_only']/weights['diagnosed_untreated']*acute*c['entrant_naive_to_rescue_ratio']*p['untreated_xr_reach']
    xr={'orladeyo':opportunity*p['oral_destination_capture']/norm,'other_prevention':inj/norm,'acute_only':acute/norm,'diagnosed_untreated':naive/norm}
    ekannual=1-(1-c['ekterly_first8_forms']/c['ekterly_launch_addressable'])**(12/8)
    ir={'ekterly':opportunity*p['ir_oral_switch_relative']*p['oral_destination_capture']/norm,
        'injectable':ekannual*p['ir_injectable_opportunity']*irchoice,
        'none':p['ir_no_rescue_opportunity']*c['rescue_oral_preference']*irchoice}
    def factor(age):
        f=min(1.,max(0.,age)/p['choice_months'])
        return (1+(p['switch_mature_multiplier']-1)*f)*(p['hazard_year6_multiplier'] if age>=60 else 1)
    return ({k:xr[k.split('|')[0]]*factor(age_xr) for k in pools},
            {k:ir[k.split('|')[1]]*factor(age_ir) for k in pools})
