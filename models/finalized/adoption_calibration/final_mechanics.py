"""Auditable starting-stock, adoption-timing and net-inflow mechanics.

No financial-engine or Monte Carlo integration is claimed by this module.
"""
import json
from pathlib import Path
from choice_share_schedule import oral_choice_share

HERE = Path(__file__).resolve().parent
INITIAL = {'orladeyo': 1700., 'other_prevention': 4080.,
           'acute_only': 2720., 'diagnosed_untreated': 1500.}
GROWTH = 0.015
# July 2025-February 2026. Source gives two-month increments; equal split assumed.
FORMS = [230., 230., 238.5, 238.5, 190.5, 190.5, 192., 192.]
CASES = {
    'lower_Ekterly': {'future_forms_monthly': 100., 'paid_conversion': .65,
                     'annual_continuation': .80, 'unique_form_fraction': .90},
    'working': {'future_forms_monthly': 192., 'paid_conversion': .80,
                'annual_continuation': .90, 'unique_form_fraction': 1.},
    'higher_Ekterly': {'future_forms_monthly': 250., 'paid_conversion': .90,
                      'annual_continuation': .95, 'unique_form_fraction': 1.},
}


def ekterly_stock(config):
    # End April 2027, before assumed PHVS May launch:22months since July2025.
    # One-month conversion delay. April forms are pending, not active.
    forms = FORMS + [config['future_forms_monthly']] * 14
    survival = config['annual_continuation'] ** (1 / 12)
    active = 0.
    rows = []
    for m, f in enumerate(forms):
        active *= survival
        new_paid = (forms[m - 1] * config['paid_conversion'] *
                    config['unique_form_fraction']) if m else 0.
        active += new_paid
        rows.append({'month_index': m, 'forms': f, 'paid_starts': new_paid,
                     'active_end': active})
    return active, rows


def starting_joint_stock(ekterly, growth=GROWTH, ltp_no_rescue_fraction=0.):
    # July2026 baseline to beginningMay2027 =10monthly population increments.
    pools = {k: v * (1 + growth) ** (10 / 12) for k, v in INITIAL.items()}
    without = pools['diagnosed_untreated'] + ltp_no_rescue_fraction * (
        pools['orladeyo'] + pools['other_prevention'])
    rescue_eligible = sum(pools.values()) - without
    assert ekterly <= rescue_eligible
    joint = {}
    for origin, n in pools.items():
        no = n if origin == 'diagnosed_untreated' else (
            n * ltp_no_rescue_fraction if origin != 'acute_only' else 0.)
        oral = (n - no) / rescue_eligible * ekterly
        joint[origin] = {'Ekterly': oral, 'injectable_rescue': n - no - oral,
                         'no_current_rescue': no}
        assert abs(sum(joint[origin].values()) - n) < 1e-8
    columns = {k: sum(v[k] for v in joint.values()) for k in next(iter(joint.values()))}
    assert abs(sum(columns.values()) - sum(pools.values())) < 1e-8
    assert abs(columns['Ekterly'] - ekterly) < 1e-8
    return {'population': sum(pools.values()), 'prevention_pools': pools,
            'rescue_pools': columns, 'joint_cells': joint}


def cohort_timing_check(case='working', extra_hazard_floor=1., growth=GROWTH):
    # Isolate timing for1000injectablepreventive users with burden pooled.
    # This is an acquisition/access ledger, not active-treatment retention/revenue.
    initial_rx = (50 / 328) * .54 / .8
    available, pending, cumulative_paid = 1000., 0., 0.
    cooldown = {}
    total = 1000.
    years = []
    annual = dict(rx=0., paid=0., net_additions=0.)
    for m in range(120):
        increment = total * ((1 + growth) ** (1 / 12) - 1)
        total += increment
        available += increment + cooldown.pop(m, 0.)
        paid = pending
        cumulative_paid += paid
        progress = min(m / 60, 1.)
        switch_factor = 1 + (extra_hazard_floor - 1) * progress
        annual_rx_probability = initial_rx * oral_choice_share(case, m) * switch_factor
        q = 1 - (1 - annual_rx_probability) ** (1 / 12)
        attempts = available * q
        available -= attempts
        pending = attempts * .8
        cooldown[m + 12] = cooldown.get(m + 12, 0.) + attempts * .2
        assert abs(available + pending + sum(cooldown.values()) + cumulative_paid - total) < 1e-7
        annual['rx'] += attempts
        annual['paid'] += paid
        annual['net_additions'] += increment
        if m % 12 == 11:
            years.append(dict(year=m // 12 + 1, **annual, available_end=available,
                              cumulative_paid=cumulative_paid, population=total))
            annual = dict(rx=0., paid=0., net_additions=0.)
    return years


def build():
    stocks = {}
    for case, config in CASES.items():
        active, monthly = ekterly_stock(config)
        stocks[case] = {'assumptions': config, 'launch_mix': starting_joint_stock(active),
                        'ekterly_projection': monthly}
    base_stock = stocks['working']['launch_mix']['rescue_pools']['Ekterly']
    no_rescue_stress = starting_joint_stock(base_stock, ltp_no_rescue_fraction=.10)
    timing = {name: cohort_timing_check(extra_hazard_floor=factor)
              for name, factor in [('depletion_only', 1.), ('extra_taper', .5),
                                   ('expanding_consideration', 1.25)]}
    zero = cohort_timing_check(growth=0.)
    assert all(abs(x['population'] - 1000.) < 1e-8 for x in zero)
    assert abs(timing['depletion_only'][0]['net_additions'] - 15.) < 1e-8
    output = {'as_of': '2026-10-09', 'PHVS_IR_launch': '2027-05', 'stocks': stocks,
              'ten_percent_LTP_no_rescue_stress': no_rescue_stress,
              'twenty_percent_extra_Ekterly_stock_for_free_supply_stress': starting_joint_stock(base_stock * 1.2),
              'timing_per1000_initial_injectable_prevention': timing,
              'net_growth': GROWTH,
              'inflow_rule': 'Add only N*((1+g)^(1/12)-1) once per month; allocate to existing joint origin proportions. Separate NEW conditional-choice rows disabled, not additive.',
              'validation': 'Stock row/column sums; Ekterly capacity; monthly pending/failed-access/paid conservation; zero-growth case; annual1.5%addition verified.',
              'status': 'Three mechanics resolved as reproducible underwriting choices; no updated DCF or full segmented-engine integration.'}
    (HERE / 'final_mechanics_results.json').write_text(json.dumps(output, indent=2) + '\n')
    for case, data in stocks.items():
        print(case, {k: round(v) for k, v in data['launch_mix']['rescue_pools'].items()})
    for row in timing['depletion_only']:
        if row['year'] in (1, 3, 5, 6, 10):
            print('year', row['year'], 'new paid', round(row['paid'], 1))
    return output


if __name__ == '__main__':
    build()
