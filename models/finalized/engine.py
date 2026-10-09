"""Independent PHVS underwriting engine; standard library only. See README.txt.
Cash amounts USD millions; quantities fractional people, doses, patient-years.
No broker forecast, spreadsheet formula or fitted S-curve is imported.
"""
import calendar
import copy
import datetime as dt
import json
import math
from pathlib import Path

BASE = Path(__file__).resolve().parent

def read(name):
    return json.loads((BASE / name).read_text())

def write(name, value):
    (BASE / name).write_text(json.dumps(value, indent=2, allow_nan=False) + '\n')

def month_index(s):
    return int(s[:4]) * 12 + int(s[5:7]) - 1

def month_string(m):
    return f'{m // 12:04}-{m % 12 + 1:02}'

def dates(m):
    y, mo = m // 12, m % 12 + 1
    return dt.date(y, mo, 1), dt.date(y, mo, calendar.monthrange(y, mo)[1])

def resolve(cfg, case='working', overrides=None):
    p = {k: v['value'] for k, v in cfg['parameters'].items()}
    p.update(cfg['cases'][case])
    p.update(overrides or {})
    assert p['shared_success'] >= max(p['ir_pos'], p['xr_pos'])
    assert p['discount_rate'] > cfg['loe']['tail_growth']
    return p

class Population:
    """Each product has overlapping membership but conserves its own shared population.
    Origin labels remain fixed; re-entry changes attempt counts, never disease count.
    """
    def __init__(self, pools):
        self.total = dict(pools)
        self.available = dict(pools)
        self.retired = {k: 0.0 for k in pools}
        self.pending, self.cooldown, self.cohorts = [], [], []
        self.max_error = 0.0

    def step(self, m, growth, rates, p, first, mature, launched):
        out = {k: dict(rx=0., starts=0., py=0., first_py=0., active=0., exits=0.) for k in self.total}
        for k in self.total:
            n = self.total[k] * growth
            self.total[k] += n
            self.available[k] += n
        keep = []
        for due, k, n in self.cooldown:
            if due <= m:
                self.available[k] += n
            else:
                keep.append((due, k, n))
        self.cooldown = keep
        if launched:
            for k in self.total:
                assert 0 <= rates[k] < 1
                n = self.available[k] * (1 - (1 - rates[k]) ** (1 / 12))
                self.available[k] -= n
                out[k]['rx'] = n
                self.pending.append((m + int(p['delay']), k, n * p['conversion']))
                self.cooldown.append((m + int(p['retry_months']), k, n * (1 - p['conversion'])))
        keep = []
        for due, k, n in self.pending:
            if due <= m:
                self.cohorts.append([k, n, 0])
                out[k]['starts'] += n
            else:
                keep.append((due, k, n))
        self.pending = keep
        for c in self.cohorts:
            k, n, age = c
            hazard = -math.log(first if age < 12 else mature)
            survivors = n * math.exp(-hazard / 12)
            py = n / 12 if hazard == 0 else (n - survivors) / hazard
            exits = n - survivors
            out[k]['py'] += py
            out[k]['first_py'] += py if age < 12 else 0
            out[k]['active'] += survivors
            out[k]['exits'] += exits
            self.retired[k] += exits * (1 - p['reentry_fraction'])
            if p['reentry_fraction']:
                self.cooldown.append((m + int(p['reentry_delay']), k, exits * p['reentry_fraction']))
            c[1], c[2] = survivors, age + 1
        for k in self.total:
            accounted = self.available[k] + self.retired[k] + out[k]['active']
            accounted += sum(n for _, origin, n in self.pending + self.cooldown if origin == k)
            error = abs(accounted - self.total[k])
            self.max_error = max(self.max_error, error)
            assert error < 1e-6, (k, m, error)
            assert self.available[k] >= -1e-8
        return out

def erosion(cfg, p, year, product):
    entry = int(p.get('loe_year_' + product, p['loe_year']))
    if year < entry:
        return 1., 1., 1.
    z = cfg['loe']; i = min(year - entry, 3)
    w = z['study_weight']
    r = w * z['study_p'][i] * z['study_q'][i] + (1 - w) * z['rare_revenue'][i]
    q = w * z['study_q'][i] + (1 - w) * z['rare_revenue'][i] / z['rare_p'][i]
    price = r / q
    if p['loe_price_cap']:
        price = min(1., price)
    if year - entry > 3:
        q *= (1 + z['tail_growth']) ** (year - entry - 3)
    return price, q, price * q

def commercial(cfg, p, success=(True, True)):
    start = month_index(cfg['start_month'])
    finish = (cfg['end_year'] + 1) * 12
    launch = {k: month_index(v['commercial_month']) + int(p.get('launch_delay_' + k, 0)) for k, v in cfg['launches'].items()}
    pools = {k: p['population'] * w for k, w in cfg['pool_weights'].items()}
    xr, ir = Population(pools), Population(pools)
    g = (1 + p['growth']) ** (1 / 12) - 1
    rows, history = [], {}
    # Counterfactual pre-generic volume freezes at the last protected calendar-year
    # average. Erosion is applied once; no population growth perpetuity after LOE.
    for m in range(start, finish):
        y = m // 12
        age_ir, age_xr = m - launch['IR'], m - launch['XR']
        ir_factor = p['hazard_year6_multiplier'] if age_ir >= 60 else 1.
        xr_factor = p['hazard_year6_multiplier'] if age_xr >= 60 else 1.
        rates_xr = {k: p['xr_' + k] * xr_factor for k in pools}
        rates_ir = {k: p['ir_rx'] * ir_factor * (p['ir_untreated_relative'] if k == 'diagnosed_untreated' else 1) for k in pools}
        xs = xr.step(m, g, rates_xr, p, p['xr_first_retention'], p['xr_mature_retention'], age_xr >= 0 and success[1])
        ins = ir.step(m, g, rates_ir, p, p['ir_retention'], p['ir_retention'], age_ir >= 0 and success[0])
        r = dict(month=month_string(m), year=y, quarter=f'{y}Q{m % 12 // 3 + 1}',
                 population=sum(xr.total.values()), xr_rx=0., ir_rx=0., xr_starts=0., ir_access_starts=0.,
                 xr_patient_years=0., ir_member_years=0., ir_initial_doses=0., ir_free_initial_doses=0., ir_replacement_doses=0.,
                 xr_active=0., ir_access_active=0., overlap_members=0.)
        for k in pools:
            x, i = xs[k], ins[k]
            overlap = min(1., x['py'] * 12 / xr.total[k])
            migration = p['nonphvs_ltp_shift'] * min(1., max(0, age_ir + 1) / 120)
            baseline = p['ir_doses_ltp'] if k in ('orladeyo', 'other_prevention') else (1 - migration) * p['ir_doses_other'] + migration * p['ir_doses_ltp']
            budget = (1 - overlap) * baseline + overlap * p['ir_doses_ltp'] * p['dual_use']
            initial = p['ir_initial_doses'] * (1 - overlap + overlap * p['dual_use'])
            assert budget >= initial - 1e-10
            r['ir_initial_doses'] += i['starts'] * initial * p['ir_initial_paid_fraction']
            r['ir_free_initial_doses'] += i['starts'] * initial * (1 - p['ir_initial_paid_fraction'])
            r['ir_replacement_doses'] += i['first_py'] * (budget - initial) + (i['py'] - i['first_py']) * budget
            for dest, value in [('xr_rx', x['rx']), ('ir_rx', i['rx']), ('xr_starts', x['starts']), ('ir_access_starts', i['starts']), ('xr_patient_years', x['py']), ('ir_member_years', i['py']), ('xr_active', x['active']), ('ir_access_active', i['active'])]:
                r[dest] += value
            r['overlap_members'] += overlap * i['py'] * 12
        raw_units = {'IR': r['ir_initial_doses'] + r['ir_replacement_doses'], 'XR': r['xr_patient_years'] * p['xr_paid_intensity']}
        history[m] = {**raw_units, 'IR_free': r['ir_free_initial_doses']}
        for prod in ('IR', 'XR'):
            entry = int(p.get('loe_year_' + prod, p['loe_year']))
            loe_p, loe_q, _ = erosion(cfg, p, y, prod)
            units = raw_units[prod]
            if y >= entry:
                units = sum(history[n][prod] for n in range((entry - 1) * 12, entry * 12)) / 12 * loe_q
            price = p['ir_wac_dose'] * p['ir_net_fraction'] if prod == 'IR' else p['xr_wac_28days'] * 365 / 28 * p['xr_net_fraction']
            r[prod + '_price_usd'] = price * loe_p
            r[prod + '_units'] = units
            if prod == 'IR':
                r['IR_free_units'] = r['ir_free_initial_doses'] if y < entry else sum(history[n]['IR_free'] for n in range((entry-1)*12, entry*12))/12*loe_q
            r[prod + '_US'] = units * price * loe_p / 1e6
            r[prod + '_loe_price_index'], r[prod + '_loe_volume_index'] = loe_p, loe_q
        rows.append(r)
    by_month = {month_index(r['month']): r for r in rows}
    # Use product-specific lag but current-calendar generic exposure internationally:
    # avoids inadvertently delaying LOE itself by the commercial launch lag.
    for r in rows:
        m, y = month_index(r['month']), r['year']
        for prod in ('IR', 'XR'):
            lag = int(p['exus_lag']); old = by_month.get(m - lag)
            ratio = p[prod.lower() + '_exus_ratio']
            phase = min(1., max(0., (m - launch[prod] - lag + 1) / p['exus_phase']))
            demand = 0. if old is None else old[prod + '_US'] * ratio * phase
            entry = int(p.get('loe_year_' + prod, p['loe_year']))
            if y >= entry:
                anchor = 0.
                for n in range((entry - 1) * 12, entry * 12):
                    old_anchor = by_month.get(n - lag)
                    phase_anchor = min(1., max(0., (n - launch[prod] - lag + 1) / p['exus_phase']))
                    anchor += (old_anchor[prod + '_US'] if old_anchor else 0.) * ratio * phase_anchor / 12
                demand = anchor * erosion(cfg, p, y, prod)[2]
            r[prod + '_exUS_endmarket'] = demand
            r[prod + '_exUS_direct'] = demand * p['exus_direct_fraction']
            r[prod + '_exUS_royalty'] = demand * (1 - p['exus_direct_fraction']) * p['partner_royalty']
            r[prod + '_recognized'] = r[prod + '_US'] + r[prod + '_exUS_direct'] + r[prod + '_exUS_royalty']
        r['cohort_counterfactual_after_loe'] = y >= min(int(p.get('loe_year_IR', p['loe_year'])), int(p.get('loe_year_XR', p['loe_year'])))
        r['revenue'] = r['IR_recognized'] + r['XR_recognized']
        r['direct_sales'] = r['IR_US'] + r['XR_US'] + r['IR_exUS_direct'] + r['XR_exUS_direct']
        r['endmarket_sales'] = r['IR_US'] + r['XR_US'] + r['IR_exUS_endmarket'] + r['XR_exUS_endmarket']
        assert r['overlap_members'] <= min(r['population'], r['ir_member_years'] * 12) + 1e-7
    return rows, max(ir.max_error, xr.max_error)

def financial(cfg, p, rows, success):
    opening = read('history.json')['opening_2026_06_30']
    policy = {k:v['value'] for k,v in cfg['financial_policy'].items() if isinstance(v,dict)}
    cash = opening['cash_eurm'] * p['fx']
    nwc0 = opening['nwc_eurm'] * p['fx']; nwc = nwc0
    ppe = opening['ppe_and_rou_eurm'] * p['fx']
    other = opening['other_net_assets_eurm'] * p['fx']
    equity = opening['equity_eurm'] * p['fx']
    shares = opening['shares_m']; nol = 0.
    last_failure = month_index(cfg['launches']['XR']['approval_assumption'])
    shutdown = last_failure + policy['winddown_months'] if not any(success) else None
    out = []
    for a in rows:
        r = dict(a); m, y = month_index(r['month']), r['year']
        status = []
        for ix, prod in enumerate(('IR', 'XR')):
            decision = month_index(cfg['launches'][prod]['approval_assumption']) + int(p.get('launch_delay_' + prod, 0))
            status.append(success[ix] or m <= decision)
        ongoing = shutdown is None or m <= shutdown
        if not ongoing:
            # Cash is returned in the shutdown month. No never-ending public shell.
            r.update(rd=0., commercial=0., ga=0., manufacture=0., license=0., depreciation=0., tax=0., capex=0., delta_nwc=0., operating_income=0., net_income=0., fcf=0., funding_gross=0., funding_net=0., funding_fee=0., cash_begin=0., cash_end=0., shares_end=shares, financing_shares=0., nwc_end=0., ppe_end=0., other_net_assets=0., equity_end=0., balance_error=0., distribution=0., shutdown=True)
            out.append(r); continue
        schedule = cfg['cost_schedule_usdm'][str(min(y, 2031))]
        inflation = (1 + p['cost_inflation']) ** max(0, min(y, int(p['loe_year']) - 1) - 2031)
        # Fixed expenses follow prior-year erosion, allowing a one-year cost lag.
        cost_runoff = sum(erosion(cfg, p, max(2026, y - 1), prod)[2] for prod in ('IR', 'XR')) / 2
        active_products = sum(status)
        platform = 1. if active_products == 2 else policy['single_product_platform_fraction'] if active_products == 1 else policy['winddown_platform_fraction']
        rd_factor = policy['rd_IR_fraction'] * status[0] + (1-policy['rd_IR_fraction']) * status[1]
        factor = inflation * cost_runoff * p['cost_scale']
        r['rd'] = schedule['rd'] * factor * rd_factor / 12
        r['commercial'] = schedule['commercial_fixed'] * factor * platform / 12
        r['commercial'] += (r['IR_US'] + r['XR_US']) * p['commercial_variable_us']
        r['commercial'] += (r['IR_exUS_direct'] + r['XR_exUS_direct']) * p['commercial_variable_exus']
        r['ga'] = schedule['ga'] * factor * platform / 12
        r['manufacture'] = r['direct_sales'] * p['manufacture_rate']
        # Free initial units still have a manufacturing cost; net-price-based cost
        # ratio is a proxy, not a claimed per-capsule production cost.
        r['manufacture'] += r['IR_free_units'] * r['IR_price_usd'] * p['manufacture_rate'] / 1e6
        r['license'] = r['endmarket_sales'] * p['license_rate']
        # Remaining EUR8m minus EUR1m paid in H1: reserve EUR7m at IR approval.
        if success[0] and m == month_index(cfg['launches']['IR']['approval_assumption']) + int(p.get('launch_delay_IR', 0)):
            r['rd'] += policy['remaining_milestone_eurm'] * p['fx']
        r['capex'] = policy['capex_fixed_usdm'] / 12 * cost_runoff + p['capex_fraction'] * r['direct_sales']
        r['depreciation'] = min(ppe + r['capex'], ppe / policy['depreciation_months'] + r['capex'] / (2*policy['depreciation_months']))
        ppe += r['capex'] - r['depreciation']
        r['operating_income'] = r['revenue'] - sum(r[k] for k in ['manufacture', 'license', 'rd', 'commercial', 'ga', 'depreciation'])
        # Interest/FX omitted from unlevered planning forecast, not set equal to
        # reported finance income. Tax module is a transparent pooled proxy.
        profit = r['operating_income']
        used = min(nol, max(0., profit) * policy['nol_utilization_cap'])
        nol = nol - used + max(0., -profit)
        r['tax'] = max(policy['tax_floor_usdm'] / 12 * cost_runoff, max(0., profit - used) * p['tax_rate'])
        r['nol_end'] = nol
        r['net_income'] = profit - r['tax']
        next_nwc = nwc0 + p['nwc_fraction'] * r['direct_sales'] * 12
        r['delta_nwc'] = next_nwc - nwc; nwc = next_nwc
        r['fcf'] = r['net_income'] + r['depreciation'] - r['capex'] - r['delta_nwc']
        r['cash_begin'] = cash
        cash += r['fcf']
        r['funding_gross'] = r['funding_net'] = r['funding_fee'] = r['financing_shares'] = 0.
        # A hypothetical financing plan, never represented as committed funding.
        if cash < p['minimum_cash_m'] and active_products:
            net = policy['cash_raise_target_multiple'] * p['minimum_cash_m'] - cash
            gross = net / (1 - p['issue_fee'])
            r['funding_gross'], r['funding_net'] = gross, net
            r['funding_fee'] = gross - net
            r['financing_shares'] = gross / p['issue_price']
            shares += r['financing_shares']; cash += net
        equity += r['net_income'] + r['funding_net']
        r['distribution'] = 0.; r['shutdown'] = False
        if shutdown is not None and m == shutdown:
            # Settle working capital, recover no goodwill/PP&E, return residual.
            residual = cash + nwc + other
            r['distribution'] = max(0., residual)
            r['liquidation_shortfall'] = max(0., -residual)
            cash, nwc, ppe, other, equity = 0., 0., 0., 0., 0.
            r['shutdown'] = True
        r['cash_end'], r['shares_end'] = cash, shares
        r['nwc_end'], r['ppe_end'], r['other_net_assets'], r['equity_end'] = nwc, ppe, other, equity
        r['balance_error'] = cash + nwc + ppe + other - equity
        assert abs(r['balance_error']) < 1e-6, (r['month'], r['balance_error'])
        out.append(r)
    return out

def elapsed_fraction(month, date):
    begin, end = dates(month_index(month))
    if date <= begin:
        return 0.
    if date > end:
        return 1.
    return (date - begin).days / ((end - begin).days + 1)

def point_balance(cfg, p, rows, when):
    date = dt.date.fromisoformat(when)
    opening = read('history.json')['opening_2026_06_30']
    cash, shares = opening['cash_eurm'] * p['fx'], opening['shares_m']
    for r in rows:
        fraction = elapsed_fraction(r['month'], date)
        cash += r['fcf'] * fraction
        if fraction == 1:
            cash += r['funding_net'] - r['distribution']
            shares = r['shares_end']
        if fraction < 1:
            break
    return cash, shares

def diluted_price(equity, shares, p):
    o = read('history.json')['opening_2026_06_30']
    basic_plus_rsu = shares + o['rsus_m']
    strike = o['option_strike_eur'] * p['fx']
    # Treasury-stock algebra using model value, not loss-period diluted EPS.
    if equity / basic_plus_rsu <= strike:
        return max(0., equity / basic_plus_rsu), basic_plus_rsu
    price = (equity + o['options_m'] * strike) / (basic_plus_rsu + o['options_m'])
    diluted = basic_plus_rsu + o['options_m'] * (1 - strike / price)
    return price, diluted

def value_at(cfg, p, rows, when, success):
    date = dt.date.fromisoformat(when)
    cash, shares = point_balance(cfg, p, rows, when)
    pv = fees_pv = 0.
    is_failure = not any(success)
    for r in rows:
        begin, end = dates(month_index(r['month']))
        if end < date:
            continue
        remain = 1 - elapsed_fraction(r['month'], date)
        t = max(0., (end - date).days / 365.25)
        # Future financing creates no enterprise value. Fees are a real cost.
        pv += r['fcf'] * remain / (1 + p['discount_rate']) ** t
        fees_pv += r['funding_fee'] / (1 + p['discount_rate']) ** t
    terminal = 0.
    if not is_failure:
        # 2055 cash flows are already in stable runoff; advance one year.
        last = [r for r in rows if r['year'] == cfg['end_year']]
        terminal = sum(r['fcf'] for r in last) * (1 + cfg['loe']['tail_growth']) / (p['discount_rate'] - cfg['loe']['tail_growth'])
        end = dt.date(cfg['end_year'], 12, 31)
        terminal /= (1 + p['discount_rate']) ** ((end - date).days / 365.25)
        equity = max(0., cash + pv + terminal - fees_pv)
    else:
        # Shareholders receive a dated liquidation distribution. Recompute
        # before liquidation, without adding operating losses a second time.
        equity = 0.
        for r in rows:
            end = dates(month_index(r['month']))[1]
            if end >= date and r['distribution']:
                equity += r['distribution'] / (1 + p['discount_rate']) ** ((end - date).days / 365.25)
        # If intervening funding occurred, current owners keep only their share.
        final_shares = rows[-1]['shares_end']
        equity *= shares / final_shares
    price, fd = diluted_price(equity, shares, p)
    return dict(date=when, price=price, equity_usdm=equity, operating_pv_usdm=pv, cash_usdm=cash, terminal_pv_usdm=terminal, financing_fee_pv_usdm=fees_pv, basic_shares_m=shares, diluted_shares_m=fd,
                future_funding_net_usdm=sum(r['funding_net'] for r in rows if dates(month_index(r['month']))[1] >= date))

def clinical_weights(p, mode='shared_gate'):
    ir, xr = p['ir_pos'], p['xr_pos']
    both = ir * xr / p['shared_success'] if mode == 'shared_gate' else ir * xr if mode == 'independent' else min(ir, xr)
    return {'both': both, 'IR_only': ir - both, 'XR_only': xr - both, 'neither': 1 - ir - xr + both}

STATES = {'both': (True, True), 'IR_only': (True, False), 'XR_only': (False, True), 'neither': (False, False)}

def aggregate(rows, field):
    groups = {}
    stocks = {'population', 'xr_active', 'ir_access_active', 'cash_end', 'shares_end', 'nwc_end', 'ppe_end', 'other_net_assets', 'equity_end', 'nol_end'}
    skip = {'month', 'quarter', 'year', 'shutdown', 'IR_price_usd', 'XR_price_usd', 'IR_loe_price_index', 'XR_loe_price_index', 'IR_loe_volume_index', 'XR_loe_volume_index'}
    for r in rows:
        key = str(r[field]); a = groups.setdefault(key, {field: r[field]})
        for k, v in r.items():
            if k in skip or isinstance(v, bool):
                continue
            if k in stocks:
                a[k] = v
            elif k == 'cash_begin':
                a.setdefault(k, v)
            elif isinstance(v, (float, int)):
                a[k] = a.get(k, 0.) + v
    return list(groups.values())

def evaluate(cfg, case='working', overrides=None, retain=False):
    p = resolve(cfg, case, overrides); weights = clinical_weights(p)
    valuations, months, checks = {}, {}, {}
    for name, flags in STATES.items():
        c, error = commercial(cfg, p, flags)
        f = financial(cfg, p, c, flags)
        valuations[name] = {d: value_at(cfg, p, f, cfg[d + '_date'], flags) for d in ('valuation', 'target')}
        checks[name] = {'max_population_error': error, 'max_balance_error': max(abs(r['balance_error']) for r in f)}
        if retain or name == 'both':
            months[name] = f
    annual = aggregate(months['both'], 'year')
    peak = max(annual, key=lambda r: r['revenue'])
    out = {'case': case, 'clinical_weights': weights, 'conditional_valuations': valuations,
           'risk_adjusted_value_today': sum(weights[s] * valuations[s]['valuation']['price'] for s in STATES),
           'expected_12month_value': sum(weights[s] * valuations[s]['target']['price'] for s in STATES),
           'conditional_peak': {'year': peak['year'], 'recognized_revenue_usdm': peak['revenue']},
           'annual_both_success': annual, 'checks': checks}
    if retain:
        out['monthly_by_clinical_state'] = months
    return out

def acquisition_outcome(standalone_target, consideration_per_share, close_date, target_date, discount_rate=.12, reinvestment_rate=0., completed=True):
    """One terminal outcome replaces standalone; no asserted probability or price.
    If incomplete, caller must use separate completion/failure scenarios. Later
    completion is discounted back; early cash may earn explicit reinvestment.
    """
    if not completed:
        return standalone_target
    days = (dt.date.fromisoformat(close_date) - dt.date.fromisoformat(target_date)).days
    if days > 0:
        return consideration_per_share / (1 + discount_rate) ** (days / 365.25)
    return consideration_per_share * (1 + reinvestment_rate) ** (-days / 365.25)

def buyer_value_interface(standalone_equity, incremental_revenue_pv, avoidable_cost_pv, integration_pv, incumbent_cannibalization_pv, shareholder_synergy_share):
    """All PVs at same transaction date, after tax and funding. No automatic bid."""
    net_synergy = incremental_revenue_pv + avoidable_cost_pv - integration_pv - incumbent_cannibalization_pv
    return {'buyer_economic_value': standalone_equity + net_synergy,
            'negotiated_price_scenario': standalone_equity + shareholder_synergy_share * max(0., net_synergy),
            'net_synergy': net_synergy,
            'economically_acceptable_at_reservation': net_synergy >= 0}

if __name__ == '__main__':
    cfg = read('inputs.json')
    results = {}
    for case in cfg['cases']:
        r = evaluate(cfg, case, retain=True)
        write(case + '_monthly.json', r.pop('monthly_by_clinical_state'))
        results[case] = r
    write('results.json', results)
    write('working_quarterly.json', [r for r in aggregate(read('working_monthly.json')['both'], 'quarter') if int(r['quarter'][:4]) <= 2031])
    print(json.dumps({k: {f: v[f] for f in ('risk_adjusted_value_today', 'expected_12month_value', 'conditional_peak')} for k, v in results.items()}, indent=2))
