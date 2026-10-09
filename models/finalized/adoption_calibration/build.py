"""Reproduce source-group underwriting bridges, not an empirical fit or valuation."""
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
CASES = ('cautious', 'working', 'stronger')
SWITCH = (0.10, 50 / 328, 0.20)
ORAL_SHARE = (0.50, 0.60, 0.70)
PAID = 0.80  # held fixed to isolate demand; access sensitivity is separate
EKTERLY_ANNUAL = 1 - (1 - 1702 / 9000) ** (12 / 8)


def build():
    rows = []
    for i, case in enumerate(CASES):
        switch, brand = SWITCH[i], ORAL_SHARE[i]
        injectable = switch * 0.54 * brand
        oral = switch * (0.40, 0.60, 0.80)[i]
        rescue_only = 4080 * injectable * (32 / 52) / 2720
        untreated = 2720 * rescue_only * (17 / 32) * (0.25, 0.50, 0.75)[i] / 1500
        ir_injectable_rx = EKTERLY_ANNUAL * (0.75, 1.00, 1.25)[i] * brand
        # LTP switching is deliberately a WEAK cross-setting proxy for oral rescue.
        ir_oral = switch * (0.40, 0.60, 0.80)[i]
        ir_no_rescue_rx = (0.10, 0.20, 0.30)[i] * 0.84 * brand
        values = {
            'XR_orladeyo': (oral, oral / PAID),
            'XR_injectable_prevention': (injectable, injectable / PAID),
            'XR_older_prevention': (injectable, injectable / PAID),
            'XR_rescue_only': (rescue_only, rescue_only / PAID),
            'XR_diagnosed_untreated': (untreated, untreated / PAID),
            'IR_injectable_rescue': (ir_injectable_rx * PAID, ir_injectable_rx),
            'IR_ekterly': (ir_oral, ir_oral / PAID),
            'IR_no_current_rescue': (ir_no_rescue_rx * PAID, ir_no_rescue_rx),
            # These are conditional choices at a treatment-selection event, NOT
            # annual hazards for all newly diagnosed people or additive inflows.
            'NEW_XR_given_prevention_selection': ((0.25, 0.35, 0.45)[i] * PAID,
                                                  (0.25, 0.35, 0.45)[i]),
            'NEW_IR_given_rescue_selection': (0.84 * brand * PAID, 0.84 * brand),
        }
        for group, (paid, rx) in values.items():
            rows.append(dict(group=group, case=case,
                             period='selection event' if group.startswith('NEW') else 'one year',
                             completed_start_proxy=paid,
                             static_Rx_equivalent=rx,
                             paid_starts_per_1000=1000 * paid))
            assert 0 <= paid <= rx <= 1
            assert abs(rx * PAID - paid) < 1e-12
    result = dict(as_of='2026-10-09', status='Proposed underwriting, not fitted PHVS estimates',
                  annualized_Ekterly_forms_proxy=EKTERLY_ANNUAL,
                  annualization_caveat='Constant-hazard equivalent of forms / broad9000population; not measured injectable-only annual acquisition; eight-month approximation.',
                  access_normalization=PAID, cases=rows)
    (HERE / 'results.json').write_text(json.dumps(result, indent=2) + '\n')
    return result


if __name__ == '__main__':
    r = build()
    for group in dict.fromkeys(x['group'] for x in r['cases']):
        a = [x for x in r['cases'] if x['group'] == group]
        print(group + ': ' + ' / '.join(f"{100*x['completed_start_proxy']:.2f}%" for x in a))
