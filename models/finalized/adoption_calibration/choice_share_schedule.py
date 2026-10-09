"""Time-dependent oral-choice calibration; not yet a full patient-stock forecast."""
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent
LAUNCH_SHARE = {'cautious': 0.25, 'working': 0.40, 'stronger': 0.60}
OLD_SHARE = {'cautious': 0.50, 'working': 0.60, 'stronger': 0.70}
MATURE_SHARE = 0.60
CONVERGENCE_MONTHS = 60
ORAL_CHOICE_GROUPS = {
    'XR_injectable_prevention', 'XR_older_prevention', 'XR_rescue_only',
    'XR_diagnosed_untreated', 'IR_injectable_rescue', 'IR_no_current_rescue',
    'NEW_IR_given_rescue_selection',
}


def oral_choice_share(case, months_since_launch):
    if months_since_launch < 0:
        raise ValueError('No commercial choice before product launch')
    fraction = min(months_since_launch / CONVERGENCE_MONTHS, 1.0)
    return LAUNCH_SHARE[case] + (MATURE_SHARE - LAUNCH_SHARE[case]) * fraction


def build_schedule():
    original = json.loads((HERE / 'results.json').read_text())
    paths = []
    for month in range(121):
        for row in original['cases']:
            share = oral_choice_share(row['case'], month)
            adjusted = row['group'] in ORAL_CHOICE_GROUPS
            factor = share / OLD_SHARE[row['case']] if adjusted else 1.0
            paths.append(dict(
                group=row['group'], case=row['case'], months_since_product_launch=month,
                oral_choice_share=share if adjusted else None,
                rate_basis=row['period'],
                completed_start_proxy=row['completed_start_proxy'] * factor,
                static_Rx_equivalent=row['static_Rx_equivalent'] * factor,
            ))
    for case in LAUNCH_SHARE:
        series = [oral_choice_share(case, m) for m in range(121)]
        assert abs(series[0] - LAUNCH_SHARE[case]) < 1e-12
        assert all(a <= b + 1e-12 for a, b in zip(series, series[1:]))
        assert all(abs(x - MATURE_SHARE) < 1e-12 for x in series[60:])
    assert all(0 <= x['completed_start_proxy'] <= x['static_Rx_equivalent'] <= 1
               for x in paths)
    output = dict(
        as_of='2026-10-09', launch_share=LAUNCH_SHARE, mature_share=MATURE_SHARE,
        convergence_months=CONVERGENCE_MONTHS,
        method='Linear interpolation from launch to month60, then flat.',
        status='User-directed common60%mature oral-choice assumption; timing and initialshares underwriting, not empirical fit. Replaces constant oral-choice factors in static results for intended implementation.',
        timing='Use months since each formulation launches: IR2027-05,XR2028-06 under existing calendar. Annual probabilities here are instantaneous annualized propensities, not actual year-one starts. Convert to monthly before cohort simulation.',
        unchanged='All-destination capture from Orladeyo/Ekterly and new-prevention overall share remain distinct assumptions; no60%equilibrium implied for them.',
        caution='No patient-stock depletion, launch-origin mix or nationalDCF recomputed. Engine.py integration remains pending. Improving choice does not prevent total starts tapering.',
        paths=paths,
    )
    (HERE / 'choice_share_timepath.json').write_text(json.dumps(output, indent=2) + '\n')
    return output


if __name__ == '__main__':
    build_schedule()
    for month in (0, 12, 24, 36, 48, 60, 120):
        print(month, {c: round(100 * oral_choice_share(c, month), 1) for c in LAUNCH_SHARE})
