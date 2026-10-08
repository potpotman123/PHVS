"""Reproduce calibration diagnostics, not a PHVS revenue forecast."""
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
MODEL = ROOT / 'models/patient_adoption'
ledger = json.loads((MODEL / 'parameter_ledger.json').read_text())
evidence = json.loads((ROOT / 'research/PHVS_commercial_calibration_2026-10-08.json').read_text())
params = {p['id']: p for p in ledger['parameters']}
assert len(params) == len(ledger['parameters'])
facts = {f['id']: f for f in evidence['new_facts']}
old = json.loads((ROOT / 'research/PHVS_adoption_evidence_2026-10-07.json').read_text())
old_ids = {o['id'] for o in old['observations']}
valid_refs = set(evidence['sources']) | set(facts) | old_ids
for p in params.values():
    assert all(k in p for k in ['name', 'definition', 'unit', 'period', 'population', 'numerator',
                               'denominator', 'source_refs', 'status', 'dependencies',
                               'double_counting_risk', 'numerical_use', 'PHVS_transferability'])
    assert set(p['source_refs']) <= valid_refs, p['id']
    if p['status'] == 'unknown':
        assert p['value'] is None and p['numerical_use'] == 'not_ready'

hist = {h['period']: h for h in evidence['historical_calibration']}
assert abs(sum(hist['2021 '+q]['revenue_USDm'] for q in ['Q1', 'Q2', 'Q3', 'Q4 / year-end']) - 122.6) < 1e-8
assert abs(sum(hist['2022 '+q]['revenue_USDm'] for q in ['Q1', 'Q2', 'Q3', 'Q4']) - 251.6) < 1e-8

results = {
    'as_of': '2026-10-08',
    'status': 'Arithmetic and source-definition checks only; no external validation or national forecast',
    'conditional_12m_retained_per_100_starts': {
        'claims_oral_benchmark': 100 * params['T1']['value']['oral'],
        'selected_paid_benchmark': 100 * params['T3']['value']['paid'],
        'management_dropout_illustration_if_annual': 100 * (1 - params['T4']['value']),
        'caution': 'Different underlying populations/definitions. Alternative scenario checkpoints, not treatment comparisons.'
    },
    'relative_retained_stock_if_same_starts_and_other_conditions': {
        '95_vs_61': .95 / .61,
        '95_vs_70': .95 / .70,
        'caution': 'Not annual revenue uplift, causal efficacy benefit or evidence PHVS achieves95%.'
    },
    'five_year_new_entry_gross_ceiling': {
        'patients': [5 * n for n in params['N3']['value']],
        'assumptions': 'Constant management-estimated entry flow;100% eligible and captured; zero exits. Ceiling for this hypothetical flow only, not entire PHVS opportunity.'
    },
    'coarse_target_load': {
        'clinicians_per_planned_field_member': params['R1']['value'] / params['R2']['value'],
        'caution': 'Simple workload scale only; no geography, call frequency, travel, account effort or filled-position verification.'
    },
    'historical_paid_stock_same_denominator_2023': {
        'active': 1104, 'paid_fraction': .715, 'approx_paid': 1104 * .715,
        'approx_long_term_free': 1104 * (1 - .715),
        'caution': 'Rounded company stock, excludes QuickStart. No similar YE2025 multiplication allowed.'
    },
    'checks': ['2021 revenue sums to122.6m', '2022 revenue sums to251.6m',
               'Source IDs resolve', 'Unknown parameters remain unset',
               'No observed first-year endpoint auto-converted to mature annual hazard'],
    'forecast_ready': False,
    'unresolved_parameters': [p['id'] for p in params.values() if p['numerical_use'] == 'not_ready']
}
(MODEL / 'calibration/check_results.json').write_text(json.dumps(results, indent=2) + '\n')
print(json.dumps({'checks': 'passed', 'parameters': len(params),
                  'unresolved': len(results['unresolved_parameters']), 'forecast_ready': False}))
