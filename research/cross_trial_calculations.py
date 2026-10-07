"""Reproduce descriptive and explicitly assumption-conditional audit arithmetic.
No patient-level adjustment or comparative superiority test is performed.
"""
from pathlib import Path
import json, math
root = Path(__file__).resolve().parents[1]
r = {
 'classification': 'Derived arithmetic; not a validated comparative treatment effect',
 'konfident_placebo_relief_fraction_all_attacks': 41/84,
 'konfident_placebo_fraction_if_eight_time_zero_censored_removed': 41/(84-8),
 'warning_on_missingness': 'Fraction illustrates the 50% threshold mechanism, not reconstruction of Kaplan-Meier event times.',
 'acute_median_difference_hours_descriptive_only': 1.99-1.28,
 'acute_median_difference_minutes_descriptive_only': (1.99-1.28)*60,
 'rescue_24h': {'rapide3_active':6/88,'rapide3_placebo':32/88,'konfident_600mg':10/93,'konfident_placebo':27/84},
 'xr_reported_rr':0.17,'berotralstat_rr_derived_from_44_2_percent_reduction':0.558,
 'xr_ratio_of_placebo_relative_rate_ratios_conditional':0.17/0.558,
 'multiplicative_bias_factor_to_neutralize_point_ratio':0.558/0.17,
 'conditions': ['Comparable endpoint and follow-up estimands', 'Exchangeable distributions of treatment-effect modifiers across trials', 'Comparable handling of discontinuation and missing data', 'Reported estimates and uncertainty valid'],
 'model_status': 'Within-trial estimates usable with source labels; cross-trial ratio not approved as a sales-model efficacy input.'
}
se1=(math.log(.28)-math.log(.10))/(2*1.96)
se2=(math.log(.770)-math.log(.405))/(2*1.96)
se=math.sqrt(se1**2+se2**2); z=math.log(.17/.558)
r['conditional_sampling_only_95_interval']=[math.exp(z-1.96*se),math.exp(z+1.96*se)]
r['interval_warning']='Approximate log-normal propagation using rounded published intervals. Excludes cross-trial bias, is not an adjusted superiority claim, and must not be quoted alone.'
(root/'output/PHVS_cross_trial_arithmetic_2026-10-07.json').write_text(json.dumps(r,indent=2)+'\n')
print(json.dumps(r,indent=2))
