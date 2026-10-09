LATEST WORKING MODEL UPDATE — 8 October2026
working_assumptions.json now drives sensitivity.py and illustrative_sensitivity.json.
XR R1=61%+39%*(20/30)*fraction_of_clinical_exits_avoided:61/74/87% for0/50/100%
assumed repair.74%is the working normalized-cohort case, not measured PHVSretention.
French cause mix transferred toUSclaims is a proxy. Mature90% remains assumed.
IR17%vs38% and17%vs22.3% are parallel cross-population comparisons; no universal
paid-revenue haircut. Survey rates are not conversion. Original threshold inversion
is archived unchanged; national adoption and reverseDCF have not been repriced.
This supersedes the older fixed61% sensitivity description below.

CONDITIONAL THRESHOLDS (8 October 2026)
threshold_inputs.json contains explicit scenarios; threshold_analysis.py reads the
saved DCF verification without changing valuation.30 XR allocations solve monthly
nonnegative first prescriptions against every annual2028–36 revenue target.
Minimum month-to-month changes defines the illustrative path; a linear-program
minimum total prescription count is also reported. No empirical adoption fit.
IR requirements are annual paid users/dispensing at assumed net prices.
Report: output/PHVS_conditional_thresholds_2026-10-08.txt; chart alongside.
Run with dependencies in threshold_requirements.txt, then validate_thresholds.py.
Results and checks saved as threshold_results.json / threshold_validation.json.
No national TAM or scenario probabilities calibrated; DCF not repriced.

LATEST CALIBRATION PASS (8 October 2026)
parameter_ledger.json:26 definitions with source/denominator/transferability and
unknown inputs preserved. flow.mmd: expanded diagram. calibration/check_calibration.py
reproduces arithmetic and source checks; it does not implement a national forecast.
Full assessment: output/PHVS_commercial_calibration_2026-10-08.txt.

START HERE: ARCHITECTURE.txt is the canonical expanded patient-flow design and
research handoff. The executable sensitivity remains illustrative only.

Patient adoption build — resumed 8 October 2026

Question: can incremental clinical benefit, tolerability and convenience generate
sufficient persistent paid demand in a competitive HAE market? Comparative clinical
superiority and market saturation are hypotheses, not established inputs.

Model geography: US first. Reverse-DCF endpoints are worldwide; no US patient
penetration may be calculated from worldwide revenue without a geographic split.

XR source states (mutually exclusive at each date):
1. Orladeyo, satisfied/control adequate.
2. Orladeyo, dissatisfied/control inadequate or intolerant.
3. Injectable prevention, satisfied/control adequate.
4. Injectable prevention, dissatisfied or administration burden unacceptable.
5. Acute-only, clinically suitable and willing to initiate prevention.
6. Other prevention/unknown. Preserve rather than allocate without evidence.
Clinical control and satisfaction are distinct; categorization needs a hierarchy,
not summed efficacy/GI/burden percentages. New diagnoses enter relevant states.

For each state and month:
Eligible remaining population -> clinical discussion -> choice/prescription ->
access/paid fill or bridge/free supply -> retained paid treatment months.
Use conditional transitions or one measured paid-transition rate, never both
as overlapping haircuts. Discontinuers can return, choose another product or
stop prevention. Competition determines destination shares, not just exit rates.
Allow a mature stable cohort to retain differently from new initiators.

Evidence mapping (source IDs from research/PHVS_adoption_evidence_2026-10-07.json):
A01: >1600 Orladeyo on-therapy US patients atYE2025, includes free supply. Lower bound.
A06:15.2% switched any LTP within12months after initiation. Calibration check only;
not an annual stable-incumbent switch probability or PHVS capture rate.
A04:61/58/53% one-year initiation persistence, definition45day gap. Historical
benchmarks only, not long-term incumbent retention or comparative efficacy.
A03:81% paid share excluding QuickStart atYE2025. Mature stock check only;
not first-paid conversion probability, and denominator differs fromA01.
A16:cessation reasons conditional on discontinuers, nonexclusive. Mechanisms only.
A18–21:survey satisfaction/preferences support heterogeneity, not US segment weights.
AlphaSense E1:review frequency/admin workload explains slow switching among stable
patients. E2/E4:better oral efficacy/tolerability supports switching interest.
E3/E5 may repeat an expert; no independent vote counts or averaged market shares.

IR layer:
Split every XR/current-prevention state by current rescue regimen and residual
burden. Estimate attacks, proportion treated, PHVS choice, doses per attack,
paid doses and net price. Keep initial home inventory and replacement separate.
A07–09:67.1% had rescue claims,11.8doses/year overall,8.0minimal-gap subgroup.
These are dispensing benchmarks, not consumed attacks or PHVS forecasts.
Do not infer prevention removes rescue ownership or multiply the same patients
by full untreated attack burden after successful XR. Same-drug dual use is
conditional on clinical evidence/label; no automatic cross-selling benefit.

FIRST COMPUTABLE CHECK
sensitivity.py generates illustrative_sensitivity.json for1000 incumbent patients.
It varies annual attempt rate5/10/20/30% and initial-paid probability60/80/100%.
Retains61% through each initiation year (historical benchmark used hypothetically)
and90% thereafter (assumed). No reattempt, new diagnosis, competing destination,
bridge conversion or within-year timing. Not a fitted national forecast.
At80%initial-paid, year5active counts are89/160/257/309 respectively.
Year-end active patients are not annual patient-years, so no annual revenue is
calculated from those snapshots. This demonstrates why headline willingness
cannot be copied into peak market share; no scenario is selected as base case.

NEXT BUILD / STOP RULE
Highest-value input is aligned current treatment stocks, followed by paid
persistent transitions conditional on previous regimen/control. Need source-tagged
ranges for those, net price and a US/worldwide revenue split before selecting a
base-case peak. Existing public evidence does not identify these precisely.
Use expert calls to define mechanism and bounds; forum anecdotes may reveal
omitted friction but not prevalence. No new bulk search or paid collection.
Next quantitative step: implement monthly treatment cohorts and threshold curves
showing required capture vs persistence/net price for each segment. User chooses
no probabilities by default; Monte Carlo waits until dependencies/ranges are defensible.
