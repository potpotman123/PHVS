# Current research state — 7 October 2026

## Investment question

Latest framing: [IR/XR patient segmentation, access, interactions and historical analogue screen](output/PHVS_SOTP_patient_segmentation_framework_2026-10-07.txt). This is a research/modeling structure, not a completed TAM estimate. [Historical prophylaxis comparator analysis](output/Orladeyo_historical_comparator_delta_2026-10-07.txt) compares the launch-era clinical propositions and documents early switching concerns.

Can a clinically attractive oral HAE drug generate enough sustained, reimbursed demand to exceed the expectations embedded in PHVS? The developing short case focuses on commercial differentiation, access, persistence, treatment intensity and execution. A successful approval or efficacy readout alone does not resolve valuation. Conversely, commercial friction alone does not establish an overvalued stock.

The strongest long case is a convenient oral treatment with compelling attack control, switching from dissatisfied incumbents, conversion of acute-only patients to prophylaxis, and execution improvements that turn free starts into paid patients. Test that case as seriously as the short. No completed consensus bridge, market-implied valuation or defensible price target has yet been produced.

## Clinical audit and corrections

[CHAPTER-3 amendment investigation](output/CHAPTER3_analytical_amendment_investigation_2026-10-07.txt): original October2024 registration confirms unchanged headline primary endpoint. Exact pre-amendment analytical rules remain unavailable; EMA confirms public CTIS cannot provide prior document versions. Revised switch-to-prophylaxis penalty can favor either arm depending on switches, timing and assigned rates. A possible severe-attack hierarchy move is an inference from registry ordering, not a verified old/new redline. No efficacy inflation established.

[Culper-method follow-up](output/PHVS_Culper_method_audit_2026-10-07.txt): early 2024 announcements match the eventual pivotal primary endpoints. CHAPTER-3 Amendment 1.5 nevertheless explicitly changed intercurrent-event handling and one endpoint’s testing hierarchy; RAPIDe-3 changed planned LTP mix and clarified qualification/retraining. Exact prior-version redlines and numerical impact remain unresolved. These are diligence leads, not a demonstrated manipulation case.

Latest: [cross-trial clinical audit](output/PHVS_cross_trial_clinical_audit_2026-10-07.txt) and [reproducible arithmetic](output/PHVS_cross_trial_arithmetic_2026-10-07.json). Commercial efficacy assumptions are pending this audit's unresolved comparability questions.

**Material correction:** Ekterly's published placebo initial-relief median of 6.72 hours is not the FDA-revised result. FDA moved missing-assessment censoring from time zero to 12 hours; its placebo median was not estimable, and sebetralstat 600 mg median was 1.99 hours. Thus the earlier “PHVS placebo takes over twice as long” concern is not established on aligned regulatory analyses. Both drugs retain evidence of efficacy against placebo; comparative superiority remains unresolved.

**New CHAPTER-3 protocol:** CTIS Amendment 1.5 (19 May 2025) explicitly specifies overdispersion-scaled Poisson standard errors, a negative-binomial supplementary analysis, and missing-data sensitivity analyses. Do not allege that variability was ignored. Their numerical results and the final SAP remain outstanding. The XR versus berotralstat efficacy gap is a strong directional signal, not a validated causal switching coefficient.

The original suspicion concerned 134 randomized patients versus 88 paired participants in RAPIDe-3. Recovered CTIS documents materially weaken the missing-patient short argument:

- 113 participants treated 201 attacks; 88 treated both attacks, leaving 25 single-attack participants and 21 randomized participants without a treated attack.
- 46 incomplete participants comprise 35 sponsor-closure cases and 11 other reasons. They are not 46 toxicity or lack-of-efficacy withdrawals.
- The protocol permits final-analysis closure after enough paired participants; the lay summary confirms sufficient information as the closure explanation. The protocol target was at least 60 paired participants; 88 were obtained.
- PEAS is the paired efficacy analysis set; SEAS includes participants with at least one treated attack. The protocol specifies supportive SEAS analyses, but their numerical treatment estimates were not recovered.
- Remaining requests: final dated SAP, numerical SEAS and treatment-policy results, first-period-only comparison, missingness/attack-count cross-tabs, and documentation of closure timing and blinding.
- The CTIS results PDF contains an inconsistent NCT identifier; other identifiers and reported trial data match RAPIDe-3. Treat as an apparent administrative discrepancy, not proof of fabrication.
- Normal-C1INH eligibility expansion is documented in January 2025 amendments. No evidence establishes a post-readout eligibility change. A full amendment-by-amendment account of every substantial modification remains incomplete.

Source and denominators: [detailed audit](output/RAPIDe-3_granular_data_audit_2026-10-06.txt), [official-source manifest](research/raw/rapide3_ctis_source_manifest.json), [disposition](output/RAPIDe-3_patient_disposition.csv). These supersede the initial memo's statement that the protocol/disposition could not be located.

## Generalizability, attack frequency and prophylaxis

Qualifying treated attacks are not all attacks. Actual time at risk differs across participants; an enrollment maximum does not imply everyone contributed that duration. A population mean attack rate cannot establish the probability that a particular participant has zero or one qualifying attack without accounting for exposure, heterogeneity and background treatment. The Poisson calculation is illustrative, not fitted validation.

Background long-term prophylaxis was 31/134 (23.1%) in the randomized population; do not automatically apply that denominator to the paired subset. Placebo control helps estimate the randomized contrast but does not erase questions about interactions, representativeness and future patient mix.

See [generalizability note](research/rapide3_generalizability_2026-10-06.txt) and [illustration](research/rapide3_poisson_illustration.json).

## Safety and mechanism

Keep treatment-related events separate from all treatment-emergent events and serious events. The clinical notes retain formulation-specific safety tables and limitations. Acute short-window exposure cannot establish the same safety profile as chronic prevention. Small trials cannot rule out rare events.

Deucrictibant and icatibant both antagonize the bradykinin B2 receptor, but molecular size or structural resemblance alone does not establish specificity, clinical superiority, exposure or safety. A full comparative selectivity/structure analysis was deferred while prioritizing commercial work. Kallikrein inhibitors act upstream and are not B2 antagonists.

The concern about using the same molecule/mechanism for prevention and rescue is a physician preference and a testable pharmacology/label question, not established futility. Additional antagonist exposure can matter; combined-use evidence and permitted dosing require direct evaluation. Avoid calling speculative concurrent-use advantages proven.

## Payer access: key correction

One UHC Orladeyo medical-necessity pathway requires prior injectable steps, but a separate notification pathway does not list the same sequence. Therefore, it is incorrect to say all UHC patients must fail all injectables. Benefit-specific coverage, exceptions, prior treatment history and population weights matter.

Andembry's favorable positioning in a reviewed UHC policy and injectable Dawnzera's nonpreferred position in a reviewed Cigna policy argue against universal insurer hostility to oral drugs or all new drugs. Rebate/net-price explanations are plausible, but confidential contracting terms have not been established. Manufacturing cost, list price, manufacturer net revenue and payer net cost are different quantities.

For PHVS prevention, Orladeyo is a useful access scenario. For acute treatment, Ekterly is a closer policy analogue. No evidence justifies an across-the-board 50% TAM cut. Some payer restrictions on normal-C1INH HAE concern evidence/eligibility; they are not proof of genetic nonresponse to deucrictibant.

See [payer audit and exact policy links](output/HAE_payer_access_audit_2026-10-06.txt). Next needed: policy-to-benefit mapping, HAE patient weights, completed prior steps, exceptions, first-paid-fill delays, and net pricing. More unweighted policy PDFs alone will not answer those questions.

## Expert calls: useful signals, limited denominators

All five supplied calls were reviewed. Main themes: stable patients' reluctance to disturb control; dissatisfaction as a switch trigger; convenience of longer-interval injections; physician and staff authorization burden; enthusiasm for potent oral prophylaxis; uncertainty about same-mechanism rescue; and interactions between easier acute treatment and preventive treatment decisions.

Small practice samples, inconsistent patient counts and possible repeated interviews with one expert prevent treating these calls as independent representative votes. An Optum-affiliated physician's prescribing view is not evidence of UHC contracting policy. Predicted willingness to try is not paid adoption or persistence.

See [25-theme long/short audit](output/PHVS_AlphaSense_long_short_audit_2026-10-07.txt) and [timestamped evidence](research/alphasense_calls_evidence_2026-10-07.json). The five originals are in `references/user_supplied/`.

## Quantified launch evidence

Orladeyo net revenue: $122.6m in2021, $251.6m in2022, $326.0m in2023 and $437.7m in2024. More than half of initial starts had prior prophylaxis. This is meaningful disconfirming evidence against a blanket claim that HAE switching inertia blocks oral uptake.

At December2023, 1,104 US patients were on paid or long-term free product; 71.5% were paid (approximately789 paid/315 free, derived). Paid share reached74.4% in June2024. Initial approximately70% retention and later >60% one-year paid persistence use different populations and cannot be read as a deterioration series.

Cerdelga following Zavesca in Gaucher is a structurally useful second-oral analogue: global sales EUR66m/106m/126m in2015/2016/2017. First-line positioning, eligibility, incumbent commercial infrastructure and a different disease limit read-through. It is not a matched clinical efficacy-delta experiment. Livmarli/PFIC brand-level revenue is confounded by its other indication. Danazol is an older oral HAE treatment, so Orladeyo was not literally the first oral prophylaxis.

See [launch audit with quarterly data and sources](output/Orladeyo_launch_analogue_2026-10-07.txt). A full public prescription/patient cohort series was not recovered. Revenue does not identify a causal switching-cost parameter.

## Modeling framework and next decisions

Separate dissatisfied oral users, stable oral users, injectable users, and acute-only/new-to-prevention patients. For each monthly start cohort, track persistence, free/paid status, time to coverage and net revenue. Do not multiply year-end paid share and retention into annual revenue as if those rates applied uniformly all year.

Acute revenue depends on residual attacks, propensity to treat, doses per attack, paid dispensing, initial stocking and replenishment. Better prophylaxis reduces rescue consumption. Easier rescue may increase treatment of mild attacks or change willingness to start prevention. Do not independently maximize all favorable assumptions.

The next highest-information evidence would be longitudinal specialty-pharmacy/hub or claims cohorts with prior treatment, first dispense, paid status, persistence and subsequent switching. Free product and specialty-channel incompleteness must be assessed. No paid data purchase or external interview has been performed.

Remaining valuation work: current expectations/consensus, net price, cohort adoption/persistence, commercial spending, funding/dilution, territorial economics and correlated IR/XR risks. Original scope explicitly deferred a complete DCF and final price target.

## Reference pitches and historical provenance

The Daniel Bao deck/workbook concern **Phathom (PHAT)**, not Pharvaris. Their assumptions and cached valuation are not PHVS data. The reference audit records useful thesis construction and material model limitations. M Loses Capital articles and Culper PRAX are examples of skeptical methods, not verified authorities on PHVS. SLGL was provided in the chat at https://riskinhibitor.substack.com/p/slgl-zero; its methodological takeaways are captured in [reference analysis](research/references.txt).

This guide summarizes the discussions and recovered files. It is not a verbatim chat transcript. Dates, later corrections and original uncertainties are retained rather than silently rewriting historical reports.
