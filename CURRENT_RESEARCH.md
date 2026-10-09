Linked US revenue ramp (9October): output/PHVS_US_revenue_ramp_2026-10-09.txt; models/patient_adoption/revenue_ramp.py with inputs/results/table. Practical evidence bridges in research/PHVS_multiyear_revenue_bridges_2026-10-09.json. Relative full model years,IRone-yearlead,notfairvalue. ProcessG37.

Central range/citation register: research/PHVS_adoption_range_source_register.json; readable index: output/PHVS_adoption_range_sources.txt. Includes source URLs, evidence locations, adjustments and explicit source gaps.

Latest XR origin/access work (9 October): [first-year underwriting ranges and paid-flow results](output/PHVS_XR_entry_access_2026-10-09.txt). xr_entry_model.py:working~460Rx->338firstpaid->295yearendactive/141paidpatientyears. Rates explicitlyassumed; new empiricalcapture notidentified. ProcessG34.

Population simplification: working1.5%net annual growth,0–3%scenario range, anchored to HAE historical treated1.53%/diagnosed2.63%CAGRs. Supersedes detailed demographic proposal; do not stack separate diagnosis/death flows. population_growth.py supplies net scaling only. ProcessG33.

Full-text prevalence audit: research/PHVS_prevalence_fulltext_audit_2026-10-08.json. Historical medication-ascertained population, not undiagnosed reservoir; statisticalCIs retained. Demographic exits added to architecture, not yet national runnable model. ProcessG32.

Latest adoption reconciliation: [population/payment gaps and monthly cohort results](output/PHVS_adoption_reconciliation_2026-10-08.txt). Conditional359–549paid/PAP state entries explain~255net growth; not observed prescriptions. Provisional population grid and MonteCarlo dependency register added. ProcessG31.

Latest population/launch calibration: [dated patient pools and net additions](output/PHVS_population_launch_calibration_2026-10-08.txt). OrladeyoYE2023~789paid;~255net paid+PAP additions,321includingQuickStart. Origin52/32/17 is historical trier mix, not conversion. Runnable market_calibration.py; no national forecast selected. ProcessG30.

Latest runnable adoption adjustment: working_assumptions.json drives normalized cohort sensitivity; XR61/74/87%first-year retention scenarios,74%working assumption; IRpivotal andLTPextension comparisons keptseparate. No nationalforecast orDCF repricing. ProcessG29.

Latest three-hypothesis audit: [XR switching/retention and IR consumption](output/PHVS_three_adoption_hypotheses_2026-10-08.txt). New H1–H4 links; KON-S LTP22.3% repeat dosing qualifies permanent38%assumption. Berolife20/80efficacy/AE exits notPHVSretentionuplift. ProcessG28.

Latest public-data discovery: [pharmacy/claims source audit](output/PHVS_public_pharmacy_sources_2026-10-08.txt). Optime supplement1919first/1746repeat dispensings; descriptive only, notpaidconversion. StateDUR requests differ from patients. SourcesPHS1–6 linked in ledger; processG27.

Latest adoption update (8 October): [flow calibration](output/PHVS_adoption_flow_calibration_2026-10-08.txt); F01-F05 benchmarks linked in patient-adoption ledger. PHVS13.8%extension vsEkterly38%pivotal redosing retained with denominator/window caveats. No national transition probabilities or valuation changed. GenAI logG26.

# Current research state — 7 October 2026

## GenAI process documentation — added 8 October 2026

User requested ongoing documentation for a pitch competition. See
[structured process log](research/genai_process_log.json) and
[rubric-organized pitch notes](research/genai_pitch_notes.txt).
Repository [working instructions](AGENTS.md) require selective updates for substantive
research turns, crediting the user's hypotheses and pushback and recording corrections,
validation and limitations. Backfilled events are labeled retrospective; unknown model
settings, dates and efficiency gains are not invented. These records document method,
not independent investment evidence.

## Investment question

8 October redosing expert follow-up: [note](output/PHVS_redosing_expert_followup_2026-10-08.txt). Supplied June25doctor discussed single-capsule results but called comparable tosebetralstat; management emphasizes pricing/payer value. Final465attack RAPIDe2poster supports low repeat use, not independent head-to-head gap. No model changes. ProcessG25.


8 October redosing audit: [note](output/PHVS_redosing_audit_2026-10-08.txt), [arithmetic](output/PHVS_redosing_arithmetic_2026-10-08.json). Matched12h observed PHVS17%vsEkterly38% implies15.2%conditional dose-volume gap, not a proven sales haircut. Long-term Ekterly19.3%is a median trend summary, not pooled average; mixed-dose PHVS extension also not directly comparable. No model coefficients changed. Process G24.


8 October same-MOA audit: [evidence](output/PHVS_same_MOA_rescue_audit_2026-10-08.txt). Direct same-target HAE rescue evidence and PHVS April2026 combination analyses weaken categorical XR-to-IR exclusion. Prior audit missed the PHVS disclosure. Specific drug safety restrictions are not proof of class failure after breakthrough. No numeric adoption changes; prospective XR+IR and observed preference remain unresolved. Process G23.


8 October revenue-comparability audit: [note](output/PHVS_revenue_comparability_audit_2026-10-08.txt). XR/Orladeyo and IR/Ekterly are useful paired analogues, conditional on setting, paid use, dosing/pack, age, launch era, access/net price and geography/accounting. User redirected next work toward observed adoption inputs, rather than more hypothetical thresholds. No model coefficients changed. Process G22.


8 October conditional thresholds: [results](output/PHVS_conditional_thresholds_2026-10-08.txt), [chart](output/PHVS_conditional_thresholds_2026-10-08.png), [editable assumptions](models/patient_adoption/threshold_inputs.json). Latest saved$31.83 reverse-DCF spine retained. Illustrative80%US/$418kXR/80%conversion/80%year1/95%mature case needs1188paid patient-years in2036 and2314first prescriptions over2028–36, including900assumed captured new diagnoses and1414existing patients. IR at80%US requires~$802m:4000annual paid users imply~$200knet annual spend each. These are conditional, not calibrated forecasts.30monthly XR scenarios reconcile; initial year-flat allocation artifact corrected before delivery. Product mix/valuation unchanged. Process G21.

8 October commercial calibration: [assessment](output/PHVS_commercial_calibration_2026-10-08.txt), [source/history table](research/PHVS_commercial_calibration_2026-10-08.json), [26-parameter ledger](models/patient_adoption/parameter_ledger.json) and [updated flow](models/patient_adoption/flow.mmd). PHVS disclosed a focused field organization and integrated launch plans. Evidence supports credible design, not proven conversion/persistence. Management dropout illustration is explicitly a scenario, not measured retention. Current national treatment partition, paid transitions, mature survival and net prices remain unset; no central TAM/peak forecast or reverse-DCF change. Arithmetic checks are separate from market validation. Process log G20.

Patient-adoption architecture: read [canonical design and handoff](models/patient_adoption/ARCHITECTURE.txt) first. Includes new diagnoses/untreated patients, XR cohorts, IR inventory and replenishment, access states, commercial parameter mapping and accounting rules. Expanded design is documented, not yet implemented or calibrated.

8 October user BioCryst archive: [bounded first-pass extraction](output/BioCryst_launch_archive_first_pass_2026-10-08.txt),23 PDFs indexed locally with hash manifest. Recovered planned30–50US reps, top200 access relationships, broader-prescriber contribution and conditional switcher retention. Management disclosures require denominator matching; no adoption coefficients changed. Process log G18.

8 October AlphaSense commercial search: [four-interview synthesis](output/PHVS_commercial_expert_search_2026-10-08.txt) and [timecoded evidence](research/alphasense_commercial_experts_2026-10-08.json). Former BioCryst, Takeda and CSL perspectives support competing switching/access mechanisms; CSL hub friction adds prescription-to-fill execution risk. Qualitative only: no adoption coefficients changed. Exact territories/headcount, conversion rates and Ekterly operations evidence remain open. Process log G17.

8 October resumed adoption work: [segment specification and evidence mapping](models/patient_adoption/README.txt).
Initial cohort sensitivity is illustrative only, not a calibrated forecast; national
stocks, PHVS paid transitions, net prices and geography remain to be estimated.


8 October injectable extension: [revenue-height audit](output/HAE_injectable_revenue_audit_2026-10-08.txt).
Takhzyro reached~$1.5bn; Cinryze~$699m recovered2017 high; HaegardaFY2025$490m.
GuggenheimXR~$614m is not unprecedented in HAE; IR~$992m exceeds Firazyr2017$663m.
These are historical benchmarks, not certified lifetime peaks or patient-demand validation.


8 October ramp audit: [results](output/Guggenheim_ramp_audit_2026-10-08.txt) and
[inputs/provenance](research/guggenheim_ramp_audit_2026-10-08.json). IR early shape
resembles Orladeyo; XR resembles Cerdelga after normalization. No obvious shape
anomaly established. Guggenheim2036 still grows11.7%IR/6.9%XR; our plateau is an
assumption, not independently estimated peak demand. Workbook unchanged.


8 October methodological correction: user identified that slower early adoption
and higher later peak sales can yield the same valuation. The reverse model has
no independently fitted S-curve; it scales a fixed Guggenheim annual path. Four-year
XR outperformance is conditional, not uniquely priced-in expectations or sufficient
short evidence. Ramp-speed/peak tradeoff and independent patient-based ceilings
remain to be implemented. See process event G12 and reverse-model README.

Latest framing: [IR/XR patient segmentation, access, interactions and historical analogue screen](output/PHVS_SOTP_patient_segmentation_framework_2026-10-07.txt). This is a research/modeling structure, not a completed TAM estimate. [Historical prophylaxis comparator analysis](output/Orladeyo_historical_comparator_delta_2026-10-07.txt) compares the launch-era clinical propositions and documents early switching concerns.

Can a clinically attractive oral HAE drug generate enough sustained, reimbursed demand to exceed the expectations embedded in PHVS? The developing short case focuses on commercial differentiation, access, persistence, treatment intensity and execution. A successful approval or efficacy readout alone does not resolve valuation. Conversely, commercial friction alone does not establish an overvalued stock.

The strongest long case is a convenient oral treatment with compelling attack control, switching from dissatisfied incumbents, conversion of acute-only patients to prophylaxis, and execution improvements that turn free starts into paid patients. Test that case as seriously as the short. A conditional reverse-valuation skeleton is now available (below); it is not an identified consensus bridge or defensible final price target.

## Clinical audit and corrections

[Bounded granular-data search](output/PHVS_granular_data_search_2026-10-07.txt): recovered September2026 RAPIDe-3 regional analyses including unpaired treated participants, with similar active relief medians but regional placebo differences; partially narrows the missing SEAS evidence gap. Endpoint-observed denominators differ from paired overall results and are not additive. New sebetralstat pooled subgroup paper provides modifier evidence but explicitly omits within-person correlation in exploratory hazard-ratio estimates. No downloadable linked patient/attack records or CHAPTER-3 switch sensitivities recovered. Registry/sharing-policy barriers documented. More aggregate fitting is not justified as a patient-adjusted superiority estimate.

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

7 October mechanistic comparison: downstream B2 blockade plus sustained XR exposure is a credible efficacy rationale, but does not establish comparative superiority. Sebetralstat also rapidly inhibits its target; berotralstat has a long half-life, so routine daily washout is unsupported. HAE-PLG offers a specific bypass-pathway hypothesis, not a broad normal-C1 population or TAM claim. Evidence and source links: [mechanistic comparison](output/PHVS_mechanistic_comparison_2026-10-07.txt). No quantitative commercial inputs changed.

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

Latest: [bounded patient-adoption evidence](output/PHVS_patient_adoption_evidence_2026-10-07.txt), [26-observation ledger](research/PHVS_adoption_evidence_2026-10-07.json) and [search boundary/handoff](research/PHVS_adoption_search_handoff_2026-10-07.txt). No agents, paid collection, external contacts, valuation or Monte Carlo were used in this pass.

- BioCryst's January2026 primary presentation confirms **>1,600 US Orladeyo patients on therapy at year-end2025**, including QuickStart/paid/PAP, and **>3,500 who have tried it** cumulatively. Neither their difference nor ratio identifies discontinued patients or retention. Paid share was **84% at Q1-end2025 and 81% at year-end2025**, excluding QuickStart. Do not multiply81% by the active count; denominators differ. These supersede the earlier unverified84% reference with dated primary evidence.
- A US linked EHR/claims study reports **12-month persistence61% berotralstat (n90),58% lanadelumab (n189),53% SC-pdC1INH (n78)**, using a45-day gap definition. These are new-start cohort benchmarks, not stable-incumbent annual retention or causal product rankings. BioCryst repeats this study; do not count it twice.
- A historical commercial-claims cohort records **50/328 (15.2%) switching LTP within12months of initiation**. It does not identify PHVS capture or stable-user switching. In the same cohort, **220/328 (67.1%) had a rescue claim**; mean post-index claims-derived doses were11.8/year overall and8.0 in the minimal-gap subgroup. Dispensing is not attack consumption.
- Recent pooled claims and clinic registry studies corroborate turnover and multiple switch reasons, but use different gap definitions and follow-up periods. Do not pool their rates. A preference experiment supports heterogeneous priorities, not national segment weights or paid adoption probabilities.
- Current injectable/acute-only stocks, new diagnoses, conditional paid switching, bridge conversion and rescue inventory replacement remain unmeasured. Historical2020 treated projections6,595/8,694 and a company~9,000 diagnosed estimate are differently defined anchors, not a current TAM confidence interval.

Next highest-value unknown: paid, persistent treatment transitions by prior regimen, tenure and control, with the denominator who could switch but did not. Public data supports scenario calibration; targeted collection is worth considering only if plausible transition assumptions change the investment decision and linked coverage is feasible. BioCryst announced a CareMed sole-source pharmacy transition inQ3 2026; historical Optime-only records could miss continuation across the transfer.

Separate dissatisfied oral users, stable oral users, injectable users, and acute-only/new-to-prevention patients. For each monthly start cohort, track persistence, free/paid status, time to coverage and net revenue. Do not multiply year-end paid share and retention into annual revenue as if those rates applied uniformly all year.

Acute revenue depends on residual attacks, propensity to treat, doses per attack, paid dispensing, initial stocking and replenishment. Better prophylaxis reduces rescue consumption. Easier rescue may increase treatment of mild attacks or change willingness to start prevention. Do not independently maximize all favorable assumptions.

The next highest-information evidence would be longitudinal specialty-pharmacy/hub or claims cohorts with prior treatment, first dispense, paid status, persistence and subsequent switching. Free product and specialty-channel incompleteness must be assessed. No paid data purchase or external interview has been performed.

7 October reverse SOTP extension: user supplied two ZIPs containing13 PHVS and6 BCRX reports and authorized a reverse DCF skeleton. [Model method, source audit and initial results](models/reverse_dcf/README.txt); [editable workbook](outputs/01a111f7-26b0-7fc2-b98f-41db76412d40/PHVS_reverse_SOTP_DCF.xlsx). Twelve distinct PHVS broker targets are retained, with same-day stale Leerink figures superseded. No agents or paid collection were used.

The initial normalized scenario uses Guggenheim's90%/90%PoS,12.5%discount rate and negative5%terminal growth with separate IR/XR ramps. Other undisclosed inputs are labeled assumptions. Source revenue probability basis remains unconfirmed; the model exposes both interpretations. Printed Guggenheim tax arithmetic is inconsistent, so the source table is preserved separately and is not blindly replicated. Morgan Stanley's explicit XR90%update supersedes a stale70%sidebar. Broker methods and parameters must not be averaged into synthetic consensus.

The initial perpetual5%cash-flow decline after2036 is now superseded by a user-requested explicit patent-erosion extension. Same workbook, annual forecasts through2055, and a declining residual from2056 onward. Kuvan worldwide2020–24 revenue calibrates retained revenue62.4/49.7/39.5/26.4%in the first four modeled generic-exposed years, followed by an assumed continuation of its24.9%2021–24 annual decline. Xenazine and Orfadin corroborate substantial continuing erosion but are not pooled. Geographic timing, partial-year exposure and competing treatments limit these proxies.

Default first full generic exposure is2039 for both formulations, separately editable. The indexed FY2025 primary filing states patents expireNovember2038 or later; full PDF retrieval failed, and no claim-by-claim IP audit was performed.2039 is a timing scenario, not established generic entry. Revenue plateaus at2036 levels before erosion; shared costs/capex/depreciation shrink with average retained revenue at a one-year lag after both products face generics (assumed cost response).

Updated dated$31.83case:0.98677x ramp, peak worldwide success-case$978mIR/$606mXR=$1.585bn; cumulative2027–36$8.526bn; lifetime undiscounted sales$15.780bn. The$44 one-year PT needs1.13786x,$1.827bn peak and$9.832bn cumulative2027–36. Residual2056 onward isEUR0.296m PV,0.0142%of equity value. This small residual reflects moving2037–55 into explicit forecast years, not removing their value. Source data, assumptions, URLs and sensitivity results are retained in the model README and source tab. These remain conditional requirements, not uniquely inferred market beliefs. Ekterly's roughly six commercial months are not a matched PHVS launch year.

Formula recalculation, scenario repricing, independent cash-flow bisection, native chart bindings and affected tab previews were checked. A500-year tail sum independently matches the residual perpetuity. Native Excel execution was not tested. No final price target, cohort adoption model or Monte Carlo is complete. Remaining valuation work: net price, patient adoption/persistence, spending, funding/dilution, territorial economics, effective generic-entry timing, cost runoff and joint IR/XR success/failure states. Broader adoption collection remains deferred until these reverse-valuation assumptions are assessed.

## Reference pitches and historical provenance

IR launch comparison implemented on workbook Launch comparison55–76: Ekterly
Q3 2025 13.7m, Q4~35.4m derived, Q1 2026 39.2m; cumulative88.3mUSD. Annualized
Q1run rate156.8m is an arithmetic benchmark, not annual actual/guidance. PHVS
IR2027–31 forecasts shown separately;2028first full calendar after launch.
Ratios to fixedQ1run rate are not age-matched ramps. Bounded search did not recover
later Ekterly actuals. No alteration to IR/XR mix or valuation assumptions.

LATEST50/50 implementation: Health Affairs appendix recovered via browser;
S-4/S-5 pp8–9 give rounded net-price and volume labels for452 formulation-specific
drug definitions in the same271-compound paper. Workbook now uses50% synthetic
study P x Q and50% Kuvan revenue. The former is an assumed representative path,
not measured mean revenue (pooled means/covariance and sample differences).
Kuvan net P assumed flat, Q implied; CMS gross P/Q not substituted for global net.
Inputs66–88 owns build; DCF112–119 shows IR/XR indices; entry dates independent.
Revenue retention47.72/36.24/28.75/21.84%; post-year4 P flat,24.91% annual runoff
assigned to Q. Peak requirement at31.83spotUSD1.623bn;2027–36 cumulative8.732bn.
Independent valuation and P/Q checks pass. Earlier blocked/Kuvan-only notes below
are superseded; full caveats in models/reverse_dcf/README.txt.

Subsequent user choice: combine the271-compound study with rare-disease analogues.
Working specification is50% study brand-only net-spending retention /50% Kuvan
revenue retention, with25/75 and endpoints as sensitivity weights (assumptions).
Xenazine remains a cross-check; Orfadin is qualitative. Study values must be rebased
to pre-entry NET spending, not its plotted retail denominator. Publisher Exhibit3
failed to load in browser; direct figure/supplement retrieval failed. No annual
brand-only values were fabricated and the workbook remains Kuvan-only pending
that extraction. This is a recorded model-design decision, not a completed model
change. Details are at the beginning of the generic-erosion audit.

7 October post-generic audit: [analogue and pooled evidence audit](output/PHVS_generic_erosion_audit_2026-10-07.txt) and [reproducible CMS observations](research/loe_pxq_evidence_2026-10-07.json). The workbook's Kuvan-only erosion curve is now explicitly considered provisional, not a validated central estimate; no workbook values changed. Orfadin's formulation-led growth despite generics and Xenazine's later Austedo competition weaken clean generic-erosion interpretation. Broad published studies (356 NMEs, 530 US substances, and a271-compound Part D net-price study; overlapping populations, different outcomes) offer benchmarks, not three independent revenue curves. A new matched manufacturer revenue panel was not assembled.

CMS2020–24 gross spending decomposes into units × spending/raw unit: Kuvan -88.7% = -88.2% units with -3.9% unit-spend; Xenazine -74.8% = -82.6% units with +44.7% unit-spend (multiplicative, not additive). These are small Medicare populations, mixed strengths/forms, and gross claims amounts before rebates, not national manufacturer net prices. Orfadin has49 claims in2024 and is excluded from quantitative extrapolation. Do not apply generic price discounts to brand prices or treat molecule-wide spending as PHVS revenue. Further matching should align actual generic launch, geography/formulation, initial dollar sales and entrant count; calculate each molecule's revenue curve before aggregating.

The Daniel Bao deck/workbook concern **Phathom (PHAT)**, not Pharvaris. Their assumptions and cached valuation are not PHVS data. The reference audit records useful thesis construction and material model limitations. M Loses Capital articles and Culper PRAX are examples of skeptical methods, not verified authorities on PHVS. SLGL was provided in the chat at https://riskinhibitor.substack.com/p/slgl-zero; its methodological takeaways are captured in [reference analysis](research/references.txt).

This guide summarizes the discussions and recovered files. It is not a verbatim chat transcript. Dates, later corrections and original uncertainties are retained rather than silently rewriting historical reports.
