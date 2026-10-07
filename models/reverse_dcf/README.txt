PHVS reverse SOTP / DCF skeleton — 7 October 2026

POST-GENERIC CALIBRATION AUDIT (subsequent research, same date)
LATEST: Workbook now implements50/50 study-informed / Kuvan P x Q scenario.
The prior limitation described below is superseded. Appendix recovered via browser.
Inputs66–88 holds the build; DCF112–119 shows separate IR/XR calendar paths.
Health Affairs appendix pp8–9, S-4/S-5:452 formulation-specific drug definitions,
same paper as271-compound analysis. Brand net-price labels64 at year-1 and
88/112/128/138 at years1–4; brand30-day Rx100 at year-1 and24/13/9/8 at years1–4.
Source labels visually checked; preserved on Street evidence112–120 and local PDF
research/raw/Health_Affairs_2026_00257_appendix.pdf.

Study R=(P/64)*(Q/100):33/22.75/18/17.25%. This is a SYNTHETIC scenario, not
measured average revenue: separate pooled averages omit covariance and have
different surviving samples. Kuvan revenue retains62.44/49.72/39.49/26.44%.
Rare-disease net P is assumed flat (Inputs E75:H75); Q is implied from revenue/P.
Do not combine worldwide net revenue with Medicare Q. Changing rare P reallocates
P/Q while preserving the observed revenue anchor; it does not change that revenue.

Blend R=.5*study R+.5*Kuvan R; blend Q=.5*study Q+.5*rare Q; effective P=R/Q.
Revenue retention47.72/36.24/28.75/21.84%; P110.41/115.54/118.56/126.86%;
Q43.22/31.36/24.25/17.22%. Indexed, not absolute dose/patient estimates. IR Q is
paid doses; XR Q is paid treatment-years. Same proxy, independent entry dates.
US/global transfer and partial Kuvan anchor remain limitations. After year4,
P is flat and existing24.91% annual revenue runoff is assigned to Q (assumed).

New spot case: peakUSD1,622.77m; cumulative2027–36USD8,732.00m; lifetime
USD15,229.46m; scale1.010578. Blend endpoints, rare P change, identities and
independent cash-flow solve pass.2 native charts retained; native Excel untested.
Earlier numerical results and Kuvan-only descriptions below are historical.

PREVIOUS AUDIT STATUS (superseded):
The saved workbook's Kuvan-based erosion inputs remain a provisional scenario,
not a validated PHVS central estimate. Broader pooled studies support substantial
brand-volume erosion with potentially stable/rising branded prices. They do not
provide an interchangeable manufacturer-net-revenue curve. The audit did not
change workbook values. See output/PHVS_generic_erosion_audit_2026-10-07.txt and
research/loe_pxq_evidence_2026-10-07.json for analogue limitations, Medicare gross
price/quantity checks and the proposed event-aligned aggregation method.

PURPOSE AND OUTPUT
Estimate the conditional IR and XR revenue paths supporting a selected stock price
or sell-side target, holding approval probabilities and other assumptions fixed.
Workbook: outputs/01a111f7-26b0-7fc2-b98f-41db76412d40/PHVS_reverse_SOTP_DCF.xlsx
Five tabs: Reverse SOTP, Inputs, DCF, Launch comparison, Street evidence.
Blue cells on Inputs are editable; downstream outputs and charts are formula-linked.

This is a normalized conditional valuation, NOT an exact broker-model replication,
consensus forecast, identified buy-side belief, investment recommendation or final PT.
The supplied PDFs are dated September2026; the quote is an October7 intraday snapshot.

HOW TO USE
Inputs E5: 1 = $31.83 spot snapshot; 2 = Guggenheim $44 PT; 3 = custom price.
E14: scale both products together, XR alone with IR fixed, or IR alone with XR fixed.
E15: treat source revenues as success-case or as already probability weighted.
E16/E17: approval probabilities. Defaults90%/90% from Guggenheim p1.
E20–E22: discount rate, annual revenue tail change and residual-value toggle.
E54/E55: first full generic-exposed year for IR/XR independently (default2039).
E56:E59: revenue retained in exposure years1–4, linked to the50/50 P/Q build.
E21: subsequent annual revenue change, source-linked to Kuvan2021–24 CAGR.
Changing a linked driver to an assumption does not overwrite the observed data.
E22=0 excludes2056 onward only, not all post-expiry cash flows.
E38 stays0 until the source revenue probability basis has been verified.
Use the source-shaped path as a starting scenario. Editing annual source data or
costs is possible; changes need their own evidence and justification.

SOURCE CHOICE AND AUDIT
Two ZIPs contained13 PHVS and6 BCRX notes. Manifest, original filenames, SHA256s and
PDF-page-delimited text are under research/raw/sellside_2026-10-07/. Immutable PDF
copies are under references/user_supplied/sellside_2026-10-07/.
Review targeted valuation, commercial and financial sections, not legal appendices.

Twelve distinct PHVS broker targets span$42–65. Closest lower targets are Wolfe42,
BofA44, Guggenheim44, KBC49 and Wedbush50. Guggenheim provides the most usable
separate revenue path among these lower-target reports. Its early revenue and PoS
assumptions remain together; the subsequent user-requested patent erosion model
replaces its terminal-growth convention. This is not the broker's complete model.
All12 brokers remain in Street evidence; absence of numerical PoS is marked n.a.

Guggenheim124306319 pp1–3:90%/90%PoS,12.5%WACC,-5%terminal growth,15%tax starting2031.
Printed product revenues are EUR thousands, converted to EUR millions. Whether
they already incorporate PoS is not unambiguously disclosed. Basis switch exposes
this; equal90% probabilities mean unweighting changes the reference multiplier
but not the final fitted sales in the common-scale solve.
The p3 printed GAAP EBIT follows pre-tax-income figures from p2; printed after-tax
EBIT adds positive tax entries in2031–34. Source figures are preserved separately.
The normalized DCF subtracts taxes. No plug forces the normalized baseline to$44.
This flags source arithmetic; it does not prove the analyst's underlying file has
identical errors or that the published target was generated from those exact rows.

MS124295137 p2 updates XRPoS to90%; p3 retains stale70%. Use the explicit update.
Leerink124293821 late8Sep updates PT65 and XR95%; earlier124292103 still shows50/60%.
Do not count the two reports as independent brokers or use the stale appendix.
Wedbush124292125 p5 product table heading saysEUR000s but dollar signs and explicit
USD-to-EUR conversion rows identify those product revenues asUSD000s. Image checked.
RBC62 blends a DCF with an EPS multiple: it cannot be reverse-solved as pure DCF.
Wells Fargo says no terminal growth, which does not mean no terminal value.

MECHANICS
EUR millions internally; sales comparisons convert to worldwide USD millions.
Current cash estimate rolls June30 reported cash to October7 using H1 cash burn;
2026 future costs include only the remaining stub. This avoids counting elapsed
burn both in reduced cash and future expense. Actual Q3 cash is not known here.
Year-end discounting; target horizon0 for spot,1 for broker PT (convention assumed),
custom horizon editable. A future equity target is discounted before solving.
One-year broker convention is measured from this valuation date, not a precise
remaining horizon from the September note. No interim equity issuance is modeled.

Let A_IR and A_XR be PVs of each product's probability-weighted marginal cash flow,
C be the negative PV of shared costs/investment, and N be cash+other assets-debt.
Then target equity PV = k_IR*A_IR + k_XR*A_XR + C + N.
For a common scale k, k=(target equity PV-C-N)/(A_IR+A_XR).
For XR-only, k_XR=(target equity PV-C-N-A_IR)/A_XR, and vice versa.
Cash flow is independently rebuilt with non-refundable taxes to verify repricing.
The analytical expression assumes taxable-year profits. Losses after the selected
tax start can invalidate it; the reconciliation check visibly fails in that case.
Do not use outputs marked INVALID. Negative required ramps are not valid forecasts.

Shared costs are deducted once and not PoS-weighted. This is an expected-revenue
skeleton, not a full success/failure-state rNPV; taxes on expected profits are not
necessarily expected taxes. No joint IR/XR approval-independence assumption is made.
A full scenario tree is needed to value failure-state savings and state taxes.
No NOL roll-forward, future financing, dilution schedule,
territorial economics, cohort switching or IR/XR rescue-use interaction is modeled.
AAE/other pipeline value is set to an explicit zero scenario, not found valueless.
Reported R&D and G&A used through2030; held flat until both products face generics.
Thereafter R&D/G&A/capex/depreciation follow the mean IR/XR revenue retention factors
with a one-year lag. This is an assumed asset-runoff cost response, not observed PHVS
cost behavior or a forecast for an indefinitely reinvesting pharmaceutical company.
FX1.16USD/EUR, WC5%of sales, annual capex and depreciationEUR0.6m are assumptions.
76.125m diluted shares are Guggenheim's projection, not reported70.204506m basic shares.
The explicit forecast now extends through2055. Revenue plateaus at2036 levels until
the assumed entry dates, then declines. Once all components share the same mature
decline rate, end2055 residual value=2056 FCF/(discount rate minus annual change).
Perpetuity means an infinite horizon, not30 or50 years. Its present value is finite.
The2055 cutoff is for explicit spreadsheet presentation; it is not a cash-flow stop.
Undiscounted lifetime sales also include the geometric revenue sum from2056 onward.

POST-GENERIC EVIDENCE AND CALIBRATION — 7 OCTOBER EXTENSION
Kuvan (sapropterin, PKU) is the main oral rare-disease small-molecule proxy. Reported
worldwide revenueUSDm:2019 463.4;2020 457.7;2021 285.776;2022 227.577;2023 180.767;
2024 approximately121. US generic entry occurred October2020.2020 is a transition
anchor, not a completely unexposed baseline; relative to2019, retention would be
slightly lower. Geographic timing, alternative therapies and pricing confound a
pure generic effect. This is a revenue proxy, not a causal or US-only estimate.
First four retained-revenue factors=2021–24 sales/2020 sales:62.44/49.72/39.49/26.44%.
Subsequent annual decline24.91%=(121/285.776)^(1/3)-1. That observed interval supports
a sensitivity anchor; continuing it indefinitely is an assumption, not an observation.
Source URLs and formulas are on Street evidence rows77–92.

Cross-checks, not pooled: Lundbeck's Xenazine product revenueDKKm2201 in2015,
1571 in2016 and1049 in2017 (28.6% and33.2% annual declines); generic entryQ3 2015
makes2015 partly exposed. Its reports distinguish US geographic sales from total
product revenue; the model uses the consistent total-product series. FX, distribution
economics and subsequent branded competition limit comparisons. Orfadin worldwide
revenueSEKm665 in2020 to459 in2021 (-31.0%); both are already post-entry years, so
this is corroboration of continuing erosion, not an initial generic-entry curve.
Firazyr is a peptide injectable, so it is not used as a small-molecule oral analogue.
Orfadin and Xenazine lack a clean matched full four-year cohort in this bounded pass.

Patent timing: indexed FY2025 PHVS annual-report disclosure says patents expire
November23,2038 or later. Default2039 is an explicitly assumed first full exposed
year. It is not a confirmed generic launch, legal opinion or guarantee of protection.
Full current PDF retrieval failed; the indexed primary excerpt was recovered, not
a full claim-by-claim patent audit. Patent extensions, later formulation claims,
litigation and territory-specific protection are unresolved. The2038 partial-year
generic stub is omitted by the annual convention. Source URL is retained in row109.
Entry years may differ for IR/XR.2037–2050 is the supported input domain so the tail
is fully in its mature decline phase by2055. No delayed-entry benefit is assumed by
default. These timing assumptions require a separate IP audit before pitch use.

INITIAL CONDITIONAL RESULTS
Using the updated erosion assumptions, the$31.83 spot requires0.98677x both ramps.
The$44 one-year PT requires1.13786x; normalized1x present value is$32.47.
At spot, required success-case worldwide sales:
Year       IR USDm      XR USDm       Total USDm
2027          62.3          0.0             62.3
2028         125.2         74.4            199.5
2030         347.8        252.8            600.6
2033         674.3        464.6          1,138.9
2036         978.4        606.1          1,584.5
2037–38 repeat2036;2039 first erosion year combined989.3;2055 combined10.1.
Cumulative2027–36:$8.526bn;2027–55:$15.750bn; lifetime:$15.780bn. All are undiscounted
success-case worldwide revenues, not equity values. At$44 future PT, peak$1.827bn,
2027–36 cumulative$9.832bn and lifetime$18.196bn.
The old5%cash-flow tail yielded$1.351bn required peak; that case is superseded,
not a directly comparable post-patent curve. It remains recoverable in Git1676752.
Residual cash-flow PV from2056 onward isEUR0.296m,0.0142%of equity value. Removing
only that residual raises required peak to$1.58468bn. Post2036 cash flows have moved
into explicit years; the tiny residual does NOT mean all distant cash flows vanished.
Tail-decline sensitivities:10%annual erosion needs$1.555bn peak;40%needs$1.598bn.
Delay IR entry to2042 alone:$1.492bn peak; delay XR alone:$1.536bn. All else fixed.

XR first full calendar years2029–32 require$167.3/252.8/331.0/401.6m versus observed
Orladeyo2021–24$122.6/251.6/326.0/437.7m. Launch months,
prices, territories and competition differ. This is not a causal efficacy premium
or calibrated second-oral adoption discount. US paid volumes must be separated next.
Ekterly$49.1m reflects roughly six commercial months in2025; PHVS2027 calendar sales
are not an identical exposure period. No precise acute ramp ratio is asserted.
Launch primary URLs are retained in output/Orladeyo_launch_analogue_2026-10-07.txt
and the workbook source tab. Financial filing: https://ir.pharvaris.com/node/12571/html

BCRX CROSS-CHECKS
RBC124314632 p1 models~30%out-year erosion and$400–450m tail; prior~30-specialist
survey suggests~35%. These are modeled/reported survey intentions, not actual losses.
Evercore124342693 pp1–2 reverse case at$8.40 assumes Orladeyo$685m2027 peak declining
to~$550m2037 and paid patients~1,200 to~700. Its operating-margin descriptions differ.
That reverse case is separate from its$17 target. Neither case directly identifies
PHVS capture: price, new users, other competitors and free/paid shifts intervene.

VERIFICATION AND LIMITS
Artifact Tool recalculation: base, broker target, XR-only, source-basis change and
no-residual, entry-date and tail-rate scenarios reprice within1e-6. A zero-target
boundary flags invalidity.
Independent Python year-by-year cash flow plus bisection matches spot, PT and no-TV
solutions within1e-9. A500-year cash-flow sum matches the residual perpetuity formula.
Exported caches contain no cell errors; two native charts
retain formula bindings. All five worksheet previews visually reviewed; overlapping
notes/charts corrected. Native Microsoft Excel recalculation was not exercised.
verification.json and independent_verification.json retain results.

NEXT DECISION
The skeleton can enter scenario analysis now; inferred investor adoption beliefs
cannot. Highest-value next work: explicit net-price/territory/paid-patient bridge,
then compare required paid adoption to observed launch cohorts and plausible switch
stocks. Refine effective generic-entry timing, cost runoff and funding before
inferring an efficacy premium. Additional weak analogue searches were stopped once
Kuvan's usable series and two cross-checks were established; no paid data or agents.
Additional broad searching or Monte Carlo before resolving these definitions would
add apparent precision. Patient forum / expanded adoption work remains deferred.
