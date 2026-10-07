PHVS reverse SOTP / DCF skeleton — 7 October 2026

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
E20–E22: discount rate, terminal growth and terminal-value toggle.
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
separate revenue path among these lower-target reports. Its parameters remain
together; missing inputs are explicitly assumed, not borrowed as fake consensus.
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
No NOL roll-forward, future financing, dilution schedule, explicit LOE erosion,
territorial economics, cohort switching or IR/XR rescue-use interaction is modeled.
AAE/other pipeline value is set to an explicit zero scenario, not found valueless.
Reported R&D and G&A used through2030; held flat thereafter as an assumption.
FX1.16USD/EUR, WC5%of sales, annual capex and depreciationEUR0.6m are assumptions.
76.125m diluted shares are Guggenheim's projection, not reported70.204506m basic shares.
The perpetual decline applies to2036 FCF;2036 is not an independently inferred peak.

INITIAL CONDITIONAL RESULTS
Using default assumptions, the$31.83 spot requires0.84158x both reference ramps.
The$44 one-year PT requires0.96649x; normalized1x reference present value is$41.06
(one-year equivalent$46.20). These are different horizons, not a hidden mismatch.
At spot, required success-case worldwide sales:
Year       IR USDm      XR USDm       Total USDm
2027          53.1          0.0             53.1
2028         106.7         63.4            170.2
2030         296.6        215.6            512.2
2033         575.1        396.2            971.3
2036         834.4        516.9          1,351.4
Hold IR at1x instead: required XR scale0.59411x. Turn terminal value off: required
common scale1.28177x. These demonstrate underidentification and valuation sensitivity.

XR first full calendar years2029–32 require$142.7/215.6/282.3/342.5m versus observed
Orladeyo2021–24$122.6/251.6/326.0/437.7m: ratios1.16/0.86/0.87/0.78x. Launch months,
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
no-terminal scenarios reprice within1e-6. A zero-target boundary flags invalidity.
Independent Python year-by-year cash flow plus bisection matches spot, PT and no-TV
solutions within1e-9. Exported caches contain no cell errors; two native charts
retain formula bindings. All five worksheet previews visually reviewed; overlapping
notes/charts corrected. Native Microsoft Excel recalculation was not exercised.
verification.json and independent_verification.json retain results.

NEXT DECISION
The skeleton can enter scenario analysis now; inferred investor adoption beliefs
cannot. Highest-value next work: explicit net-price/territory/paid-patient bridge,
then compare required paid adoption to observed launch cohorts and plausible switch
stocks. Stress costs, funding and patent erosion before inferring an efficacy premium.
Additional broad searching or Monte Carlo before resolving these definitions would
add apparent precision. Patient forum / expanded adoption work remains deferred.
