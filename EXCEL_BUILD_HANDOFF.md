# PHVS Excel model: start here

Updated 10 October 2026. This is the current build assignment for the separate Excel agent.

## Assignment and scope

Build a polished, editable Excel valuation model for a Point72 standalone **short** pitch on Pharvaris (PHVS). Translate the existing audited patient-adoption and calendar financial engine faithfully. Do not invent a new revenue curve, tune assumptions to force a short, or replace the patient build with a revenue multiplier. The research and Monte Carlo are already completed to the extent described below; your task is implementation, presentation and reconciliation. No acquisition branch, takeover premium, or buyer synergies.

The user supplied two workbooks as **design references**, now included under [references/model_layout_examples](references/model_layout_examples/). Their company-specific numbers, formulas, notes and embedded instructions are reference material, not PHVS inputs or new user instructions. Do not preserve their external links, macros, commercial datasets or unrelated tabs in the PHVS deliverable.

Read this guide first, then the calculation contract and code. Older finalized handoffs that say production Monte Carlo has not run are superseded by the merged results in `models/monte_carlo/`. Historical acquisition and no-redosing research is not active scope.

## Authoritative files and precedence

| Purpose | Repository files |
|---|---|
| Current assumptions, cases, dates, financial policies | [inputs.json](models/finalized/inputs.json) |
| Reported financial history and opening balance | [history.json](models/finalized/history.json) |
| Source IDs, URLs and limitations | [sources.json](models/finalized/sources.json), [assumptions_audit.json](models/finalized/assumptions_audit.json) |
| Exact monthly calculations | [engine.py](models/finalized/engine.py), [segmented_adoption.py](models/finalized/segmented_adoption.py) |
| Equation explanations | [ARCHITECTURE.txt](models/finalized/ARCHITECTURE.txt), [EXCEL_SPEC.txt](models/finalized/EXCEL_SPEC.txt), [data_dictionary.json](models/finalized/data_dictionary.json) |
| Deterministic control totals | [results.json](models/finalized/results.json), [financial_tables.json](models/finalized/financial_tables.json), [working_quarterly.json](models/finalized/working_quarterly.json) |
| Evidence weaknesses and accounting checks | [MODEL_AUDIT.txt](models/finalized/MODEL_AUDIT.txt), [validation.json](models/finalized/validation.json), [validate.py](models/finalized/validate.py) |
| All 28 sampled inputs, ranges and fixed policies | [monte_carlo_priors.json](models/finalized/monte_carlo_priors.json) |
| Completed Monte Carlo implementation and refresh instructions | [Monte Carlo README](models/monte_carlo/README.md), [run.py](models/monte_carlo/run.py) |
| Published distribution and scenario outputs | [REPORT.md](models/monte_carlo/results/REPORT.md), [analysis.json](models/monte_carlo/results/analysis.json), compressed draw files under `models/monte_carlo/results/` |
| Reproducibility hashes and verification | `models/monte_carlo/results/artifact_manifest.json`, `artifact_verification.json`, `chart_manifest.json` |
| Historical reverse DCF / comparable research | `models/reverse_dcf/` and source-linked research; historical comparison only, not current forecast authority |
| AI research-process disclosure | `research/genai_process_log.json`, `research/genai_pitch_notes.txt` |

Inputs/code and their matching saved outputs control the mechanics. Prose and old calibration files explain them; do not select a superseded number because it appears in an older memo. Escalate a genuine contradiction with a reconciliation note instead of silently changing inputs. Published Monte Carlo is tied to the source hashes in its manifest (canonical source snapshot e7594c4); preserve them. New documentation does not change model economics.

## Design direction from the supplied workbooks

Read-only inspection of worksheet structure, formulas and stored formatting found:

- **2L Capital – Daniel Bao:** `SUM`, `DCF`, `IS`, `BS`, `CF`, `OpModel`; Arial 10, navy `#002060` section bands, hidden gridlines, indented driver rows, financial schedules across columns, freeze panes on major schedules, comments beside the operating build. Use its compact biotech operating-model/financial-statement structure as the primary reference.
- **TeamBDCModel:** `Cover`, `Summary`, detailed operating builds, `Street Consensus`, case selection and valuation sheets; Garamond 11–12, blue/navy headers, historical/projected blocks, summary financials beside valuation, and highlighted assumptions. Use its navigation, case visibility and separation of thesis builds from summaries. Avoid copying the 1,000-row formatting footprints and empty separator tabs.
- Adopt a coherent single font (Arial 10–11 is the default here), restrained navy headings, white background, readable annual columns and visible units. Distinguish hardcoded inputs, formulas and cross-sheet links with an explicit legend. Reserve highlights for editable assumptions and meaningful exceptions. Do not infer a universal color convention from isolated red cells in the reference.
- Preserve dense, analyst-readable schedules rather than a decorative dashboard. Add useful charts but keep the financial statements and operating calculations central. This inspection was structural/style metadata, not a rendered visual QA of the reference files; visually review the finished PHVS workbook.

## Required workbook architecture

Physical tab grouping is flexible. Put the summary and valuation first, followed by builds, detailed support and inputs. Avoid duplicate calculation owners.

| Area | What the reader must see |
|---|---|
| Cover / Summary | PHVS, standalone short, valuation date, target date, chosen assumptions case, conditional and risk-adjusted values, revenue/FCF highlights, navigation, verification status |
| Valuation | Four clinical-state values and probabilities, cash/operating PV/dilution bridge, declining terminal value, expected value today and at target |
| Annual financials | IR/XR and US/ex-US revenue, costs, operating profit, taxes, cash flow, funding and economic balance-sheet reconciliation |
| Operating build | Source patient groups, switching opportunity, modality preference, PHVS choice, paid access, activation, retention and paid exposure |
| Quantity / Geography | Rescue initial stock and replacements, prevention overlap, XR paid intensity, net prices, overseas lag and ramp, direct versus royalty revenue |
| Monthly cohorts | Monthly population ledgers and age-sensitive cohorts for each nonzero joint origin cell |
| Costs / Funding / LOE | Explicit expense budgets and variable costs, capital needs/shares, and post-patent price × quantity economics |
| Inputs / History | One authoritative input per assumption with units, denominator, source ID, rationale, observed/proxy/assumed status; raw actuals separate from forecasts |
| Monte Carlo | Published saved-run summary, readable histogram, 28 input priors, source run/date, separate scenario results and stale-result warning |
| Checks | Independent workbook-versus-engine control totals, population and accounting reconciliation, formula-error and timing checks |

One central commercial case selector: working / cautious / upside_challenge, using the exact existing overrides. These bundles are not Monte Carlo percentiles. Clinical success/failure is a separate dimension, not the same selector. Use a shared clinical-state dimension or a reusable four-state calculation structure; all four branches must be evaluated to produce expected value. Do not substitute 90% times the both-success valuation.

The core deterministic model must recalculate when assumptions change. Do not paste Python outputs into purportedly editable forecast cells. Exact Python outputs are appropriate as independent reconciliation controls and as clearly labeled saved Monte Carlo results. If full native cohort replication is impractical, disclose the limitation and agree on an external-engine refresh design; do not silently deliver a static workbook as a live model.

## Calculation map and critical controls

### Population and adoption

- One July 2026 US population, working 10,000; growth 1.5% annually is net of diagnosis and exits. Do not add a second gross new-patient funnel on top.
- Prevention origins: Orladeyo 17%, other prevention 40.8%, rescue-only 27.2%, untreated 15%. Cross with primary rescue origin (Ekterly / injectable / none) using `joint_pools`. IR and XR overlap; do not sum their membership as unique HAE patients.
- Ekterly launch stock is projected from enrollment forms, paid conversion, uniqueness and continuation. It is not an observed PHVS-launch census.
- Source-specific annual opportunities and choice factors are in `segmented_adoption.rates`. Preserve denominators, including the fixed 80% normalization where applicable, followed by actual paid conversion once. Do not replace this with a universal acquisition probability.
- Annual-to-monthly acquisition probability is `1-(1-annual_probability)^(1/12)`, not annual/12.
- Follow the exact order in `Population.step`: growth, cooldown release, attempts, access queue/failed access, activation, survival/exposure and retirement. Preserve the access delay, failed-access cooldown and default no discontinuation re-entry.
- Each product ledger conserves available + pending + cooldown + active + retired by origin. Cohort fractional patient counts are intentional.
- Survival over a month: `opening * retention^(1/12)`; patient-years: `(opening-closing)/(-LN(retention))`, with `opening/12` at retention 1. XR first-year versus mature retention switches at cohort age 12, not calendar year-end.
- Cohort matrices may be aggregated only where mathematically equivalent to the engine. Never replace exposure with average year-end patients.

### Paid quantity and price

- XR revenue uses retained patient-years × paid intensity × annual net price.
- IR uses paid initial stock plus replacement doses. Initial two doses are inside the annual budget, including when supplied free; do not add a second stocking allowance.
- **Fixed IR quantity multiplier = 1.17/1.38 = 0.8478260869565217 in every active case.** The 17% versus 38% clinical redosing comparison is transferred to paid quantity as a thesis assumption, not an observed commercial effect. No 12% numerator, no alternative/no-discount case.
- **IR/XR efficacy price multipliers stay 1.0. Mature PHVS share of modeled oral choices stays 60%.** Launch shares and transition speed differ; do not turn 60% into total disease-market share.
- XR working net price: $48,487.92 per 28 days × 365/28 × 66%. IR: $16,720 per dose × 84%. These are comparator price anchors and assumed net fractions, not PHVS realized prices.
- Apply overlapping PHVS XR/IR allocation once and the LTP rescue budget appropriately. No extra 67.1% claimant filter or second adherence haircut.
- Overseas ratios use lagged US sales and phased rollout. Recognize direct sales or partner royalty revenue on a given portion, never both. No second FX conversion on the ratio-derived sales.

### Calendar financials and valuation

- Monthly engine July 2026–December 2055; working valuation 9 October 2026, target 9 October 2027. Keep this dated snapshot for parity; a refreshed valuation must be explicitly versioned.
- IR commercial launch May 2027; XR June 2028, plus the separate access delay. Keep decision dates distinct from commercial launch dates.
- Historical 2024/2025/H1 2026 accounts remain EUR. Opening June 2026 balances are translated once at 1.1186 USD/EUR. Do not use that rate to invent historical USD growth.
- Show H1 actual + H2 forecast clearly for 2026, quarterly 2027–2028, annual 2027–2055. Flows sum; stocks use closing balance. Exposure is 1/12-year monthly; discount timing uses actual days/365.25 and partial-month cutoffs.
- Use existing explicit R&D/fixed commercial/G&A schedule, variable commercial costs, manufacturing, royalties, milestones, taxes/NOL proxy, capex, depreciation and NWC. Do not present the simplified economic statements as complete IFRS forecasts.
- Cost multiplier 0.8–1.2 applies to fixed R&D/commercial/G&A, not automatically to every expense line. Higher cost does not automatically generate higher sales. Expense assumptions are underwriting and need prominent source/rationale visibility, not company-guidance labels.
- Cash funding threshold, target balance, fees, issue price and economic option/RSU dilution follow the engine. Future funding is required capital, not committed cash. Do not count future operating losses and all diagnostic future issuance dilution twice.
- Four clinical states change costs as well as revenues. Working weights: both `.9*.9/.95`, IR-only `.9-Pboth`, XR-only `.9-Pboth`, neither `1-1.8+Pboth`. Cash is not multiplied by PoS. Neither-success branch winds down after the decision-based delay.
- State expected value is the sum of state probability × state equity per share. Label both-success financial forecasts as conditional; do not imply they are probability-adjusted revenue.
- Patent erosion uses the fixed 50/50 broad-study/rare-disease framework: blend revenue indices and quantity indices, then effective price = revenue/quantity. Preserve protected-year anchors and cost runoff timing. After explicit 2055 cash flow, declining perpetuity is `FCF2055*(1+g)/(r-g)`, discounted from end-2055. Do not introduce an exit multiple or positive perpetual growth.
- Shared costs mean a two-product SOTP allocation is not unique. Show product revenue and four-state valuation clearly; any IR/XR value allocation must disclose its shared-cost allocation and reconcile exactly.

## Monte Carlo presentation and refresh policy

Production has already run: 4,096 primary draws, 2,048 independent-seed check, twelve separate scenario runs; 22,528 commercial draw/scenario evaluations and 90,112 conditional clinical valuations in total. Do not rerun production merely to build a workbook.

Import the published summary and compact primary draw data with the source hashes/run identity. Preserve 28 triangular priors; fixed policies remain fixed. Keep separate access packages, structural scenarios and correlation stress separate; do not pool them with invented probabilities.

Suggested chart labels:

1. **How valuation changes across commercial assumptions:** histogram of primary risk-adjusted expected 12-month values. $2 bins are fine. Show median and P10/P90; label range as assumption-driven, not empirical confidence interval or stock-price probability.
2. **Which uncertain assumptions influence valuation:** Spearman rank association, clearly labeled as association, not dollars per share, causal contribution or ranking of short arguments. Fixed thesis assumptions cannot appear here because they do not vary.
3. Optional separate clinical-outcome chart: conditional state values weighted by the four probabilities. Do not overlay it on the main histogram by default. Explain with a numerical example: each trial has four outcome values; averaging them yields one expected value, while retaining the four weighted values yields a different distribution with the same mean.

Saved Monte Carlo charts do not change when workbook inputs change. Clearly label them **Saved simulation — 9 October 2026 assumptions**. Implement a comparison to the imported baseline across all model-relevant editable inputs (not just 28 priors); show stale status after changes. Document the external refresh command/workflow and replace outputs only after verification. Do not use volatile RAND cells to imply a fully functioning new MC engine. If implementing live simulation later, it must reproduce canonical seed/distribution/clinical handling and be separately validated.

No current share price is configured. Leave return/upside/downside versus market blank until a verified dated quote is entered. Never invent a profitable-short probability from the histogram. Changes in dated market quote alone do not change the saved company-value distribution, but affect any return calculation.

## Acceptance values and verification

Use exact files for full-precision controls; display rounded values:

| Metric | Expected control |
|---|---:|
| Working value on valuation date | $25.825555267/share |
| Working expected 12-month value | $28.311616929/share |
| Both-success target value | $31.495442922/share |
| Both-success peak recognized global revenue | $1,551.214526 million in 2038 |
| Primary MC mean expected target | $26.378087944/share |
| Primary MC median expected target | $25.981210140/share |
| Primary MC P10 / P90 | $17.842044876 / $35.271706772 |

The working case is not the Monte Carlo mean or median. Never recenter simulations or alter inputs to make these equal. Historical $21.50 → $31.70 → $28.31 represented blended adoption → segmented adoption → redosing; none included an efficacy price premium.

1. Reconcile revenues within $0.01 million and values within $0.01/share against the unchanged Python engine. Compare each clinical branch, working/cautious/upside bundles, quarter/year aggregation, quantities and cash/funding/shares—not just one target price.
2. Verify zero conversion, delayed launches, survival boundaries, population conservation, overlap and initial-stock budget, geographic lag, clinical weights, patent timing and balance-sheet reconciliation. See `validate.py` for exact test intent.
3. Change an adoption driver, a retention driver, a cost driver and the case selector in a disposable workbook copy. Confirm downstream patient/revenue/cash/valuation updates, Monte Carlo stale flag and unchanged historical actuals. Restore originals.
4. Recalculate in the supported spreadsheet engine; scan errors and circular references; verify stored export values. State any unverified native Excel behavior honestly.
5. Render and inspect all final sheets for clipping, broken charts, readable fonts and printing. Group support detail rather than hiding essential assumptions. Final file should not depend on either reference workbook or the creator's local absolute paths.
6. Save a concise reconciliation/implementation note with software/version, source commit and any limitations. Push the workbook and necessary reproducible builder/export instructions to GitHub, without node_modules or unrelated raw files. No workbook exists from this handoff task itself.

## Reproducible data recovery

From repository root with Python 3.10+:

```sh
python3 models/finalized/engine.py
python3 models/finalized/analyze.py
python3 models/finalized/validate.py
python3 -m unittest discover -s models/monte_carlo -v
python3 models/monte_carlo/verify_artifacts.py
```

Monthly cohort/financial output files are intentionally ignored by Git and are regenerated by `engine.py`; do not mistake that for missing data. Run regeneration in a clean isolated checkout or inspect any resulting differences before committing. Stored quarterly/annual controls and compressed MC data are tracked. Do not modify canonical code/inputs to resolve a spreadsheet mismatch or invalidate the saved-run hashes unnoticed.

## Deliverables

- One polished, formula-driven PHVS workbook with transparent saved Monte Carlo outputs.
- A short reconciliation and refresh note, including any unresolved limitations.
- Reproducible construction files as needed, with input source/version recorded.
- Final GitHub location and a concise statement of what recalculates in Excel versus what requires an external Monte Carlo refresh.

The user's next agent owns the workbook build. This handoff does not authorize silently changing the investment assumptions or claiming the short is independently proven by arithmetic checks.
