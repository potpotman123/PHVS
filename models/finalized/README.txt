PHVS CANONICAL MODEL HANDOFF — 9 October 2026

Start here. This directory supersedes the working patient_adoption modules for
forward financial valuation. The earlier modules remain historical research,
not parallel sources of current assumptions. The broker-shaped reverse DCF is
an archived research exercise and is not a submission model or market consensus.

Working investment framing: standalone short thesis, tested against a strong
long steelman; acquisition is a separate upside risk to the short.

Purpose: an independent, reproducible patient-to-cash-flow engine and Excel
specification. No workbook or slides are produced. All pivotal inputs resolve
to public evidence or an explicitly identified underwriting judgment.

Read in order:
  1. HANDOFF.txt — results, readiness, decisions and limitations.
  2. ARCHITECTURE.txt — definitions and formulas.
  3. inputs.json / history.json / sources.json — canonical input authority.
  4. EXCEL_SPEC.txt — implementation instructions and acceptance checks.
  5. monte_carlo_spec.json / acquisition_spec.json — extension interfaces.
  6. changes.json / validation.json — changes and checks.

Reproduce from repository root, Python 3 standard library only:
  python3 models/finalized/engine.py
  python3 models/finalized/analyze.py
  python3 models/finalized/validate.py

The engine writes local monthly cohort/accounting results (ignored by Git;
regenerate them), deterministic cases and quarterly tables. Analyze writes
calendar financial tables, sensitivities, launch checks and driver ranking.
Validation checks identities and edge cases; it is not commercial validation.
Inputs are not fitted to a stock price. Ranges are not confidence intervals.

Money: forecast flows USD millions; per-dose/year prices absolute USD;
historical financials EUR millions; shares millions; patient counts actual
fractional people. Historical EUR is not translated at today's FX to measure
growth. The 2026 full-year USD planning bridge is explicitly constant-FX.

Operating statements are conditional on both products being approved. Four
clinical states determine risk-adjusted valuation. Acquisition is a separate
terminal outcome, not an extra SOTP asset. No acquisition premium is included
in the working standalone value. Full simulation has not been run.

Do not change input values in output files. Change inputs.json, explain the
source/judgment and regenerate outputs. Freeze case definitions before a pitch.
