CURRENT EXCEL BUILD ENTRY POINT — 10 October 2026
Read ../../EXCEL_BUILD_HANDOFF.md. Production Monte Carlo is complete in
../monte_carlo/results/REPORT.md. Statements below that it has not run are historical.

PHVS CANONICAL STANDALONE MODEL — 9 October 2026

Current task: evidence-led SHORT pitch; no acquisition modeling.
Fixed in every active case:17%versus38%redosing discount, no efficacy pricing
premium,60%mature oral choice. MonteCarlo is for a separate agent to run.

Start with MONTE_CARLO_HANDOFF.txt. It explains what happened, constraints,
inputs, dependencies, equations, validation, simulation outputs and compute budget.
MODEL_AUDIT.txt and assumptions_audit.json document evidence strength and limits.
inputs.json/history.json/sources.json are canonical; historical calibration tables
are provenance, not parallel sources of current assumptions.

Current working expected12month value$28.31; today'svalue$25.83.
The historical$21.50->$31.70increase was segmented adoption, not a pricepremium.
Applying the fixed redosing discount produced$28.31. Do not reintroduce no-
discount cases from historical snapshots. No market-return probability computed.

Reproduce from repository root (Python3 standard library):
  python3 models/finalized/engine.py
  python3 models/finalized/analyze.py
  python3 models/finalized/validate.py
  python3 models/finalized/mc_preflight.py
Last command runs bounded implementation checks, not productionMonteCarlo.

Architecture:ARCHITECTURE.txt, segmented_adoption.py, engine.py.
Simulation preparation:monte_carlo_priors.json, monte_carlo_spec.json, mc_prepare.py.
Results:summary.txt/results.json; uncertainty:sensitivities.json/driver_ranking.json.
Checks:validation.json/mc_preflight_results.json/cohort_optimization_check.json.
Workbook contract:EXCEL_SPEC.txt. No workbook or slides produced here.
Monthly generated files are ignored byGit and reproducible. Sources resolve to
public documents or explicit judgment. Archive/ contains superseded specifications.
