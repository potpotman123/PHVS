# PHVS standalone Monte Carlo

Start with the **[results and interpretation](results/REPORT.md)**. This is an additive simulation layer over the unmodified `models/finalized/` engine at [e7594c4](https://github.com/potpotman123/PHVS/commit/e7594c4a8e7d5f4571055e52aa9c3320b39c5ab2), following its [Monte Carlo handoff](../finalized/MONTE_CARLO_HANDOFF.txt) and [prior manifest](../finalized/monte_carlo_priors.json).

The user's meeting transcript motivated Monte Carlo because patient transitions and commercial assumptions are uncertain and often transferred from small or imperfect source cohorts. This implements that uncertainty analysis for the **standalone short thesis**. IR means immediate-release rescue treatment; XR means extended-release preventive treatment. The two products share a patient population, and a preventive user can also use rescue. Neither headcounts nor revenues are simply added without overlap accounting.

## Reproduce

Python 3.10+; numerical engine, sampling, tests and packaging use only the standard library. The published run used Python 3.12. All commands below run from the repository root; the scripts also resolve model paths when called elsewhere. Use a fresh output folder if the engine or priors change: the checkpoint guard intentionally refuses to mix versions.

```sh
python3 -m unittest discover -s models/monte_carlo -v
python3 models/monte_carlo/run.py --draws 4096 --seed 7202609 --workers 4 --output /tmp/phvs-mc
python3 models/monte_carlo/run.py --draws 2048 --seed 7202610 --workers 2 --output /tmp/phvs-mc
python3 models/monte_carlo/suite.py --draws 1024 --workers 4 --output /tmp/phvs-mc
python3 models/monte_carlo/run.py --draws 4096 --structural later_LOE --workers 2 --output /tmp/phvs-mc
python3 models/monte_carlo/run.py --draws 2048 --dependence commercial_rank_linked --workers 4 --output /tmp/phvs-mc
python3 models/monte_carlo/package_results.py --input /tmp/phvs-mc
python3 models/monte_carlo/verify_artifacts.py
```

The pilot used 64 draws. Primary summaries preserve 64/128/256/512/1,024/2,048/4,096 prefixes. The later-exclusivity scenario was extended after missing the prespecified stability threshold at 1,024 and narrowly at 2,048; the primary was extended to the same 4,096 draws for a matched paired comparison. Two-seed and prefix comparisons are sampling diagnostics, not validation of investment probabilities. Execution timing in a summary measures that invocation's new draws only and excludes summary serialization. The second seed is a robustness check, not pooled into headline results. Local CPU workers run the same deterministic engine; results do not depend on scheduling.

The rank-linked dependence stress was also extended from 1,024 to 2,048 because its outcome P10 and median had not stabilized at the smaller count. Every final scenario passes the stated prefix diagnostic. No assumption or distribution was altered to obtain that result.

For a new exploratory scenario, use the manifest's named access/structural/dependence options with `run.py --help`. Never pool their outputs without a defensible probability model. `suite.py` runs twelve alternatives: two access packages, nine structural variations, and one commercial dependence stress.

Optional charts require matplotlib. After generating results, run `python3 models/monte_carlo/plot_results.py`. The chart package is not needed to reproduce any valuation. Published charts use matplotlib 3.11.2. The artifact manifest covers numerical/report files; chart file hashes are separately in `chart_manifest.json`.

## Calculation

1. Generate 28 independent primitive uniform quantiles from the manifest's seed using Python `random.Random`. Inverse triangular CDFs use the repository's bounds and modes; only duration inputs marked integer are rounded. Retain all other working inputs.
2. Apply a named access package jointly to conversion, delay and net fractions; apply at most one structural variation. The rank-linked stress shares a percentile across seven demand variables before inverse CDF transformation. All mandatory fixed policies remain fixed.
3. Run the original monthly engine through 2055 for each of four clinical states, including shared population, origin-specific opportunity/choice/access, competitor stock, retention, dose use, revenue, costs, taxes, capital needs and the post-exclusivity tail.
4. Store conditional state prices and funding. Compute each draw's expected valuation by weighting once: both 85.2632%, IR only 4.7368%, XR only 4.7368%, neither 5.2632%. These preserve each product's 90% marginal probability. Alternative clinical dependence weights reuse those same state prices; they do not trigger new commercial simulations.
5. Report separately (a) commercial uncertainty in the within-draw clinical-risk-adjusted expectation and (b) the distribution of conditional outcome values weighted by state probability divided by draw count. Their means must agree; their quantiles generally differ. Clinical outcomes are not simulated a second time on an already risk-adjusted value.

Empirical quantiles use the inverse CDF rather than interpolating across discrete clinical masses. Mean MC standard errors use independent commercial draws. Approximate quantile MC errors use eight interleaved groups of commercial draws, keeping the four clinical outcomes together. Paired scenario differences use identical primitive quantiles and the matching primary prefix. Spearman correlations use average ranks for ties; they are prior-dependent associations, not causality or variance attribution.

## Preserved controls

- Fixed rescue quantity multiplier: 1.17/1.38 = 0.847826; no efficacy-based price multiplier; mature oral choice share 60%; segmented adoption on; obsolete second hazard taper disabled.
- Prevention opportunity, oral preference, oral-pair choice, incumbent departure capture and paid conversion retain separate denominators. No universal 54% TAM haircut, second persistence haircut, or duplicate redosing haircut.
- First-year XR retention retains the updated efficacy-only allowance, with mode/upper bound 69.45%, not the superseded 61% working case or 74% broader clinical-benefit assumption. No new adverse-event retention benefit.
- July 2026 financial opening history, October 2026 valuation, October 2027 target, May 2027 IR and June 2028 XR commercial launch assumptions, phased overseas recognition, fair-financing treatment and explicit post-LOE decay are inherited unchanged.
- No M&A branch, takeover premium, buyer cost synergies, or invented scenario weights. No replacement of the engine by a response-surface approximation. No recentering priors toward a desired short conclusion.

## Audit trail

- [Recent update reconciliation](RECENT_UPDATES.md): which of the 23 new commits govern this run, including superseded findings.
- [Source coverage](source_coverage.json): every one of 102 changed paths, commit URLs, source blob hashes, available exact-byte text hashes and disposition. This inventories all changes, not a claim to independently reread every original third-party source.
- [Upstream validation](upstream_validation.json): original 152 model checks and 72 preflight checks, run on an isolated copy so source files remain unchanged.
- [Run regression tests](test_run.py): weighting, sampler boundaries, fixed controls, point-case reproduction, seeded resume, worker invariance, and stale-checkpoint rejection.
- [Analysis JSON](results/analysis.json): both distributions, replication/prefix diagnostics, paired sensitivities, rank associations, capital needs and scenario-level accounting residuals.
- Each scenario folder has a summary with source/code SHA-256 hashes and a gzip compact output table. `primitive_inputs.json.gz` records the sampled independent input values per seed. Scenario access/structural packages and rank linking are fully specified in the manifest and summary; the packager replays **every** sampled override against its checkpoint. The compact tables retain unrounded floating-point values.
- `process_event.json` is a new proposed log event crediting the user's thesis and method. It is separate because the user requested that existing repository files remain untouched; no historical log entries were edited.

Checkpoint `draws.jsonl` files contain fuller per-draw diagnostics locally. The committed gzip tables contain draw IDs, expected values, conditional state prices, state funding/diluted shares, 2030 product revenues and peak revenue/year. Full detailed paths can be regenerated from the stored seed and exact hashed source. A checkpoint failure writes `failure.json` and halts; invalid draws are never silently clipped or replaced.

The implementation does not produce a current-market-price downside probability because the canonical manifest has no verified, dated price reference. Conditional business values are not executable stock prices, and Monte Carlo does not turn judgmental priors into evidence that a short will work.
