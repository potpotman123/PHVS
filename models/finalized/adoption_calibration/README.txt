HISTORICAL CALIBRATION PROVENANCE
Current engine inputs live in ../inputs.json. The full monthly choice grid is
generated locally by choice_share_schedule.py and excluded fromGit; formulas
and compact inputs are tracked. See ../MONTE_CARLO_HANDOFF.txt for current scope.

CURRENT ORAL-CHOICE SCHEDULE (user-directed,9October2026)
Launch share: cautious25%,working40%,stronger60%. All converge linearly
to60%by60months after each formulation launch and staythere. Timingisassumed.
choice_share_schedule.py produces choice_share_timepath.json; this supersedes
constant oral-choice factors in the static tables BELOW for intended implementation.
Sharesareconditionalonoralchoice,notannualswitchingoroverallmarketshare.
All-destination capture and new-prevention selection remain separate.
Fullengine integration and valuation remain pending.

ALL-GROUP ADOPTION CALIBRATION — 9 October 2026
Proposed evidence-anchored assumptions; not empirical estimates or updated valuation.
Rates below are completed-start proxies, with access held at80%.
Cautious / Working / Stronger; annual except NEW conditional selection events.

XR_orladeyo: 4.00% / 9.15% / 16.00%
XR_injectable_prevention: 2.70% / 4.94% / 7.56%
XR_older_prevention: 2.70% / 4.94% / 7.56%
XR_rescue_only: 2.49% / 4.56% / 6.98%
XR_diagnosed_untreated: 0.60% / 2.20% / 5.04%
IR_injectable_rescue: 8.09% / 12.95% / 18.89%
IR_ekterly: 4.00% / 9.15% / 16.00%
IR_no_current_rescue: 3.36% / 8.06% / 14.11%
NEW_XR_given_prevention_selection: 20.00% / 28.00% / 36.00%
NEW_IR_given_rescue_selection: 33.60% / 40.32% / 47.04%

Read provenance.json for exact formulas, sources, judgments and limitations.
Main weak bridges: preventive-switch rate borrowed for IR Ekterly; untreated
engagement; historical firstoral entrant mix transferred to a followon oral.
These are transparent usable placeholders, not newfound empirical rates.
Newly treated choice shares are NOT annual hazards or extra population flows.
No double access haircut; claims starts and prescription forms differ.
Reproduce: python3 models/finalized/adoption_calibration/build.py
Engine integration and origin-stock/taper calibration remain pending.

CHOICE-SHARE AUDIT: research/PHVS_choice_share_comparable_audit_2026-10-09.json
finds no universal empirical60%efficacypremium. Attruby18-20%estimatedNBRx;
Scemblix23%1Lversus57-59%laterlineillustrate timing/denominator differences.
The repeated60%values are distinct assumptions,not onevalidated coefficient.
At25/40/60/80%oralchoice,injectableXRcompletedstartproxiesare2.06/3.29/4.94/6.59%.
