"""Reproduce historical calibration, not a PHVS forecast."""
import json
from pathlib import Path
ROOT = Path(__file__).resolve().parents[2]
d = json.loads((ROOT / "research/PHVS_population_launch_calibration_2026-10-08.json").read_text())
h = d["orladeyo_history"]
a, b = d["population_anchors"][:2]
for row in (a, b):
    assert row["treated"] + row["not_treated"] == row["diagnosed"]
prior = h["YE2023_paid_plus_PAP"] / (1 + h["YE2023_growth_yoy_rounded"])
out = {
    "status": "Historical checks; arithmetic validation only, not independent market validation",
    "YE2023_paid_patients_approx": round(h["YE2023_paid_plus_PAP"] * h["YE2023_paid_fraction"]),
    "YE2022_paid_plus_PAP_approx_from_rounded_growth": round(prior),
    "FY2023_net_paid_plus_PAP_additions_approx": round(h["YE2023_paid_plus_PAP"] - prior),
    "FY2023_net_all_status_additions_observed": h["FY2023_net_additions_including_QuickStart"],
    "2019_2023_average_annual_net_diagnosed_stock_change": (b["diagnosed"]-a["diagnosed"])/4,
    "2019_2023_average_annual_net_treated_stock_change": (b["treated"]-a["treated"])/4,
    "FY2023_net_paid_additions": None,
    "FY2023_gross_paid_starts": None,
    "note": "Net diagnosis stock change is not incident diagnosis. Paid additions cannot be identified without prior aligned paid stock; gross starts additionally require exits and payment transitions."
}
(ROOT / "models/patient_adoption/market_calibration_results.json").write_text(json.dumps(out, indent=2)+"\n")
print(json.dumps(out, indent=2))
