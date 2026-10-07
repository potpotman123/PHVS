"""Bounded CMS gross-spending decomposition; no manufacturer net-price inference.
Usage: bundled-python research/loe_pxq_audit.py /tmp/phvs_cms_2024.csv
Source CSV URL and checksum are retained with the extracted observations.
"""
import csv
import hashlib
import json
import sys
from pathlib import Path

path = Path(sys.argv[1])
brands = {"Kuvan", "Xenazine", "Orfadin"}
rows = [r for r in csv.DictReader(path.open())
        if r["Brnd_Name"] in brands and r["Mftr_Name"] == "Overall"]
assert len(rows) == 3 and {r["Brnd_Name"] for r in rows} == brands
observations = []
for r in rows:
    years = []
    for year in range(2020, 2025):
        def val(prefix):
            x = r[f"{prefix}_{year}"]
            return float(x) if x else None
        spend, units = val("Tot_Spndng"), val("Tot_Dsg_Unts")
        years.append(dict(year=year, gross_spend_usd=spend, dosage_units=units,
                          claims=val("Tot_Clms"), beneficiaries=val("Tot_Benes"),
                          gross_spend_per_raw_unit=spend / units,
                          cms_weighted_spend_per_unit=val("Avg_Spnd_Per_Dsg_Unt_Wghtd")))
    start, end = years[0], years[-1]
    q = end["dosage_units"] / start["dosage_units"]
    p = end["gross_spend_per_raw_unit"] / start["gross_spend_per_raw_unit"]
    rev = end["gross_spend_usd"] / start["gross_spend_usd"]
    assert abs(q * p - rev) < 1e-12
    observations.append(dict(brand=r["Brnd_Name"], raw_source_row=r, annual=years,
        change_2020_to_2024=dict(gross_spending=rev-1, dosage_units=q-1,
                               gross_spend_per_raw_unit=p-1), identity_verified=True))
out = dict(as_of="2026-10-07", source_url="https://data.cms.gov/sites/default/files/2026-06/98218f98-166c-4723-8438-c344a4ef96a6/DSD_PTD_RY26_P04_V10_DY24_BGM.csv",
    source_sha256=hashlib.sha256(path.read_bytes()).hexdigest(),
    population="Medicare Part D; Overall manufacturer rows only",
    limitations=["Gross claim spending excludes confidential manufacturer rebates.",
                 "Raw dosage units mix strengths/forms; not standardized patient-days.",
                 "Gross spending per unit includes mix; not a pure price index.",
                 "Suppressed beneficiary cells are unknown, not zero.",
                 "2020 is not a clean pre-entry year for these products.",
                 "Do not divide worldwide manufacturer revenue by Medicare volume."],
    observations=observations)
destination = Path(__file__).with_name("loe_pxq_evidence_2026-10-07.json")
destination.write_text(json.dumps(out, indent=2) + "\n")
for obs in observations:
    print(obs["brand"], obs["change_2020_to_2024"])
