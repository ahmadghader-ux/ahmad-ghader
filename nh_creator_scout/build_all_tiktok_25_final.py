#!/usr/bin/env python3
"""
Builds the final 25-row all-TikTok CSV: the 18 TikTok-confirmed candidates
from build_all_tiktok_25.py, plus 7 highest-confidence TikTok creators
reused from the original round-1 scout (build_ecom_trading25.py) to reach
25. Each row is marked with a "Source" column so it's clear which are new
to this round vs. reused from round 1, per the user's explicit choice to
close the gap this way.
"""
import csv

from build_ecom_trading25 import CANDIDATES as ROUND1
from build_all_tiktok_25 import EXISTING, NEW_TIKTOK_CANDIDATES

TIKTOK_URL = "https://www.tiktok.com/@{}"

REUSED_HANDLES = ["@econabdallah", "@adeel8866", "@bin_solaiman", "@abu_danah00",
                  "@mohakhuu", "@assadalsharef", "@m_albahly"]


def main():
    rows = []

    for c in EXISTING:
        if c.get("platform") != "tiktok":
            continue
        rows.append(["New (this round)", c["name"], c["category"], c["handle"],
                     TIKTOK_URL.format(c["handle"].lstrip("@")), c["followers"],
                     c["why_fit"], c["risk_flag"], c["status"]])

    for c in NEW_TIKTOK_CANDIDATES:
        rows.append(["New (this round)", c["name"], c["category"], c["handle"],
                     TIKTOK_URL.format(c["handle"].lstrip("@")), c["followers"],
                     c["why_fit"], c["risk_flag"], c["status"]])

    by_handle = {c["handle"]: c for c in ROUND1}
    for h in REUSED_HANDLES:
        c = by_handle[h]
        rows.append(["Reused (Round 1)", c["name"], c["category"], c["handle"],
                     TIKTOK_URL.format(c["handle"].lstrip("@")), c["followers"],
                     c["why_fit"], c["risk_flag"], c["status"]])

    assert len(rows) == 25, len(rows)

    out = "/tmp/all_tiktok_25_final.csv"
    with open(out, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["#", "Source", "Name", "Niche", "Handle", "Direct Profile Link",
                    "Followers", "Why Fit", "Risk/Flag", "Status"])
        for i, r in enumerate(rows, 1):
            w.writerow([i] + r)

    ec = sum(1 for r in rows if r[2] == "E-commerce")
    tr = sum(1 for r in rows if r[2] == "Trading")
    print(f"Wrote {out}: 25 rows ({ec} E-commerce, {tr} Trading), "
          f"{sum(1 for r in rows if r[0].startswith('New'))} new, "
          f"{sum(1 for r in rows if r[0].startswith('Reused'))} reused")


if __name__ == "__main__":
    main()
