#!/usr/bin/env python3
"""
Builds a CSV of the "New 25" list narrowed to TikTok-confirmed accounts only,
per the request that every creator in this sheet be on TikTok specifically.

Reuses the 10 TikTok-platform candidates already in build_ecom_trading25_new.py
CANDIDATES, and adds 8 newly-researched TikTok-confirmed candidates (3
e-commerce, 5 trading) found in a dedicated follow-up research pass that
excluded every handle already used across all prior rounds.

Total: 18 of 25. The remaining 7 slots could not be filled with candidates
that are BOTH (a) not already used in any prior round of this project and
(b) confirmed to have an active TikTok account — the research agents
explicitly reported that most recognizable Saudi trading/e-commerce
personalities are X/Snapchat/Telegram-native, not TikTok-native. Padding to
25 would have required either reusing round-1 candidates (breaking the
"all new" requirement) or including unverifiable/fabricated entries, which
violates this project's no-fabrication contract. Left at 18, flagged
honestly, pending the user's choice on how to close the gap.
"""
import csv

from build_ecom_trading25_new import CANDIDATES as EXISTING

TIKTOK_URL = "https://www.tiktok.com/@{}"

NEW_TIKTOK_CANDIDATES = [
    # ---- E-commerce (3) ----
    dict(name="زيد الدوسري / Zaid Al-Dosari (Z Brand / زي براند)", category="E-commerce",
         handle="@zbrandsa",
         niche="Founder who scaled a Saudi fashion store into a multi-branch brand and expanded to sell via Amazon and TikTok Shop, managed through Zid",
         followers="~29.6K (per indexed search snippet — not independently re-confirmed)",
         why_fit="Featured as an official Zid merchant success story; grew from ~100K SAR seed capital to a reported ~13M SAR revenue with 6+ physical branches (Riyadh/Jeddah/Dammam/Khobar/Tabuk/Abha) — corroborated across 3 independent sources (Zid blog, X, YouTube)",
         risk_flag="This is a brand/store account, not a personal how-to education channel — confirm whether store-owner case studies fit the outreach goal vs. pure educator accounts",
         status="Brand-account caveat; Saudi nationality strongly implied (Zid is Saudi-only, branches across KSA cities) but not from a primary-ID source"),
    dict(name="the_etchh (name on account unconfirmed — possibly Abdallah Hesham)", category="E-commerce",
         handle="@the_etchh",
         niche="Teaches how to open a TikTok Shop seller account and sell into the US market while living abroad; #تجارة_الكترونية #دروبشيبينج #البيع_على_أمازون",
         followers="n/a",
         why_fit="Content niche is a direct match (TikTok Shop + dropshipping education), account confirmed active via a live video permalink",
         risk_flag="Video title's Egyptian colloquial phrasing (\"ازاي افتح حساب... وانا مقيم خارج امريكا\") is a strong signal pointing to an Egyptian creator rather than Saudi",
         status="Nationality unconfirmed — likely Egyptian, not Saudi; verify before outreach"),
    dict(name="وليد صويتي / Walid Swaiti", category="E-commerce",
         handle="@walid_swaiti",
         niche="\"How to earn from Noon with no capital,\" affiliate marketing (#تسويق_بالعمولة #التجارة_الالكترونية), tagged with both #السعودية and #الامارات hashtags",
         followers="n/a",
         why_fit="Direct niche match (Noon selling + affiliate/e-commerce marketing), multiple confirmed active video permalinks",
         risk_flag="Targets both Saudi and UAE audiences simultaneously; surname \"Swaiti\" is a common Levantine/Jordanian surname — cannot rule out non-Saudi origin",
         status="Nationality unconfirmed — surname suggests possible Levantine/Jordanian origin, not Saudi; verify before outreach"),

    # ---- Trading (5) ----
    dict(name="منصة المضارب الذكية 1M / cci_555", category="Trading",
         handle="@cci_555",
         niche="Saudi stock market (TASI) trading \"signals\"/AI-assisted picks, plus gold; cross-posts to Telegram and YouTube, runs a paid subscription store",
         followers="n/a",
         why_fit="Directly on-niche, Saudi-market-specific hashtags (#تداول_الاسهم_السعودية #منصة_المضارب_الذكية_1m #توصيات), active multi-video TikTok presence",
         risk_flag="HIGH — appears to be a paid stock-recommendation subscription service; textbook unlicensed-advice pattern the Saudi CMA has prosecuted; reads as a brand/team account, not one identifiable individual",
         status="Brand account, not individual — unlicensed-signals risk"),
    dict(name="منصة العز للأسهم / ezshares", category="Trading",
         handle="@ezshares",
         niche="Saudi stock market (TASI) content, drives followers to an off-platform Telegram group (\"قروب العز\")",
         followers="n/a",
         why_fit="TASI/Saudi-stock hashtags directly, TikTok presence with multiple confirmed posts",
         risk_flag="Funnels viewers to an unverified off-platform Telegram group; a secondary source's attribution of a specific personal name to this account could not be substantiated and is explicitly not included",
         status="Individual identity unconfirmed — do not treat secondary-source name attributions as verified"),
    dict(name="الاسهم السعودية 🇸🇦🇺🇸 / ashomksa", category="Trading",
         handle="@ashomksa",
         niche="Saudi stock market recommendations/analysis, including a dedicated \"Saudi stock recommendations\" (توصيات) playlist",
         followers="n/a",
         why_fit="Explicit Saudi-flag branding, dedicated Saudi-stock-recommendations content",
         risk_flag="MEDIUM-HIGH — a dedicated recommendations playlist is a direct unlicensed-tipping red flag; an unverified Facebook complaint against the account surfaced but its specifics could not be confirmed; account also posts unrelated non-finance content (mixed-content account)",
         status="Unverified FB complaint against account + mixed-content, not pure finance niche"),
    dict(name="ahmedalbadry06", category="Trading",
         handle="@ahmedalbadry06",
         niche="Gold and forex trading explainer content",
         followers="n/a",
         why_fit="Direct forex/gold trading content, confirmed active TikTok video",
         risk_flag="Only one video surfaced, insufficient to assess signal-selling risk",
         status="Saudi nationality unconfirmed — possible non-Saudi (Egyptian/Levantine surname pattern); weak evidence overall"),
    dict(name="forex.knowledge.1", category="Trading",
         handle="@forex.knowledge.1",
         niche="Forex/options trading education, including a \"free trading bot\" giveaway video",
         followers="n/a",
         why_fit="Forex trading niche, confirmed active TikTok presence",
         risk_flag="MEDIUM — a \"free trading bot\" giveaway is a common lure pattern in forex-scam content; treat with caution",
         status="Saudi nationality unconfirmed — generic pan-Arab forex account, weakest candidate of the batch"),
]


def main():
    rows = []
    for c in EXISTING:
        if c.get("platform") != "tiktok":
            continue
        rows.append([c["name"], c["category"], c["handle"], TIKTOK_URL.format(c["handle"].lstrip("@")),
                     c["followers"], c["why_fit"], c["risk_flag"], c["status"]])
    for c in NEW_TIKTOK_CANDIDATES:
        rows.append([c["name"], c["category"], c["handle"], TIKTOK_URL.format(c["handle"].lstrip("@")),
                     c["followers"], c["why_fit"], c["risk_flag"], c["status"]])

    out = "/tmp/all_tiktok_25.csv"
    with open(out, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["#", "Name", "Niche", "Handle", "Direct Profile Link", "Followers", "Why Fit", "Risk/Flag", "Status"])
        for i, r in enumerate(rows, 1):
            w.writerow([i] + r)
    print(f"Wrote {out}: {len(rows)} TikTok-confirmed candidates "
          f"({sum(1 for r in rows if r[1]=='E-commerce')} E-commerce, "
          f"{sum(1 for r in rows if r[1]=='Trading')} Trading)")


if __name__ == "__main__":
    main()
