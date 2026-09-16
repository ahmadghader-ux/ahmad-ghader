# NH Expert/Creator Scout — reconstructed

`NH_Expert_Creator_Scout.xlsx` is a rebuilt version of the Northouse (NH)
creator-scouting spreadsheet referenced in a Claude Code session transcript
(`conversation_export_2026-08-13.txt`, shared via Slack). That session ran six
parallel research agents to find real Saudi experts across leadership,
speaking, AI-for-professionals, personal finance, career verticals
(medical/nursing/accounting/PMP/etc.), and "hidden experts" (high authority,
near-zero online distribution) — candidates Northouse could acquire or build
into paid-community creators.

## Provenance

The original workbook was built live in that session via a Python/`openpyxl`
script and is not recoverable — only the chat transcript survived. This file
is reconstructed from the six scout agents' raw pipe-delimited result
payloads as they were quoted back into that chat, deduplicated and organized
into tabs. It has **not** been re-verified against live sources; treat it as
a faithful transcript reconstruction, not a fresh research pass.

## Tabs

- **All Candidates** — every candidate transcribed from the six scouts
  (121 rows), with category, credentials, socials, selling status, an
  expert-depth score (1-10, as assessed by the sourcing agent), and source
  URLs. `Status / Flags` notes disqualifications (`DQ: ...`), cross-listings
  (the same person found independently by two scouts — a convergence
  signal), and nationality/verification flags.
- **Top Picks** — the explicitly named "top pick" / "top 5" candidates from
  each scout's own summary (27 rows).
- **Verticals & Gaps** — the career-vertical priority ranking, honest gaps
  (niches with no credible named individual found), and candidates the
  scouts explicitly excluded during research (namesake traps, wrong-country
  matches, etc.).
- **Method** — sourcing rules (no-fabrication contract, namesake-trap
  checks), row counts, and caveats (e.g. Favikon-sourced entries show an
  influence score but no public follower count).

## Rebuilding

`build_scout.py` regenerates the workbook from the transcribed data:

```
pip install openpyxl
python3 build_scout.py
```

## Saudi TikTok creator batches (trading / e-commerce / real estate)

- **Batch 1** (`build_tiktok_scout.py` → `NH_TikTok_Ecommerce_Trading_RealEstate_Scout.xlsx`,
  plus the Google-Sheets/PDF/age-format/Final25 variants): 25 candidates
  across e-commerce, trading, and real estate.
- **Batch 2** (`build_tiktok_scout_batch2.py` → `NH_TikTok_Scout_Batch2_EcommerceTrading.xlsx`):
  24 NEW candidates, trading and e-commerce only, excluding every handle
  already used in Batch 1. One short of the requested 25 — the research
  passes declined to invent a 25th name to hit the round number; see the
  "Method" tab for the shortfall note.

Both batches were sourced the same way: no direct TikTok/X/Instagram/Snapchat
scraping was possible (this environment's egress proxy blocks those
domains), so every candidate was found and cross-verified via web search,
under a no-fabrication contract — every claim needs a citable source URL,
unknown fields are "n/a", and nationality/niche-mismatch traps are flagged
rather than silently resolved.
