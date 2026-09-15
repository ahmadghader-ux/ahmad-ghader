#!/usr/bin/env python3
"""
Builds NH_TikTok_EcommerceTrading_Only25.xlsx — 25 Saudi(-market) creators
restricted to two niches only (E-commerce, Trading), dropping Real Estate
entirely per the request that generated this file.

Sourced in two passes:
  1. Reused from the earlier 3-niche scout (build_tiktok_scout.py): the 5
     original E-commerce candidates and 10 original Trading candidates,
     carried over unchanged (same handles, same flags/status).
  2. A fresh research pass (this session) adding 7 new E-commerce and 10
     new Trading leads, narrowed to 7 E-commerce + 3 Trading survivors after
     dropping one duplicate of an already-excluded lead (see EXCLUDED below)
     and one manual follow-up search to fill the last e-commerce slot.

Same no-fabrication contract as the prior scout: every claim needs a
citable source, unverified nationality is flagged rather than assumed, and
namesake/wrong-country traps are called out explicitly rather than silently
resolved.
"""
import re
from collections import Counter

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

FONT_NAME = "Arial"
NAVY = "1F4E78"
BAR_COLORS = {"E-commerce": "2E7D32", "Trading": "C62828"}

CANDIDATES = [
    # ================= E-COMMERCE (12) =================
    # ---- reused from the original 3-niche scout (unchanged) ----
    dict(name="عبدالله الفوزان / Abdullah Al-Fawzan", category="E-commerce",
         handle="@econabdallah", platform="tiktok",
         niche="Dropshipping, launching an online store \"in 24 hours,\" TikTok paid-ads strategy",
         credentials="Describes achieving \"financial freedom\" before age 30; Riyadh-based",
         socials="X @econabdullah; IG @econabdullah; YouTube econAbdullah; Snapchat @econabdullah; econabdullah.com",
         followers="~178.1K (947.9K likes)",
         why_fit="Clearest, best-documented Saudi e-commerce-education creator, with consistent Riyadh/Saudi signals across every platform",
         risk_flag="", depth_confidence="High",
         sources="https://www.tiktok.com/@econabdallah ; https://x.com/econabdullah ; https://www.instagram.com/econabdullah/",
         status=""),
    dict(name="عادل صالح / Adel Saleh", category="E-commerce",
         handle="@adeel8866", platform="tiktok",
         niche="Dropshipping supplier lists, Chinese import/shopping sites, \"e-commerce A-to-Z\" tutorials",
         credentials="n/a (no CV found)", socials="Telegram t.me/Adelacademy",
         followers="n/a",
         why_fit="Content explicitly hashtagged #السعودية and #adelsaleh; practical step-by-step dropshipping tutorials",
         risk_flag="", depth_confidence="Medium — Saudi market-focus clear from hashtags, personal nationality not independently confirmed",
         sources="https://www.tiktok.com/@adeel8866/video/7542995492918119698", status=""),
    dict(name="ملاذ المصري / Malaz Al-Masri", category="E-commerce",
         handle="@malazmarketing", platform="tiktok",
         niche="E-commerce/digital marketing training, Salla & Shopify store setup, TikTok/Meta ad campaigns",
         credentials="BA marketing (King Abdulaziz Univ, Jeddah); MA e-marketing (Cairo Chamber of Commerce); runs ad agency \"MalazooAds\"; 15-18 yrs in Saudi & Turkish markets",
         socials="IG/X @malazooads; YouTube; malazmarketing.com", followers="n/a",
         why_fit="Deepest professional credentials of any e-commerce candidate found; long track record training Saudi merchants on Salla/Shopify",
         risk_flag="", depth_confidence="Flagged — nationality uncertain: surname \"Al-Masri\" (\"the Egyptian\") + Cairo credential suggest possible Egyptian origin working/educated in Saudi Arabia; his own site markets him as training \"in Saudi Arabia,\" not explicitly as a Saudi national",
         sources="https://malazmarketing.com/", status="Flag: nationality uncertain (possibly Egyptian, Saudi-market based)"),
    dict(name="Ecommerce Academy / Katib School (أيت كاتب محمد)", category="E-commerce",
         handle="@katibschool", platform="tiktok",
         niche="Free e-commerce training, \"from A to Z\"",
         credentials="IG bio cites 9 yrs marketing/e-commerce experience; founder identified as \"أيت كاتب محمد\" (Ait Katib Mohammed)",
         socials="IG @katibschool (~55K); Facebook; katibschool.com", followers="n/a on TikTok",
         why_fit="Established free e-commerce training brand", risk_flag="",
         depth_confidence="Flagged — likely non-Saudi: \"Ait Katib\" is an Amazigh/North-African (commonly Moroccan) naming pattern, content appears pan-Arab rather than Saudi-specific",
         sources="https://www.instagram.com/katibschool/ ; https://katibschool.com/",
         status="Flag: nationality likely non-Saudi (Moroccan naming pattern), not confirmed Saudi"),
    dict(name="Omar — \"تجارة الكترونية بالعربي\"", category="E-commerce",
         handle="@maketingomar", platform="tiktok",
         niche="Arabic-language e-commerce content", credentials="n/a",
         socials="n/a found", followers="n/a",
         why_fit="Only verifiable TikTok presence found for this exact search angle",
         risk_flag="", depth_confidence="Low — no biographical detail, nationality, or credentials found beyond the handle itself",
         sources="https://www.tiktok.com/@maketingomar", status="Insufficient verification — bare handle only"),

    # ---- new this pass ----
    dict(name="ياسر الدبيخي / Yasser Al-Dubaikhi", category="E-commerce",
         handle="@yasseraldbikhi", platform="instagram",
         niche="Dropshipping/e-commerce education — Shopify store-building, supplier sourcing, product selection, logistics",
         credentials="Aggregator-listed full name \"ياسر عمر بن ابراهيم الدبيخي\" (Al-Dubaikhi is a known Najdi Saudi tribal name); master's in international trade (UEA, UK) per search summary",
         socials="YouTube @YasserAldbikhi; X @YasserAldbikhi", followers="~100K (search-summarized, not independently re-confirmed — TikTok fetch blocked)",
         why_fit="Multi-platform dropshipping/e-commerce educator with a Saudi-pattern full name and a credentialed international-trade background",
         risk_flag="Follower count and full name sourced via search-engine summary of a biography aggregator (manhom.com), not a directly-loaded profile",
         depth_confidence="Medium-High — Saudi-style full name across platforms, but no primary-source nationality document",
         sources="https://www.instagram.com/yasseraldbikhi/ ; https://www.youtube.com/@YasserAldbikhi ; https://x.com/yasseraldbikhi ; https://manhom.com/شخصيات/ياسر-عمر-بن-ابراهيم-الدبيخي/",
         status="Verified only via search snippets / third-party aggregator — not independently re-confirmed"),
    dict(name="Abdulrahman Khaled", category="E-commerce",
         handle="@abdulrahman_khaalid", platform="tiktok",
         niche="Driving product sales/revenue directly through TikTok content (organic growth, not paid ads) — adjacent to e-commerce education",
         credentials="Display name carries a 🇸🇦 flag emoji; self-reported claim of \"20M+ SAR in sales from content without paid promotion\"; runs free workshops on the method",
         socials="n/a confirmed beyond TikTok", followers="~3.6M / 22.1M likes (search-summarized, not independently re-confirmed)",
         why_fit="Large Saudi-flagged TikTok account teaching e-commerce sellers how to grow revenue through content",
         risk_flag="Niche is closer to TikTok growth/sales coaching than pure store-building education — borderline fit; stats unverified beyond search summary",
         depth_confidence="Medium — handle and self-identified flag solid, niche specificity and stats not hard-confirmed",
         sources="https://www.tiktok.com/@abdulrahman_khaalid", status="Niche borderline; stats unverified"),
    dict(name="فراس عثمان / Firas Othman", category="E-commerce",
         handle="@firasothmanacademy", platform="tiktok",
         niche="Shopify/e-commerce academy — full paid courses on store setup, marketing, e-commerce fundamentals",
         credentials="Self-described e-commerce expert since 2017; runs \"Firas Othman Academy\" (Shopify Masterclass, Complete Shopify Course)",
         socials="YouTube @firasothmanacademy; Facebook facebook.com/firas.othman.ent; firasothman.com",
         followers="n/a",
         why_fit="Full-time, multi-platform e-commerce educator with a dedicated academy brand",
         risk_flag="Namesake trap confirmed: an unrelated Saudi healthcare executive, Dr. \"Firas Othman Al-Ghamdi,\" shares a similar name — do not conflate; this creator's own nationality could not be confirmed in any source found",
         depth_confidence="Low-Medium — real, active academy, but nationality unconfirmed",
         sources="https://firasothman.com/ ; https://www.youtube.com/@firasothmanacademy ; https://www.tiktok.com/@firasothmanacademy/video/7081975066686278917 ; namesake contrast: https://manhom.com/شخصيات/فراس-الغامدي/",
         status="Nationality uncertain — namesake trap flagged, do not assume Saudi"),
    dict(name="إيهاب أبو دية / Ehab Abu Dayeh", category="E-commerce",
         handle="@ehababudayeh", platform="instagram",
         niche="E-commerce training academy (\"The 7 Planets\") — dropshipping/online-store education for Arab beginners",
         credentials="Search summary: trained 15,000+ students; founder of an e-commerce academy; started by selling cake balls before moving into e-commerce",
         socials="Facebook facebook.com/EhabAbuDayeh; the7planets.com", followers="~517K (search-summarized, not independently re-confirmed)",
         why_fit="Large, established e-commerce education following with a named academy brand",
         risk_flag="No Saudi nationality indicator found anywhere; academy branding/site language reads regionally Jordan-associated; content targets the broad Arab market rather than Saudi specifically",
         depth_confidence="Medium — real, large, findable account, but likely non-Saudi",
         sources="https://www.instagram.com/ehababudayeh/ ; https://the7planets.com/en/about-us/ ; https://www.facebook.com/EhabAbuDayeh/",
         status="Nationality uncertain / likely non-Saudi (regionally Jordan-linked) — not confirmed Saudi"),
    dict(name="محمود ورده / Mahmoud Wardeh", category="E-commerce",
         handle="@wardehecom", platform="instagram",
         niche="E-commerce personal-brand-building and store-scaling education",
         credentials="Search summary: personally built 12+ e-commerce brands", socials="YouTube @MahmoudWardeh",
         followers="~196K (search-summarized, not independently re-confirmed)",
         why_fit="Sizable, dedicated e-commerce-focused personal brand with cross-platform presence",
         risk_flag="No Saudi nationality indicator found; \"Wardeh\" is a common Levantine (Jordanian/Palestinian) surname; no source ties him to Saudi Arabia specifically",
         depth_confidence="Low-Medium — real findable account, but likely non-Saudi",
         sources="https://www.instagram.com/wardehecom/ ; https://www.youtube.com/@MahmoudWardeh",
         status="Nationality uncertain / likely non-Saudi (Levantine surname) — insufficiently verified as Saudi"),
    dict(name="بدر / Badr", category="E-commerce",
         handle="@marketingwithbadr", platform="tiktok",
         niche="TikTok marketing/e-commerce content — TikTok Shop tutorials, ad-campaign tips, engagement/growth for sellers",
         credentials="n/a beyond content itself", socials="n/a found",
         followers="Reported ~1.2M in an earlier search summary — could not be independently re-confirmed against a primary source",
         why_fit="Active, dedicated TikTok account specifically teaching e-commerce marketing/TikTok Shop selling",
         risk_flag="Weakest identity verification on this list — no confirmed full surname, no nationality indicator; essentially a bare-handle entry",
         depth_confidence="Low — real, active handle confirmed via multiple video URLs, but no verified full name, credentials, or nationality",
         sources="https://www.tiktok.com/@marketingwithbadr/video/7517334023706791188 ; https://www.tiktok.com/@marketingwithbadr/video/7509590945566952722 ; https://www.tiktok.com/@marketingwithbadr/video/7509918191267728658",
         status="Insufficient verification — bare handle only; full name and nationality unconfirmed"),
    dict(name="محمد شعبان / Mohamed Shaban", category="E-commerce",
         handle="@marketingwithshaban", platform="tiktok",
         niche="Digital-marketing/e-commerce education — Salla & Zid (Saudi e-commerce platforms), TikTok ads, WhatsApp ads",
         credentials="Instagram bio cites 10+ yrs marketing/advertising experience; search summary describes the account as \"based in Riyadh, Saudi Arabia\"",
         socials="Instagram @marketingwithshaban (~205K); possible secondary TikTok @marketingwithshaban24 (unconfirmed if same operator)",
         followers="~334K on TikTok per an influencer-ranking aggregator (not independently re-confirmed — TikTok fetch blocked)",
         why_fit="Active Riyadh-based creator whose content specifically covers Saudi e-commerce platforms (Salla/Zid) and Saudi-market TikTok ad tactics",
         risk_flag="\"Shaban\" is a common Egyptian surname — plausible expatriate-in-Saudi rather than Saudi national; nationality not independently confirmed beyond \"Riyadh-based\" aggregator language",
         depth_confidence="Medium — real, active, cross-platform account with a specific Saudi-market niche fit, but nationality unconfirmed and plausibly non-Saudi",
         sources="https://www.tiktok.com/@marketingwithshaban ; https://www.instagram.com/marketingwithshaban/ ; https://www.tiktok.com/@marketingwithshaban/video/7503239282119871751",
         status="Nationality uncertain / plausibly non-Saudi (Egyptian surname, Riyadh-based) — not confirmed Saudi"),

    # ================= TRADING (13) =================
    # ---- reused from the original 3-niche scout (unchanged) ----
    dict(name="فؤاد الحربي / Fuad Al-Harbi", category="Trading",
         handle="@fuad8k", platform="tiktok",
         niche="Saudi stock market (Tadawul) beginner education — explicitly no crypto, no fund management",
         credentials="Runs \"Fuad Academy\"; paid beginner course (349 SAR) with completion certificate",
         socials="X @FUAD7333; YouTube @FuadAcademy; fuadacademy.com", followers="n/a",
         why_fit="Explicit \"stocks only, no crypto, no money management\" positioning suggests deliberate compliance awareness",
         risk_flag="Paid course sold via social media — verify he isn't implying licensed advisory status", depth_confidence="Medium",
         sources="https://fuadacademy.com ; https://x.com/fuad7333", status=""),
    dict(name="تركي الحربي / Turki Al-Harbi", category="Trading",
         handle="@kii31t", platform="tiktok",
         niche="Personal investing/savings education; sells an \"intelligent investment package\" (261 SAR)",
         credentials="Self-described trainer; no license mentioned", socials="Snapchat (~64K subs); X @kii31t",
         followers="n/a on TikTok",
         why_fit="Cross-platform Saudi investment educator",
         risk_flag="Sells a paid \"investment package\" without disclosed licensing — potential unlicensed-advice risk",
         depth_confidence="Medium", sources="Search snippets for \"kii31t تركي الحربي\"; Snapchat profile page", status=""),
    dict(name="@deserttrader6906", category="Trading",
         handle="@deserttrader6906", platform="tiktok",
         niche="US stock market + crypto (Bitcoin, altcoins), general investing",
         credentials="n/a", socials="n/a found", followers="~52K (per influencer-listing snippet)",
         why_fit="Bio explicitly states content is personal opinion, not financial advice — a compliance-conscious disclaimer",
         risk_flag="None major noted", depth_confidence="Flagged — Saudi-base claim rests on a third-party aggregator listing (Heepsy-style), not a primary-source bio",
         sources="Heepsy/aggregator listing describing account as \"based in Saudi Arabia\"", status="Flag: nationality verified only via third-party aggregator"),
    dict(name="د. اسعد رزق / Dr. Assad Rizq", category="Trading",
         handle="@assadalsharef", platform="tiktok",
         niche="Crypto/blockchain education (Bitcoin, altcoins), \"how to build wealth with crypto\"",
         credentials="Described as a former Saudi neurologist who left medicine after discovering Bitcoin in 2013 while studying in Hungary; claims 9+ yrs experience, 10,000+ students",
         socials="n/a confirmed", followers="~1.1M (per search snippet)",
         why_fit="Unusually strong personal-credibility narrative (doctor-turned-crypto-educator), large following, Saudi hashtags (#السعودية)",
         risk_flag="Crypto education content — verify no unlicensed investment-scheme promotion; \"build wealth\" framing leans aspirational-marketing",
         depth_confidence="Medium-High", sources="YouTube video \"قصة د. أسعد رزق\"; TikTok search snippets for \"@assadalsharef\"", status=""),
    dict(name="بسام بن سليمان العبيد / Bassam bin Solaiman Al-Obaid", category="Trading",
         handle="@bin_solaiman", platform="tiktok",
         niche="US stock/options trading strategies, also covers Saudi market (#الاسهم_السعودية)",
         credentials="Writer for Al-Eqtisadiah (Saudi financial newspaper), consultant/analyst for TickerChart",
         socials="X @BinSolaiman; YouTube \"Bassam Alobid\"", followers="n/a",
         why_fit="Strongest verifiable professional credential of the set (mainstream financial-press byline)",
         risk_flag="Options-trading content carries inherent leverage/risk; no red flags found beyond that",
         depth_confidence="High", sources="TikTok video snippets tagged #تداول #الاسهم_السعودية; TickerChart profile; YouTube channel", status=""),
    dict(name="محمد الباهلي / Muhammad Al-Bahli", category="Trading",
         handle="@m_albahly", platform="tiktok",
         niche="Saudi stock market (Tadawul) news/education — dividend stocks, sukuk, automated \"rebound zone\" chart screens",
         credentials="n/a beyond content itself", socials="n/a found", followers="n/a",
         why_fit="Consistently labels content \"not a recommendation\" (ليست توصية) across videos — better compliance hygiene than most",
         risk_flag="Automated \"screening\" videos could be read as stock tips despite disclaimers",
         depth_confidence="Medium", sources="Multiple TikTok video URLs under @m_albahly (#تاسي #السوق_السعودي)", status=""),
    dict(name="المستر سلطان المضحي / Mr. Sultan Al-Mudhi", category="Trading",
         handle="@mrwallstreets.com", platform="tiktok",
         niche="US stocks, options contracts, gold, oil, crypto — framed around Saudi CMA-approved trading platforms",
         credentials="Self-described advisor/educator; no license cited",
         socials="IG @mrwallstreet.official; X @mrofwallstreat", followers="n/a",
         why_fit="Content directly addresses Saudi regulatory approval status of brokers — audience-relevant and somewhat compliance-aware",
         risk_flag="Options/leveraged-instrument education for retail audience carries elevated risk profile even when not \"hype\" style",
         depth_confidence="Medium", sources="TikTok video URLs under @mrwallstreets.com (#المستر_سلطان #الاسهم_الامريكية #عقود_الاوبشن)", status=""),
    dict(name="العمدة أبو أحمد الاسهم / Al-Omda Abu Ahmad", category="Trading",
         handle="@abo_ahmed_al3mda", platform="tiktok",
         niche="US stock technical analysis / \"fractal equations\"; explicitly \"not investment recommendations\"",
         credentials="n/a", socials="YouTube \"قناة أبو أحمد للأسهم\"", followers="~11.9K",
         why_fit="Niche technical-analysis angle",
         risk_flag="", depth_confidence="Flagged — nationality/Saudi base not confirmed: Gulf-Arabic content but no source ties this account specifically to Saudi Arabia rather than another Gulf state",
         sources="TikTok profile/video snippets; YouTube channel listing", status="Flag: Gulf-Arabic content, Saudi base unconfirmed (could be another Gulf state)"),
    dict(name="رعد الحربي (أبو دانه) / Raad Al-Harbi (\"Abu Danah\")", category="Trading",
         handle="@abu_danah00", platform="tiktok",
         niche="Saudi (Tadawul) stock and fund trading education — technical/fundamental analysis, portfolio building, beginner-to-pro courses",
         credentials="Runs \"Abu Danah Al-Harbi Academy\"; bachelor's degree in finance/economics (per search summaries); content emphasizes trading \"with awareness\" to avoid heavy losses",
         socials="X @abu_danah01 (\"اسهم وصناديق\"); YouTube @Abudanahlearn; abudanahacademy.com",
         followers="n/a",
         why_fit="Named individual (tribal surname Al-Harbi) with self-declared Saudi flag; cross-verified via 3 independent properties beyond the TikTok page (X, YouTube, own academy site); risk-aware framing, not signal-selling",
         risk_flag="", depth_confidence="High — real name + 3 independent corroborating sources",
         sources="https://www.tiktok.com/@abu_danah00 ; https://x.com/abu_danah01 ; https://www.youtube.com/@Abudanahlearn ; https://abudanahacademy.com/",
         status=""),
    dict(name="محمد الخضير / Mohammed Alkhudhair", category="Trading",
         handle="@mohakhuu", platform="tiktok",
         niche="Retail investing/fintech education — comparing Saudi investment/trading apps (Derayah, Tamra, Malaa, Drahim, Abyan, Al Rajhi Capital), tracking Tadawul holdings",
         credentials="Self-described financial advisor; master's degree in law; publicly invested SAR 250,000 across 5 licensed Saudi investment apps for a year to compare real returns",
         socials="X @Mohakhu (~120K+); Instagram @mokhco",
         followers="~114K",
         why_fit="Named individual with a disclosed, non-hype comparison methodology; strongly Saudi-contextualized content (PIF, Saudia Airlines, Saudi banks, CMA-licensed apps); cross-verified via X + Instagram under consistent name/photo",
         risk_flag="", depth_confidence="High — real name + disclosed methodology + 2 independent corroborating platforms",
         sources="https://www.tiktok.com/@mohakhuu ; https://x.com/Mohakhu ; https://www.instagram.com/mokhco/",
         status=""),

    # ---- new this pass ----
    dict(name="أبو فهد العتيبي / Abu Fahd Al-Otaibi", category="Trading",
         handle="@aaghhhhhrr", platform="tiktok",
         niche="Daily/live technical analysis and buy-sell commentary on Tadawul-listed stocks (\"تحليل اسهم\" livestreams)",
         credentials="Self-described \"Saudi stock market expert\"; claims either 20 or 25 years' experience (sources conflict); runs a companion Telegram channel and lists a Saudi mobile number for ad inquiries",
         socials="Snapchat @aasdf2932 (~17.6K); X @aaghhhhhrr; Telegram channel; possible second TikTok @_e5ti (unconfirmed if same operator)",
         followers="~73.4K (TikTok, per search-result description); ~17.6K (Snapchat)",
         why_fit="Active, TikTok-native creator whose entire content focus is live Tadawul stock analysis, with a real cross-platform footprint",
         risk_flag="Runs a paid-inquiry Telegram channel alongside stock analysis — possible monetized recommendations/signals; experience-length claim inconsistent across sources; nationality inferred from name/mobile prefix/content only, not independently confirmed",
         depth_confidence="Medium — activity and follower counts corroborated across 4 independent listings, but no formal biography or licensing record found",
         sources="https://www.tiktok.com/@aaghhhhhrr ; https://www.tiktok.com/@aaghhhhhrr/video/7554543587338259730 ; https://www.snapchat.com/@aasdf2932 ; https://nicegram.app/hub/channel/aaghhhhhrrr ; https://x.com/aaghhhhhrr?lang=ar",
         status="Possible unlicensed signal/advice distribution via Telegram; experience claim inconsistent; nationality inferred, not independently confirmed"),
    dict(name="أنس الراجحي / Anas Al-Rajhi", category="Trading",
         handle="@Anas_S_Alrajhi", platform="x",
         niche="Financial/fundamental and technical analysis of TASI (Tadawul) stocks; investor-education commentary on Saudi and global markets",
         credentials="Identified in multiple independent Saudi news outlets as a \"Saudi financial markets analyst\"; described as a certified lecturer (\"محاضر معتمد\") at Money Experts Institute for Training; sources claim he \"holds a trusted license,\" not independently verified against a regulator",
         socials="Snapchat @anas_alrajhi; quoted/featured on ajel.sa, okaz.com.sa, and Asharq News (now.asharq.com)",
         followers="n/a — no numeric follower count found",
         why_fit="Repeatedly quoted by mainstream Saudi press specifically on Tadawul/insurance-sector stock commentary and investor education — strong independent corroboration of both Saudi nationality and niche fit",
         risk_flag="No TikTok account found (X/Snapchat only); a third-party influencer-rating aggregator (BrokerTrust) marks him \"unverified\" with a trust score of 43, which conflicts with other sources' \"trusted license\" claim — licensing status should be treated as unconfirmed/mixed",
         depth_confidence="High for identity/nationality (multiple mainstream Saudi press citations under full name); Medium for platform/follower specifics",
         sources="https://x.com/anas_s_alrajhi ; https://www.snapchat.com/@anas_alrajhi?locale=en-US ; https://ajel.sa/local/yuvg564r8d ; https://ajel.sa/local/wnxrrf7tqk ; https://www.okaz.com.sa/author/1581/1/أنس-الراجحي ; https://now.asharq.com/clips/1390911/ ; https://brokertrust.co/financial-influencers/anas-al-rajhi/",
         status="TikTok presence not found (X/Snapchat only); licensing claim conflicts with third-party trust-score aggregator"),
    dict(name="علي العمري / Ali Alamri", category="Trading",
         handle="@d.alialamri", platform="tiktok",
         niche="Saudi retail-investor content — trading-platform comparisons, TASI stock tips/hashtagged content, savings/investment motivation",
         credentials="n/a — no independent biography or licensing info found for the account operator specifically",
         socials="Snapchat @d.alialamri (bio references interest in international business administration and TASI investment, per search summary)",
         followers="n/a",
         why_fit="Content is squarely and consistently focused on the Saudi stock market (TASI) and Saudi-specific trading-platform comparisons",
         risk_flag="Namesake trap: a separate TikTok account @dr.alialamri (\"د.علي العمري\") also exists, and a real, separately-verified Saudi business consultant, Dr. Ali bin Hamid Al-Amri, shares a similar name — no evidence links that consultant to this account, do not conflate; at least one video is tagged \"#توصيات\" (stock recommendations) alongside #تاسي #تداول, a possible unlicensed-recommendation risk that could not be confirmed or ruled out",
         depth_confidence="Low — identity, credentials, and nationality not independently confirmed beyond the account's own content/hashtags; namesake collision adds confusion risk",
         sources="https://www.tiktok.com/@d.alialamri/video/7492412102863260945 ; https://www.tiktok.com/@d.alialamri/video/7651850824406158593 ; https://www.snapchat.com/@d.alialamri?locale=en-US ; namesake contrast: https://www.tiktok.com/@dr.alialamri",
         status="Namesake trap flagged (do not conflate with Dr. Ali bin Hamid Al-Amri, business consultant); nationality uncertain; possible unlicensed-recommendation risk via #توصيات hashtag"),
]

EXCLUDED = [
    dict(name="Hussam Ansari (@hussam.ansari)", category="E-commerce",
         reason="Teaches Amazon FBA \"in Saudi Arabia and UAE,\" but own hashtags (#ukpakistani, #overseaspakistani) indicate Pakistani nationality/heritage — market-focus-vs-nationality trap. Carried over from the prior 3-niche scout."),
    dict(name="Zuhair Abu Al-Reesh", category="E-commerce",
         reason="Genuine, well-documented Saudi-market digital-marketing/TikTok-ads trainer (23 yrs experience), but no personal TikTok creator account located — fails identifiable-account bar. Carried over from the prior 3-niche scout."),
    dict(name="Yazeed (@YazeedKM_)", category="E-commerce",
         reason="Founder of e-commerce training platform \"Adatik,\" active/credentialed on X only. Excluded in the prior scout for lacking a TikTok handle; this pass's research independently resurfaced the same handle and it is excluded again for the same reason — no TikTok presence, thin single-platform profile, no confirmed full name or nationality."),
    dict(name="Ali Al-Hamed (علي الحامد / @alialhamed.97)", category="Trading",
         reason="Well-documented halal-investing financial influencer, but multiple sources (including his own TikTok bio) identify him as Emirati (\"مؤثر مالي مرخص في الإمارات\"), not Saudi, and his content is not Saudi-market-specific — wrong-country trap, re-confirmed independently in this pass."),
    dict(name="Tariq Al-Harbi (@tariq_alharbi80)", category="Trading",
         reason="Name/hashtag matched trading searches, but account's own listed profession is \"Actor\" with entertainment content — wrong niche despite name match. Carried over from the prior 3-niche scout."),
    dict(name="Mohammed Al-Shumaimari (أبو سهم الشميمري)", category="Trading",
         reason="Genuinely credentialed Saudi markets analyst (34-yr track record, TV/press commentator), but no TikTok handle could be verified despite targeted searches. Carried over from the prior 3-niche scout."),
    dict(name="Rateel Alshehri", category="Trading",
         reason="Confirmed Saudi, TikTok-verified, but content niche is youth entrepreneurship/podcasting, not trading — niche mismatch, not nationality. Carried over from the prior 3-niche scout."),
    dict(name="Thaal indicator (@thaalpro)", category="Trading",
         reason="A paid TradingView technical-indicator/subscription product, not an individual creator; one promotional clip is framed around \"gold recommendations in the UAE\" — both a wrong-country signal and a company/product mismatch."),
    dict(name="Zaid Othman (@zo14777)", category="Trading",
         reason="Active forex/gold trading-course creator with a real TikTok/YouTube presence, but no source found any indication of his nationality one way or the other (audience discussion surfaced in a Jordan-focused forum, which isn't evidence of his own nationality) — left out as an unresolved lead rather than guessed at."),
]

METHOD_NOTES = [
    "No-fabrication contract: every candidate is a real, findable account verified via web search; every claim (follower count, credentials, nationality) is sourced. No name, handle, or statistic was invented.",
    "Unverified nationality is flagged, not silently assumed or used to exclude a candidate — see the Status/Flag column on each row.",
    "Namesake and wrong-country traps (same/similar name as an unrelated person; creators targeting the Saudi market without being Saudi nationals) are called out explicitly in Risk/Flag.",
    "This environment's network egress proxy blocked direct page loads of tiktok.com, instagram.com, snapchat.com, and most influencer-listing aggregators (hypeauditor.com, iqfluence.io, manhom.com, brokertrust.co, etc.) during this research pass — all data below comes from search-engine result snippets/summaries, not directly re-loaded profile pages. Follower counts in particular should be treated as indicative only and re-verified by Northouse directly before outreach or valuation decisions.",
    "Real Estate entries from the earlier 3-niche scout (10 candidates) were dropped entirely — this file covers E-commerce and Trading only, per request.",
]


TIKTOK_URL_RE = re.compile(r"https?://(?:www\.)?tiktok\.com/@[A-Za-z0-9_.]+")
X_URL_RE = re.compile(r"https?://(?:www\.)?x\.com/[A-Za-z0-9_]+")
HANDLE_RE = re.compile(r"@[A-Za-z0-9_.]+")


def account_url(c):
    """Prefer a URL for the candidate's stated primary platform already
    present in the sourced citations; otherwise derive it deterministically
    from the verified handle."""
    if c["platform"] == "x":
        m = X_URL_RE.search(c["sources"])
        if m:
            return m.group(0)
        h = HANDLE_RE.search(c["handle"])
        return f"https://x.com/{h.group(0)[1:]}" if h else None
    if c["platform"] == "instagram":
        h = HANDLE_RE.search(c["handle"])
        return f"https://www.instagram.com/{h.group(0)[1:]}/" if h else None
    m = TIKTOK_URL_RE.search(c["sources"])
    if m:
        return m.group(0)
    h = HANDLE_RE.search(c["handle"])
    return f"https://www.tiktok.com/{h.group(0)}" if h else None


def style_header(ws, row, ncols, fill=NAVY):
    header_font = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor=fill)
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)


def autosize(ws, min_width=8, max_width=48):
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            col = cell.column_letter
            length = len(str(cell.value))
            widths[col] = min(max(widths.get(col, min_width), length + 2), max_width)
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def parse_followers(followers_field):
    if not followers_field:
        return None
    s = followers_field.strip()
    if s.lower().startswith("n/a") or s == "":
        return None
    m = re.search(r"([\d.]+)\s*([KM])", s, re.IGNORECASE)
    if not m:
        return None
    num = float(m.group(1))
    mult = 1_000_000 if m.group(2).upper() == "M" else 1_000
    return num * mult


def main():
    assert len(CANDIDATES) == 25, f"Expected 25 candidates, got {len(CANDIDATES)}"
    counts = Counter(c["category"] for c in CANDIDATES)
    assert set(counts) == {"E-commerce", "Trading"}, f"Unexpected categories: {set(counts)}"

    wb = openpyxl.Workbook()
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ---------------- Creators tab ----------------
    ws = wb.active
    ws.title = "Creators"
    headers = ["#", "Name (AR/EN)", "Niche", "Platform", "Account (click)", "Followers",
               "Content Focus", "Why Notable/Fit", "Risk/Flag", "Status", "Sources"]
    ws.append(headers)
    style_header(ws, 1, len(headers))
    ws.freeze_panes = "A2"

    for i, c in enumerate(CANDIDATES, 1):
        url = account_url(c)
        ws.append([i, c["name"], c["category"], c["platform"].capitalize(), c["handle"],
                   c["followers"], c["niche"], c["why_fit"], c["risk_flag"], c["status"], c["sources"]])
        r = ws.max_row
        if url:
            cell = ws.cell(row=r, column=5)
            cell.hyperlink = url
            cell.style = "Hyperlink"

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers)):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    autosize(ws, max_width=42)

    # ---------------- Chart tab (text-rendered bars — survives Sheets conversion) ----------------
    ws2 = wb.create_sheet("Chart")
    ws2["A1"] = "25 Candidates by Niche (E-commerce / Trading only)"
    ws2["A1"].font = Font(name=FONT_NAME, bold=True, size=13)
    ws2.append([])
    ws2.append(["Niche", "Count", "Bar"])
    style_header(ws2, ws2.max_row, 3)

    niches = ["E-commerce", "Trading"]
    BAR_UNIT = 3
    for niche in niches:
        n = counts.get(niche, 0)
        bar = "█" * (n * BAR_UNIT)
        ws2.append([niche, n, bar])
        r = ws2.max_row
        ws2.cell(row=r, column=3).font = Font(name="Consolas", color=BAR_COLORS.get(niche, "444444"), size=12)

    ws2.append([])
    ws2.cell(row=ws2.max_row + 1, column=1, value="Disclosed Followers (thousands)").font = Font(name=FONT_NAME, bold=True, size=13)
    ws2.append([])
    ws2.append(["Handle", "Niche", "Followers (K)", "Bar"])
    style_header(ws2, ws2.max_row, 4)

    follower_rows = []
    for c in CANDIDATES:
        f = parse_followers(c["followers"])
        if f is not None:
            follower_rows.append((c["handle"], c["category"], f / 1000))
    follower_rows.sort(key=lambda x: -x[2])

    if follower_rows:
        max_k = max(v for _, _, v in follower_rows)
        for handle, cat, k in follower_rows:
            bar_len = max(1, round((k / max_k) * 40))
            bar = "█" * bar_len
            ws2.append([handle, cat, round(k, 1), bar])
            r = ws2.max_row
            ws2.cell(row=r, column=4).font = Font(name="Consolas", color=BAR_COLORS.get(cat, "444444"), size=12)

    for row in ws2.iter_rows(min_row=1, max_row=ws2.max_row, max_col=4):
        for cell in row:
            if cell.value is not None and cell.column != 3:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws2.column_dimensions["A"].width = 24
    ws2.column_dimensions["B"].width = 14
    ws2.column_dimensions["C"].width = 12
    ws2.column_dimensions["D"].width = 45

    # ---------------- Excluded / Notes tab ----------------
    ws3 = wb.create_sheet("Excluded & Method")
    ws3["A1"] = "Excluded Leads (real but disqualified)"
    ws3["A1"].font = Font(name=FONT_NAME, bold=True, size=13)
    ws3.append([])
    ws3.append(["Name", "Niche", "Reason Excluded"])
    style_header(ws3, ws3.max_row, 3)
    for e in EXCLUDED:
        ws3.append([e["name"], e["category"], e["reason"]])

    ws3.append([])
    r = ws3.max_row + 1
    ws3.cell(row=r, column=1, value="Method Notes").font = Font(name=FONT_NAME, bold=True, size=13)
    ws3.append([])
    for note in METHOD_NOTES:
        ws3.append([note])
        ws3.merge_cells(start_row=ws3.max_row, start_column=1, end_row=ws3.max_row, end_column=3)

    for row in ws3.iter_rows(min_row=1, max_row=ws3.max_row, max_col=3):
        for cell in row:
            if cell.value is not None:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws3.column_dimensions["A"].width = 30
    ws3.column_dimensions["B"].width = 12
    ws3.column_dimensions["C"].width = 70

    for name in ("Creators", "Chart", "Excluded & Method"):
        for row in wb[name].iter_rows():
            for cell in row:
                if cell.font.name != "Consolas":
                    existing = cell.font
                    cell.font = Font(name=FONT_NAME, bold=existing.bold, color=existing.color,
                                      size=existing.size if existing.size else 10)

    out = "NH_TikTok_EcommerceTrading_Only25.xlsx"
    wb.save(out)
    print(f"Wrote {out}: {len(CANDIDATES)} candidates "
          f"({counts['E-commerce']} E-commerce, {counts['Trading']} Trading), "
          f"{len(follower_rows)} with a disclosed follower count.")


if __name__ == "__main__":
    main()
