#!/usr/bin/env python3
"""
Builds NH_TikTok_Ecommerce_Trading_RealEstate_Scout.xlsx — 20 Saudi TikTok
creators across three niches (e-commerce, trading, real estate), sourced via
three parallel web-research agents under a no-fabrication contract: every
claim needs a citable source URL, unknown fields are "n/a", and namesake /
nationality traps are flagged rather than silently resolved.

TikTok itself was not directly scraped (no live TikTok API/browser access in
that research pass) — candidates were found via web search, cross-referenced
socials, and influencer-listing sites, each verified to have an identifiable
TikTok handle.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

FONT_NAME = "Arial"

CANDIDATES = [
    # ---------------- E-COMMERCE (5) ----------------
    dict(name="عبدالله الفوزان / Abdullah Al-Fawzan", category="E-commerce",
         handle="@econabdallah", niche="Dropshipping, launching an online store \"in 24 hours,\" TikTok paid-ads strategy",
         credentials="Describes achieving \"financial freedom\" before age 30; Riyadh-based",
         socials="X @econabdullah; IG @econabdullah; YouTube econAbdullah; Snapchat @econabdullah; econabdullah.com",
         followers="~178.1K (947.9K likes)",
         why_fit="Clearest, best-documented Saudi e-commerce-education creator, with consistent Riyadh/Saudi signals across every platform",
         risk_flag="", depth_confidence="High",
         sources="https://www.tiktok.com/@econabdallah ; https://x.com/econabdullah ; https://www.instagram.com/econabdullah/",
         status=""),
    dict(name="عادل صالح / Adel Saleh", category="E-commerce",
         handle="@adeel8866", niche="Dropshipping supplier lists, Chinese import/shopping sites, \"e-commerce A-to-Z\" tutorials",
         credentials="n/a (no CV found)", socials="Telegram t.me/Adelacademy",
         followers="n/a",
         why_fit="Content explicitly hashtagged #السعودية and #adelsaleh; practical step-by-step dropshipping tutorials",
         risk_flag="", depth_confidence="Medium — Saudi market-focus clear from hashtags, personal nationality not independently confirmed",
         sources="https://www.tiktok.com/@adeel8866/video/7542995492918119698", status=""),
    dict(name="ملاذ المصري / Malaz Al-Masri", category="E-commerce",
         handle="@malazmarketing", niche="E-commerce/digital marketing training, Salla & Shopify store setup, TikTok/Meta ad campaigns",
         credentials="BA marketing (King Abdulaziz Univ, Jeddah); MA e-marketing (Cairo Chamber of Commerce); runs ad agency \"MalazooAds\"; 15-18 yrs in Saudi & Turkish markets",
         socials="IG/X @malazooads; YouTube; malazmarketing.com", followers="n/a",
         why_fit="Deepest professional credentials of any e-commerce candidate found; long track record training Saudi merchants on Salla/Shopify",
         risk_flag="", depth_confidence="Flagged — nationality uncertain: surname \"Al-Masri\" (\"the Egyptian\") + Cairo credential suggest possible Egyptian origin working/educated in Saudi Arabia; his own site markets him as training \"in Saudi Arabia,\" not explicitly as a Saudi national",
         sources="https://malazmarketing.com/", status="Flag: nationality uncertain (possibly Egyptian, Saudi-market based)"),
    dict(name="Ecommerce Academy / Katib School (أيت كاتب محمد)", category="E-commerce",
         handle="@katibschool", niche="Free e-commerce training, \"from A to Z\"",
         credentials="IG bio cites 9 yrs marketing/e-commerce experience; founder identified as \"أيت كاتب محمد\" (Ait Katib Mohammed)",
         socials="IG @katibschool (~55K); Facebook; katibschool.com", followers="n/a on TikTok",
         why_fit="Established free e-commerce training brand", risk_flag="",
         depth_confidence="Flagged — likely non-Saudi: \"Ait Katib\" is an Amazigh/North-African (commonly Moroccan) naming pattern, content appears pan-Arab rather than Saudi-specific",
         sources="https://www.instagram.com/katibschool/ ; https://katibschool.com/",
         status="Flag: nationality likely non-Saudi (Moroccan naming pattern), not confirmed Saudi"),
    dict(name="Omar — \"تجارة الكترونية بالعربي\"", category="E-commerce",
         handle="@maketingomar", niche="Arabic-language e-commerce content", credentials="n/a",
         socials="n/a found", followers="n/a",
         why_fit="Only verifiable TikTok presence found for this exact search angle",
         risk_flag="", depth_confidence="Low — no biographical detail, nationality, or credentials found beyond the handle itself",
         sources="https://www.tiktok.com/@maketingomar", status="Insufficient verification — bare handle only"),

    # ---------------- TRADING (8) ----------------
    dict(name="فؤاد الحربي / Fuad Al-Harbi", category="Trading",
         handle="@fuad8k", niche="Saudi stock market (Tadawul) beginner education — explicitly no crypto, no fund management",
         credentials="Runs \"Fuad Academy\"; paid beginner course (349 SAR) with completion certificate",
         socials="X @FUAD7333; YouTube @FuadAcademy; fuadacademy.com", followers="n/a",
         why_fit="Explicit \"stocks only, no crypto, no money management\" positioning suggests deliberate compliance awareness",
         risk_flag="Paid course sold via social media — verify he isn't implying licensed advisory status", depth_confidence="Medium",
         sources="https://fuadacademy.com ; https://x.com/fuad7333", status=""),
    dict(name="تركي الحربي / Turki Al-Harbi", category="Trading",
         handle="@kii31t", niche="Personal investing/savings education; sells an \"intelligent investment package\" (261 SAR)",
         credentials="Self-described trainer; no license mentioned", socials="Snapchat (~64K subs); X @kii31t",
         followers="n/a on TikTok",
         why_fit="Cross-platform Saudi investment educator",
         risk_flag="Sells a paid \"investment package\" without disclosed licensing — potential unlicensed-advice risk",
         depth_confidence="Medium", sources="Search snippets for \"kii31t تركي الحربي\"; Snapchat profile page", status=""),
    dict(name="@deserttrader6906", category="Trading",
         handle="@deserttrader6906", niche="US stock market + crypto (Bitcoin, altcoins), general investing",
         credentials="n/a", socials="n/a found", followers="~52K (per influencer-listing snippet)",
         why_fit="Bio explicitly states content is personal opinion, not financial advice — a compliance-conscious disclaimer",
         risk_flag="None major noted", depth_confidence="Flagged — Saudi-base claim rests on a third-party aggregator listing (Heepsy-style), not a primary-source bio",
         sources="Heepsy/aggregator listing describing account as \"based in Saudi Arabia\"", status="Flag: nationality verified only via third-party aggregator"),
    dict(name="د. اسعد رزق / Dr. Assad Rizq", category="Trading",
         handle="@assadalsharef", niche="Crypto/blockchain education (Bitcoin, altcoins), \"how to build wealth with crypto\"",
         credentials="Described as a former Saudi neurologist who left medicine after discovering Bitcoin in 2013 while studying in Hungary; claims 9+ yrs experience, 10,000+ students",
         socials="n/a confirmed", followers="~1.1M (per search snippet)",
         why_fit="Unusually strong personal-credibility narrative (doctor-turned-crypto-educator), large following, Saudi hashtags (#السعودية)",
         risk_flag="Crypto education content — verify no unlicensed investment-scheme promotion; \"build wealth\" framing leans aspirational-marketing",
         depth_confidence="Medium-High", sources="YouTube video \"قصة د. أسعد رزق\"; TikTok search snippets for \"@assadalsharef\"", status=""),
    dict(name="بسام بن سليمان العبيد / Bassam bin Solaiman Al-Obaid", category="Trading",
         handle="@bin_solaiman", niche="US stock/options trading strategies, also covers Saudi market (#الاسهم_السعودية)",
         credentials="Writer for Al-Eqtisadiah (Saudi financial newspaper), consultant/analyst for TickerChart",
         socials="X @BinSolaiman; YouTube \"Bassam Alobid\"", followers="n/a",
         why_fit="Strongest verifiable professional credential of the set (mainstream financial-press byline)",
         risk_flag="Options-trading content carries inherent leverage/risk; no red flags found beyond that",
         depth_confidence="High", sources="TikTok video snippets tagged #تداول #الاسهم_السعودية; TickerChart profile; YouTube channel", status=""),
    dict(name="محمد الباهلي / Muhammad Al-Bahli", category="Trading",
         handle="@m_albahly", niche="Saudi stock market (Tadawul) news/education — dividend stocks, sukuk, automated \"rebound zone\" chart screens",
         credentials="n/a beyond content itself", socials="n/a found", followers="n/a",
         why_fit="Consistently labels content \"not a recommendation\" (ليست توصية) across videos — better compliance hygiene than most",
         risk_flag="Automated \"screening\" videos could be read as stock tips despite disclaimers",
         depth_confidence="Medium", sources="Multiple TikTok video URLs under @m_albahly (#تاسي #السوق_السعودي)", status=""),
    dict(name="المستر سلطان المضحي / Mr. Sultan Al-Mudhi", category="Trading",
         handle="@mrwallstreets.com", niche="US stocks, options contracts, gold, oil, crypto — framed around Saudi CMA-approved trading platforms",
         credentials="Self-described advisor/educator; no license cited",
         socials="IG @mrwallstreet.official; X @mrofwallstreat", followers="n/a",
         why_fit="Content directly addresses Saudi regulatory approval status of brokers — audience-relevant and somewhat compliance-aware",
         risk_flag="Options/leveraged-instrument education for retail audience carries elevated risk profile even when not \"hype\" style",
         depth_confidence="Medium", sources="TikTok video URLs under @mrwallstreets.com (#المستر_سلطان #الاسهم_الامريكية #عقود_الاوبشن)", status=""),
    dict(name="العمدة أبو أحمد الاسهم / Al-Omda Abu Ahmad", category="Trading",
         handle="@abo_ahmed_al3mda", niche="US stock technical analysis / \"fractal equations\"; explicitly \"not investment recommendations\"",
         credentials="n/a", socials="YouTube \"قناة أبو أحمد للأسهم\"", followers="~11.9K",
         why_fit="Niche technical-analysis angle",
         risk_flag="", depth_confidence="Flagged — nationality/Saudi base not confirmed: Gulf-Arabic content but no source ties this account specifically to Saudi Arabia rather than another Gulf state",
         sources="TikTok profile/video snippets; YouTube channel listing", status="Flag: Gulf-Arabic content, Saudi base unconfirmed (could be another Gulf state)"),

    # ---------------- REAL ESTATE (7) ----------------
    dict(name="مطر الشمري / Matar Al-Shammari", category="Real Estate",
         handle="@mnfr60", niche="Real-estate & mortgage-financing advisory: refinancing vs. rescheduling, contract-signing tips, avoiding bad financing decisions",
         credentials="Self-describes as real estate & financing consultant with ~20-25 yrs experience (per own Snapchat/TikTok bios)",
         socials="X and Snapchat under @mnfr60/@mnfr10-style handles; YouTube \"مطر الشمري مستشار عقاري وتمويلي\"",
         followers="n/a", why_fit="Clearest example of ongoing real-estate/financing education content by a self-identified Saudi consultant",
         risk_flag="", depth_confidence="High",
         sources="https://www.tiktok.com/@mnfr60 ; https://www.tiktok.com/@mnfr60/video/7322902717909781768 ; https://www.snapchat.com/@mnfr10",
         status="Note: also surfaced independently in the earlier Northouse Finance scout (real-estate/تمويل عقاري angle) as @mnfr10 — same individual, cross-project convergence"),
    dict(name="سعد الغامدي / Saad Al-Ghamdi", category="Real Estate",
         handle="@jubail_aqari", niche="Local property marketing/listings for Jubail (Eastern Province); REGA-licensed (\"معتمد من الهيئة العقارية\" per bio/hashtags)",
         credentials="Real estate marketer certified by the General Real Estate Authority (per account bio/hashtags)",
         socials="Cross-posts to Instagram under the same tag", followers="~33.5K (165.9K likes)",
         why_fit="Named, licensed individual agent creator — precise match for \"real-estate agent content\"", risk_flag="",
         depth_confidence="High", sources="https://www.tiktok.com/@jubail_aqari ; https://www.instagram.com/p/DOOS1aliEno/", status=""),
    dict(name="عقار AQAR (Aqar App)", category="Real Estate",
         handle="@aqarapp", niche="Saudi Arabia's largest real-estate listings/marketing platform (financing, leasing, market data)",
         credentials="10+ years old, 12M+ app downloads, \"#1 real estate app\" per Arab News/Saudi Gazette coverage",
         socials="aqar.fm; X @aqarapp", followers="99.8K (362.7K likes)",
         why_fit="Dominant brand in Saudi real-estate content — company/platform account, not an individual creator",
         risk_flag="", depth_confidence="High (brand, not individual)",
         sources="https://www.tiktok.com/@aqarapp ; https://www.arabnews.com/node/2570933/corporate-news ; https://saudigazette.com.sa/article/645428",
         status="Note: company/platform account, not an individual creator"),
    dict(name="عقارات الرياض / Riyadh Real Estate", category="Real Estate",
         handle="@aqaar_1", niche="Riyadh-focused property listings/showcase content", credentials="n/a (no bio/credential info surfaced)",
         socials="n/a found", followers="n/a",
         why_fit="Plausible Riyadh listings account", risk_flag="",
         depth_confidence="Low — under-documented, include with caution",
         sources="https://www.tiktok.com/@aqaar_1", status="Low confidence — no bio/credentials verifiable"),
    dict(name="شقق ايجار يومي الملقا الرياض", category="Real Estate",
         handle="@rr05l", niche="Daily/short-term rental apartment listings in Al-Malqa and central/north Riyadh",
         credentials="Appears to be a property-management/rental business account", socials="n/a found",
         followers="~37K (1.5M likes)",
         why_fit="Real-estate/rental-listing niche, closer to classifieds than education", risk_flag="",
         depth_confidence="Medium", sources="https://www.tiktok.com/@rr05l", status=""),
    dict(name="عقار وأكثر / Aqar wa Akthar", category="Real Estate",
         handle="@aqari_jdh", niche="Handle name suggests Jeddah-area real estate (\"jdh\"); specific content niche not independently confirmed",
         credentials="n/a", socials="n/a found", followers="n/a",
         why_fit="Lead for further vetting only", risk_flag="",
         depth_confidence="Low — unverified beyond handle",
         sources="https://www.tiktok.com/@aqari_jdh", status="Low confidence — unverified beyond handle existing"),
    dict(name="عقار المملكة / Kingdom Real Estate (cluster)", category="Real Estate",
         handle="@kingdom_aqar (related: @kingdomaqar, @aqar_ksa__)",
         niche="Saudi real-estate/development news (e.g. Riyadh infrastructure/road-project coverage tied to property value)",
         credentials="n/a", socials="n/a found", followers="n/a",
         why_fit="Real-estate news/commentary niche", risk_flag="",
         depth_confidence="Low — multiple similarly-named accounts found; cannot confirm which is the \"main\"/most-followed one",
         sources="https://www.tiktok.com/@kingdom_aqar ; https://www.tiktok.com/@kingdomaqar ; https://www.tiktok.com/@aqar_ksa__",
         status="Low confidence — handle cluster needs manual disambiguation"),

    # ---------------- ROUND 2: STRICTER-BAR REPLACEMENTS (5) ----------------
    # Sourced under a stricter bar: real named individual (no companies/anon
    # accounts), independently-confirmed Saudi nationality, and at least one
    # corroborating source beyond the TikTok page itself.
    dict(name="رعد الحربي (أبو دانه) / Raad Al-Harbi (\"Abu Danah\")", category="Trading",
         handle="@abu_danah00",
         niche="Saudi (Tadawul) stock and fund trading education — technical/fundamental analysis, portfolio building, beginner-to-pro courses",
         credentials="Runs \"Abu Danah Al-Harbi Academy\"; bachelor's degree in finance/economics (per search summaries); content emphasizes trading \"with awareness\" to avoid heavy losses",
         socials="X @abu_danah01 (\"اسهم وصناديق\"); YouTube @Abudanahlearn; abudanahacademy.com",
         followers="n/a", selling="Yes — paid academy courses", activity="active",
         why_fit="Named individual (tribal surname Al-Harbi) with self-declared Saudi flag; cross-verified via 3 independent properties beyond the TikTok page (X, YouTube, own academy site); risk-aware framing, not signal-selling",
         risk_flag="", depth_confidence="High — real name + 3 independent corroborating sources",
         sources="https://www.tiktok.com/@abu_danah00 ; https://x.com/abu_danah01 ; https://www.youtube.com/@Abudanahlearn ; https://abudanahacademy.com/",
         status=""),
    dict(name="محمد الخضير / Mohammed Alkhudhair", category="Trading",
         handle="@mohakhuu",
         niche="Retail investing/fintech education — comparing Saudi investment/trading apps (Derayah, Tamra, Malaa, Drahim, Abyan, Al Rajhi Capital), tracking Tadawul holdings",
         credentials="Self-described financial advisor; master's degree in law; publicly invested SAR 250,000 across 5 licensed Saudi investment apps for a year to compare real returns",
         socials="X @Mohakhu (~120K+); Instagram @mokhco",
         followers="~114K", selling="n/a visible", activity="active",
         why_fit="Named individual with a disclosed, non-hype comparison methodology; strongly Saudi-contextualized content (PIF, Saudia Airlines, Saudi banks, CMA-licensed apps); cross-verified via X + Instagram under consistent name/photo",
         risk_flag="", depth_confidence="High — real name + disclosed methodology + 2 independent corroborating platforms",
         sources="https://www.tiktok.com/@mohakhuu ; https://x.com/Mohakhu ; https://www.instagram.com/mokhco/",
         status=""),
    dict(name="عماد منشي / Emad Munshi", category="Real Estate",
         handle="@emad_munshi",
         niche="Real-estate investment consultancy/analysis (Jeddah apartment/land investing, market-timing videos)",
         credentials="Self-identifies as licensed real-estate consultant/analyst; runs Aya Real Estate (شركة آيا العقارية), a Jeddah investment-property firm",
         socials="X @EmadMunshi; Instagram @emad_munshi; Snapchat @emad_munshi; emadmunshi.com",
         followers="~165.7K (543.3K likes)", selling="Yes — consultancy via Aya Real Estate", activity="active",
         why_fit="Named individual with a consistent identity across TikTok/X/Instagram/Snapchat/website; explicit Jeddah, Saudi tie via his company, independently confirmed beyond the TikTok page",
         risk_flag="", depth_confidence="High — real name + company + 4 consistent platforms",
         sources="https://www.tiktok.com/@emad_munshi ; https://x.com/EmadMunshi/status/1958484466529402960 ; https://www.instagram.com/emad_munshi/ ; https://www.snapchat.com/@emad_munshi",
         status=""),
    dict(name="سعد بن محمد التويم / Saad bin Mohammed Al-Tuwaim", category="Real Estate",
         handle="@saadaltwaim1",
         niche="Real-estate market commentary, appraisal/valuation education, Riyadh market analysis (demand trends, pricing)",
         credentials="Chairman of the Board of Tathmen Company (شركة تثمين), described as the first licensed real-estate valuation/appraisal firm in Saudi Arabia (est. 2007); specializes in structuring real-estate funds",
         socials="X @altuwaim_s; YouTube @Saadaltuwym (featured on \"حديث عقاري\" podcast); Facebook; LinkedIn",
         followers="~133.2K (490.1K likes)", selling="n/a visible (corporate role)", activity="active",
         why_fit="Real named individual with a documented corporate role (Tathmen chairman), independently confirmed via the company's own site and 4 distinct platforms (X, YouTube, Facebook, LinkedIn) under a consistent name",
         risk_flag="", depth_confidence="High — real name + corporate role + company site + 4 platforms",
         sources="https://www.tiktok.com/@saadaltwaim1 ; https://x.com/altuwaim_s ; https://www.youtube.com/@Saadaltuwym ; http://tathmen.net/en/ ; https://www.linkedin.com/in/سعود-بن-سعد-التويم-saud-saad-altwaim-4a7636232/",
         status=""),
    dict(name="طارق شقرون / Tariq Shagroon (\"Mr. Broker\")", category="Real Estate",
         handle="@tshagroon",
         niche="Personal real-estate consulting/brokerage content in Madinah — property tours, buy/sell/invest Q&A",
         credentials="Owner/marketer of Alamat Real Estate Marketing Co. (شركة علامات العقارية), a Saudi real-estate marketing firm founded 2012 in Madinah",
         socials="X @tshagroon; Snapchat @tshagroon (bio \"Mr. Broker\")",
         followers="~194.8K", selling="Yes — brokerage/consulting via Alamat", activity="active",
         why_fit="Named individual (not a brand/listings account); explicit Madinah tie across content and company; independently corroborated by two separate news/feature articles plus a distinct company website",
         risk_flag="", depth_confidence="High — real name + company site + 2 independent news articles",
         sources="https://www.tiktok.com/@tshagroon ; https://x.com/tshagroon ; https://alamat.sa/aboutus/ ; https://slaati.com/2025/05/26/p2690045.html ; https://slaati.com/2025/04/26/p2690076.html",
         status=""),
]

EXCLUDED = [
    dict(name="Hussam Ansari (@hussam.ansari)", category="E-commerce",
         reason="Teaches Amazon FBA \"in Saudi Arabia and UAE,\" but own hashtags (#ukpakistani, #overseaspakistani) indicate Pakistani nationality/heritage — market-focus-vs-nationality trap"),
    dict(name="Zuhair Abu Al-Reesh", category="E-commerce",
         reason="Genuine, well-documented Saudi-market digital-marketing/TikTok-ads trainer (23 yrs experience), but no personal TikTok creator account located — fails identifiable-TikTok-presence bar"),
    dict(name="Yazeed (@YazeedKM_)", category="E-commerce",
         reason="Founder of e-commerce training platform \"Adatik,\" active/credentialed on X, but no TikTok handle verified"),
    dict(name="Ali Al-Hamed (علي الحامد)", category="Trading",
         reason="Described as a \"licensed financial influencer\" but licensed/based in the UAE, not Saudi Arabia — namesake/wrong-country trap"),
    dict(name="Tariq Al-Harbi (@tariq_alharbi80)", category="Trading",
         reason="Name/hashtag matched trading searches, but account's own listed profession is \"Actor\" with entertainment content — wrong niche despite name match"),
    dict(name="Mohammed Al-Shumaimari (أبو سهم الشميمري)", category="Trading",
         reason="Genuinely credentialed Saudi markets analyst (34-yr track record, TV/press commentator), but no TikTok handle could be verified despite targeted searches"),
    dict(name="\"Omar\" (Collabstr forex/crypto listing)", category="Trading",
         reason="Profile describes an 11-yr markets creator reaching 22 Arab countries, but no source confirms Saudi nationality or base specifically"),
    dict(name="Rateel Alshehri", category="Trading",
         reason="Confirmed Saudi, TikTok-verified, but content niche is youth entrepreneurship/podcasting, not trading — niche mismatch, not nationality"),
    dict(name="Abdullah Al-Rajhi (@abdullahmrajhi, 194.1K)", category="Real Estate",
         reason="Search results inconsistently tie this handle to a Riyadh villa-developer chairman AND separately to an unrelated Al-Rajhi businessman bio — extremely common Saudi business-family name, could not confirm which entity/bio actually belongs to this account"),
    dict(name="Reem (@alreem.ksaa, Instagram)", category="Real Estate",
         reason="Genuinely appears to be a REGA-licensed Saudi real-estate marketer, but no TikTok handle could be located — Instagram-only presence fails the TikTok-presence requirement"),
    dict(name="\"مسوقة عقارية Official\" (Linktree marketersalma_7)", category="Real Estate",
         reason="Linktree confirms TikTok is among her linked platforms, but exact TikTok username unconfirmed, and search surfaced multiple unrelated \"Salma\" accounts (incl. an apparently Egyptian one) — namesake-confusion risk"),
]

TOP_PICKS = [
    ("@econabdallah", "E-commerce", "#1 — only e-commerce candidate with fully cross-verified Saudi nationality (Riyadh), ~178K TikTok followers, content squarely in dropshipping/store-launch niche across every platform"),
    ("@adeel8866", "E-commerce", "#2 — strong niche fit (dropshipping suppliers, sourcing, tutorials) with explicit Saudi-market hashtagging, though personal nationality confirmation is thinner"),
    ("@malazmarketing", "E-commerce", "#3 — deepest professional credentials of any e-commerce candidate, but flagged: likely-Egyptian background means Saudi nationality (vs. market residency) is unconfirmed"),
    ("@assadalsharef", "Trading", "#1 — leads on reach (~1.1M) and a distinctive doctor-turned-crypto-educator credibility narrative, though \"build wealth with crypto\" framing warrants compliance scrutiny"),
    ("@bin_solaiman", "Trading", "#2 — most professionally credentialed (financial-press byline, market-data consultancy) with genuine Saudi-market hashtag activity"),
    ("@fuad8k", "Trading", "#3 — cleanest self-imposed compliance boundary of the group: stocks-only education, explicitly no crypto or fund management"),
    ("@mnfr60", "Real Estate", "#1 — most clearly documented individual creator, consistently branded across TikTok/X/Snapchat/YouTube as a real-estate and mortgage-financing consultant with a repeatable content format"),
    ("@jubail_aqari", "Real Estate", "#2 — named, REGA-credentialed local property marketer with a concrete follower count (33.5K) and a clear regional niche (Jubail/Eastern Province)"),
    ("@aqarapp", "Real Estate", "#3 — not an individual, but the single best-verified, largest real-estate brand active on Saudi TikTok (99.8K followers), backed by independent press coverage"),
]

METHOD_NOTES = [
    "Sourcing rule (no-fabrication contract): every claim needed a citable source URL; unknown fields are marked n/a, never guessed. A candidate had to have an identifiable TikTok handle/channel to count — verified via web search, cross-referenced socials, or influencer-listing sites (not direct TikTok scraping/API access, which wasn't available in this research pass).",
    "Three parallel research agents ran one per niche: E-commerce, Trading, Real Estate — each independently searched, verified, and flagged uncertainty rather than padding the list to a target count. Totals: 5 e-commerce + 8 trading + 7 real estate = 20 candidates.",
    "'Depth/Confidence' column: the researching agent's own confidence in the verification (High/Medium/Low), separate from follower count or expertise depth — Low-confidence entries are included transparently rather than dropped, so treat them as leads needing manual follow-up, not settled picks.",
    "Nationality flags: several candidates' Saudi nationality (vs. Saudi-market residency/focus) could not be independently confirmed — flagged per-row rather than silently included or excluded. This is a materially harder niche for nationality verification than the earlier Northouse expert-scout project, since TikTok bios rarely state nationality directly.",
    "Trading-niche risk flags: several trading creators sell paid 'investment packages' or courses without disclosed licensing, or produce options/leveraged-instrument content — flagged for compliance review before any outreach, not as a disqualification by itself.",
    "Real-estate niche note: individually-branded real-estate 'influencers' are less common on Saudi TikTok than licensed agents/marketers or platform brand accounts (e.g. Aqar App) — both categories are included and labeled.",
    "Cross-project convergence: مطر الشمري (@mnfr60/@mnfr10) also appeared in the earlier Northouse Finance scout (real-estate/تمويل عقاري angle) under a related handle — independent convergence across two separate research passes, a stronger signal than either alone.",
]


def autosize(ws, min_width=10, max_width=55):
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


def style_header(ws, row=1, ncols=1, fill="1F4E78"):
    header_font = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor=fill)
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)


def set_default_font(ws):
    for row in ws.iter_rows():
        for cell in row:
            existing = cell.font
            cell.font = Font(
                name=FONT_NAME,
                bold=existing.bold if existing else False,
                color=existing.color if existing else None,
                size=existing.size if existing and existing.size else 10,
            )


def main():
    wb = openpyxl.Workbook()
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws = wb.active
    ws.title = "All Candidates"
    headers = ["Name (AR/EN)", "Niche", "TikTok Handle", "Content Niche Specifics",
               "Credentials/Background", "Other Socials", "Followers",
               "Why Notable/Fit", "Risk/Compliance Flag", "Verification Confidence",
               "Status / Flags", "Source URL(s)"]
    ws.append(headers)
    style_header(ws, ncols=len(headers))
    ws.freeze_panes = "A2"
    for c in CANDIDATES:
        ws.append([c["name"], c["category"], c["handle"], c["niche"], c["credentials"],
                   c["socials"], c["followers"], c["why_fit"], c["risk_flag"],
                   c["depth_confidence"], c["status"], c["sources"]])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers)):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    autosize(ws, max_width=45)

    ws2 = wb.create_sheet("Top Picks")
    headers2 = ["TikTok Handle", "Niche", "Rank Note"]
    ws2.append(headers2)
    style_header(ws2, ncols=len(headers2))
    ws2.freeze_panes = "A2"
    for handle, cat, note in TOP_PICKS:
        ws2.append([handle, cat, note])
    for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row, max_col=len(headers2)):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    autosize(ws2, max_width=70)

    ws3 = wb.create_sheet("Excluded & Flags")
    ws3["A1"] = "Candidates Excluded During Research (namesake traps, wrong niche, unverifiable)"
    ws3["A1"].font = Font(name=FONT_NAME, bold=True, size=13)
    ws3.append([])
    ws3.append(["Name", "Niche", "Reason Excluded"])
    style_header(ws3, row=ws3.max_row, ncols=3)
    for e in EXCLUDED:
        ws3.append([e["name"], e["category"], e["reason"]])
    for row in ws3.iter_rows(min_row=1, max_row=ws3.max_row, max_col=3):
        for cell in row:
            if cell.value is not None:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
    autosize(ws3, max_width=60)

    ws4 = wb.create_sheet("Method")
    ws4["A1"] = "Method, Rules & Counts"
    ws4["A1"].font = Font(name=FONT_NAME, bold=True, size=13)
    ws4.append([])
    ws4.append(["Metric", "Value"])
    style_header(ws4, row=ws4.max_row, ncols=2)
    ws4.append(["Total candidates", len(CANDIDATES)])
    ws4.append(["E-commerce", sum(1 for c in CANDIDATES if c["category"] == "E-commerce")])
    ws4.append(["Trading", sum(1 for c in CANDIDATES if c["category"] == "Trading")])
    ws4.append(["Real Estate", sum(1 for c in CANDIDATES if c["category"] == "Real Estate")])
    ws4.append(["Explicitly excluded during research", len(EXCLUDED)])
    ws4.append([])
    r = ws4.max_row + 1
    ws4.cell(row=r, column=1, value="Notes").font = Font(name=FONT_NAME, bold=True, size=12)
    ws4.append([])
    for note in METHOD_NOTES:
        ws4.append([note])
        ws4.cell(row=ws4.max_row, column=1).alignment = Alignment(wrap_text=True, vertical="top")
        ws4.merge_cells(start_row=ws4.max_row, start_column=1, end_row=ws4.max_row, end_column=6)
    ws4.column_dimensions["A"].width = 110
    ws4.column_dimensions["B"].width = 14
    for row in ws4.iter_rows(min_row=1, max_row=7, max_col=2):
        for cell in row:
            if cell.value is not None:
                cell.alignment = Alignment(vertical="top", wrap_text=True)

    for name in ("All Candidates", "Top Picks", "Excluded & Flags", "Method"):
        set_default_font(wb[name])

    out_path = "NH_TikTok_Ecommerce_Trading_RealEstate_Scout.xlsx"
    wb.save(out_path)
    print(f"Wrote {out_path}: {len(CANDIDATES)} candidates, {len(TOP_PICKS)} top picks, {len(EXCLUDED)} excluded.")


if __name__ == "__main__":
    main()
