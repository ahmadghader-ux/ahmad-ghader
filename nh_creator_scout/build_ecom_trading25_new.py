#!/usr/bin/env python3
"""
Builds NH_TikTok_EcommerceTrading_New25.xlsx — 25 FULLY NEW Saudi(-market)
creators, e-commerce and trading only, none of which overlap with any
candidate shown in the two earlier files in this project
(NH_TikTok_Ecommerce_Trading_RealEstate_Scout.xlsx's 25, and
NH_TikTok_EcommerceTrading_Only25.xlsx's 25 — 50 previously-shown handles
total, all excluded from this round by construction).

Sourced via 4 parallel research agents (2 e-commerce angles, 2 trading
angles) plus ~14 follow-up manual searches to close gaps left by the
agents' search-budget limits and to dedupe overlapping finds between the
two e-commerce agents (both independently surfaced the same 5 people).

Same no-fabrication contract as the earlier files: every candidate is a
real, verified account with citable sources; unconfirmed nationality is
flagged rather than assumed or silently excluded; namesake/wrong-country
traps are called out explicitly. Several entries here are lower-confidence
than the earlier files' best entries — the Saudi-national, e-commerce/
trading, social-media-creator pool is finite, and a second from-scratch
pass necessarily digs into thinner leads. This is noted per-row, not
glossed over.
"""
import re
from collections import Counter

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

FONT_NAME = "Arial"
NAVY = "1F4E78"
BAR_COLORS = {"E-commerce": "2E7D32", "Trading": "C62828"}

CANDIDATES = [
    # ================= E-COMMERCE (11) =================
    dict(name="سهل مهدي / Sahel Mahdi", category="E-commerce",
         handle="@sahelmahdi", platform="tiktok",
         niche="E-commerce/digital-marketing educator — Instagram e-commerce strategy, dropshipping, integrated marketing courses",
         credentials="Master's in marketing (reportedly pursuing a doctorate in digital marketing); founder of \"Sahel Rich,\" \"Mukhtalif Store,\" and \"Mastery Academy\" (hosts his own e-commerce/dropshipping courses)",
         socials="Instagram @sahelmahdi; YouTube; Snapchat; X @Sahelmahdi; Telegram; sahelmahdi.com",
         followers="~277.7K (TikTok, per search-engine summary — not independently re-confirmed, site unreachable this session)",
         why_fit="Directly teaches e-commerce/dropshipping/Instagram-commerce as core content and monetizes via his own academy — largest, best cross-corroborated candidate found this round",
         risk_flag="Nationality asserted across multiple secondary sources/snippets but no single authoritative bio page was directly loadable to confirm",
         depth_confidence="Medium-High — identity, businesses, and course catalog corroborated across many independent search hits",
         sources="https://sahelmahdi.com/ ; https://www.instagram.com/sahelmahdi/?hl=en ; https://www.tiktok.com/@sahelmahdi ; https://www.emasteryacademy.com/courses/course?coursetitle=Instagram_E-Commerce_Strategy&id=4 ; https://x.com/Sahelmahdi",
         status="Verify nationality and follower counts directly on-platform before outreach"),
    dict(name="د. عبدالكريم عنف / Dr. Abdulkarim Anaf", category="E-commerce",
         handle="@anafabdulkareem", platform="tiktok",
         niche="Amazon FBA / China-import / international e-commerce education, incl. content titled \"Amazon FBA Saudi Arabia — Easy Selling Guide\"",
         credentials="PhD in international e-commerce (Nankai University, China); Master's/MBA (Shandong University); 13+ yrs in the China trade/e-commerce market; executive manager of the Arab Businessmen's Forum in China",
         socials="YouTube \"Dr. Anaf Abdulkarem\"; anafabdulkarem.com",
         followers="n/a",
         why_fit="Content directly and explicitly targets Saudi Amazon FBA sellers — squarely matches the Amazon/Noon-selling niche",
         risk_flag="Nationality NOT confirmed Saudi — his entire career/education footprint is China-based, serving Arab entrepreneurs broadly; no source confirms Saudi nationality specifically",
         depth_confidence="Low-Medium — real, active, clear niche fit; nationality genuinely uncertain",
         sources="https://www.tiktok.com/@anafabdulkareem ; https://www.youtube.com/@ANAFABDULKAREM ; https://anafabdulkarem.com/امازون-fba-السعودية-دليل-البيع-السهل/",
         status="Nationality uncertain — high priority to verify before outreach"),
    dict(name="LLC4ARAB", category="E-commerce",
         handle="@llc4arab", platform="tiktok",
         niche="Helps Arab (Saudi/UAE-targeted) e-commerce/dropshipping sellers form US LLCs and open US bank accounts to sell internationally — an ecosystem player frequently cited by dropshippers",
         credentials="Cited as having helped 200+ businesses form US companies; active multi-platform content operation",
         socials="Instagram @llc4arab (~83K); YouTube @llc4arab929; linktr.ee/llc4arab",
         followers="TikTok ~25.7K (58K likes)",
         why_fit="Core audience is Gulf/Saudi e-commerce sellers; content is educational (company setup, e-commerce compliance)",
         risk_flag="This is a brand/company account, not a clearly identified individual creator — no founder name surfaced; Facebook listing shows a Delaware, US business address (likely a registered-agent address, not personal residency)",
         depth_confidence="Low — real and active, but individual identity/nationality behind the account unconfirmed",
         sources="https://www.tiktok.com/@llc4arab ; https://www.instagram.com/llc4arab/ ; https://www.facebook.com/llc4arab/",
         status="Brand/company account — founder identity unconfirmed"),
    dict(name="ابراهيم البحراوي / Ibrahim Al-Bahrawy", category="E-commerce",
         handle="bhrawy.com", platform="website",
         niche="E-commerce/digital-marketing consultant and trainer — publishes guides on opening TikTok Shop stores and Salla/e-commerce marketing in Saudi Arabia",
         credentials="Self-described digital-marketing consultant/trainer with 20+ years' experience; positions himself as \"best e-commerce marketing expert in Saudi Arabia 2026\"",
         socials="Possible Instagram @ibrahim_bhrawy (match unconfirmed)",
         followers="n/a",
         why_fit="Content is squarely e-commerce/TikTok-store education targeted at the Saudi market",
         risk_flag="Likely non-Saudi — his own site describes him as based \"in both Egypt and Saudi Arabia\"; \"Bahrawy\"/\"El-Bahrawy\" is a common Egyptian surname; site markets him as serving the Saudi market, not stating Saudi nationality",
         depth_confidence="Low-Medium — real, findable website/content; personal social handle and nationality unconfirmed",
         sources="https://bhrawy.com/ ; https://bhrawy.com/المدونة/كيفية-إعداد-متجر-التيك-توك-في-السعودية/",
         status="Nationality uncertain / likely non-Saudi (Egypt-linked) — not confirmed Saudi; no confirmed personal social handle"),
    dict(name="علاء الحسن / Alaa Al-Hassan", category="E-commerce",
         handle="alaaalhasan.com", platform="website",
         niche="Amazon FBA education — full beginner-to-advanced courses, blog content on FBA selling",
         credentials="Runs a dedicated Amazon FBA course platform with student reviews cited",
         socials="n/a confirmed beyond website",
         followers="n/a",
         why_fit="Squarely in the Amazon FBA education niche",
         risk_flag="Nationality entirely unconfirmed; one of his own video titles references \"Amazon FBA Germany,\" suggesting his content may target the German Amazon marketplace rather than Saudi/Gulf sellers specifically",
         depth_confidence="Low — real educator with clear e-commerce niche; nationality and market focus both unverified",
         sources="https://alaaalhasan.com/ ; https://alaaalhasan.com/category/amazon-fba/ ; https://alaaalhasan.com/amazon-fba-course-beginner-to-advanced/",
         status="Nationality uncertain; possible market-focus mismatch (Germany) — verify before outreach"),
    dict(name="أحمد مطر / Ahmed Matar", category="E-commerce",
         handle="ahmedmtr.com", platform="website",
         niche="E-commerce and digital-marketing consulting/education",
         credentials="Master's in e-commerce (Canada); Google/Meta ads certifications; 10-15+ yrs experience; claims work with major Saudi companies (Extra/الشركة المتحدة للإلكترونيات, Al-Jomaih Automotive)",
         socials="No verified personal TikTok/Instagram/X handle found — several same-name accounts exist (incl. an unrelated Iraqi poet) but none confirmed as this consultant",
         followers="n/a",
         why_fit="Professional background/claimed client list strongly suggests genuine Saudi-market e-commerce expertise",
         risk_flag="No verified social-media account — reads as an e-commerce consultant with a website rather than a confirmed social-media creator; include with caution",
         depth_confidence="Low — real, findable professional via website/bio, but fails the \"active identifiable social account\" bar until a handle is confirmed",
         sources="https://ahmedmtr.com/ ; https://ahmedmtr.com/عن-أحمد-مطر/",
         status="No verified social handle — weakest e-commerce candidate, verify before use"),
    dict(name="@ifast.sa", category="E-commerce",
         handle="@ifast.sa", platform="tiktok",
         niche="Beginner e-commerce education — TikTok business-account verification, online store setup, \"مشروعي الصغير\" (my small project) content",
         credentials="n/a — real name of the operator could not be identified",
         socials="n/a found",
         followers="n/a",
         why_fit="Content and hashtags (#تجارة_إلكترونية_للمبتدئين #متاجر_الكترونية) are squarely in scope; the \".sa\" handle plus content framing suggest a Saudi-market-focused account",
         risk_flag="Real name/identity of the creator unverified — may be a small agency/brand account rather than a named individual",
         depth_confidence="Low — account niche fit clear; personal identity and nationality unconfirmed",
         sources="https://www.tiktok.com/@ifast.sa/video/7290119503982120200",
         status="Identity unconfirmed — bare handle, verified only via search snippet"),
    dict(name="فهيد الفهيد / Fahad Al-Fahad", category="E-commerce",
         handle="@fh_afiliate0", platform="tiktok",
         niche="Affiliate-marketing and TikTok-ads education — affiliate marketing, TikTok Ads Manager, e-commerce fundamentals",
         credentials="Founder/owner of \"دليلك للتسويق\" (Daleelak Marketing) agency, based in Al-Qassim, Saudi Arabia; claims 5 yrs e-commerce experience, 3,000+ merchants helped",
         socials="YouTube @Fheed1; Snapchat @ecfheed",
         followers="Reported 1.7M+ combined across platforms (per search summary — not independently verified per-platform)",
         why_fit="Active TikTok content specifically on affiliate marketing, TikTok ad campaigns for e-commerce, and merchant coaching",
         risk_flag="Follower/merchant-count figures come from a secondary summary, not a primary bio directly viewed — treat numbers as approximate pending verification",
         depth_confidence="Medium-High — Al-Qassim location and agency name recur consistently across independent search hits",
         sources="https://www.tiktok.com/@fh_afiliate0 ; https://www.youtube.com/@Fheed1 ; https://www.snapchat.com/@ecfheed",
         status="Verify follower counts directly on-platform"),
    dict(name="بيان لنجكاوي / Bayan Lanjakawi", category="E-commerce",
         handle="31 Store (handle unconfirmed)", platform="unknown",
         niche="Lifestyle/fashion e-commerce — founder & executive director of \"31 Store,\" an e-commerce platform for lifestyle/fashion products",
         credentials="Founder & Executive Director of 31 Store, per a Saudi women's-lifestyle magazine feature",
         socials="Not directly verified this session",
         followers="n/a",
         why_fit="A named Saudi female founder of an e-commerce (fashion/lifestyle) brand, per an editorial feature — diversifies the candidate pool by gender and vertical",
         risk_flag="Rests on a single magazine citation rather than a directly-viewed social profile; personal TikTok/Instagram handle not confirmed before search budget ran out",
         depth_confidence="Low — single-source attribution, handle/follower data unverified",
         sources="https://www.lahamag.com/article/229986-سعوديات-رائدات-في-عالم-السوشيال-ميديا-موضة-وتسويق-وأعمال",
         status="Unverified handle — needs manual lookup before outreach"),
    dict(name="بزنس فاطمة / Fatemah Business", category="E-commerce",
         handle="@fatemahbusiness", platform="tiktok",
         niche="E-commerce/digital-marketing education — \"a complete team that helps you succeed in your e-commerce business and marketing it\" (per own bio)",
         credentials="n/a beyond own positioning; active cross-platform content operation",
         socials="Instagram @fatemahbusiness (~1.3K); Snapchat @fatemahbusiness",
         followers="n/a",
         why_fit="Active, cross-platform (TikTok/Instagram/Snapchat) e-commerce education account with consistent Arabic content",
         risk_flag="Own bio describes \"a complete team,\" suggesting this is a small agency/brand account rather than one named individual; nationality not explicitly confirmed as Saudi",
         depth_confidence="Low-Medium — real, active, cross-platform account; identity is a team, not a confirmed individual, and nationality unconfirmed",
         sources="https://www.instagram.com/fatemahbusiness/ ; https://www.snapchat.com/@fatemahbusiness ; https://www.tiktok.com/@fatemahbusiness",
         status="Team/brand account, not a confirmed individual; nationality unconfirmed"),
    dict(name="عالية العالي / Aliya Al-Ali (No Added Sugar)", category="E-commerce",
         handle="@noaddedsugar.me", platform="instagram",
         niche="Founder of a Saudi-designed wellness/lifestyle product brand (\"No Added Sugar\"), selling via Amazon, TikTok, and her own store from one control panel",
         credentials="Cited by name in Zid's own official Saudi e-commerce-platform success-stories page as a merchant who sells across Amazon/TikTok/own store",
         socials="Facebook (100090599246583); noaddedsugar.me",
         followers="~26K (Instagram)",
         why_fit="Strongest independent corroboration of any e-commerce candidate this round — named directly by Zid (a real, major Saudi e-commerce platform) as a merchant success story, confirming both Saudi-market basis and multi-channel e-commerce operation",
         risk_flag="Brand-account-primary (Instagram is the product brand, not a clearly personal creator account); personal social presence separate from the brand not confirmed",
         depth_confidence="Medium — strong third-party (platform-level) corroboration via Zid, but no personal (non-brand) social account confirmed",
         sources="https://zid.sa/ar/success-stories/ ; https://www.instagram.com/noaddedsugar.me/ ; https://noaddedsugar.me/products",
         status="Brand account, not a personal creator account — verify a personal handle before outreach"),

    # ================= TRADING (14) =================
    dict(name="أبو ريان / Abu Rayan", category="Trading",
         handle="@aaboryan", platform="snapchat",
         niche="Technical analysis/signals for the Saudi stock market (TASI), multi-platform presence",
         credentials="No CMA-license claim found; explicitly located in Abha, Saudi Arabia per search results",
         socials="TikTok (\"محلل الاسهم ابو ريان\"); YouTube playlist \"ابو ريان تحليل السوق السعودي للاسهم\"; Telegram channel",
         followers="n/a (not retrievable via search snippets)",
         why_fit="Real, findable, multi-platform Saudi-market-focused creator matching the Tadawul/TASI niche, explicitly located in Abha",
         risk_flag="Marketed as an \"easy and profitable\" method with a Telegram technical-analysis channel — borderline unlicensed signal-style promotion, no CMA license found; multiple similarly-named \"Abu Rayan\" accounts exist (namesake risk)",
         depth_confidence="Medium — city and Saudi-market focus corroborated across Snapchat/TikTok/YouTube/Telegram references; real legal name not found",
         sources="https://www.snapchat.com/@aaboryan ; https://www.tiktok.com/discover/محلل-الاسهم-ابو-ريان ; https://www.youtube.com/playlist?list=PLTctY8U_WjvWVNkZ1jUsv3cDl7m8HKsGZ ; https://x.com/29_shg",
         status="Verify follower count on-platform; namesake risk (multiple \"Abu Rayan\" accounts) noted"),
    dict(name="أبو إبراهيم / Abu Ibrahim (\"Wolfe Wave\")", category="Trading",
         handle="@wolfe_wave", platform="tiktok",
         niche="Technical analysis of Saudi (TASI) stocks, branded around \"Wolfe Wave\" chart patterns",
         credentials="n/a",
         socials="n/a confirmed beyond TikTok",
         followers="n/a",
         why_fit="Confirmed live TikTok account (indexed video URL); bio self-identifies as covering Saudi stocks; independently cited by a third-party X user as worth following for Saudi stock analysis",
         risk_flag="Possible — free live technical-analysis content with no disclosed license, consistent with the CMA's standing warnings about unlicensed social-media financial content, though no explicit paid-signal evidence found",
         depth_confidence="Medium — account existence/bio confirmed via indexed video, corroborated by a third-party tweet; nationality inferred from bio wording and market focus, not an explicit personal statement",
         sources="https://www.tiktok.com/@wolfe_wave/video/7297841818538872082 ; https://x.com/RahafStock/status/2038159848882209113",
         status="Real name not found; only \"Abu Ibrahim\" alias"),
    dict(name="مكهرب الأسهم (\"أبو عبدالله\") / Mkahrib Al-Ashom", category="Trading",
         handle="@mokhrepashom", platform="x",
         niche="Table-based technical analysis exclusively for the Saudi stock market since ~2006; explicitly frames output as \"views, not buy/sell recommendations\"",
         credentials="Long-tenured account (active since ~2006, per references)",
         socials="n/a confirmed",
         followers="n/a",
         why_fit="Confirmed active X account, Saudi-market-exclusive focus, frequently cited in \"best Saudi stock Twitter analysts\" round-ups",
         risk_flag="Explicitly disclaims giving recommendations, but the semi-automated technical-table format with entry conditions functions similarly to trade signals for a lay audience",
         depth_confidence="Low-Medium — account/niche confirmed via search snippets citing it as a well-known Saudi-market analyst; live profile not loadable to confirm follower count, real name, or current bio",
         sources="https://x.com/mokhrepashom",
         status="Nationality inferred, not confirmed; real name unknown"),
    dict(name="خبير أسهم / Khabeer Ashom (\"Stock Expert\")", category="Trading",
         handle="@khabeer999", platform="x",
         niche="Daily Saudi stock-market reads, market-opportunity alerts, technical/fundamental commentary",
         credentials="Provides a personal contact email in bio per search snippet; no license found",
         socials="n/a confirmed",
         followers="n/a",
         why_fit="Confirmed active X account with multiple indexed tweets on Saudi stock movements/dividends; described in search results as a \"Saudi Arabian stocks analyst account\"",
         risk_flag="Posts market \"alerts\"/opportunity calls to followers without stated licensing",
         depth_confidence="Low-Medium — Saudi description is a third-party search-engine characterization, not a primary self-statement directly verified; profile not loadable to confirm current status",
         sources="https://x.com/khabeer999 ; https://x.com/khabeer999/status/1332601280976588802",
         status="Nationality flagged as inferred/unconfirmed"),
    dict(name="قناة تاسي / TASI_y", category="Trading",
         handle="@tasi_y", platform="youtube",
         niche="Dedicated channel analyzing \"all Saudi stocks\" and the TASI general index",
         credentials="n/a", socials="X @tasi_y", followers="n/a",
         why_fit="Channel name and content are explicitly and exclusively about the Saudi Tadawul All Share Index",
         risk_flag="General technical/fundamental commentary channel; no explicit evidence of paid signal-selling, but stock-specific-call content carries standard unlicensed-advice risk common to the niche",
         depth_confidence="Low — surfaced only via generic search aggregation, not independently cross-referenced by a third party; operator's real identity/nationality not established beyond the channel's exclusive Saudi-market subject matter",
         sources="https://www.youtube.com/@tasi_y ; https://x.com/tasi_y?lang=ar",
         status="Operator identity unknown — lower-confidence lead, verify directly"),
    dict(name="أبو نوره / KsaTrader (Abu Nourah)", category="Trading",
         handle="@KsaTrader", platform="x",
         niche="Daily automated/algorithmic technical analysis of Saudi stocks, #تاسي #تداول #الاسهم_السعودية",
         credentials="n/a", socials="Cited as also having Telegram and Snapchat channels", followers="n/a",
         why_fit="Confirmed active X account; handle signals a Saudi-market identity (\"KSA Trader\"); multi-platform footprint (X + Telegram + Snapchat)",
         risk_flag="Daily algorithmic-style stock calls disclaimed as \"not official recommendations\" — a pattern the CMA has warned against for informal social-media investment tips",
         depth_confidence="Low-Medium — handle/hashtag usage strongly suggest Saudi-market focus, but personal identity/nationality not independently confirmed beyond self-selected \"KSA\" branding",
         sources="https://x.com/KsaTrader",
         status="Nationality inferred from self-branding only — flagged as unconfirmed"),
    dict(name="محمد الميموني / Mohammad Al-Maimoni", category="Trading",
         handle="@mohammad_fx", platform="x",
         niche="Forex and Saudi stock trading educator/commentator — risk management (stop-loss), technical/fundamental/news-based trading",
         credentials="Described in podcast coverage as trading financial markets since 2004, a \"certified trainer\" (certifying body unconfirmed), based in Riyadh, Saudi Arabia, age ~46",
         socials="Analyst/signals profile on arabictrader.com; guest on the Saudi \"بترولي\" (Bertoli) podcast (YouTube/Spotify)",
         followers="~516 on X per an indexed search snippet (low; may be stale — treat as approximate)",
         why_fit="Most independently corroborated trading candidate this round — own active X account with original Saudi-market commentary, profiled by a real Saudi podcast with documented forex-trading history, and a page on the arabictrader.com analyst community",
         risk_flag="arabictrader.com hosts a \"توصيات محمد الميموني\" (Mohammad Al-Maimoni's recommendations) page — i.e. he distributes trade recommendations/signals; no evidence found of CMA licensing for investment advice",
         depth_confidence="Medium-High — identity corroborated across own account, third-party podcast profile with biographical detail, and a trading-community signals page; Saudi location (Riyadh) explicitly stated in a source, not merely inferred",
         sources="https://x.com/mohammad_fx ; https://www.arabictrader.com/ar/signal/mohammad_fx ; https://www.arabictrader.com/ar/analysis/user/mohammad_fx ; https://www.youtube.com/watch?v=8gKBZrD_s98",
         status="Explicit unlicensed-recommendations risk flag — verify current follower count before outreach"),
    dict(name="شؤون الأسهم / Stock Affairs", category="Trading",
         handle="@Stock_Affairs", platform="x",
         niche="Financial/technical analysis with specific focus on Saudi IPOs and investment tracking; site claims AI-assisted analysis \"from a licensed company\"",
         credentials="Site claims 15+ yrs stock-market experience and a \"licensed company\" behind the AI tool (licensing body/number not confirmed against the CMA's licensed-entities list)",
         socials="Website stocksaffairsa.com", followers="n/a",
         why_fit="Active, findable X account plus a linked commercial website, distinct niche (IPO tracking)",
         risk_flag="Markets \"AI-powered trading services\" and investment recommendations based on an unverified \"licensed\" claim — should be treated as an unverified/potentially misleading licensing claim pending confirmation",
         depth_confidence="Low — reads more like a brand/media outlet than one identifiable individual creator; no personal name surfaced, no TikTok/Instagram presence found",
         sources="https://x.com/Stock_Affairs ; https://stocksaffairsa.com/ ; https://stocksaffairsa.com/about-us/",
         status="Unclear if individual or small team/brand; \"licensed company\" claim unverified — treat cautiously"),
    dict(name="د. عبدالله المزراقي / Dr. Abdullah Al-Mezraqi", category="Trading",
         handle="@dr.mezraqi", platform="tiktok",
         niche="TASI/Tadawul stock analysis plus American-market options commentary, under the branded hashtag #المزراقي_تحليل_الأسهم",
         credentials="Uses the honorific \"Dr.\" in his display name; no verifiable academic/professional credential found beyond that self-description",
         socials="n/a confirmed beyond TikTok",
         followers="n/a",
         why_fit="Active, dedicated TikTok account with a distinct content brand around Saudi TASI stock analysis and American-market options, using Saudi-specific hashtags/flag emoji",
         risk_flag="One source indicates content is framed as analysis \"without buy/sell recommendations,\" which is lower-risk than signal-selling, but this self-description is unverified",
         depth_confidence="Medium — account/handle and Saudi-market-specific content verified via multiple TikTok video listings; Saudi nationality inferred from content/hashtags, not an explicit self-statement; no follower count retrievable",
         sources="https://www.tiktok.com/@dr.mezraqi ; https://www.tiktok.com/@dr.mezraqi/video/7578628902151572754 ; https://www.tiktok.com/@dr.mezraqi/video/7490323957678132488 ; https://www.tiktok.com/@dr.mezraqi/video/7514827546895060232",
         status=""),
    dict(name="USA Stock Hunter (صائد الأسهم الأمريكية)", category="Trading",
         handle="@usastockhunter", platform="tiktok",
         niche="American stock-market day-trading and investing education for Arabic-speaking beginners; paid courses via usastockhunteracademy.com",
         credentials="Self-described \"10+ years\" experience trading US markets; no independently verifiable license or credential found",
         socials="Instagram, YouTube, Facebook, X (all @usastockhunter); Telegram t.me/stockhunteralerts (\"free alerts\")",
         followers="~20K (TikTok, ~103.5K likes, per search results — not independently confirmed on-profile)",
         why_fit="Genuine, multi-platform, actively producing American-stock trading education content with a clear brand and following",
         risk_flag="HIGH — explicitly offers stock recommendations (\"توصيات\"), indicators, and a Telegram \"alerts\" channel alongside paid courses; free-alerts-funnel + paid-academy + recommendations is a textbook signal-selling pattern",
         depth_confidence="Low-Medium — account and cross-platform presence well-documented, but Saudi nationality could not be independently confirmed (no explicit Saudi self-identification, Saudi flag, or TASI content found — content is exclusively US-market); flagging explicitly, possible wrong-nationality risk",
         sources="https://www.tiktok.com/@usastockhunter ; https://www.instagram.com/usastockhunter/ ; https://www.youtube.com/@usastockhunter ; https://x.com/usastockhunter ; https://usastockhunteracademy.com/author/usastockhunter/",
         status="Nationality unconfirmed — no Saudi-specific content signal found despite multi-platform presence"),
    dict(name="أكاديمية بندر المحسن المالية / Bandar Al-Mohsin Financial Academy", category="Trading",
         handle="@BAlmohsin1", platform="x",
         niche="American stocks and Sharia-compliant options-contract (\"عقود الأوبشن الشرعية\") trading education, live interactive analysis broadcasts, paid beginner courses",
         credentials="Full name found as \"بندر بن عبدالله المحسن\" (Bandar bin Abdullah Al-Mohsin) — a Saudi-pattern tribal naming convention; no independently verifiable financial license/credential located",
         socials="Telegram t.me/Balmohsin; website balmohsin.com",
         followers="n/a",
         why_fit="One of the few creators found teaching American options contracts to an Arabic audience with an explicit halal/Sharia-compliant-options framing — a distinct sub-niche",
         risk_flag="HIGH — explicitly sells an options-analysis + interactive-live-stream package and requires naming a specific stock/contract symbol for personalized replies; advice/signal-adjacent commercial activity around a leveraged, high-risk instrument",
         depth_confidence="Medium — account, brand, and offering well-documented across X, Telegram, and a dedicated website; Saudi nationality inferred fairly strongly from the full name pattern but not explicitly self-stated; no TikTok presence found (X/Telegram-first)",
         sources="https://x.com/BAlmohsin1 ; https://x.com/bandaralmohsin ; https://balmohsin.com/ ; https://t.me/Balmohsin ; https://x.com/BAlmohsin1/status/1901294317743882446",
         status="Nationality inferred from name, not confirmed; explicit unlicensed-advice risk around options"),
    dict(name="Atradez (\"Ahmed\")", category="Trading",
         handle="@atradez", platform="youtube",
         niche="Personal finance, saving, and American stock-market investing education, with specific focus on penny stocks and building a trading community for market newcomers",
         credentials="Founder referred to only as \"Ahmed\" (أحمد) in search snippets; no surname, license, or further biographical detail found",
         socials="Website atradez.pro (paid community/products page)",
         followers="n/a",
         why_fit="Matches a personal-finance-creator-who-covers-investing profile (channel centers on money management, saving methods, and general market investing, not just stocks)",
         risk_flag="MEDIUM-HIGH — specializes in promoting penny stocks, a category flagged industry-wide for high volatility, manipulation risk, and unsuitability for retail beginners, combined with a paid community",
         depth_confidence="Low — Saudi nationality, full name, and any Saudi-specific content could not be independently confirmed; channel appears US-market-focused with no Saudi-specific signals found (no TASI content, no Saudi hashtags/flag)",
         sources="https://www.youtube.com/@atradez ; https://atradez.pro/",
         status="Nationality unconfirmed — weak Saudi signal, needs direct verification before outreach"),
    dict(name="صالح القحطاني / Saleh Al-Qahtani", category="Trading",
         handle="@s-qht11", platform="snapchat",
         niche="Snapchat Spotlight investment-opportunity content — e.g. promoting a specific car-services-company investment offer (\"ساطع للخدمات السيارة\")",
         credentials="n/a — no license or formal credential found",
         socials="n/a confirmed beyond Snapchat Spotlight",
         followers="n/a",
         why_fit="Real, dated (2026) Snapchat Spotlight post promoting an investment opportunity to a Saudi audience; \"Al-Qahtani\" is a Saudi tribal surname",
         risk_flag="Content promotes a specific named investment opportunity/company rather than general trading education — closer to investment-opportunity marketing than personal trading-signal selling, but should be checked for proper disclosure and whether the promoted opportunity is CMA-regulated",
         depth_confidence="Low — a single dated Spotlight post is the only direct evidence found; no broader profile, follower count, or biography retrievable this session",
         sources="https://www.snapchat.com/@s-qht11/spotlight/W7_EDlXWTBiXAEEniNoMPwAAYa2pueGF3enN6AZ5XBFbPAZ5XBA49AAAAAQ",
         status="Thin evidence — single post found; verify full profile before outreach"),
    dict(name="محمد الحديثي / Mohammed Al-Hudaithi", category="Trading",
         handle="@mhmh1199", platform="tiktok",
         niche="Saudi stock market/investment-fund education, incl. content on calculating zakat on investment funds (#صناديق_الاستثمار #اسهم)",
         credentials="n/a — no license or formal credential found",
         socials="n/a confirmed beyond TikTok",
         followers="~142.7K (336.9K likes)",
         why_fit="Sizable, active TikTok account with content specifically on Saudi investment funds and zakat compliance — a distinct sub-niche (Islamic-finance-compliant investing) not covered by earlier candidates",
         risk_flag="No explicit nationality statement found; \"Al-Hudaithi\" is a tribal surname found among both Saudi/Najdi and Iraqi families — flagging as inferred, not confirmed",
         depth_confidence="Medium — sizable, verifiable following and consistent investment-fund/zakat content niche; nationality inferred from name only",
         sources="https://www.tiktok.com/@mhmh1199 ; https://www.tiktok.com/@mhmh1199/video/7454104777165688082 ; https://www.tiktok.com/@mhmh1199/video/7162640661404028161",
         status="Nationality inferred from name, not confirmed"),
]

METHOD_NOTES = [
    "No-fabrication contract: every candidate is a real, findable account verified via web search; every claim (follower count, credentials, nationality) is sourced. No name, handle, or statistic was invented.",
    "This is a second, from-scratch research pass explicitly excluding all 25 candidates already shown in NH_TikTok_Ecommerce_Trading_RealEstate_Scout.xlsx and all 25 in NH_TikTok_EcommerceTrading_Only25.xlsx (50 handles total) — none of those 50 appear here.",
    "Because the pool of clearly-Saudi, e-commerce/trading, social-media-first creators is finite, this second pass necessarily includes thinner leads than the first — several rows here have Low depth_confidence and unresolved nationality; these are flagged per-row, not smoothed over.",
    "Two of the four research agents run for this pass independently surfaced the same 5 e-commerce leads (Sahel Mahdi, Dr. Abdulkarim Anaf, Ibrahim Al-Bahrawy, Ahmed Matar, @ifast.sa) — deduplicated to one entry each, keeping the richer sourcing from both passes.",
    "One trading lead found this round (@saadaltwaim1) was dropped because it is the same handle already shown to the user as a Real Estate entry in the first scout — not counted as 'new'. One e-commerce lead (Tarek Al-Arabi, @al.araby85) was dropped because his own TikTok bio confirms Egyptian nationality (🇪🇬) — moved out rather than included as merely 'flagged'.",
    "This environment's network egress proxy blocked direct page loads of tiktok.com, instagram.com, x.com, snapchat.com, and most influencer-listing aggregators during this research pass — all data comes from search-engine result snippets/summaries, not directly re-loaded profile pages. Follower counts especially should be treated as indicative only and re-verified by Northouse directly before outreach or valuation decisions.",
]

TIKTOK_URL_RE = re.compile(r"https?://(?:www\.)?tiktok\.com/@[A-Za-z0-9_.]+")
X_URL_RE = re.compile(r"https?://(?:www\.)?x\.com/[A-Za-z0-9_]+")
SNAP_URL_RE = re.compile(r"https?://(?:www\.)?snapchat\.com/@[A-Za-z0-9_.\-]+")
YT_URL_RE = re.compile(r"https?://(?:www\.)?youtube\.com/@[A-Za-z0-9_.\-]+")
IG_URL_RE = re.compile(r"https?://(?:www\.)?instagram\.com/[A-Za-z0-9_.\-]+/?")
HANDLE_RE = re.compile(r"@[A-Za-z0-9_.\-]+")


def account_url(c):
    plat = c["platform"]
    src = c["sources"]
    if plat == "x":
        m = X_URL_RE.search(src)
        if m:
            return m.group(0)
    elif plat == "snapchat":
        m = SNAP_URL_RE.search(src)
        if m:
            return m.group(0)
    elif plat == "youtube":
        m = YT_URL_RE.search(src)
        if m:
            return m.group(0)
    elif plat == "instagram":
        m = IG_URL_RE.search(src)
        if m:
            return m.group(0)
    elif plat == "website":
        return c["handle"] if c["handle"].startswith(("http://", "https://")) else f"https://{c['handle']}"
    elif plat == "unknown":
        return None
    m = TIKTOK_URL_RE.search(src)
    if m:
        return m.group(0)
    h = HANDLE_RE.search(c["handle"])
    if h and plat == "tiktok":
        return f"https://www.tiktok.com/{h.group(0)}"
    return None


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

    ws2 = wb.create_sheet("Chart")
    ws2["A1"] = "25 NEW Candidates by Niche (E-commerce / Trading only)"
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

    # ---------------- All 25 profiles, with clickable links, on the Chart tab ----------------
    ws2.append([])
    r = ws2.max_row + 1
    ws2.cell(row=r, column=1, value="All 25 Profiles").font = Font(name=FONT_NAME, bold=True, size=13)
    ws2.append([])
    ws2.append(["#", "Name", "Niche", "Profile Link"])
    style_header(ws2, ws2.max_row, 4)

    for i, c in enumerate(CANDIDATES, 1):
        url = account_url(c)
        ws2.append([i, c["name"], c["category"], url or "No confirmed handle — see Creators tab"])
        r = ws2.max_row
        if url:
            cell = ws2.cell(row=r, column=4)
            cell.hyperlink = url
            cell.style = "Hyperlink"

    for row in ws2.iter_rows(min_row=ws2.max_row - len(CANDIDATES), max_row=ws2.max_row, max_col=4):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    ws3 = wb.create_sheet("Method")
    ws3["A1"] = "Method Notes"
    ws3["A1"].font = Font(name=FONT_NAME, bold=True, size=13)
    ws3.append([])
    for note in METHOD_NOTES:
        ws3.append([note])
        ws3.merge_cells(start_row=ws3.max_row, start_column=1, end_row=ws3.max_row, end_column=3)
    for row in ws3.iter_rows(min_row=1, max_row=ws3.max_row, max_col=3):
        for cell in row:
            if cell.value is not None:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws3.column_dimensions["A"].width = 100

    for name in ("Creators", "Chart", "Method"):
        for row in wb[name].iter_rows():
            for cell in row:
                if cell.font.name != "Consolas":
                    existing = cell.font
                    cell.font = Font(name=FONT_NAME, bold=existing.bold, color=existing.color,
                                      size=existing.size if existing.size else 10)

    out = "NH_TikTok_EcommerceTrading_New25.xlsx"
    wb.save(out)
    print(f"Wrote {out}: {len(CANDIDATES)} candidates "
          f"({counts['E-commerce']} E-commerce, {counts['Trading']} Trading), "
          f"{len(follower_rows)} with a disclosed follower count.")


if __name__ == "__main__":
    main()
