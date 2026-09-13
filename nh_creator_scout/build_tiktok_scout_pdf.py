#!/usr/bin/env python3
"""
Renders NH_TikTok_Ecommerce_Trading_RealEstate_Scout.pdf — a readable report
version of the same 20-candidate Saudi TikTok scout data as the xlsx
(build_tiktok_scout.py), for sharing/printing rather than filtering/sorting.
"""
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, PageBreak,
)
from reportlab.lib.enums import TA_LEFT

from build_tiktok_scout import CANDIDATES, EXCLUDED, TOP_PICKS, METHOD_NOTES

NAVY = colors.HexColor("#1F4E78")
LIGHT_GREY = colors.HexColor("#D9D9D9")
DARK_GREY = colors.HexColor("#444444")
FLAG_RED = colors.HexColor("#B00020")

styles = getSampleStyleSheet()
title_style = ParagraphStyle("TitleX", parent=styles["Title"], textColor=NAVY, spaceAfter=4)
subtitle_style = ParagraphStyle("Subtitle", parent=styles["Normal"], textColor=DARK_GREY,
                                 fontSize=10, spaceAfter=14, leading=13)
h1 = ParagraphStyle("H1", parent=styles["Heading1"], textColor=colors.white, backColor=NAVY,
                     fontSize=13, leading=16, spaceBefore=0, spaceAfter=10,
                     leftIndent=6, borderPadding=(4, 4, 4, 4))
h2 = ParagraphStyle("H2", parent=styles["Heading2"], textColor=NAVY, fontSize=12, spaceBefore=6, spaceAfter=6)
name_style = ParagraphStyle("Name", parent=styles["Normal"], fontSize=11, fontName="Helvetica-Bold",
                             textColor=colors.black, spaceAfter=2)
field_style = ParagraphStyle("Field", parent=styles["Normal"], fontSize=8.7, leading=11.5, spaceAfter=1)
flag_style = ParagraphStyle("Flag", parent=field_style, textColor=FLAG_RED, fontName="Helvetica-Oblique")
source_style = ParagraphStyle("Source", parent=field_style, fontSize=7.5, textColor=colors.HexColor("#555577"))
note_style = ParagraphStyle("Note", parent=styles["Normal"], fontSize=9, leading=12.5, spaceAfter=6)
label = lambda s: f"<b>{s}</b>"


def candidate_block(c, idx):
    flow = []
    flow.append(Paragraph(f"{idx}. {c['name']}  —  <font color='#1F4E78'>{c['handle']}</font>", name_style))
    flow.append(Paragraph(f"{label('Niche:')} {c['niche']}", field_style))
    if c["credentials"] and c["credentials"] != "n/a":
        flow.append(Paragraph(f"{label('Background:')} {c['credentials']}", field_style))
    if c["socials"] and c["socials"] not in ("n/a", "n/a found"):
        flow.append(Paragraph(f"{label('Other socials:')} {c['socials']}", field_style))
    flow.append(Paragraph(f"{label('Followers:')} {c['followers']}  &nbsp;&nbsp; {label('Confidence:')} {c['depth_confidence'] if len(c['depth_confidence']) < 40 else c['depth_confidence'].split(' — ')[0]}", field_style))
    flow.append(Paragraph(f"{label('Why fit:')} {c['why_fit']}", field_style))
    if c["risk_flag"]:
        flow.append(Paragraph(f"⚠ Risk/compliance: {c['risk_flag']}", flag_style))
    if c["status"]:
        flow.append(Paragraph(f"⚑ {c['status']}", flag_style))
    flow.append(Paragraph(f"Source(s): {c['sources']}", source_style))
    flow.append(Spacer(1, 8))
    return flow


def build():
    doc = SimpleDocTemplate(
        "NH_TikTok_Ecommerce_Trading_RealEstate_Scout.pdf",
        pagesize=letter,
        leftMargin=0.7 * inch, rightMargin=0.7 * inch,
        topMargin=0.7 * inch, bottomMargin=0.6 * inch,
        title="Saudi TikTok Creator Scout — E-commerce / Trading / Real Estate",
    )
    story = []

    story.append(Paragraph("Saudi TikTok Creator Scout", title_style))
    story.append(Paragraph("E-commerce &middot; Trading &middot; Real Estate — 20 candidates", subtitle_style))
    story.append(Paragraph(
        "Sourced via three parallel web-research passes (not direct TikTok scraping/API access). "
        "Every candidate needed an identifiable TikTok handle and a citable source; unknown fields are "
        "marked n/a rather than guessed, and nationality/compliance uncertainty is flagged in red rather "
        "than silently resolved.",
        note_style,
    ))
    story.append(HRFlowable(width="100%", color=LIGHT_GREY, thickness=1))
    story.append(Spacer(1, 8))

    categories = ["E-commerce", "Trading", "Real Estate"]
    for cat in categories:
        story.append(Paragraph(f"&nbsp;{cat}", h1))
        items = [c for c in CANDIDATES if c["category"] == cat]
        for i, c in enumerate(items, 1):
            story.extend(candidate_block(c, i))
        story.append(Spacer(1, 4))

    # ---------------- Top Picks ----------------
    story.append(PageBreak())
    story.append(Paragraph("&nbsp;Top Picks", h1))
    tp_data = [["Handle", "Niche", "Why"]]
    for handle, cat, note in TOP_PICKS:
        tp_data.append([
            Paragraph(handle, field_style),
            Paragraph(cat, field_style),
            Paragraph(note, field_style),
        ])
    tp_table = Table(tp_data, colWidths=[1.1 * inch, 0.95 * inch, 4.75 * inch], repeatRows=1)
    tp_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("GRID", (0, 0), (-1, -1), 0.5, LIGHT_GREY),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F4F7FA")]),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(tp_table)
    story.append(Spacer(1, 16))

    # ---------------- Excluded & Flags ----------------
    story.append(Paragraph("&nbsp;Excluded During Research", h1))
    story.append(Paragraph(
        "Namesake traps, wrong-country matches, niche mismatches, or names with no verifiable TikTok "
        "presence — kept here for transparency rather than silently dropped.", note_style))
    ex_data = [["Name", "Niche", "Reason excluded"]]
    for e in EXCLUDED:
        ex_data.append([
            Paragraph(e["name"], field_style),
            Paragraph(e["category"], field_style),
            Paragraph(e["reason"], field_style),
        ])
    ex_table = Table(ex_data, colWidths=[1.7 * inch, 0.95 * inch, 4.15 * inch], repeatRows=1)
    ex_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("GRID", (0, 0), (-1, -1), 0.5, LIGHT_GREY),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F4F7FA")]),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(ex_table)
    story.append(Spacer(1, 16))

    # ---------------- Method ----------------
    story.append(Paragraph("&nbsp;Method &amp; Notes", h1))
    counts = f"{len(CANDIDATES)} total &mdash; {sum(1 for c in CANDIDATES if c['category']=='E-commerce')} e-commerce, {sum(1 for c in CANDIDATES if c['category']=='Trading')} trading, {sum(1 for c in CANDIDATES if c['category']=='Real Estate')} real estate. {len(EXCLUDED)} excluded during research."
    story.append(Paragraph(f"<b>Counts:</b> {counts}", note_style))
    for n in METHOD_NOTES:
        story.append(Paragraph(f"&bull; {n}", note_style))

    doc.build(story)
    print("Wrote NH_TikTok_Ecommerce_Trading_RealEstate_Scout.pdf")


if __name__ == "__main__":
    build()
