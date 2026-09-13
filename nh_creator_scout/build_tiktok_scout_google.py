#!/usr/bin/env python3
"""
Builds NH_TikTok_Scout_ForGoogleSheets.xlsx — same 20-candidate dataset as
build_tiktok_scout.py, reshaped for Google Sheets: one flat, chart-friendly
tab with clickable hyperlinks straight to each creator's TikTok account, plus
two embedded bar charts (candidates by niche; followers for the candidates
with a disclosed count).

Google Sheets opens .xlsx directly (File > Open, or drag into Drive) and
keeps the hyperlinks and charts as native Sheets objects.
"""
import re

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList

from build_tiktok_scout import CANDIDATES, TOP_PICKS

FONT_NAME = "Arial"
NAVY = "1F4E78"

# Followers with a disclosed approximate count, transcribed from the source
# text (thousands). Candidates not listed here had no disclosed count ("n/a").
FOLLOWERS_K = {
    "@econabdallah": 178.1,
    "@deserttrader6906": 52.0,
    "@assadalsharef": 1100.0,  # ~1.1M
    "@abo_ahmed_al3mda": 11.9,
    "@jubail_aqari": 33.5,
    "@aqarapp": 99.8,
    "@rr05l": 37.0,
}

TIKTOK_URL_RE = re.compile(r"https?://(?:www\.)?tiktok\.com/@[A-Za-z0-9_.]+")
HANDLE_RE = re.compile(r"@[A-Za-z0-9_.]+")


def account_url(candidate):
    """Prefer a tiktok.com URL already present in the sourced citations;
    otherwise derive the profile URL directly from the verified handle
    (tiktok.com/@handle is a deterministic mapping, not a guess)."""
    m = TIKTOK_URL_RE.search(candidate["sources"])
    if m:
        return m.group(0), "Cited directly in research sources"
    h = HANDLE_RE.search(candidate["handle"])
    if h:
        return f"https://www.tiktok.com/{h.group(0)}", "Derived from verified handle (not independently re-clicked)"
    return None, "No handle found"


def style_header(ws, row, ncols, fill=NAVY):
    header_font = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor=fill)
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)


def autosize(ws, min_width=10, max_width=48):
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


def main():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Creators"

    headers = ["#", "Name (AR/EN)", "Niche", "TikTok Account (click)", "Followers (approx.)",
               "Content Focus", "Why Notable/Fit", "Risk/Flag", "Link Basis"]
    ws.append(headers)
    style_header(ws, 1, len(headers))
    ws.freeze_panes = "A2"

    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    niche_counts = {}
    for i, c in enumerate(CANDIDATES, 1):
        url, basis = account_url(c)
        niche_counts[c["category"]] = niche_counts.get(c["category"], 0) + 1
        row = [i, c["name"], c["category"], c["handle"], c["followers"],
               c["niche"], c["why_fit"], c["risk_flag"] or c["status"], basis]
        ws.append(row)
        r = ws.max_row
        if url:
            cell = ws.cell(row=r, column=4)
            cell.hyperlink = url
            cell.style = "Hyperlink"

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers)):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    autosize(ws, max_width=42)

    # ---------------- helper data for charts (kept on a plain sheet so the
    # charts remain live/editable if data changes) ----------------
    ws_data = wb.create_sheet("Chart Data")
    ws_data.append(["Niche", "Candidates"])
    for niche in ["E-commerce", "Trading", "Real Estate"]:
        ws_data.append([niche, niche_counts.get(niche, 0)])

    ws_data.append([])
    ws_data.append(["Handle", "Followers (thousands)"])
    for c in CANDIDATES:
        if c["handle"] in FOLLOWERS_K:
            ws_data.append([c["handle"], FOLLOWERS_K[c["handle"]]])

    style_header(ws_data, 1, 2)
    style_header(ws_data, 6, 2)
    autosize(ws_data)

    # ---------------- chart 1: candidates by niche ----------------
    chart1 = BarChart()
    chart1.type = "col"
    chart1.title = "Saudi TikTok Creators Found, by Niche"
    chart1.y_axis.title = "Candidates"
    chart1.x_axis.title = "Niche"
    chart1.style = 10
    data = Reference(ws_data, min_col=2, min_row=1, max_row=4)
    cats = Reference(ws_data, min_col=1, min_row=2, max_row=4)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.dataLabels = DataLabelList()
    chart1.dataLabels.showVal = True
    chart1.width, chart1.height = 14, 8
    ws_data.add_chart(chart1, "E2")

    # ---------------- chart 2: followers, where disclosed ----------------
    n_followers = len(FOLLOWERS_K)
    chart2 = BarChart()
    chart2.type = "bar"
    chart2.title = "Disclosed TikTok Followers (thousands) — 7 of 20 candidates"
    chart2.y_axis.title = "Handle"
    chart2.x_axis.title = "Followers (K)"
    chart2.style = 11
    start_row = 7
    end_row = 6 + n_followers
    data2 = Reference(ws_data, min_col=2, min_row=start_row, max_row=end_row)
    cats2 = Reference(ws_data, min_col=1, min_row=start_row + 1, max_row=end_row)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats2)
    chart2.dataLabels = DataLabelList()
    chart2.dataLabels.showVal = True
    chart2.width, chart2.height = 16, 10
    ws_data.add_chart(chart2, "E20")

    # ---------------- Top Picks tab, with the same clickable links ----------------
    ws_top = wb.create_sheet("Top Picks")
    ws_top.append(["Niche", "TikTok Account (click)", "Rank Note"])
    style_header(ws_top, 1, 3)
    ws_top.freeze_panes = "A2"
    by_handle = {c["handle"]: c for c in CANDIDATES}
    for handle, cat, note in TOP_PICKS:
        c = by_handle.get(handle)
        url, _ = account_url(c) if c else (None, None)
        ws_top.append([cat, handle, note])
        r = ws_top.max_row
        if url:
            cell = ws_top.cell(row=r, column=2)
            cell.hyperlink = url
            cell.style = "Hyperlink"
    for row in ws_top.iter_rows(min_row=2, max_row=ws_top.max_row, max_col=3):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    autosize(ws_top, max_width=70)

    out = "NH_TikTok_Scout_ForGoogleSheets.xlsx"
    wb.save(out)
    print(f"Wrote {out}: {len(CANDIDATES)} rows, {len(FOLLOWERS_K)} with disclosed followers.")


if __name__ == "__main__":
    main()
