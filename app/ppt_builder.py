"""Builds a .pptx presentation from a DeckData object.

Layout constants are tuned to the 16:9 canvas defined in theme.py. Every
build_* function draws one slide (or a small family of near-identical
slides, e.g. one line-of-business deep-dive per LoB found in the data).
"""
from __future__ import annotations

from collections import OrderedDict
from typing import List, Optional

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

from . import theme
from .ppt_helpers import (
    add_background, add_rect, add_text, add_multirun_text, add_table,
    set_cell, status_pill, status_dot, add_kpi_tile, fmt_num, rgb, no_line,
)
from .models import DeckData, Initiative, Project

W = theme.SLIDE_W
H = theme.SLIDE_H
M = theme.MARGIN


# --------------------------------------------------------------------------
# Frame: background, header, footer
# --------------------------------------------------------------------------

def new_slide(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide, theme.WHITE)
    return slide


def new_dark_slide(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide, theme.INK)
    return slide


def add_header(slide, headline: str, subheadline: str = ""):
    add_text(slide, M, Emu(274320), W - 2 * M, Emu(520000), headline,
              size=19, color=theme.INK, bold=True)
    top = Emu(274320) + Emu(500000)
    if subheadline:
        add_text(slide, M, top, W - 2 * M, Emu(340000), subheadline,
                  size=11, color=theme.MUTED)
    rule_top = Emu(1150000)
    add_rect(slide, M, rule_top, W - 2 * M, Emu(9525), fill=theme.LINE)
    return rule_top + Emu(150000)


def add_footer(slide, footer_text: str, page_num: int):
    y = H - Emu(320000)
    add_text(slide, M, y, Emu(6000000), Emu(240000), footer_text or "",
              size=9, color=theme.MUTED)
    add_text(slide, W - M - Emu(600000), y, Emu(600000), Emu(240000), str(page_num),
              size=9, color=theme.MUTED, align=PP_ALIGN.RIGHT)


def finalize(slide, footer_text: str, page_num: int):
    add_footer(slide, footer_text, page_num)


# --------------------------------------------------------------------------
# 1. Cover
# --------------------------------------------------------------------------

def build_cover(prs, data: DeckData):
    slide = new_dark_slide(prs)
    cy = Emu(2700000)
    if data.cover.subtitle:
        add_text(slide, M, cy - Emu(500000), W - 2 * M, Emu(400000),
                  data.cover.subtitle.upper(), size=14, color=theme.MUTED, bold=True)
    add_text(slide, M, cy, W - 2 * M, Emu(900000), data.cover.title or "Untitled Deck",
              size=36, color=theme.WHITE, bold=True)
    if data.cover.date_label:
        add_text(slide, M, cy + Emu(950000), W - 2 * M, Emu(350000),
                  data.cover.date_label, size=13, color=theme.MUTED)
    if data.cover.footer:
        add_text(slide, M, H - Emu(500000), W - 2 * M, Emu(300000),
                  data.cover.footer, size=10, color=theme.MUTED)
    add_text(slide, W - M - Emu(600000), H - Emu(500000), Emu(600000), Emu(300000),
              "1", size=10, color=theme.MUTED, align=PP_ALIGN.RIGHT)
    return slide


# --------------------------------------------------------------------------
# 2. Table of contents
# --------------------------------------------------------------------------

def build_toc(prs, data: DeckData, page_num: int):
    slide = new_slide(prs)
    add_header(slide, "Contents")
    top = Emu(1450000)
    row_h = Emu(720000)
    for entry in data.toc:
        add_text(slide, M, top, Emu(700000), row_h, entry.number, size=20, color=theme.PRIMARY, bold=True)
        add_text(slide, M + Emu(800000), top, Emu(7500000), Emu(400000), entry.section, size=15, bold=True, color=theme.INK)
        add_text(slide, M + Emu(800000), top + Emu(380000), Emu(7500000), Emu(320000), entry.description, size=10, color=theme.MUTED)
        add_text(slide, W - M - Emu(700000), top, Emu(700000), row_h, entry.page, size=13, color=theme.MUTED, align=PP_ALIGN.RIGHT)
        top += row_h
    finalize(slide, data.cover.footer, page_num)
    return slide


# --------------------------------------------------------------------------
# 3. Section divider
# --------------------------------------------------------------------------

def build_section_divider(prs, number: str, title: str, subtitle: str, data: DeckData, page_num: int):
    slide = new_dark_slide(prs)
    add_text(slide, M, Emu(2400000), Emu(1200000), Emu(900000), number, size=44, color=theme.PRIMARY, bold=True)
    add_text(slide, M, Emu(3200000), W - 2 * M, Emu(600000), title, size=26, color=theme.WHITE, bold=True)
    if subtitle:
        add_text(slide, M, Emu(3850000), W - 2 * M, Emu(400000), subtitle, size=13, color=theme.MUTED)
    finalize(slide, data.cover.footer, page_num)
    return slide


# --------------------------------------------------------------------------
# 4. Executive summary
# --------------------------------------------------------------------------

def build_exec_summary(prs, data: DeckData, page_num: int):
    es = data.exec_summary
    slide = new_slide(prs)
    content_top = add_header(slide, es.headline or "Executive Summary", es.subheadline)

    col_w = Emu(3750000)
    gap = Emu(180000)
    left1 = M

    add_kpi_tile(slide, left1, content_top, col_w, Emu(1150000),
                 es.committed_label or "Committed benefit", f"SAR {fmt_num(es.committed_bri_m)}m")

    # IPI tiers stacked bar beneath tile 1
    tiers_top = content_top + Emu(1150000) + Emu(180000)
    add_text(slide, left1, tiers_top, col_w, Emu(260000), "BY EXECUTION TIER", size=9, bold=True, color=theme.MUTED)
    bar_top = tiers_top + Emu(300000)
    bar_h = Emu(500000)
    total = sum(t.bri_m for t in es.ipi_tiers) or 1
    x = left1
    for t in es.ipi_tiers:
        seg_w = Emu(int(col_w * (t.bri_m / total)))
        if seg_w > 0:
            add_rect(slide, x, bar_top, seg_w, bar_h, fill=theme.status_color(t.color_key))
            if seg_w > Emu(500000):
                add_text(slide, x, bar_top, seg_w, bar_h, fmt_num(t.bri_m), size=10, bold=True,
                          color=theme.WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            x += seg_w
    label_top = bar_top + bar_h + Emu(80000)
    label_w = Emu(int(col_w / max(len(es.ipi_tiers), 1)))
    lx = left1
    for t in es.ipi_tiers:
        add_text(slide, lx, label_top, label_w, Emu(240000), t.label, size=8, color=theme.MUTED, align=PP_ALIGN.LEFT)
        lx += label_w

    # Savings tile
    savings_top = tiers_top + Emu(300000) + bar_h + Emu(520000)
    add_kpi_tile(slide, left1, savings_top, col_w, Emu(1000000),
                 "Savings BRI", f"SAR {fmt_num(es.savings_total_m)}m", value_color=theme.status_color("on_track"))
    sb_top = savings_top + Emu(1050000)
    add_text(slide, left1, sb_top, col_w, Emu(240000), "WHERE IT COMES FROM (SAR m)", size=9, bold=True, color=theme.MUTED)
    row_top = sb_top + Emu(320000)
    for item in es.savings_breakdown[:4]:
        add_text(slide, left1, row_top, Emu(2400000), Emu(280000), item.category, size=10, color=theme.INK)
        add_text(slide, left1 + Emu(2400000), row_top, Emu(1200000), Emu(280000), fmt_num(item.amount_m), size=10, bold=True, color=theme.PRIMARY, align=PP_ALIGN.RIGHT)
        row_top += Emu(300000)
        if item.detail:
            add_text(slide, left1, row_top, col_w, Emu(260000), item.detail, size=8, color=theme.MUTED)
            row_top += Emu(280000)

    # Right column: benefit status by LoB
    left2 = left1 + col_w + gap
    col2_w = W - M - left2
    add_text(slide, left2, content_top, col2_w, Emu(280000), "BENEFIT STATUS BY LINE OF BUSINESS", size=10, bold=True, color=theme.MUTED)
    row_top = content_top + Emu(400000)
    row_h = Emu(620000)
    max_bri = max([s.bri_m for s in es.benefit_status_by_lob] + [1])
    for s in es.benefit_status_by_lob:
        add_text(slide, left2, row_top, Emu(1800000), Emu(320000), s.lob, size=12, bold=True, color=theme.INK)
        add_text(slide, left2 + Emu(1800000), row_top, Emu(900000), Emu(320000), f"{fmt_num(s.percent_on_track)}%", size=12, bold=True, color=theme.PRIMARY, align=PP_ALIGN.RIGHT)
        track_top = row_top + Emu(360000)
        track_w = col2_w
        add_rect(slide, left2, track_top, track_w, Emu(160000), fill=theme.SURFACE_ALT)
        fill_w = Emu(int(track_w * (s.bri_m / max_bri)))
        if fill_w > 0:
            add_rect(slide, left2, track_top, fill_w, Emu(160000), fill=theme.PRIMARY)
        add_text(slide, left2, track_top + Emu(180000), track_w, Emu(240000), f"SAR {fmt_num(s.bri_m)}m", size=8, color=theme.MUTED)
        row_top += row_h
    finalize(slide, data.cover.footer, page_num)
    return slide


# --------------------------------------------------------------------------
# 5. Context / GWP charts
# --------------------------------------------------------------------------

def _style_chart(chart, show_legend=True):
    chart.has_legend = show_legend
    if show_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(9)
        chart.legend.font.name = theme.FONT
    chart.font.size = Pt(9)
    chart.font.name = theme.FONT
    try:
        chart.category_axis.tick_labels.font.size = Pt(9)
        chart.value_axis.tick_labels.font.size = Pt(9)
    except Exception:
        pass


def build_context_gwp(prs, data: DeckData, page_num: int):
    ctx = data.context_gwp
    slide = new_slide(prs)
    content_top = add_header(slide, ctx.headline or "Context — GWP Trajectory", ctx.subheadline)

    col_w = Emu(6800000)
    if ctx.years:
        add_text(slide, M, content_top, col_w, Emu(260000), "GWP TRAJECTORY (SAR B)", size=9, bold=True, color=theme.MUTED)
        chart_data = CategoryChartData()
        chart_data.categories = [y.year for y in ctx.years]
        chart_data.add_series("BAU base", [y.bau_m for y in ctx.years])
        chart_data.add_series("Initiative layer", [y.initiative_m for y in ctx.years])
        gframe = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_STACKED, M, content_top + Emu(300000), col_w, Emu(3900000), chart_data
        )
        chart = gframe.chart
        _style_chart(chart)
        chart.plots[0].series[0].format.fill.solid()
        chart.plots[0].series[0].format.fill.fore_color.rgb = rgb("C9C6E8")
        chart.plots[0].series[1].format.fill.solid()
        chart.plots[0].series[1].format.fill.fore_color.rgb = rgb(theme.PRIMARY)

    left2 = M + col_w + Emu(200000)
    col2_w = W - M - left2
    if ctx.by_lob:
        label = ctx.by_lob_year_label or "GWP BY LINE OF BUSINESS"
        add_text(slide, left2, content_top, col2_w, Emu(260000), label.upper(), size=9, bold=True, color=theme.MUTED)
        row_top = content_top + Emu(380000)
        row_h = Emu(560000)
        max_v = max([b.gwp_m for b in ctx.by_lob] + [1])
        for b in ctx.by_lob:
            add_text(slide, left2, row_top, col2_w, Emu(260000), b.lob, size=10, color=theme.INK, bold=True)
            track_top = row_top + Emu(280000)
            add_rect(slide, left2, track_top, col2_w, Emu(160000), fill=theme.SURFACE_ALT)
            fw = Emu(int(col2_w * (b.gwp_m / max_v)))
            if fw > 0:
                add_rect(slide, left2, track_top, fw, Emu(160000), fill=theme.PRIMARY)
            add_text(slide, left2, track_top + Emu(180000), col2_w, Emu(220000), f"{fmt_num(b.gwp_m)}m", size=8, color=theme.MUTED)
            row_top += row_h

    finalize(slide, data.cover.footer, page_num)
    return slide


# --------------------------------------------------------------------------
# 6. IPI by sector (custom horizontal bars with an on-track reference line)
# --------------------------------------------------------------------------

def build_ipi_sector(prs, data: DeckData, page_num: int):
    ipi = data.ipi_sector
    slide = new_slide(prs)
    content_top = add_header(slide, ipi.headline or "The Execution Signal — IPI by Sector", ipi.subheadline)

    chart_left = M + Emu(1600000)
    chart_w = W - 2 * M - Emu(1600000) - Emu(700000)
    max_scale = 5.0
    threshold = 3.0
    row_h = Emu(620000)
    top = content_top + Emu(150000)

    for item in ipi.items:
        add_text(slide, M, top, Emu(1500000), row_h, item.lob, size=12, bold=True, color=theme.INK, anchor=MSO_ANCHOR.MIDDLE)
        track_top = top + Emu(120000)
        track_h = Emu(360000)
        add_rect(slide, chart_left, track_top, chart_w, track_h, fill=theme.SURFACE_ALT)
        frac = max(0.0, min(1.0, item.ipi / max_scale))
        bar_w = Emu(int(chart_w * frac))
        color = theme.status_color("on_track") if item.ipi >= threshold else (
            theme.status_color("cautious") if item.ipi >= threshold - 0.5 else theme.status_color("at_risk"))
        if bar_w > 0:
            add_rect(slide, chart_left, track_top, bar_w, track_h, fill=color)
        add_text(slide, chart_left + chart_w + Emu(60000), track_top, Emu(650000), track_h,
                  f"{item.ipi:.2f}", size=11, bold=True, color=theme.INK, anchor=MSO_ANCHOR.MIDDLE)
        top += row_h

    # threshold reference line
    thresh_x = chart_left + Emu(int(chart_w * (threshold / max_scale)))
    add_rect(slide, thresh_x, content_top + Emu(150000), Emu(12700), top - content_top - Emu(150000), fill=theme.INK)
    add_text(slide, thresh_x - Emu(400000), top, Emu(800000), Emu(240000), f"{threshold:.1f} on-track", size=8, color=theme.MUTED, align=PP_ALIGN.CENTER)

    if ipi.read_note:
        add_text(slide, M, top + Emu(350000), W - 2 * M, Emu(700000), ipi.read_note, size=10, color=theme.MUTED, italic=True)

    finalize(slide, data.cover.footer, page_num)
    return slide


# --------------------------------------------------------------------------
# 7. Line-of-business overview (narrative + table)
# --------------------------------------------------------------------------

def build_lob_overview(prs, data: DeckData, page_num: int):
    lo = data.lob_overview
    slide = new_slide(prs)
    content_top = add_header(slide, lo.headline or "Where Benefit Concentrates by Line of Business", "")
    top = content_top
    if lo.narrative:
        add_text(slide, M, top, W - 2 * M, Emu(700000), lo.narrative, size=11, color=theme.INK_SOFT)
        top += Emu(800000)

    if lo.rows:
        headers = ["LINE OF BUSINESS", "STATUS SUMMARY", "IPI", "COMMITTED BRI, SAR M"]
        col_widths = [Emu(2100000), Emu(5600000), Emu(1200000), Emu(1900000)]
        table = add_table(slide, M, top, sum(col_widths, Emu(0)), Emu(400000) * (len(lo.rows) + 1), col_widths, len(lo.rows) + 1)
        for c, htext in enumerate(headers):
            set_cell(table.cell(0, c), htext, size=9, bold=True, color=theme.WHITE, fill=theme.INK,
                      align=PP_ALIGN.CENTER if c != 0 and c != 1 else PP_ALIGN.LEFT)
        for r, row in enumerate(lo.rows, start=1):
            fill = theme.SURFACE if r % 2 == 0 else theme.WHITE
            set_cell(table.cell(r, 0), row.lob, bold=True, fill=fill)
            set_cell(table.cell(r, 1), row.status_summary, fill=fill)
            set_cell(table.cell(r, 2), f"{row.ipi:.2f}", align=PP_ALIGN.CENTER, fill=fill)
            set_cell(table.cell(r, 3), fmt_num(row.bri_m), align=PP_ALIGN.CENTER, fill=fill, bold=True, color=theme.PRIMARY)

    finalize(slide, data.cover.footer, page_num)
    return slide


# --------------------------------------------------------------------------
# 8/9. Initiative & project deep-dive tables (one slide per LoB)
# --------------------------------------------------------------------------

def _grouped_by_lob(rows):
    grouped = OrderedDict()
    for r in rows:
        grouped.setdefault(r.lob, []).append(r)
    return grouped


def build_initiative_slide(prs, lob: str, initiatives: List[Initiative], data: DeckData, page_num: int):
    slide = new_slide(prs)
    headline = f"{lob} — Initiative Execution"
    subheadline = "Per-initiative IPI/TI vs. target, committed vs. actual BRI, and status."
    content_top = add_header(slide, headline, subheadline)

    headers = ["INITIATIVE", "IPI (vs 3.0)", "TI (vs 3.0)", "BRI COMMITTED, SAR M", "BRI ACTUAL, SAR M", "STATUS", "COMMENTS"]
    col_widths = [Emu(2450000), Emu(950000), Emu(950000), Emu(1350000), Emu(1250000), Emu(1050000), Emu(3192000)]
    n_rows = len(initiatives) + 1
    row_h = min(Emu(560000), Emu(int((H - content_top - Emu(500000)) / max(n_rows, 1))))
    table = add_table(slide, M, content_top, sum(col_widths, Emu(0)), row_h * n_rows, col_widths, n_rows)
    for c, htext in enumerate(headers):
        set_cell(table.cell(0, c), htext, size=8, bold=True, color=theme.WHITE, fill=theme.INK,
                  align=PP_ALIGN.CENTER if c != 0 and c != 6 else PP_ALIGN.LEFT)
    for r, it in enumerate(initiatives, start=1):
        fill = theme.SURFACE if r % 2 == 0 else theme.WHITE
        name = it.initiative + (f"\n{it.subtitle}" if it.subtitle else "")
        set_cell(table.cell(r, 0), name, size=9, bold=True, fill=fill)
        set_cell(table.cell(r, 1), f"{it.ipi:.2f}", align=PP_ALIGN.CENTER, fill=fill)
        set_cell(table.cell(r, 2), f"{it.ti:.2f}", align=PP_ALIGN.CENTER, fill=fill)
        set_cell(table.cell(r, 3), fmt_num(it.bri_committed_m), align=PP_ALIGN.CENTER, fill=fill)
        set_cell(table.cell(r, 4), fmt_num(it.bri_actual_m), align=PP_ALIGN.CENTER, fill=fill, bold=True, color=theme.PRIMARY)
        set_cell(table.cell(r, 5), theme.status_label(it.status), align=PP_ALIGN.CENTER, fill=theme.status_color(it.status), color=theme.WHITE, bold=True, size=8)
        set_cell(table.cell(r, 6), it.comments, size=8, fill=fill)

    finalize(slide, data.cover.footer, page_num)
    return slide


def build_project_slide(prs, lob: str, projects: List[Project], data: DeckData, page_num: int):
    slide = new_slide(prs)
    headline = f"{lob} — Projects Under Each Initiative"
    subheadline = "Every strategic initiative broken into its delivery projects, scored on IPI and Time (vs 3.0)."
    content_top = add_header(slide, headline, subheadline)

    headers = ["INITIATIVE / PROJECT", "IPI (vs 3.0)", "TI (vs 3.0)", "COMMENTS"]
    col_widths = [Emu(3200000), Emu(1100000), Emu(1100000), Emu(4792000)]
    n_rows = len(projects) + 1
    row_h = min(Emu(560000), Emu(int((H - content_top - Emu(500000)) / max(n_rows, 1))))
    table = add_table(slide, M, content_top, sum(col_widths, Emu(0)), row_h * n_rows, col_widths, n_rows)
    for c, htext in enumerate(headers):
        set_cell(table.cell(0, c), htext, size=8, bold=True, color=theme.WHITE, fill=theme.INK,
                  align=PP_ALIGN.CENTER if c != 0 and c != 3 else PP_ALIGN.LEFT)
    for r, p in enumerate(projects, start=1):
        fill = theme.SURFACE if r % 2 == 0 else theme.WHITE
        name = f"{p.initiative}\n{p.project}" if p.initiative else p.project
        set_cell(table.cell(r, 0), name, size=9, fill=fill)
        set_cell(table.cell(r, 1), f"{p.ipi:.2f}", align=PP_ALIGN.CENTER, fill=fill)
        set_cell(table.cell(r, 2), f"{p.ti:.2f}", align=PP_ALIGN.CENTER, fill=fill)
        set_cell(table.cell(r, 3), p.comments, size=8, fill=fill)

    finalize(slide, data.cover.footer, page_num)
    return slide


# --------------------------------------------------------------------------
# 10/11. NPS summary & detail
# --------------------------------------------------------------------------

def build_nps_summary(prs, data: DeckData, page_num: int):
    nps = data.nps_summary
    slide = new_slide(prs)
    content_top = add_header(slide, nps.headline or "Customer Experience — NPS", nps.subheadline)

    add_kpi_tile(slide, M, content_top, Emu(3200000), Emu(1100000), "Companywide NPS",
                 f"{fmt_num(nps.company_actual)}", sublabel=f"vs target {fmt_num(nps.company_target)}")

    left2 = M + Emu(3200000) + Emu(300000)
    col2_w = W - M - left2
    add_text(slide, left2, content_top, col2_w, Emu(260000), "NPS BY LINE OF BUSINESS — ACTUAL VS TARGET", size=9, bold=True, color=theme.MUTED)
    row_top = content_top + Emu(380000)
    row_h = Emu(600000)
    max_v = max([b.target for b in nps.by_lob] + [b.actual for b in nps.by_lob] + [1])
    for b in nps.by_lob:
        add_text(slide, left2, row_top, Emu(1600000), Emu(300000), b.lob, size=11, bold=True, color=theme.INK)
        gap = b.actual - b.target
        gap_color = theme.status_color("on_track") if gap >= 0 else theme.status_color("at_risk")
        add_text(slide, left2 + Emu(1600000), row_top, Emu(1600000), Emu(300000),
                  f"{fmt_num(b.actual)} / {fmt_num(b.target)}  ({'+' if gap >= 0 else ''}{fmt_num(gap)})",
                  size=10, color=gap_color, bold=True, align=PP_ALIGN.RIGHT)
        track_top = row_top + Emu(330000)
        track_w = col2_w
        add_rect(slide, left2, track_top, track_w, Emu(140000), fill=theme.SURFACE_ALT)
        fw = Emu(int(track_w * max(0, min(1, b.actual / max_v))))
        if fw > 0:
            add_rect(slide, left2, track_top, fw, Emu(140000), fill=theme.PRIMARY)
        target_x = left2 + Emu(int(track_w * max(0, min(1, b.target / max_v))))
        add_rect(slide, target_x, track_top - Emu(40000), Emu(12700), Emu(220000), fill=theme.INK)
        row_top += row_h

    finalize(slide, data.cover.footer, page_num)
    return slide


def build_nps_detail(prs, data: DeckData, page_num: int):
    npsd = data.nps_detail
    slide = new_slide(prs)
    content_top = add_header(slide, npsd.headline or "Customer Experience — NPS by Segment", npsd.narrative)

    headers = ["LINE OF BUSINESS", "SEGMENT", "ACTUAL", "TARGET", "GAP"]
    col_widths = [Emu(2400000), Emu(2600000), Emu(1500000), Emu(1500000), Emu(1700000)]
    n_rows = len(npsd.rows) + 1
    row_h = min(Emu(500000), Emu(int((H - content_top - Emu(500000)) / max(n_rows, 1))))
    table = add_table(slide, M, content_top, sum(col_widths, Emu(0)), row_h * n_rows, col_widths, n_rows)
    for c, htext in enumerate(headers):
        set_cell(table.cell(0, c), htext, size=9, bold=True, color=theme.WHITE, fill=theme.INK, align=PP_ALIGN.CENTER if c > 1 else PP_ALIGN.LEFT)
    for r, row in enumerate(npsd.rows, start=1):
        fill = theme.SURFACE if r % 2 == 0 else theme.WHITE
        gap = row.actual - row.target
        gap_color = theme.status_color("on_track") if gap >= 0 else theme.status_color("at_risk")
        set_cell(table.cell(r, 0), row.lob, bold=True, fill=fill)
        set_cell(table.cell(r, 1), row.segment, fill=fill)
        set_cell(table.cell(r, 2), fmt_num(row.actual), align=PP_ALIGN.CENTER, fill=fill)
        set_cell(table.cell(r, 3), fmt_num(row.target), align=PP_ALIGN.CENTER, fill=fill)
        set_cell(table.cell(r, 4), f"{'+' if gap >= 0 else ''}{fmt_num(gap)}", align=PP_ALIGN.CENTER, fill=fill, bold=True, color=gap_color)

    finalize(slide, data.cover.footer, page_num)
    return slide


# --------------------------------------------------------------------------
# 12. Recovery tracker (paginated table)
# --------------------------------------------------------------------------

def build_recovery_tracker(prs, data: DeckData, page_num: int) -> int:
    rt = data.recovery_tracker
    headers = ["#", "INITIATIVE / LOB", "TYPE", "BRI, SAR M", "STATUS", "IPI / TI", "ROOT CAUSE", "CORRECTIVE ACTION", "OWNER"]
    col_widths = [Emu(350000), Emu(1750000), Emu(750000), Emu(900000), Emu(850000), Emu(750000), Emu(2100000), Emu(2400000), Emu(1342000)]
    per_slide = 7
    items = rt.items
    chunks = [items[i:i + per_slide] for i in range(0, len(items), per_slide)] or [[]]
    for idx, chunk in enumerate(chunks):
        slide = new_slide(prs)
        suffix = f" (cont. {idx + 1}/{len(chunks)})" if len(chunks) > 1 else ""
        content_top = add_header(slide, (rt.headline or "EPMO Recovery Tracker") + suffix, rt.subheadline if idx == 0 else "")
        n_rows = len(chunk) + 1
        row_h = min(Emu(650000), Emu(int((H - content_top - Emu(500000)) / max(n_rows, 1))))
        table = add_table(slide, M, content_top, sum(col_widths, Emu(0)), row_h * n_rows, col_widths, n_rows)
        for c, htext in enumerate(headers):
            set_cell(table.cell(0, c), htext, size=8, bold=True, color=theme.WHITE, fill=theme.INK,
                      align=PP_ALIGN.CENTER if c not in (1, 6, 7) else PP_ALIGN.LEFT)
        for r, it in enumerate(chunk, start=1):
            fill = theme.SURFACE if r % 2 == 0 else theme.WHITE
            set_cell(table.cell(r, 0), it.rank, align=PP_ALIGN.CENTER, fill=fill)
            set_cell(table.cell(r, 1), f"{it.initiative}\n{it.lob}", size=9, bold=True, fill=fill)
            set_cell(table.cell(r, 2), it.item_type, size=8, align=PP_ALIGN.CENTER, fill=fill)
            set_cell(table.cell(r, 3), fmt_num(it.bri_m), align=PP_ALIGN.CENTER, bold=True, color=theme.PRIMARY, fill=fill)
            set_cell(table.cell(r, 4), theme.status_label(it.status), align=PP_ALIGN.CENTER, fill=theme.status_color(it.status), color=theme.WHITE, bold=True, size=8)
            set_cell(table.cell(r, 5), f"{it.ipi:.2f} / {it.ti:.2f}", size=8, align=PP_ALIGN.CENTER, fill=fill)
            set_cell(table.cell(r, 6), it.root_cause, size=8, fill=fill)
            set_cell(table.cell(r, 7), it.corrective_action, size=8, fill=fill)
            set_cell(table.cell(r, 8), it.owner, size=8, fill=fill)
        finalize(slide, data.cover.footer, page_num)
        page_num += 1
    return page_num


# --------------------------------------------------------------------------
# 13. Scenarios
# --------------------------------------------------------------------------

def build_scenarios(prs, data: DeckData, page_num: int):
    sc = data.scenarios
    slide = new_slide(prs)
    content_top = add_header(slide, sc.headline or "Probability of Achieving the Committed BRI", sc.subheadline)
    if sc.committed_m:
        add_text(slide, M, content_top, W - 2 * M, Emu(300000), f"Committed  SAR {fmt_num(sc.committed_m)}m", size=12, bold=True, color=theme.INK)
        content_top += Emu(420000)

    n = max(len(sc.items), 1)
    gap = Emu(200000)
    card_w = Emu(int((W - 2 * M - gap * (n - 1)) / n))
    card_h = Emu(3300000)
    x = M
    for item in sc.items:
        add_rect(slide, x, content_top, card_w, card_h, fill=theme.SURFACE_ALT)
        add_text(slide, x + Emu(150000), content_top + Emu(150000), card_w - Emu(300000), Emu(400000), item.name, size=13, bold=True, color=theme.INK)
        add_text(slide, x + Emu(150000), content_top + Emu(650000), card_w - Emu(300000), Emu(700000), f"SAR {fmt_num(item.value_m)}m", size=22, bold=True, color=theme.PRIMARY)
        add_text(slide, x + Emu(150000), content_top + Emu(1350000), card_w - Emu(300000), Emu(400000), f"{fmt_num(item.percent)}% of committed", size=11, color=theme.MUTED)
        add_text(slide, x + Emu(150000), content_top + Emu(1800000), card_w - Emu(300000), card_h - Emu(1950000), item.description, size=9, color=theme.INK_SOFT)
        x += card_w + gap

    finalize(slide, data.cover.footer, page_num)
    return slide


# --------------------------------------------------------------------------
# 14. Recommended actions
# --------------------------------------------------------------------------

def build_recommended_actions(prs, data: DeckData, page_num: int):
    ra = data.recommended_actions
    slide = new_slide(prs)
    content_top = add_header(slide, ra.headline or "Summary & Recommended Actions", ra.subheadline)

    n = max(len(ra.items), 1)
    cols = min(n, 3)
    rows = (n + cols - 1) // cols
    gap = Emu(180000)
    card_w = Emu(int((W - 2 * M - gap * (cols - 1)) / cols))
    card_h = Emu(int((H - content_top - Emu(500000) - gap * (rows - 1)) / rows))
    for i, item in enumerate(ra.items):
        r, c = divmod(i, cols)
        x = M + c * (card_w + gap)
        y = content_top + r * (card_h + gap)
        add_rect(slide, x, y, card_w, card_h, fill=theme.SURFACE_ALT)
        add_rect(slide, x, y, Emu(80000), card_h, fill=theme.PRIMARY)
        add_text(slide, x + Emu(220000), y + Emu(150000), card_w - Emu(370000), Emu(240000), f"{i + 1:02d}", size=11, bold=True, color=theme.PRIMARY)
        add_text(slide, x + Emu(220000), y + Emu(420000), card_w - Emu(370000), Emu(500000), item.title, size=12, bold=True, color=theme.INK)
        add_text(slide, x + Emu(220000), y + Emu(950000), card_w - Emu(370000), card_h - Emu(1100000), item.description, size=9, color=theme.INK_SOFT)

    finalize(slide, data.cover.footer, page_num)
    return slide


# --------------------------------------------------------------------------
# 15. Closing
# --------------------------------------------------------------------------

def build_closing(prs, data: DeckData, page_num: int):
    slide = new_dark_slide(prs)
    add_text(slide, M, Emu(2900000), W - 2 * M, Emu(800000), data.closing.message or "Thank you", size=32, color=theme.WHITE, bold=True)
    footer = data.closing.footer or data.cover.footer
    if footer:
        add_text(slide, M, Emu(3800000), W - 2 * M, Emu(350000), footer, size=12, color=theme.MUTED)
    add_text(slide, W - M - Emu(600000), H - Emu(500000), Emu(600000), Emu(300000), str(page_num), size=10, color=theme.MUTED, align=PP_ALIGN.RIGHT)
    return slide


# --------------------------------------------------------------------------
# Orchestration
# --------------------------------------------------------------------------

def build_deck(data: DeckData) -> Presentation:
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H

    page = 1
    build_cover(prs, data)
    page += 1

    if data.toc:
        build_toc(prs, data, page)
        page += 1

    def toc_title(idx, default):
        if idx < len(data.toc) and data.toc[idx].section:
            return data.toc[idx].section
        return default

    has_exec = any([
        data.exec_summary.committed_bri_m, data.exec_summary.ipi_tiers,
        data.exec_summary.benefit_status_by_lob, data.exec_summary.savings_breakdown,
    ])
    if has_exec:
        build_section_divider(prs, "01", toc_title(0, "Executive Summary"), "", data, page)
        page += 1
        build_exec_summary(prs, data, page)
        page += 1

    has_context = bool(data.context_gwp.years or data.context_gwp.by_lob)
    if has_context:
        build_section_divider(prs, "02", toc_title(1, "Context"), "", data, page)
        page += 1
        build_context_gwp(prs, data, page)
        page += 1

    has_ipi = bool(data.ipi_sector.items)
    has_lob_overview = bool(data.lob_overview.rows)
    if has_ipi or has_lob_overview:
        build_section_divider(prs, "03", toc_title(2, "Execution Signal"), "", data, page)
        page += 1
        if has_ipi:
            build_ipi_sector(prs, data, page)
            page += 1
        if has_lob_overview:
            build_lob_overview(prs, data, page)
            page += 1

    initiatives_by_lob = _grouped_by_lob(data.initiatives)
    projects_by_lob = _grouped_by_lob(data.projects)
    has_nps = bool(data.nps_summary.by_lob or data.nps_detail.rows)
    has_recovery = bool(data.recovery_tracker.items)

    if initiatives_by_lob or projects_by_lob or has_nps or has_recovery:
        build_section_divider(prs, "04", toc_title(3, "Line-of-Business Deep-Dives"), "", data, page)
        page += 1
        seen_lobs = list(initiatives_by_lob.keys())
        for lob in seen_lobs:
            build_initiative_slide(prs, lob, initiatives_by_lob[lob], data, page)
            page += 1
            if lob in projects_by_lob:
                build_project_slide(prs, lob, projects_by_lob[lob], data, page)
                page += 1
        for lob in projects_by_lob:
            if lob not in initiatives_by_lob:
                build_project_slide(prs, lob, projects_by_lob[lob], data, page)
                page += 1
        if data.nps_summary.by_lob:
            build_nps_summary(prs, data, page)
            page += 1
        if data.nps_detail.rows:
            build_nps_detail(prs, data, page)
            page += 1
        if has_recovery:
            page = build_recovery_tracker(prs, data, page)

    has_scenarios = bool(data.scenarios.items)
    has_actions = bool(data.recommended_actions.items)
    if has_scenarios or has_actions:
        build_section_divider(prs, "05", toc_title(4, "Summary & Recommended Actions"), "", data, page)
        page += 1
        if has_scenarios:
            build_scenarios(prs, data, page)
            page += 1
        if has_actions:
            build_recommended_actions(prs, data, page)
            page += 1

    build_closing(prs, data, page)
    return prs
