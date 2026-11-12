"""Builds samples/q3-review.pptx, the neutral demo deck used by the README screenshots.

Run from the repository root (python-pptx is the only dependency):

    uv run --with python-pptx python scripts/make_sample_deck.py

Shape names are meaningful on purpose: the canvas reports the shapes under a mark or comment to the agent.
"""

from datetime import datetime
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

OUT = Path(__file__).resolve().parent.parent / "samples" / "q3-review.pptx"

NAVY = "1B2B4B"
TEAL = "0F9D8A"
AMBER = "F5A524"
INK = "232B3A"
MUTED = "5B6475"
LINE = "DCE1EA"
CARD = "FFFFFF"
PAPER = "F4F6FA"
GREEN = "1A7F37"
HEAD_FONT = "Segoe UI Semibold"
BODY_FONT = "Segoe UI"


def rgb(value: str) -> RGBColor:
    return RGBColor.from_string(value)


def background(slide, color: str) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb(color)


def text(slide, name, x, y, w, h, value, size=16, color=INK, bold=False, font=BODY_FONT,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    box.name = name
    frame = box.text_frame
    frame.word_wrap = True
    frame.vertical_anchor = anchor
    frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = 0
    lines = value if isinstance(value, list) else [value]
    for index, line in enumerate(lines):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.alignment = align
        run = paragraph.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.name = font
        run.font.color.rgb = rgb(color)
    return box


def shape(slide, name, kind, x, y, w, h, fill, line=None, radius=None):
    item = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    item.name = name
    item.fill.solid()
    item.fill.fore_color.rgb = rgb(fill)
    if line:
        item.line.color.rgb = rgb(line)
        item.line.width = Pt(1)
    else:
        item.line.fill.background()
    item.shadow.inherit = False
    if radius is not None and kind == MSO_SHAPE.ROUNDED_RECTANGLE:
        item.adjustments[0] = radius
    return item


def card(slide, name, x, y, w, h):
    return shape(slide, name, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, CARD, LINE, radius=0.06)


def title(slide, value, color=NAVY):
    placeholder = slide.shapes.title
    placeholder.left, placeholder.top = Inches(0.7), Inches(0.45)
    placeholder.width, placeholder.height = Inches(11.9), Inches(0.9)
    frame = placeholder.text_frame
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    frame.margin_left = frame.margin_right = 0
    paragraph = frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.LEFT
    run = paragraph.add_run()
    run.text = value
    run.font.size = Pt(32)
    run.font.name = HEAD_FONT
    run.font.color.rgb = rgb(color)
    return placeholder


def accent_bar(slide):
    shape(slide, "accent-bar", MSO_SHAPE.RECTANGLE, 0.7, 0.32, 0.9, 0.07, TEAL)


def notes(slide, value: str) -> None:
    slide.notes_slide.notes_text_frame.text = value


def new_slide(prs, fill=PAPER):
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only
    background(slide, fill)
    return slide


def slide_cover(prs):
    slide = new_slide(prs, NAVY)
    for index, (x, y, d, color) in enumerate([(9.4, 0.9, 3.6, "24406E"), (10.6, 3.6, 2.4, TEAL), (8.6, 4.6, 1.5, AMBER)]):
        shape(slide, f"cover-circle-{index + 1}", MSO_SHAPE.OVAL, x, y, d, d, color)
    placeholder = title(slide, "Q3 product review", "FFFFFF")
    placeholder.left, placeholder.top, placeholder.width, placeholder.height = Inches(0.9), Inches(2.5), Inches(8.2), Inches(1.4)
    placeholder.text_frame.paragraphs[0].runs[0].font.size = Pt(54)
    text(slide, "cover-subtitle", 0.9, 4.0, 8.0, 0.5, "Lumen  ·  Growth team  ·  Draft for review", 20, "B9C6E0")
    shape(slide, "cover-accent", MSO_SHAPE.RECTANGLE, 0.9, 2.2, 1.1, 0.08, TEAL)
    notes(slide, "Welcome. This is a draft: please mark up anything that is unclear before Friday.")


def slide_kpis(prs):
    slide = new_slide(prs)
    accent_bar(slide)
    title(slide, "Q3 at a glance")
    kpis = [
        ("kpi-active-users", "Weekly active users", "128k", "+18% vs Q2", "Up in every month of the quarter."),
        ("kpi-retention", "Day-30 retention", "42%", "+3 pts vs Q2", "Driven by the new onboarding flow."),
        ("kpi-nps", "Net promoter score", "54", "+6 vs Q2", "Highest score since launch."),
    ]
    for index, (name, label, value, delta, note) in enumerate(kpis):
        x = 0.7 + index * 4.1
        card(slide, f"{name}-card", x, 1.75, 3.7, 3.55)
        text(slide, f"{name}-label", x + 0.35, 2.1, 3.0, 0.4, label.upper(), 12, MUTED, True)
        text(slide, f"{name}-value", x + 0.35, 2.6, 3.0, 1.2, value, 64, NAVY, False, HEAD_FONT)
        text(slide, f"{name}-delta", x + 0.35, 3.95, 3.0, 0.4, delta, 18, GREEN, True)
        text(slide, f"{name}-note", x + 0.35, 4.5, 3.0, 0.6, note, 14, MUTED)
    text(slide, "kpi-source", 0.7, 6.55, 11.9, 0.35, "Source: product analytics, July to September. Retention is measured per signup cohort.", 12, MUTED)
    notes(slide, "Three numbers to remember: 128k weekly actives, 42% day-30 retention, and an NPS of 54.")


def slide_chart(prs):
    slide = new_slide(prs)
    accent_bar(slide)
    title(slide, "Weekly active users grew every month")
    card(slide, "chart-card", 0.7, 1.7, 7.9, 4.9)
    data = CategoryChartData()
    data.categories = ["Apr", "May", "Jun", "Jul", "Aug", "Sep"]
    data.add_series("Weekly active users (k)", (92, 99, 108, 114, 121, 128))
    holder = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.95), Inches(1.95), Inches(7.4), Inches(4.4), data)
    holder.name = "chart-active-users"
    chart = holder.chart
    chart.has_legend = False
    chart.has_title = False
    chart.value_axis.visible = False
    chart.value_axis.has_major_gridlines = False
    chart.category_axis.tick_labels.font.size = Pt(14)
    chart.category_axis.tick_labels.font.name = BODY_FONT
    chart.category_axis.format.line.color.rgb = rgb(LINE)
    plot = chart.plots[0]
    plot.gap_width = 55
    plot.has_data_labels = True
    plot.data_labels.font.size = Pt(14)
    plot.data_labels.font.name = BODY_FONT
    plot.data_labels.number_format = '0"k"'
    plot.data_labels.number_format_is_linked = False
    plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    series = plot.series[0]
    series.format.fill.solid()
    series.format.fill.fore_color.rgb = rgb("B7C3DA")
    for index in (3, 4, 5):
        point = series.points[index]
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = rgb(TEAL)
    points = [
        ("takeaway-growth", "+18%", "Sep vs Jun: the strongest quarter this year."),
        ("takeaway-source", "62%", "of new users arrived through invites."),
        ("takeaway-risk", "Watch", "Server cost per active user is rising."),
    ]
    for index, (name, big, small) in enumerate(points):
        y = 1.7 + index * 1.7
        card(slide, f"{name}-card", 8.95, y, 3.65, 1.5)
        text(slide, f"{name}-big", 9.25, y + 0.2, 3.1, 0.55, big, 26, TEAL if index < 2 else AMBER, True, HEAD_FONT)
        text(slide, f"{name}-text", 9.25, y + 0.8, 3.1, 0.6, small, 14, MUTED)
    notes(slide, "Highlight: July to September are the teal bars. Growth came mostly from invites, not paid acquisition.")


def slide_shipped(prs):
    slide = new_slide(prs)
    accent_bar(slide)
    title(slide, "What we shipped")
    items = [
        ("ship-onboarding", "1", "Faster onboarding", "Setup dropped from 9 steps to 3. Median time to first project: 4 minutes."),
        ("ship-workspaces", "2", "Shared workspaces", "Teams can invite colleagues and share projects with one link."),
        ("ship-offline", "3", "Offline mode", "Edits sync automatically when the connection comes back."),
    ]
    for index, (name, number, heading, body) in enumerate(items):
        x = 0.7 + index * 4.1
        card(slide, f"{name}-card", x, 1.75, 3.7, 4.4)
        badge = shape(slide, f"{name}-badge", MSO_SHAPE.OVAL, x + 0.35, 2.15, 0.75, 0.75, TEAL)
        frame = badge.text_frame
        frame.margin_left = frame.margin_right = 0
        paragraph = frame.paragraphs[0]
        paragraph.alignment = PP_ALIGN.CENTER
        run = paragraph.add_run()
        run.text = number
        run.font.size = Pt(24)
        run.font.bold = True
        run.font.name = HEAD_FONT
        run.font.color.rgb = rgb("FFFFFF")
        frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        text(slide, f"{name}-heading", x + 0.35, 3.25, 3.0, 0.5, heading, 22, NAVY, False, HEAD_FONT)
        text(slide, f"{name}-body", x + 0.35, 3.9, 3.0, 1.9, body, 16, MUTED)
    notes(slide, "Each item maps to a roadmap commitment from Q2.")


def slide_roadmap(prs):
    slide = new_slide(prs)
    accent_bar(slide)
    title(slide, "Roadmap for the next four months")
    shape(slide, "roadmap-line", MSO_SHAPE.RECTANGLE, 1.0, 3.35, 11.3, 0.06, LINE)
    milestones = [
        ("Oct", "Mobile beta", "Invite-only beta for iOS and Android."),
        ("Nov", "Team plans", "Seats, roles and billing for teams."),
        ("Dec", "Public API", "REST API and webhooks for partners."),
        ("Jan", "Enterprise SSO", "SAML sign-in and audit logs."),
    ]
    for index, (month, heading, body) in enumerate(milestones):
        x = 1.0 + index * 3.0
        name = f"milestone-{month.lower()}"
        shape(slide, f"{name}-dot", MSO_SHAPE.OVAL, x, 3.15, 0.46, 0.46, TEAL if index == 0 else NAVY)
        text(slide, f"{name}-month", x, 2.4, 2.6, 0.5, month.upper(), 16, MUTED, True)
        text(slide, f"{name}-heading", x, 3.95, 2.6, 0.5, heading, 22, NAVY, False, HEAD_FONT)
        text(slide, f"{name}-body", x, 4.55, 2.6, 1.2, body, 16, MUTED)
    notes(slide, "Dates are targets, not commitments, until the pricing review on October 12.")


def slide_next(prs):
    slide = new_slide(prs)
    accent_bar(slide)
    title(slide, "Next steps")
    rows = [
        ("Finalize team pricing", "Mina", "Oct 12"),
        ("Publish the mobile beta invite page", "Alex", "Oct 18"),
        ("Draft the public API reference", "Sam", "Oct 31"),
        ("Plan the enterprise pilot with two customers", "Priya", "Nov 15"),
    ]
    for index, (task, owner, due) in enumerate(rows):
        y = 1.7 + index * 1.2
        name = f"step-{index + 1}"
        card(slide, f"{name}-card", 0.7, y, 11.9, 1.0)
        text(slide, f"{name}-number", 1.0, y, 0.6, 1.0, str(index + 1), 26, TEAL, True, HEAD_FONT, anchor=MSO_ANCHOR.MIDDLE)
        text(slide, f"{name}-task", 1.8, y, 6.8, 1.0, task, 20, INK, anchor=MSO_ANCHOR.MIDDLE)
        text(slide, f"{name}-owner", 8.9, y, 1.9, 1.0, owner, 18, MUTED, anchor=MSO_ANCHOR.MIDDLE)
        text(slide, f"{name}-due", 10.9, y, 1.4, 1.0, due, 18, NAVY, True, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT)
    notes(slide, "Owners confirm their dates in the review thread by Friday.")


def main() -> None:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    core = prs.core_properties
    core.title = "Q3 product review (sample)"
    core.author = core.last_modified_by = "Copilot PPTX Review Canvas"
    core.comments = "Neutral sample deck for the Copilot PPTX Review Canvas README."
    core.created = core.modified = datetime(2026, 11, 12)
    for build in (slide_cover, slide_kpis, slide_chart, slide_shipped, slide_roadmap, slide_next):
        build(prs)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(OUT)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    main()
