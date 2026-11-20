"""Render the same approved English report content to Word and PDF."""
from io import BytesIO
from xml.sax.saxutils import escape

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import LongTable, Paragraph, SimpleDocTemplate, Spacer, TableStyle

from .report_data import SECTIONS, display_quantity


TITLES = {"MC10": "Daily Site Report", "MC02": "Weekly Progress Report"}
HEADERS = ("Workfront", "Activity", "Subcontractor", "Unit", "Planned", "Reported", "Verified")


def content(payload, approved_by=""):
    dates = payload["start"] if payload["start"] == payload["end"] else f'{payload["start"]} to {payload["end"]}'
    blocks = [
        ("title", TITLES[payload["command"]]),
        ("body", f'{payload["command"]} | {dates} | Revision {payload["revision"]:02d} | Synthetic data'),
        ("body", f'Snapshot SHA-256: {payload["fingerprint"]}'),
        ("body", f'Prepared by {payload["author"]} | ' + (f"Approved by {approved_by}" if approved_by else "Draft for review")),
        ("heading", "Planned, reported and verified quantities"),
        ("body", "Period quantities come from dated site records. Different units stay separate. "
                 "A fraction shows how many records supplied that value. Zero is a recorded quantity."),
    ]
    for group in payload["comparison"]:
        blocks.append(("quantity", tuple(value or "Not recorded" for value in group["dimensions"])
                       + tuple(display_quantity(group["sums"][field], group["counts"][field], group["records"])
                               for field in ("planned", "reported", "verified"))))
    if not payload["comparison"]:
        blocks.append(("body", "No dated site records in this period."))
    blocks.append(("body", "Open items and current controls reflect the approved snapshot. Dated items after "
                   "the period end are excluded. A recorded release is displayed as register information."))
    for section, heading in SECTIONS.items():
        entries = payload["sections"][section]
        blocks.append(("heading", f"{heading} ({len(entries)})"))
        if not entries:
            blocks.append(("body", "No matching records."))
        for row in entries:
            fields = [row["id"], row["title"], row["status"], row["day"],
                      row["workfront"] if row["workfront"] != row["id"] else ""]
            headline = " | ".join(dict.fromkeys(value for value in fields if value))
            blocks.append(("record", f'{headline} [sheet {row["sheet"]}, row {row["row"]}]'))
            details = list(row["details"])
            if section == "progress":
                details += [(field.capitalize(), display_quantity(row[field], 1, 1))
                            for field in ("planned", "reported", "verified")]
                if row["unit"]:
                    details.append(("Unit", row["unit"]))
            if details:
                blocks.append(("body", " | ".join(f"{label}: {value}" for label, value in details)))
    return blocks


def _blocks(payload, approver):
    blocks = content(payload, approver)
    rows = [value for kind, value in blocks if kind == "quantity"]
    ready = []
    for kind, value in blocks:
        if kind == "quantity":
            if rows:
                ready.append(("table", [HEADERS, *rows]))
                rows = []
        else:
            ready.append((kind, value))
    return ready


def word(payload, approver):
    doc = Document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Inches(0.7)
    sec.top_margin = sec.bottom_margin = Inches(0.65)
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10)
    style.font.color.rgb = RGBColor.from_string("21354A")
    for kind, value in _blocks(payload, approver):
        if kind == "table":
            table = doc.add_table(rows=0, cols=len(HEADERS))
            table.style = "Light Shading Accent 1"
            for row in value:
                for cell, item in zip(table.add_row().cells, row):
                    cell.text = item
        elif kind == "title":
            doc.add_heading(value, 0)
        elif kind == "heading":
            doc.add_heading(value, 1)
        else:
            paragraph = doc.add_paragraph()
            paragraph.add_run(value).bold = kind == "record"
            paragraph.paragraph_format.keep_with_next = kind == "record"
    doc.core_properties.title = TITLES[payload["command"]]
    doc.core_properties.author = payload["author"]
    buffer = BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def pdf(payload, approver):
    styles = getSampleStyleSheet()
    for style in ("Title", "Heading2", "BodyText"):
        styles[style].textColor = colors.HexColor("#183448")
    styles["BodyText"].spaceAfter = 6
    stories = []
    for kind, value in _blocks(payload, approver):
        if kind == "table":
            rows = [[Paragraph(escape(item), styles["BodyText"]) for item in row] for row in value]
            table = LongTable(rows, colWidths=[102, 80, 77, 34, 74, 63, 65], repeatRows=1)
            table.setStyle(TableStyle([
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d6ecec")),
                ("LINEBELOW", (0, 0), (-1, 0), 0.6, colors.HexColor("#285965")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f0f5f7")]),
            ]))
            stories.extend([table, Spacer(1, 10)])
            continue
        label = escape(value)
        if kind == "record":
            label = f"<b>{label}</b>"
        style = {"title": "Title", "heading": "Heading2"}.get(kind, "BodyText")
        paragraph = Paragraph(label, styles[style])
        if kind == "record":
            paragraph.keepWithNext = True
        stories.append(paragraph)

    def footer(canvas, runner):
        canvas.setFont("Helvetica", 9)
        canvas.setFillColor(colors.HexColor("#526679"))
        canvas.drawString(45, 29, f'{payload["command"]}  |  Revision {payload["revision"]:02d}')
        canvas.drawRightString(A4[0] - 45, 29, f"Page {runner.page}")

    buffer = BytesIO()
    runner = SimpleDocTemplate(buffer, pagesize=A4, leftMargin=45, rightMargin=45,
                               topMargin=42, bottomMargin=45, title=TITLES[payload["command"]],
                               author=payload["author"])
    runner.build(stories, onFirstPage=footer, onLaterPages=footer)
    return buffer.getvalue()
