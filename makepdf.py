from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from pathlib import Path
import os, sys

# Read the .txt files and convert to real PDFs
txt_dir = Path.home() / "E:\\aresmgmt\\material\\data\\input\\pdf_reports"
pdf_dir = Path.home() / "E:\\aresmgmt\\material\\data\\input\\pdf_reports"

NAVY  = colors.HexColor("#1A2E4A")
TEAL  = colors.HexColor("#0D7C66")
GOLD  = colors.HexColor("#F0A500")
LGREY = colors.HexColor("#F7F9FB")

styles = getSampleStyleSheet()

def make_pdf(txt_path: Path):
    pdf_path = txt_path.with_suffix(".pdf")
    text = txt_path.read_text(encoding="utf-8")

    doc = SimpleDocTemplate(str(pdf_path), pagesize=A4,
          leftMargin=20*mm, rightMargin=20*mm, topMargin=15*mm, bottomMargin=15*mm)
    story = []

    # Header band (simulated with a table)
    hdr_data = [["FUND ADMINISTRATOR SERVICES LTD", "NET ASSET VALUE STATEMENT — OFFICIAL"]]
    hdr_tbl = Table(hdr_data, colWidths=[90*mm, 90*mm])
    hdr_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), NAVY),
        ("TEXTCOLOR",  (0,0), (-1,-1), colors.white),
        ("FONTNAME",   (0,0), (-1,-1), "Helvetica-Bold"),
        ("FONTSIZE",   (0,0), (-1,-1), 10),
        ("ALIGN",      (0,0), (0,0),   "LEFT"),
        ("ALIGN",      (1,0), (1,0),   "RIGHT"),
        ("PADDING",    (0,0), (-1,-1), 6),
    ]))
    story.append(hdr_tbl)
    story.append(Spacer(1, 6*mm))

    # Parse sections from the text
    lines = [l.rstrip() for l in text.splitlines()]
    in_section = None
    section_lines = []

    for line in lines:
        # Skip decorator lines
        if set(line.strip()) <= {"=", "-", " "} and len(line.strip()) > 5:
            if section_lines and in_section:
                _flush_section(story, in_section, section_lines)
                section_lines = []
            continue
        if line.strip().startswith("SECTION"):
            in_section = line.strip()
            section_lines = []
        elif line.strip():
            section_lines.append(line)

    if section_lines and in_section:
        _flush_section(story, in_section, section_lines)

    # Confidentiality footer
    story.append(Spacer(1, 8*mm))
    story.append(HRFlowable(width="100%", thickness=1, color=NAVY))
    story.append(Spacer(1, 2*mm))
    conf = ParagraphStyle("conf", fontName="Helvetica-Oblique", fontSize=8,
                          textColor=colors.grey, alignment=TA_CENTER)
    story.append(Paragraph("CONFIDENTIAL — FOR AUTHORISED RECIPIENTS ONLY", conf))

    doc.build(story)
    print(f"  ✓ {pdf_path.name}")

def _flush_section(story, section_name, lines):
    # Section header
    sec_data = [[section_name]]
    sec_tbl = Table(sec_data, colWidths=[180*mm])
    sec_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), TEAL),
        ("TEXTCOLOR",  (0,0), (-1,-1), colors.white),
        ("FONTNAME",   (0,0), (-1,-1), "Helvetica-Bold"),
        ("FONTSIZE",   (0,0), (-1,-1), 9),
        ("PADDING",    (0,0), (-1,-1), 4),
    ]))
    story.append(sec_tbl)
    story.append(Spacer(1, 3*mm))

    # Section content as two-column key:value table
    rows = []
    for line in lines:
        if ":" in line:
            parts = line.split(":", 1)
            label = parts[0].strip()
            value = parts[1].strip() if len(parts) > 1 else ""
            if label:
                rows.append([label, value])
        elif line.strip():
            rows.append([line.strip(), ""])

    if rows:
        tbl = Table(rows, colWidths=[90*mm, 90*mm])
        style = [
            ("FONTNAME",  (0,0), (0,-1), "Helvetica-Bold"),
            ("FONTNAME",  (1,0), (1,-1), "Helvetica"),
            ("FONTSIZE",  (0,0), (-1,-1), 9),
            ("TEXTCOLOR", (0,0), (0,-1), NAVY),
            ("PADDING",   (0,0), (-1,-1), 3),
            ("ROWBACKGROUNDS", (0,0), (-1,-1), [colors.white, LGREY]),
            ("LINEBELOW", (0,0), (-1,-2), 0.25, colors.HexColor("#E2E8F0")),
        ]
        tbl.setStyle(TableStyle(style))
        story.append(tbl)
    story.append(Spacer(1, 5*mm))

txt_files = sorted(txt_dir.glob("*.txt"))
print(f"Converting {len(txt_files)} text files to PDF...")
for f in txt_files:
    make_pdf(f)
print("Done.")