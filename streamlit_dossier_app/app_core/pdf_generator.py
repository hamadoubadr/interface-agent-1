from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm, cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, Image, KeepTogether, PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import Flowable
from reportlab.pdfgen import canvas
import os
from pathlib import Path
from io import BytesIO
import re

# ── Brand colours ──────────────────────────────────────────────────────────────
NAVY   = colors.HexColor("#1B2A6B")   # dark blue
ORANGE = colors.HexColor("#E84C0E")   # vivid orange
LIGHT  = colors.HexColor("#F5F7FA")   # very light grey background
WHITE  = colors.white
GREY   = colors.HexColor("#6B7280")
DARK   = colors.HexColor("#1F2937")

PAGE_W, PAGE_H = A4
MARGIN_L = MARGIN_R = 2.0 * cm
MARGIN_T = 1.5 * cm
MARGIN_B = 2.0 * cm

LOGO_PATH = "logo_gidnai.png"  # Real logo PNG file

# ── Custom Flowables ────────────────────────────────────────────────────────────
class OrangeRule(Flowable):
    """A thick orange horizontal rule with an optional left-accent block."""
    def __init__(self, width, height=3, accent=True):
        Flowable.__init__(self)
        self.width  = width
        self.height = height
        self.accent = accent

    def draw(self):
        c = self.canv
        c.setFillColor(ORANGE)
        c.rect(0, 0, self.width, self.height, fill=1, stroke=0)
        if self.accent:
            c.setFillColor(NAVY)
            c.rect(0, 0, 8, self.height, fill=1, stroke=0)


class SectionHeader(Flowable):
    """Pill-shaped section title with navy background."""
    def __init__(self, text, width, font_size=12):
        Flowable.__init__(self)
        self.text      = text
        self.width     = width
        self.font_size = font_size
        self.height    = font_size + 14

    def draw(self):
        c = self.canv
        r = self.height / 2
        # Rounded rect (navy)
        c.setFillColor(NAVY)
        c.roundRect(0, 0, self.width, self.height, r, fill=1, stroke=0)
        # Orange left accent stripe
        c.setFillColor(ORANGE)
        c.roundRect(0, 0, 6, self.height, 3, fill=1, stroke=0)
        # Text
        c.setFillColor(WHITE)
        c.setFont("Helvetica-Bold", self.font_size)
        c.drawString(16, (self.height - self.font_size) / 2 + 2, self.text)


class FooterCanvas(canvas.Canvas):
    """Draws a branded header and footer on every page."""
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._saved_page_states)
        for i, state in enumerate(self._saved_page_states):
            self.__dict__.update(state)
            self._draw_chrome(i + 1, total)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def _draw_chrome(self, page_num, total):
        self.saveState()
        w, h = A4

        # ── Top bar ────────────────────────────────────────────────────────────
        self.setFillColor(NAVY)
        self.rect(0, h - 18*mm, w, 18*mm, fill=1, stroke=0)
        # Orange accent line below bar
        self.setFillColor(ORANGE)
        self.rect(0, h - 18*mm - 3, w, 3, fill=1, stroke=0)
        # Logo in header
        logo_path = Path(__file__).parent.parent / "assets" / LOGO_PATH
        if logo_path.exists():
            logo_w, logo_h = 55*mm, 13*mm
            self.drawImage(str(logo_path), MARGIN_L, h - 18*mm + 2.5*mm,
                           width=logo_w, height=logo_h, preserveAspectRatio=True, mask='auto')
        # Header right text
        self.setFillColor(WHITE)
        self.setFont("Helvetica", 7.5)
        self.drawRightString(w - MARGIN_R, h - 18*mm + 8*mm, "DCE — LOT FLUIDES — IND C")
        self.setFont("Helvetica-Bold", 7.5)
        self.setFillColor(ORANGE)
        self.drawRightString(w - MARGIN_R, h - 18*mm + 3*mm, "GIDNAI AI SYSTEMS")

        # ── Bottom bar ─────────────────────────────────────────────────────────
        bar_h = 12*mm
        self.setFillColor(NAVY)
        self.rect(0, 0, w, bar_h, fill=1, stroke=0)
        self.setFillColor(ORANGE)
        self.rect(0, bar_h, w, 2, fill=1, stroke=0)
        # Page number
        self.setFillColor(WHITE)
        self.setFont("Helvetica", 8)
        self.drawCentredString(w / 2, bar_h / 2 - 3, f"Page {page_num} / {total}")
        # Left footer text
        self.setFont("Helvetica-Oblique", 7)
        self.setFillColor(colors.HexColor("#9CA3AF"))
        self.drawString(MARGIN_L, bar_h / 2 - 3, "Identité du Projet — Usage Confidentiel")
        # Right footer
        self.drawRightString(w - MARGIN_R, bar_h / 2 - 3, "www.gidnai.ai")

        self.restoreState()


# ── Styles ──────────────────────────────────────────────────────────────────────
def make_styles():
    base = getSampleStyleSheet()

    def s(name, **kw):
        return ParagraphStyle(name, **kw)

    styles = {
        "cover_title": s("cover_title",
            fontName="Helvetica-Bold", fontSize=22,
            textColor=WHITE, alignment=TA_CENTER, spaceAfter=4),
        "cover_sub": s("cover_sub",
            fontName="Helvetica", fontSize=11,
            textColor=ORANGE, alignment=TA_CENTER, spaceAfter=2),
        "cover_meta": s("cover_meta",
            fontName="Helvetica-Oblique", fontSize=9,
            textColor=colors.HexColor("#D1D5DB"), alignment=TA_CENTER),
        "section_intro": s("section_intro",
            fontName="Helvetica-Oblique", fontSize=9.5,
            textColor=GREY, leading=14, spaceAfter=8),
        "body": s("body",
            fontName="Helvetica", fontSize=9.5,
            textColor=DARK, leading=15, spaceAfter=4),
        "bullet": s("bullet",
            fontName="Helvetica", fontSize=9.5,
            textColor=DARK, leading=15, leftIndent=14,
            bulletIndent=4, spaceAfter=3),
        "label": s("label",
            fontName="Helvetica-Bold", fontSize=9,
            textColor=NAVY),
        "value": s("value",
            fontName="Helvetica", fontSize=9,
            textColor=DARK),
        "tag": s("tag",
            fontName="Helvetica-Bold", fontSize=8,
            textColor=WHITE),
    }
    return styles


# ── Cover page ─────────────────────────────────────────────────────────────────
class CoverPage(Flowable):
    def __init__(self, width, height, title="IDENTITÉ DU PROJET", subtitle="PROJET DE CONSTRUCTION — SIÈGE SOCIAL"):
        Flowable.__init__(self)
        self.width    = width
        self.height   = height
        self.title    = title
        self.subtitle = subtitle

    def draw(self):
        c = self.canv
        w, h = self.width, self.height

        # Background gradient-like blocks
        c.setFillColor(NAVY)
        c.rect(0, h * 0.38, w, h * 0.62, fill=1, stroke=0)
        c.setFillColor(colors.HexColor("#141f52"))
        c.rect(0, 0, w, h * 0.38, fill=1, stroke=0)

        # Decorative orange bar
        c.setFillColor(ORANGE)
        c.rect(0, h * 0.38 - 5, w, 8, fill=1, stroke=0)

        # Subtle circuit-like decoration (right side)
        c.setStrokeColor(colors.HexColor("#ffffff22"))
        c.setLineWidth(1)
        for i in range(6):
            y = h * 0.55 + i * 28
            c.line(w * 0.55, y, w * 0.9, y)
            c.circle(w * 0.9, y, 3, fill=0, stroke=1)

        # Logo
        logo_path = Path(__file__).parent.parent / "assets" / LOGO_PATH
        if logo_path.exists():
            lw, lh = 100*mm, 25*mm
            c.drawImage(str(logo_path), (w - lw) / 2, h * 0.62,
                        width=lw, height=lh, preserveAspectRatio=True, mask='auto')

        # Title block
        c.setFillColor(WHITE)
        c.setFont("Helvetica-Bold", 20)
        c.drawCentredString(w / 2, h * 0.52, self.title)
        c.setFillColor(ORANGE)
        c.setFont("Helvetica-Bold", 11)
        c.drawCentredString(w / 2, h * 0.48, self.subtitle)

        # Divider
        c.setFillColor(ORANGE)
        c.rect(w * 0.3, h * 0.455, w * 0.4, 2, fill=1, stroke=0)

        # Meta info boxes
        metas = [
            ("LOT", "FLUIDES — IND C"),
            ("LOCALISATION", "Bd Corniche / Av. de Nice"),
            ("SHOB", "1 973 m²"),
        ]
        box_w = (w - 2 * MARGIN_L) / len(metas)
        bx = MARGIN_L
        by = h * 0.30
        for label, val in metas:
            c.setFillColor(colors.HexColor("#ffffff18"))
            c.roundRect(bx + 4, by, box_w - 8, 38, 5, fill=1, stroke=0)
            c.setFillColor(ORANGE)
            c.setFont("Helvetica-Bold", 7)
            c.drawCentredString(bx + box_w / 2, by + 24, label)
            c.setFillColor(WHITE)
            c.setFont("Helvetica-Bold", 8.5)
            c.drawCentredString(bx + box_w / 2, by + 10, val)
            bx += box_w

        # Bottom tagline
        c.setFillColor(colors.HexColor("#9CA3AF"))
        c.setFont("Helvetica-Oblique", 8)
        c.drawCentredString(w / 2, h * 0.21, "Document confidentiel — Usage interne")

        # Bottom orange stripe
        c.setFillColor(ORANGE)
        c.rect(0, 0, w, 6, fill=1, stroke=0)


# ── Main builder functions ──────────────────────────────────────────────────────
def markdown_to_pdf_bytes(markdown_text: str, title: str = "Document", document_type: str = "identite") -> bytes:
    """Convert Markdown text to PDF bytes using the professional GIDNAI template."""
    def _clean_inline_markdown(text: str) -> str:
        # Remove accidental inline markdown headings like "# Identité du Projet"
        # when they appear inside bullets or plain paragraphs.
        return re.sub(r"^\s*#+\s*", "", text).strip()

    buffer = BytesIO()
    usable_w = PAGE_W - MARGIN_L - MARGIN_R
    ST = make_styles()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=MARGIN_L, rightMargin=MARGIN_R,
        topMargin=MARGIN_T + 18*mm,   # leave room for header bar
        bottomMargin=MARGIN_B + 12*mm, # leave room for footer bar
        title=title,
        author="GIDNAI AI Systems",
    )

    story = []

    # Parse markdown content
    lines = markdown_text.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line or line.startswith('# '):
            i += 1
            continue
            
        # Handle tables
        if line.startswith('|'):
            table_data = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                row_line = lines[i].strip()
                if '---' in row_line: # skip separator
                    i += 1
                    continue
                # Split cells and clean
                cells = [c.strip() for c in row_line.split('|') if c.strip() or (row_line.startswith('|') and row_line.endswith('|'))]
                # Filter out empty strings from start/end if they exist
                if row_line.startswith('|'): cells = cells[1:]
                if row_line.endswith('|'): cells = cells[:-1]
                
                if cells:
                    table_data.append([Paragraph(c, ST["value"]) for c in cells])
                i += 1
            
            if table_data:
                # Determine column widths
                num_cols = len(table_data[0])
                if num_cols == 2:
                    col_widths = [usable_w * 0.4, usable_w * 0.6]
                else:
                    col_widths = [usable_w / num_cols] * num_cols
                
                t = Table(table_data, colWidths=col_widths)
                t.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                    ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
                    ("GRID", (0, 0), (-1, -1), 0.5, GREY),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [LIGHT, WHITE]),
                    ("TOPPADDING", (0, 0), (-1, -1), 6),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ]))
                story.append(t)
                story.append(Spacer(1, 12))
            continue

        # Handle headers
        if line.startswith('## '):
            section_title = line[3:].strip().upper()
            story.append(SectionHeader(section_title, usable_w, font_size=11))
            story.append(Spacer(1, 6))
            story.append(OrangeRule(usable_w, height=2, accent=False))
            story.append(Spacer(1, 10))
            i += 1
            
        elif line.startswith('### '):
            sub_title = line[4:].strip()
            story.append(Paragraph(f"<b>{sub_title}</b>", ST["label"]))
            story.append(Spacer(1, 4))
            i += 1
            
        elif line.startswith('- ') or line.startswith('* '):
            item_text = _clean_inline_markdown(line[2:].strip())
            # Check for "Label : Value" pattern
            if " : " in item_text:
                label, val = item_text.split(" : ", 1)
                story.append(Paragraph(f"<font color='#1B2A6B'><b>{label} :</b></font> {val}", ST["body"]))
            else:
                story.append(Paragraph(f"• {item_text}", ST["body"]))
            story.append(Spacer(1, 4))
            i += 1
            
        else:
            story.append(Paragraph(_clean_inline_markdown(line), ST["body"]))
            story.append(Spacer(1, 8))
            i += 1

    def on_first_page(c, doc):
        # Same rendering as later pages (header/footer handled by FooterCanvas).
        pass

    def on_later_pages(c, doc):
        # Just the header/footer chrome (handled by FooterCanvas)
        pass

    doc.build(story, 
              canvasmaker=FooterCanvas, 
              onFirstPage=on_first_page, 
              onLaterPages=on_later_pages)
    
    buffer.seek(0)
    return buffer.getvalue()

def generate_identite_pdf(markdown_content: str, output_path: str | Path) -> None:
    """Generate a professional PDF from markdown content for identity documents."""
    pdf_bytes = markdown_to_pdf_bytes(markdown_content, "Identité du Projet", "identite")
    with open(output_path, 'wb') as f:
        f.write(pdf_bytes)

def generate_synthese_pdf(markdown_content: str, output_path: str | Path) -> None:
    """Generate a professional PDF from markdown content for detailed synthesis."""
    pdf_bytes = markdown_to_pdf_bytes(markdown_content, "Synthèse Détaillée", "synthese")
    with open(output_path, 'wb') as f:
        f.write(pdf_bytes)
