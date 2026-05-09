"""
formatter.py
────────────
Core Word document formatting engine.

Accepts a settings dict (from groq_parser.py or manual GUI input)
and applies all formatting rules to the document.

Features:
  - Page size  (A4 / Letter / Legal)
  - Page border with independent cm sizes per edge
  - Per-level fonts: unit heading / body / table header / table body / references
  - Groq run-level classification for mixed-format paragraphs
  - Italic book titles inside "double quotes"
  - Bold lecture-hour counts e.g. (9)
  - Table cell borders
  - Groq-cleaned reference text injection
"""

import re
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ══════════════════════════════════════════════════════════════════════════════
#  PAGE SIZE
# ══════════════════════════════════════════════════════════════════════════════

_PAGE_SIZES = {
    "a4":     (Inches(8.27),  Inches(11.69)),
    "letter": (Inches(8.5),   Inches(11.0)),
    "legal":  (Inches(8.5),   Inches(14.0)),
}

def apply_page_size(doc, size_name: str):
    key = size_name.lower()
    if key not in _PAGE_SIZES:
        return
    width, height = _PAGE_SIZES[key]
    for section in doc.sections:
        section.page_width  = width
        section.page_height = height


# ══════════════════════════════════════════════════════════════════════════════
#  PAGE BORDER  (independent cm size per edge)
# ══════════════════════════════════════════════════════════════════════════════

_PAGE_BORDER_MAP = {
    "single": "single",
    "double": "double",
    "shadow": "shadow",
    "thick":  "thick",
}

# 1 cm = 567 twentieths of a point (Word's internal unit for border spacing)
_CM_TO_TWENTIETHS = 567

def apply_page_border(doc, style: str, color: str, sizes: dict):
    """
    sizes dict keys: top, bottom, left, right  (values in cm as float)
    """
    if style == "none":
        return

    w_val = _PAGE_BORDER_MAP.get(style, "single")
    color = color.lstrip("#").upper() or "000000"

    for section in doc.sections:
        sectPr = section._sectPr

        for old in list(sectPr.findall(qn("w:pgBorders"))):
            sectPr.remove(old)

        pgBorders = OxmlElement("w:pgBorders")
        pgBorders.set(qn("w:offsetFrom"), "page")

        for edge in ("top", "left", "bottom", "right"):
            cm_val = float(sizes.get(edge, 1.0))
            space  = str(int(cm_val * _CM_TO_TWENTIETHS))

            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:val"),   w_val)
            el.set(qn("w:sz"),    "18")
            el.set(qn("w:space"), space)
            el.set(qn("w:color"), color)
            pgBorders.append(el)

        pgSz = sectPr.find(qn("w:pgSz"))
        if pgSz is not None:
            sectPr.insert(list(sectPr).index(pgSz), pgBorders)
        else:
            sectPr.append(pgBorders)


# ══════════════════════════════════════════════════════════════════════════════
#  TABLE CELL BORDERS
# ══════════════════════════════════════════════════════════════════════════════

def _apply_cell_borders(cell, style: str, color: str):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for old in list(tcPr.findall(qn("w:tcBorders"))):
        tcPr.remove(old)

    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        tag = OxmlElement(f"w:{edge}")
        tag.set(qn("w:val"),   style)
        tag.set(qn("w:sz"),    "4")
        tag.set(qn("w:color"), color.lstrip("#"))
        tcBorders.append(tag)
    tcPr.append(tcBorders)

def apply_table_borders(table, style: str, color: str):
    seen = set()
    for row in table.rows:
        for cell in row.cells:
            if id(cell) not in seen:
                seen.add(id(cell))
                _apply_cell_borders(cell, style, color)


# ══════════════════════════════════════════════════════════════════════════════
#  ITALIC BOOK TITLES
# ══════════════════════════════════════════════════════════════════════════════

def italicise_quoted_titles(paragraph, font_name: str, font_pt):
    full_text = paragraph.text
    if '"' not in full_text:
        return

    p = paragraph._p
    for r in list(p.findall(qn("w:r"))):
        p.remove(r)

    parts = re.split(r'("(?:[^"]*)")', full_text)
    for part in parts:
        if not part:
            continue
        run = paragraph.add_run(part)
        run.font.name = font_name
        run.font.size = font_pt
        run.italic    = part.startswith('"') and part.endswith('"')


# ══════════════════════════════════════════════════════════════════════════════
#  RUN FONT HELPER
# ══════════════════════════════════════════════════════════════════════════════

def _set_run_font(run, font_name: str, font_pt, bold=None):
    run.font.name = font_name
    run.font.size = font_pt
    if bold is not None:
        run.bold = bold


# ══════════════════════════════════════════════════════════════════════════════
#  PARAGRAPH CLASSIFIERS
# ══════════════════════════════════════════════════════════════════════════════

_UNIT_HEADING_RE  = re.compile(r"^[A-Z][A-Z\s,/]+:\s")
_LECTURE_HOURS_RE = re.compile(r"\(\d+\)\s*$")
_REF_KEYS         = ["TEXT BOOKS", "TEXTBOOKS", "REFERENCES"]
_END_KEYS         = ["COURSE OUTCOMES", "MAPPING COURSE", "ARTICULATION", "TOTAL L"]

def _is_unit_heading(para) -> bool:
    return bool(_UNIT_HEADING_RE.match(para.text.strip()))

def _is_ref_start(upper: str) -> bool:
    return any(k in upper for k in _REF_KEYS)

def _is_ref_end(upper: str) -> bool:
    return any(k in upper for k in _END_KEYS)

def _has_lecture_hours(para) -> bool:
    return bool(_LECTURE_HOURS_RE.search(para.text))


# ══════════════════════════════════════════════════════════════════════════════
#  REFERENCE EXTRACTOR
# ══════════════════════════════════════════════════════════════════════════════

def extract_references(docx_path: str) -> list:
    doc = Document(docx_path)
    refs, in_ref = [], False
    for para in doc.paragraphs:
        text  = para.text.strip()
        upper = text.upper()
        if _is_ref_start(upper):
            in_ref = True
            continue
        if in_ref:
            if _is_ref_end(upper):
                in_ref = False
                continue
            if text and len(text) > 10:
                refs.append(text)
    return refs


# ══════════════════════════════════════════════════════════════════════════════
#  GROQ RUN-LEVEL FORMATTING
# ══════════════════════════════════════════════════════════════════════════════

def apply_groq_run_classification(para, classified_runs: list, settings: dict):
    """
    Rebuild paragraph runs using Groq's classification.
    Each run gets the font/size matching its detected type.
    """
    uh_font  = settings.get("unit_heading_font",  "Times New Roman")
    uh_pt    = Pt(settings.get("unit_heading_size", 12))
    body_font = settings.get("body_font",   "Times New Roman")
    body_pt   = Pt(settings.get("body_size", 11))
    ref_font  = settings.get("reference_font",  "Times New Roman")
    ref_pt    = Pt(settings.get("reference_size", 11))
    th_font   = settings.get("table_header_font", "Arial")
    th_pt     = Pt(settings.get("table_header_size", 10))

    _TYPE_MAP = {
        "unit_heading":  (uh_font,   uh_pt,   True),
        "section_label": (uh_font,   uh_pt,   True),
        "body":          (body_font, body_pt, None),
        "lecture_hours": (body_font, body_pt, True),
        "reference":     (ref_font,  ref_pt,  None),
        "table_header":  (th_font,   th_pt,   True),
        "table_body":    (th_font,   th_pt,   None),
    }

    # Remove all runs and rebuild from classification
    p = para._p
    for r in list(p.findall(qn("w:r"))):
        p.remove(r)

    for cr in classified_runs:
        run_type = cr.get("type", "body")
        font_name, font_pt, bold = _TYPE_MAP.get(run_type, (body_font, body_pt, None))
        run = para.add_run(cr.get("text", ""))
        _set_run_font(run, font_name, font_pt, bold)
        if cr.get("italic"):
            run.italic = True


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN FORMATTER
# ══════════════════════════════════════════════════════════════════════════════

def format_document(
    input_path:        str,
    output_path:       str,
    settings:          dict,
    cleaned_refs:      list = None,
    groq_api_key:      str  = None,
    use_run_classify:  bool = False,
):
    """
    Apply all formatting to a Word document and save.

    settings keys (all optional — defaults used if missing):
      page_size, page_border_style, page_border_color
      page_border_top, page_border_bottom, page_border_left, page_border_right (cm)
      unit_heading_font, unit_heading_size
      body_font, body_size
      table_header_font, table_header_size
      table_body_font, table_body_size
      reference_font, reference_size
      alignment, bold_headings, italic_titles, bold_hours
      cell_border_style, cell_border_color

    groq_api_key      : if provided + use_run_classify=True, Groq classifies
                        each paragraph's runs for precise mixed formatting
    """
    from groq_parser import classify_paragraph_runs

    doc = Document(input_path)

    # ── Page size ─────────────────────────────────────────────────────────────
    apply_page_size(doc, settings.get("page_size", "A4"))

    # ── Page border with per-edge cm sizes ────────────────────────────────────
    apply_page_border(
        doc,
        settings.get("page_border_style", "none"),
        settings.get("page_border_color", "000000"),
        {
            "top":    settings.get("page_border_top",    1.0),
            "bottom": settings.get("page_border_bottom", 1.0),
            "left":   settings.get("page_border_left",   1.0),
            "right":  settings.get("page_border_right",  1.0),
        },
    )

    # ── Font / size objects ───────────────────────────────────────────────────
    uh_font  = settings.get("unit_heading_font",  "Times New Roman")
    uh_pt    = Pt(settings.get("unit_heading_size", 12))
    body_font = settings.get("body_font",   "Times New Roman")
    body_pt   = Pt(settings.get("body_size", 11))
    th_font   = settings.get("table_header_font", "Arial")
    th_pt     = Pt(settings.get("table_header_size", 10))
    tb_font   = settings.get("table_body_font",   "Arial")
    tb_pt     = Pt(settings.get("table_body_size", 10))
    ref_font  = settings.get("reference_font",  "Times New Roman")
    ref_pt    = Pt(settings.get("reference_size", 11))

    # ── Alignment ─────────────────────────────────────────────────────────────
    _align_map = {
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "left":    WD_ALIGN_PARAGRAPH.LEFT,
        "center":  WD_ALIGN_PARAGRAPH.CENTER,
        "right":   WD_ALIGN_PARAGRAPH.RIGHT,
    }
    body_align = _align_map.get(
        settings.get("alignment", "justify"),
        WD_ALIGN_PARAGRAPH.JUSTIFY,
    )

    in_ref_section = False
    ref_idx        = 0

    # ── Process paragraphs ────────────────────────────────────────────────────
    for para in doc.paragraphs:
        text  = para.text.strip()
        upper = text.upper()

        if _is_ref_start(upper):
            in_ref_section = True
        elif _is_ref_end(upper):
            in_ref_section = False

        if not text:
            continue

        # ── Groq run-level classification (optional) ──────────────────────
        if use_run_classify and groq_api_key:
            runs_meta = [
                {
                    "text":   r.text,
                    "bold":   bool(r.bold),
                    "italic": bool(r.italic),
                }
                for r in para.runs if r.text
            ]
            if runs_meta:
                classified = classify_paragraph_runs(runs_meta, groq_api_key)
                apply_groq_run_classification(para, classified, settings)

                # Still handle lecture-hour bolding
                if settings.get("bold_hours", True) and _has_lecture_hours(para):
                    for run in para.runs:
                        if _LECTURE_HOURS_RE.search(run.text):
                            run.bold = True
                continue   # skip manual rules below — Groq handled it

        # ── Manual rule-based classification ──────────────────────────────
        if _is_unit_heading(para):
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for run in para.runs:
                if run.bold or run.text.strip().isupper():
                    _set_run_font(run, uh_font, uh_pt,
                                  bold=settings.get("bold_headings", True))
                else:
                    _set_run_font(run, body_font, body_pt)
            if settings.get("bold_hours", True) and _has_lecture_hours(para):
                for run in para.runs:
                    if _LECTURE_HOURS_RE.search(run.text):
                        run.bold = True

        elif in_ref_section:
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            is_hdr = any(k in upper for k in _REF_KEYS)

            if cleaned_refs and not is_hdr and ref_idx < len(cleaned_refs):
                cleaned_text = cleaned_refs[ref_idx]
                ref_idx += 1
                p = para._p
                for r in list(p.findall(qn("w:r"))):
                    p.remove(r)
                run = para.add_run(cleaned_text)
                _set_run_font(run, ref_font, ref_pt)

            for run in para.runs:
                _set_run_font(run, ref_font, ref_pt)

            if settings.get("italic_titles", True):
                italicise_quoted_titles(para, ref_font, ref_pt)

        else:
            para.alignment = body_align
            for run in para.runs:
                _set_run_font(run, body_font, body_pt)

    # ── Tables ────────────────────────────────────────────────────────────────
    cell_border_style = settings.get("cell_border_style", "single")
    cell_border_color = settings.get("cell_border_color", "000000")

    for table in doc.tables:
        if cell_border_style != "none":
            apply_table_borders(table, cell_border_style, cell_border_color)

        for row_idx, row in enumerate(table.rows):
            is_header = (row_idx == 0)
            font_name = th_font if is_header else tb_font
            font_pt   = th_pt   if is_header else tb_pt
            seen_cells = set()
            for cell in row.cells:
                if id(cell) in seen_cells:
                    continue
                seen_cells.add(id(cell))
                for cell_para in cell.paragraphs:
                    for run in cell_para.runs:
                        _set_run_font(run, font_name, font_pt,
                                      bold=True if is_header else run.bold)

    doc.save(output_path)
    return True
