"""
groq_parser.py
──────────────
Sends the user's natural language formatting instruction + document
paragraph metadata to Groq (Llama 3.3 70B).

Groq returns a structured JSON settings dict that formatter.py applies
directly — no dropdowns needed.

Example NLP inputs handled:
  "Make unit headings Times New Roman 13pt bold, justify body, single black border"
  "Use Arial 11pt throughout, A4 page, double border in navy"
  "Heading font Calibri 14, body Georgia 11, italic all book titles"
  "Left align everything, no page border, table headers bold Arial 10"
"""

import json
import re
import requests


# ── Default fallback settings (used if Groq fails) ───────────────────────────

DEFAULTS = {
    "page_size":           "A4",
    "page_border_style":   "none",
    "page_border_color":   "000000",
    "page_border_top":     1.0,
    "page_border_bottom":  1.0,
    "page_border_left":    1.0,
    "page_border_right":   1.0,
    "unit_heading_font":   "Times New Roman",
    "unit_heading_size":   12,
    "body_font":           "Times New Roman",
    "body_size":           11,
    "table_header_font":   "Arial",
    "table_header_size":   10,
    "table_body_font":     "Arial",
    "table_body_size":     10,
    "reference_font":      "Times New Roman",
    "reference_size":      11,
    "alignment":           "justify",
    "bold_headings":       True,
    "italic_titles":       True,
    "bold_hours":          True,
    "bold_topic_headings": True,
    "cell_border_style":   "single",
    "cell_border_color":   "000000",
    # Reference part formatting
    "author_italic":       True,
    "author_bold":         False,
    "title_italic":        False,
    "title_bold":          True,
    "publisher_italic":    False,
    "publisher_bold":      False,
    "edition_italic":      False,
    "edition_bold":        False,
    "year_italic":         False,
    "year_bold":           False,
}


# ── System prompt ─────────────────────────────────────────────────────────────

_SYSTEM_PROMPT = """You are a Word document formatting assistant.
The user gives a natural language formatting instruction for an academic syllabus document.

Extract formatting settings and return ONLY a valid JSON object — no explanation, no markdown, no backticks.

JSON keys and allowed values:
{
  "page_size":           "A4" | "Letter" | "Legal",
  "page_border_style":   "none" | "single" | "double" | "shadow" | "thick",
  "page_border_color":   "<6-digit hex e.g. 000000>",
  "page_border_top":     <cm as float e.g. 1.0>,
  "page_border_bottom":  <cm as float>,
  "page_border_left":    <cm as float>,
  "page_border_right":   <cm as float>,
  "unit_heading_font":   "<font name>",
  "unit_heading_size":   <int pt>,
  "body_font":           "<font name>",
  "body_size":           <int pt>,
  "table_header_font":   "<font name>",
  "table_header_size":   <int pt>,
  "table_body_font":     "<font name>",
  "table_body_size":     <int pt>,
  "reference_font":      "<font name>",
  "reference_size":      <int pt>,
  "alignment":           "justify" | "left" | "center" | "right",
  "bold_headings":       true | false,
  "italic_titles":       true | false,
  "bold_hours":          true | false,
  "cell_border_style":   "single" | "double" | "none",
  "cell_border_color":   "<6-digit hex>"
}

Rules:
- If the user does not mention a setting, use its default value (listed below).
- Font names must be valid Word fonts e.g. Times New Roman, Arial, Calibri, Georgia, Verdana.
- Colours: convert colour names to hex. black=000000, white=FFFFFF, navy=000080, red=FF0000, grey=808080.
- If user says "no border" or "none" for page border, set page_border_style to "none".
- Return ONLY the JSON. Nothing else.

Defaults:
page_size=A4, page_border_style=none, page_border_color=000000,
page_border_top=1.0, page_border_bottom=1.0, page_border_left=1.0, page_border_right=1.0,
unit_heading_font=Times New Roman, unit_heading_size=12,
body_font=Times New Roman, body_size=11,
table_header_font=Arial, table_header_size=10,
table_body_font=Arial, table_body_size=10,
reference_font=Times New Roman, reference_size=11,
alignment=justify, bold_headings=true, italic_titles=true, bold_hours=true,
cell_border_style=single, cell_border_color=000000
"""


# ── Paragraph classifier prompt (for run-level detection) ─────────────────────

_RUN_SYSTEM_PROMPT = """You are a Word document run-level classifier.
Given a paragraph with its text runs, classify each run into one of these types:

  unit_heading   — ALL-CAPS bold heading e.g. "CONCEPTS OF STRESS AND STRAIN:"
  topic_heading  — A named topic/sub-topic inside a unit paragraph that acts as
                   a heading for a concept. These are typically title-case or
                   sentence-case noun phrases that name a specific engineering
                   topic e.g. "Shear Stress", "Bending Moment", "Euler's Theory",
                   "Deflection of Beams". They appear before or between body text
                   and describe the subject of that section. Bold these.
  body           — Descriptive or explanatory text — not a topic name.
                   e.g. "and their applications", "using double integration method"
  lecture_hours  — A number in brackets at end of paragraph e.g. "(9)"
  reference      — A textbook or reference citation line
  table_header   — First-row table cell content
  table_body     — Non-header table cell content
  section_label  — Section labels like "TEXT BOOKS:", "REFERENCES:"

Classification rules:
- If a run is a named concept/topic heading → topic_heading (bold it)
- If a run is descriptive explanation of a concept → body
- Lecture-hour counts like (9) at end of paragraph → lecture_hours (bold)
- When in doubt between topic_heading and body, prefer topic_heading for
  short noun phrases that clearly name an engineering concept.

Return ONLY a JSON array of objects — no explanation, no markdown, no backticks.
Format:
[
  {"text": "<run text>", "type": "<type>", "bold": true|false, "italic": true|false},
  ...
]
"""


def parse_nlp_instruction(instruction: str, api_key: str) -> dict:
    """
    Send a natural language formatting instruction to Groq.
    Returns a settings dict ready for format_document().
    Falls back to DEFAULTS on any error.
    """
    if not instruction.strip() or not api_key.strip():
        return DEFAULTS.copy()

    try:
        response = requests.post(
            "https://api.groq.com/openai/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            json={
                "model": "llama-3.3-70b-versatile",
                "messages": [
                    {"role": "system", "content": _SYSTEM_PROMPT},
                    {"role": "user",   "content": instruction},
                ],
                "max_tokens": 512,
                "temperature": 0,
            },
            timeout=30,
        )

        if response.status_code != 200:
            print(f"[Groq Parser] API error {response.status_code}: {response.text}")
            return DEFAULTS.copy()

        raw = response.json()["choices"][0]["message"]["content"].strip()

        # Strip accidental markdown fences
        raw = re.sub(r"^```[a-z]*\n?", "", raw)
        raw = re.sub(r"\n?```$",       "", raw)

        parsed = json.loads(raw)

        # Merge with defaults so missing keys are always filled
        settings = DEFAULTS.copy()
        settings.update(parsed)

        # Normalise types
        settings["unit_heading_size"]  = int(settings["unit_heading_size"])
        settings["body_size"]          = int(settings["body_size"])
        settings["table_header_size"]  = int(settings["table_header_size"])
        settings["table_body_size"]    = int(settings["table_body_size"])
        settings["reference_size"]     = int(settings["reference_size"])
        settings["page_border_top"]    = float(settings["page_border_top"])
        settings["page_border_bottom"] = float(settings["page_border_bottom"])
        settings["page_border_left"]   = float(settings["page_border_left"])
        settings["page_border_right"]  = float(settings["page_border_right"])
        settings["bold_headings"]      = bool(settings["bold_headings"])
        settings["italic_titles"]      = bool(settings["italic_titles"])
        settings["bold_hours"]         = bool(settings["bold_hours"])

        # Strip any # from colour fields
        for key in ("page_border_color", "cell_border_color"):
            settings[key] = str(settings[key]).lstrip("#").upper()

        return settings

    except json.JSONDecodeError as e:
        print(f"[Groq Parser] JSON parse error: {e}\nRaw response: {raw}")
        return DEFAULTS.copy()
    except Exception as e:
        print(f"[Groq Parser] Unexpected error: {e}")
        return DEFAULTS.copy()


# ── Reference parts prompt ────────────────────────────────────────────────────

_REF_PARTS_SYSTEM_PROMPT = """You are an academic reference formatter.
Given a single reference/citation string, split it into labelled parts.

Part types:
  author    — the author name(s) at the start
  title     — the book or article title (may or may not be in quotes)
  publisher — the publisher name
  edition   — edition info e.g. "7th edition", "2nd ed."
  year      — the publication year
  separator — punctuation/spaces between parts e.g. ", " or ". "

Rules:
- Split as finely as possible — each distinct piece gets its own object.
- Keep punctuation (commas, periods, quotes) attached to the nearest part OR as a separator object.
- Return ONLY a JSON array — no explanation, no markdown fences, no backticks.

Format:
[
  {"text": "James M Gere", "part": "author"},
  {"text": ", ",            "part": "separator"},
  {"text": "Mechanics of Materials", "part": "title"},
  {"text": ", ",            "part": "separator"},
  {"text": "Cengage Learning, Inc", "part": "publisher"},
  {"text": ", ",            "part": "separator"},
  {"text": "7th edition",  "part": "edition"},
  {"text": ", ",            "part": "separator"},
  {"text": "2008.",         "part": "year"}
]
"""


def format_reference_parts(reference_text: str, api_key: str,
                            ref_format: dict) -> list:
    """
    Split one reference string into labelled parts using Groq,
    then apply bold/italic/normal based on ref_format.

    ref_format keys (all bool):
      author_italic, author_bold
      title_italic,  title_bold
      publisher_italic, publisher_bold
      edition_italic,   edition_bold
      year_italic,      year_bold

    Returns a list of dicts:
      [{"text": "...", "bold": bool, "italic": bool}, ...]
    ready to be written as Word runs.

    Falls back to single unstyled run on any error.
    """
    fallback = [{"text": reference_text, "bold": False, "italic": False}]

    if not reference_text.strip() or not api_key.strip():
        return fallback

    try:
        response = requests.post(
            "https://api.groq.com/openai/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            json={
                "model": "llama-3.3-70b-versatile",
                "messages": [
                    {"role": "system", "content": _REF_PARTS_SYSTEM_PROMPT},
                    {"role": "user",   "content": reference_text},
                ],
                "max_tokens": 512,
                "temperature": 0,
            },
            timeout=30,
        )

        if response.status_code != 200:
            return fallback

        raw = response.json()["choices"][0]["message"]["content"].strip()
        raw = re.sub(r"^```[a-z]*\n?", "", raw)
        raw = re.sub(r"\n?```$",       "", raw)

        parts = json.loads(raw)

        # Apply formatting based on part type
        result = []
        for part in parts:
            ptype  = part.get("part", "separator")
            text   = part.get("text", "")
            bold   = ref_format.get(f"{ptype}_bold",   False)
            italic = ref_format.get(f"{ptype}_italic", False)
            result.append({"text": text, "bold": bold, "italic": italic})

        return result if result else fallback

    except Exception as exc:
        print(f"[Groq Ref Parts] Error: {exc}")
        return fallback


def classify_paragraph_runs(runs: list, api_key: str) -> list:
    """
    Send a paragraph's runs to Groq for type classification.

    runs: list of dicts with keys: text, bold, italic
    Returns enriched list with 'type' field added to each run.
    Falls back to 'body' type for all runs on any error.
    """
    if not runs or not api_key.strip():
        return [{**r, "type": "body"} for r in runs]

    user_msg = json.dumps(runs, ensure_ascii=False)

    try:
        response = requests.post(
            "https://api.groq.com/openai/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            json={
                "model": "llama-3.3-70b-versatile",
                "messages": [
                    {"role": "system", "content": _RUN_SYSTEM_PROMPT},
                    {"role": "user",   "content": user_msg},
                ],
                "max_tokens": 512,
                "temperature": 0,
            },
            timeout=30,
        )

        if response.status_code != 200:
            return [{**r, "type": "body"} for r in runs]

        raw = response.json()["choices"][0]["message"]["content"].strip()
        raw = re.sub(r"^```[a-z]*\n?", "", raw)
        raw = re.sub(r"\n?```$",       "", raw)

        classified = json.loads(raw)

        # Safety: if count changed, return originals
        if len(classified) != len(runs):
            return [{**r, "type": "body"} for r in runs]

        return classified

    except Exception:
        return [{**r, "type": "body"} for r in runs]
