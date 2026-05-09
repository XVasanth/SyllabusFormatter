# 📄 Word Document Formatter — Powered by Groq

> A free, local desktop app that formats academic Word documents using **one Groq API key** for three tasks: NLP instruction parsing, run-level formatting, and reference cleanup.

![Python](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)
![Groq](https://img.shields.io/badge/Groq_API-Free_Tier-orange)

---

## ✨ What It Does

Type your formatting requirements in plain English — Groq does the rest.

```
"Times New Roman 13pt headings bold, Georgia 11pt body justified,
 Arial 10pt tables, single black border 1.5cm top 1cm sides, italic book titles"
```

The app extracts every formatting rule from that sentence and applies it to your Word document instantly.

---

## 🔑 One Groq API Key — Three Jobs

| Job | What Groq Does |
|-----|----------------|
| **NLP Parsing** | Reads your plain-English instruction → returns formatting settings JSON |
| **Run Classification** | Identifies each text run inside a paragraph (heading / body / lecture hours) |
| **Reference Cleanup** | Fixes `7thedition` → `7th edition`, spacing errors in citations |

**Free tier is enough:** each document = 1–3 API calls = ~5,000 tokens. Free limit is 1,000 requests/day.

---

## 🖥 App Layout

```
┌─────────────────────────────────────────────────────────────────┐
│  📄  Word Document Formatter  ·  Powered by Groq               │
├─────────────────────────────────────────────────────────────────┤
│  GROQ API KEY   [gsk_****************************]             │
│                                                                 │
│  INPUT FILE     [path/to/file.docx]              [Browse]      │
│                                                                 │
│  FORMATTING INSTRUCTION  —  Type in plain English              │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │ ☑ Use Groq AI to read my instruction                   │   │
│  │ [Times New Roman 12pt headings, Arial 10pt tables ...]  │   │
│  └─────────────────────────────────────────────────────────┘   │
│                                                                 │
│  MANUAL SETTINGS  (tabs: Page | Fonts | Paragraph | Tables)    │
│                                                                 │
│  EXTRA GROQ FEATURES                                           │
│  ☐ Run-level classification  ☑ Clean reference entries         │
│                                                                 │
│  OUTPUT FILE    [path/to/file_formatted.docx]    [Browse]      │
│                                                                 │
│              [     FORMAT DOCUMENT     ]                        │
│                                                                 │
│  STATUS LOG                                                     │
│  > Sending instruction to Groq ...                             │
│  > ✔  Groq returned formatting settings.                       │
│  > ✅  Done!  →  file_formatted.docx                           │
└─────────────────────────────────────────────────────────────────┘
```

---

## 🚀 Quick Start

```bash
git clone https://github.com/YOUR_USERNAME/word-doc-formatter.git
cd word-doc-formatter
python -m venv venv
source venv/bin/activate      # Windows: venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

---

## 🔑 Get a Free Groq API Key

1. Go to **[console.groq.com](https://console.groq.com)**
2. Sign up — no credit card required
3. Click **API Keys → Create API Key**
4. Copy the key (starts with `gsk_`)
5. Paste it into the app's API Key field

---

## 📁 Project Structure

```
word-doc-formatter/
├── main.py            # Tkinter GUI — entry point
├── formatter.py       # Core formatting engine
├── groq_parser.py     # NLP instruction → settings JSON + run classification
├── groq_cleaner.py    # Reference entry cleanup
├── requirements.txt   # Python dependencies
├── LICENSE            # MIT
└── README.md
```

---

## ⚙️ Formatting Options

### Via NLP (plain English)
Just type what you want. Examples:
- *"Use Calibri throughout, 14pt headings bold, 11pt body, no border"*
- *"A4 page, double navy border 2cm all sides, justify everything"*
- *"Arial 12pt headings, Times New Roman 11pt body, italic book titles, single grey table borders"*

### Via Manual Tabs (when NLP is off)
| Tab | Controls |
|-----|---------|
| Page | Page size, border style, border colour, border size per edge (cm) |
| Fonts | Font family + size for: unit heading / body / table header / table body / references |
| Paragraph | Alignment, bold headings, italic titles, bold lecture hours |
| Table Borders | Cell border style + colour |

---

## 🛠 Requirements

- Python 3.8+
- `python-docx >= 1.1.2`
- `requests >= 2.31.0`
- Tkinter (bundled with Python on Windows/macOS)

Linux: `sudo apt install python3-tk`

---

## 🤝 Contributing

1. Fork → `git checkout -b feature/my-feature` → commit → push → Pull Request

Ideas welcome: PDF export, batch processing, dark mode, font family picker.

---

## 📄 License

MIT — free to use, modify, and distribute.
