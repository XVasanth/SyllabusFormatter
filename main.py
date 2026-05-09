"""
main.py
───────
Word Document Formatter — Tkinter GUI
Groq API key powers:
  1. NLP instruction → formatting settings  (groq_parser.parse_nlp_instruction)
  2. Run-level classification per paragraph (groq_parser.classify_paragraph_runs)
  3. Reference entry cleanup               (groq_cleaner.clean_references_with_groq)
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from formatter import extract_references, format_document
from groq_cleaner import clean_references_with_groq
from groq_parser import parse_nlp_instruction, DEFAULTS

FONT_OPTIONS        = ["Times New Roman", "Arial", "Calibri", "Georgia",
                       "Verdana", "Courier New", "Cambria", "Garamond", "Helvetica"]
PAGE_SIZE_OPTIONS   = ["A4", "Letter", "Legal"]
PAGE_BORDER_OPTIONS = ["None", "Single", "Double", "Shadow", "Thick"]
ALIGNMENT_OPTIONS   = ["Justify", "Left", "Center", "Right"]
CELL_BORDER_OPTIONS = ["Single", "Double", "None"]


class FormatterApp:
    C_HEADER  = "#2c3e50"
    C_BTN     = "#3498db"
    C_RUN     = "#27ae60"
    C_NLP     = "#8e44ad"
    C_BG      = "#f0f0f0"
    C_LABEL   = "#555555"

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Word Document Formatter — Powered by Groq")
        self.root.geometry("720x820")
        self.root.resizable(False, False)
        self.root.configure(bg=self.C_BG)

        # ── Groq key (shared across all three features) ───────────────────────
        self.groq_api_key   = tk.StringVar()

        # ── NLP mode ─────────────────────────────────────────────────────────
        self.use_nlp        = tk.BooleanVar(value=True)
        self.nlp_instruction = tk.StringVar(
            value="Times New Roman 12pt headings, 11pt body, Arial 10pt tables, "
                  "justify body, single black borders, A4, italic book titles"
        )

        # ── Manual settings (used when NLP is off) ────────────────────────────
        self.page_size          = tk.StringVar(value="A4")
        self.page_border_style  = tk.StringVar(value="None")
        self.page_border_color  = tk.StringVar(value="000000")
        self.page_border_top    = tk.DoubleVar(value=1.0)
        self.page_border_bottom = tk.DoubleVar(value=1.0)
        self.page_border_left   = tk.DoubleVar(value=1.0)
        self.page_border_right  = tk.DoubleVar(value=1.0)
        self.unit_heading_font  = tk.StringVar(value="Times New Roman")
        self.unit_heading_size  = tk.IntVar(value=12)
        self.body_font          = tk.StringVar(value="Times New Roman")
        self.body_size          = tk.IntVar(value=11)
        self.table_header_font  = tk.StringVar(value="Arial")
        self.table_header_size  = tk.IntVar(value=10)
        self.table_body_font    = tk.StringVar(value="Arial")
        self.table_body_size    = tk.IntVar(value=10)
        self.reference_font     = tk.StringVar(value="Times New Roman")
        self.reference_size     = tk.IntVar(value=11)
        self.alignment          = tk.StringVar(value="Justify")
        self.bold_headings      = tk.BooleanVar(value=True)
        self.italic_titles      = tk.BooleanVar(value=True)
        self.bold_hours         = tk.BooleanVar(value=True)
        self.cell_border_style  = tk.StringVar(value="Single")
        self.cell_border_color  = tk.StringVar(value="000000")

        # ── Groq feature toggles ──────────────────────────────────────────────
        self.use_run_classify   = tk.BooleanVar(value=False)
        self.use_ref_cleanup    = tk.BooleanVar(value=True)

        # ── File paths ────────────────────────────────────────────────────────
        self.input_path  = tk.StringVar()
        self.output_path = tk.StringVar()

        self._build_ui()

    # ── UI ───────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Header
        banner = tk.Frame(self.root, bg=self.C_HEADER)
        banner.pack(fill=tk.X)
        tk.Label(
            banner,
            text="📄  Word Document Formatter  ·  Powered by Groq",
            font=("Arial", 14, "bold"), fg="white", bg=self.C_HEADER,
        ).pack(pady=12)

        main = tk.Frame(self.root, bg=self.C_BG)
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=8)

        # ── Groq API Key (top — shared by all features) ───────────────────────
        self._section_title(main, "GROQ API KEY  —  Free tier · console.groq.com")
        kf = tk.Frame(main, bg=self.C_BG)
        kf.pack(fill=tk.X, pady=2)
        tk.Label(kf, text="API Key:", bg=self.C_BG, width=8,
                 anchor="w").pack(side=tk.LEFT)
        tk.Entry(kf, textvariable=self.groq_api_key, show="*",
                 width=56, font=("Arial", 9)).pack(side=tk.LEFT, padx=4)

        # ── Input file ────────────────────────────────────────────────────────
        self._section_title(main, "INPUT FILE")
        f = tk.Frame(main, bg=self.C_BG)
        f.pack(fill=tk.X, pady=2)
        tk.Entry(f, textvariable=self.input_path, width=57,
                 font=("Arial", 9)).pack(side=tk.LEFT)
        self._btn(f, "Browse", self._browse_input).pack(side=tk.LEFT, padx=6)

        # ── NLP instruction box ───────────────────────────────────────────────
        self._section_title(main, "FORMATTING INSTRUCTION  —  Type in plain English")
        nlp_frame = tk.LabelFrame(main, bg=self.C_BG, padx=10, pady=8)
        nlp_frame.pack(fill=tk.X, pady=2)

        tk.Checkbutton(
            nlp_frame,
            text="Use Groq AI to read my instruction below  "
                 "(unchecking lets you set options manually in the tabs)",
            variable=self.use_nlp,
            bg=self.C_BG,
            command=self._toggle_nlp,
            font=("Arial", 9, "bold"),
        ).pack(anchor="w")

        self.nlp_entry = tk.Entry(
            nlp_frame,
            textvariable=self.nlp_instruction,
            width=82,
            font=("Arial", 9),
            fg="#333",
        )
        self.nlp_entry.pack(fill=tk.X, pady=(6, 2))
        tk.Label(
            nlp_frame,
            text='e.g.  "Arial 12pt headings bold, Georgia 11pt body justify, '
                 'single navy border 1.5cm top 1cm sides, italic book titles"',
            fg="#888", bg=self.C_BG, font=("Arial", 8),
        ).pack(anchor="w")

        # ── Manual settings tabs ──────────────────────────────────────────────
        self._section_title(main, "MANUAL SETTINGS  —  Used when Groq AI is off")
        self.nb = ttk.Notebook(main)
        self.nb.pack(fill=tk.BOTH, pady=2)
        self._tab_page(self.nb)
        self._tab_fonts(self.nb)
        self._tab_paragraph(self.nb)
        self._tab_tables(self.nb)

        # ── Extra Groq features ───────────────────────────────────────────────
        self._section_title(main, "EXTRA GROQ FEATURES")
        gf = tk.LabelFrame(main, bg=self.C_BG, padx=10, pady=6)
        gf.pack(fill=tk.X, pady=2)
        tk.Checkbutton(
            gf,
            text="Run-level classification  "
                 "(Groq identifies bold/body/hours per run — more precise, slower)",
            variable=self.use_run_classify,
            bg=self.C_BG,
        ).pack(anchor="w")
        tk.Checkbutton(
            gf,
            text="Clean reference entries  "
                 "(fixes '7thedition' → '7th edition' etc.)",
            variable=self.use_ref_cleanup,
            bg=self.C_BG,
        ).pack(anchor="w", pady=(4, 0))

        # ── Output file ───────────────────────────────────────────────────────
        self._section_title(main, "OUTPUT FILE")
        of = tk.Frame(main, bg=self.C_BG)
        of.pack(fill=tk.X, pady=2)
        tk.Entry(of, textvariable=self.output_path, width=57,
                 font=("Arial", 9)).pack(side=tk.LEFT)
        self._btn(of, "Browse", self._browse_output).pack(side=tk.LEFT, padx=6)

        # ── Format button ─────────────────────────────────────────────────────
        tk.Button(
            main, text="FORMAT DOCUMENT",
            command=self._run_formatter,
            bg=self.C_RUN, fg="white",
            font=("Arial", 12, "bold"),
            relief=tk.FLAT, padx=24, pady=8, cursor="hand2",
        ).pack(pady=8)

        # ── Status log ────────────────────────────────────────────────────────
        self._section_title(main, "STATUS LOG")
        self.log = tk.Text(
            main, height=5, state=tk.DISABLED,
            bg="#1e1e1e", fg="#d4d4d4",
            font=("Courier", 9), relief=tk.FLAT,
        )
        self.log.pack(fill=tk.X)
        self._log("Ready — enter your Groq API key and select a .docx file.")

        # Start with manual tabs disabled (NLP is on by default)
        self._toggle_nlp()

    # ── Tab builders ─────────────────────────────────────────────────────────

    def _tab_page(self, nb):
        t = tk.Frame(nb, bg=self.C_BG, padx=14, pady=10)
        nb.add(t, text="  Page  ")
        self._combo_row(t, "Page Size:",          self.page_size,         PAGE_SIZE_OPTIONS)
        self._combo_row(t, "Page Border Style:",  self.page_border_style, PAGE_BORDER_OPTIONS)
        r = tk.Frame(t, bg=self.C_BG); r.pack(fill=tk.X, pady=4)
        tk.Label(r, text="Border Colour (hex):", width=22,
                 anchor="w", bg=self.C_BG).pack(side=tk.LEFT)
        tk.Entry(r, textvariable=self.page_border_color, width=10,
                 font=("Arial", 9)).pack(side=tk.LEFT)

        # Per-edge border sizes
        sz = tk.Frame(t, bg=self.C_BG); sz.pack(fill=tk.X, pady=4)
        tk.Label(sz, text="Border Size (cm):", width=22,
                 anchor="w", bg=self.C_BG).pack(side=tk.LEFT)
        for label, var in [("Top", self.page_border_top),
                            ("Bottom", self.page_border_bottom),
                            ("Left", self.page_border_left),
                            ("Right", self.page_border_right)]:
            tk.Label(sz, text=f"{label}:", bg=self.C_BG).pack(side=tk.LEFT)
            tk.Spinbox(sz, from_=0.1, to=5.0, increment=0.1,
                       textvariable=var, width=5,
                       format="%.1f", font=("Arial", 9)).pack(side=tk.LEFT, padx=(0, 8))

    def _tab_fonts(self, nb):
        t = tk.Frame(nb, bg=self.C_BG, padx=14, pady=10)
        nb.add(t, text="  Fonts  ")
        rows = [
            ("Unit Heading:",   self.unit_heading_font, self.unit_heading_size,
             "ALL-CAPS bold e.g. CONCEPTS OF STRESS:"),
            ("Body / Sub-topic:", self.body_font,       self.body_size,
             "Regular paragraph text"),
            ("Table Header:",   self.table_header_font, self.table_header_size,
             "First row of every table"),
            ("Table Body:",     self.table_body_font,   self.table_body_size,
             "Remaining table rows"),
            ("References:",     self.reference_font,    self.reference_size,
             "Textbook and reference entries"),
        ]
        for label, fvar, svar, hint in rows:
            r = tk.Frame(t, bg=self.C_BG); r.pack(fill=tk.X, pady=4)
            tk.Label(r, text=label, width=17, anchor="w",
                     bg=self.C_BG).pack(side=tk.LEFT)
            ttk.Combobox(r, textvariable=fvar, values=FONT_OPTIONS,
                         width=17, state="readonly").pack(side=tk.LEFT)
            tk.Label(r, text="  Size:", bg=self.C_BG).pack(side=tk.LEFT)
            tk.Spinbox(r, from_=8, to=24, textvariable=svar,
                       width=4, font=("Arial", 9)).pack(side=tk.LEFT, padx=4)
            tk.Label(r, text=f"pt   ({hint})", fg="#888",
                     bg=self.C_BG, font=("Arial", 8)).pack(side=tk.LEFT)

    def _tab_paragraph(self, nb):
        t = tk.Frame(nb, bg=self.C_BG, padx=14, pady=10)
        nb.add(t, text="  Paragraph  ")
        self._combo_row(t, "Body Alignment:", self.alignment, ALIGNMENT_OPTIONS)
        for text, var in [
            ("Bold Unit Headings",                                         self.bold_headings),
            ('Italic Book Titles  (inside "double quotes" in references)', self.italic_titles),
            ("Bold Lecture Hours  e.g. (9) at end of unit paragraph",     self.bold_hours),
        ]:
            tk.Checkbutton(t, text=text, variable=var,
                           bg=self.C_BG).pack(anchor="w", pady=5)

    def _tab_tables(self, nb):
        t = tk.Frame(nb, bg=self.C_BG, padx=14, pady=10)
        nb.add(t, text="  Table Borders  ")
        self._combo_row(t, "Cell Border Style:", self.cell_border_style, CELL_BORDER_OPTIONS)
        r = tk.Frame(t, bg=self.C_BG); r.pack(fill=tk.X, pady=5)
        tk.Label(r, text="Cell Border Colour:", width=22,
                 anchor="w", bg=self.C_BG).pack(side=tk.LEFT)
        tk.Entry(r, textvariable=self.cell_border_color, width=10,
                 font=("Arial", 9)).pack(side=tk.LEFT)

    # ── Widget helpers ────────────────────────────────────────────────────────

    def _section_title(self, parent, text: str):
        f = tk.Frame(parent, bg=self.C_BG)
        f.pack(fill=tk.X, pady=(10, 2))
        tk.Label(f, text=text, font=("Arial", 9, "bold"),
                 fg=self.C_LABEL, bg=self.C_BG).pack(side=tk.LEFT)
        tk.Frame(f, bg="#bdc3c7", height=1).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=6)

    def _combo_row(self, parent, label: str, var, values: list):
        r = tk.Frame(parent, bg=self.C_BG); r.pack(fill=tk.X, pady=5)
        tk.Label(r, text=label, width=22, anchor="w",
                 bg=self.C_BG).pack(side=tk.LEFT)
        ttk.Combobox(r, textvariable=var, values=values,
                     width=14, state="readonly").pack(side=tk.LEFT)

    def _btn(self, parent, label: str, cmd) -> tk.Button:
        return tk.Button(
            parent, text=label, command=cmd,
            bg=self.C_BTN, fg="white",
            relief=tk.FLAT, padx=10, cursor="hand2",
        )

    def _log(self, message: str):
        self.log.config(state=tk.NORMAL)
        self.log.insert(tk.END, f"  {message}\n")
        self.log.see(tk.END)
        self.log.config(state=tk.DISABLED)

    def _toggle_nlp(self):
        """Enable/disable manual tabs based on NLP toggle."""
        state = tk.DISABLED if self.use_nlp.get() else tk.NORMAL
        self.nlp_entry.config(state=tk.NORMAL if self.use_nlp.get() else tk.DISABLED)
        for tab_id in self.nb.tabs():
            # Grey out tab label when NLP is on
            pass  # Notebook tabs can't be individually disabled in Tkinter easily
            # Instead we just let both coexist — NLP overrides manual when on

    # ── File dialogs ──────────────────────────────────────────────────────────

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Select Word Document",
            filetypes=[("Word Documents", "*.docx"), ("All files", "*.*")],
        )
        if path:
            self.input_path.set(path)
            base = os.path.splitext(path)[0]
            self.output_path.set(base + "_formatted.docx")
            self._log(f"Loaded: {os.path.basename(path)}")

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Save Formatted Document",
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx")],
        )
        if path:
            self.output_path.set(path)

    # ── Runner ────────────────────────────────────────────────────────────────

    def _run_formatter(self):
        if not self.input_path.get():
            messagebox.showerror("Missing Input", "Please select an input .docx file.")
            return
        if not self.output_path.get():
            messagebox.showerror("Missing Output", "Please set an output file path.")
            return
        if not os.path.isfile(self.input_path.get()):
            messagebox.showerror("File Not Found", "The selected input file does not exist.")
            return
        threading.Thread(target=self._do_format, daemon=True).start()

    def _do_format(self):
        self._log("─" * 54)
        api_key = self.groq_api_key.get().strip()

        # ── Step 1: Get settings — NLP or manual ─────────────────────────────
        if self.use_nlp.get():
            if not api_key:
                self._log("⚠  NLP mode needs Groq API key — using defaults.")
                settings = DEFAULTS.copy()
            else:
                self._log(f"Sending instruction to Groq …")
                self._log(f'  "{self.nlp_instruction.get()[:70]}…"')
                settings = parse_nlp_instruction(
                    self.nlp_instruction.get(), api_key
                )
                self._log("✔  Groq returned formatting settings.")
                self._log(f"   Page: {settings['page_size']}  "
                          f"Border: {settings['page_border_style']}  "
                          f"Align: {settings['alignment']}")
                self._log(f"   Heading: {settings['unit_heading_font']} "
                          f"{settings['unit_heading_size']}pt  "
                          f"Body: {settings['body_font']} {settings['body_size']}pt")
        else:
            self._log("Using manual settings …")
            settings = {
                "page_size":           self.page_size.get(),
                "page_border_style":   self.page_border_style.get().lower(),
                "page_border_color":   self.page_border_color.get().lstrip("#"),
                "page_border_top":     self.page_border_top.get(),
                "page_border_bottom":  self.page_border_bottom.get(),
                "page_border_left":    self.page_border_left.get(),
                "page_border_right":   self.page_border_right.get(),
                "unit_heading_font":   self.unit_heading_font.get(),
                "unit_heading_size":   self.unit_heading_size.get(),
                "body_font":           self.body_font.get(),
                "body_size":           self.body_size.get(),
                "table_header_font":   self.table_header_font.get(),
                "table_header_size":   self.table_header_size.get(),
                "table_body_font":     self.table_body_font.get(),
                "table_body_size":     self.table_body_size.get(),
                "reference_font":      self.reference_font.get(),
                "reference_size":      self.reference_size.get(),
                "alignment":           self.alignment.get().lower(),
                "bold_headings":       self.bold_headings.get(),
                "italic_titles":       self.italic_titles.get(),
                "bold_hours":          self.bold_hours.get(),
                "cell_border_style":   self.cell_border_style.get().lower(),
                "cell_border_color":   self.cell_border_color.get().lstrip("#"),
            }

        # ── Step 2: Reference cleanup ─────────────────────────────────────────
        cleaned_refs = None
        if self.use_ref_cleanup.get():
            if not api_key:
                self._log("⚠  Reference cleanup needs Groq API key — skipping.")
            else:
                self._log("Extracting reference entries …")
                try:
                    refs = extract_references(self.input_path.get())
                    if refs:
                        self._log(f"Found {len(refs)} entries — cleaning with Groq …")
                        cleaned_refs = clean_references_with_groq(refs, api_key)
                        self._log("✔  References cleaned.")
                    else:
                        self._log("No reference entries found.")
                except Exception as exc:
                    self._log(f"⚠  Cleanup error ({exc}) — skipping.")

        # ── Step 3: Apply formatting ──────────────────────────────────────────
        self._log("Applying formatting to document …")
        if self.use_run_classify.get() and api_key:
            self._log("  Run-level classification via Groq — this may take a moment …")

        try:
            format_document(
                self.input_path.get(),
                self.output_path.get(),
                settings,
                cleaned_refs,
                groq_api_key     = api_key if self.use_run_classify.get() else None,
                use_run_classify = self.use_run_classify.get() and bool(api_key),
            )
            out_name = os.path.basename(self.output_path.get())
            self._log(f"✅  Done!  →  {out_name}")
            messagebox.showinfo(
                "Success",
                f"Document formatted and saved!\n\n{self.output_path.get()}",
            )
        except Exception as exc:
            self._log(f"❌  Error: {exc}")
            messagebox.showerror("Formatting Error", str(exc))


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    root = tk.Tk()
    app  = FormatterApp(root)
    root.mainloop()
