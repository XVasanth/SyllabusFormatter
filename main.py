"""
main.py
───────
Word Document Formatter — Template-based approach.

Flow:
  1. User uploads a correctly-formatted template .docx
  2. Groq reads the template and extracts all formatting rules
  3. User uploads the input .docx to format
  4. User sets page settings manually (size, border)
  5. App applies extracted rules + page settings → output .docx
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from template_reader import extract_rules_from_template, DEFAULTS
from formatter import extract_references, format_document
from groq_cleaner import clean_references_with_groq

PAGE_SIZE_OPTIONS   = ["A4", "Letter", "Legal"]
PAGE_BORDER_OPTIONS = ["None", "Single", "Double", "Shadow", "Thick"]


class FormatterApp:
    C_HEADER = "#2c3e50"
    C_BTN    = "#3498db"
    C_RUN    = "#27ae60"
    C_BG     = "#f0f0f0"
    C_LABEL  = "#555555"
    C_CARD   = "#e8ecef"

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Word Document Formatter")
        self.root.geometry("680x780")
        self.root.resizable(False, False)
        self.root.configure(bg=self.C_BG)

        # Paths
        self.template_path = tk.StringVar()
        self.input_path    = tk.StringVar()
        self.output_path   = tk.StringVar()

        # Groq
        self.groq_api_key    = tk.StringVar()
        self.use_run_classify = tk.BooleanVar(value=False)
        self.use_ref_cleanup  = tk.BooleanVar(value=True)

        # Page settings (manual — not read from template)
        self.page_size          = tk.StringVar(value="A4")
        self.page_border_style  = tk.StringVar(value="None")
        self.page_border_color  = tk.StringVar(value="000000")
        self.page_border_top    = tk.DoubleVar(value=1.0)
        self.page_border_bottom = tk.DoubleVar(value=1.0)
        self.page_border_left   = tk.DoubleVar(value=1.0)
        self.page_border_right  = tk.DoubleVar(value=1.0)

        # Extracted settings (shown after template is read)
        self._extracted_settings = None

        self._build_ui()

    # ── UI ───────────────────────────────────────────────────────────────────

    def _build_ui(self):
        banner = tk.Frame(self.root, bg=self.C_HEADER)
        banner.pack(fill=tk.X)
        tk.Label(
            banner, text="📄  Word Document Formatter",
            font=("Arial", 15, "bold"), fg="white", bg=self.C_HEADER,
        ).pack(pady=12)

        main = tk.Frame(self.root, bg=self.C_BG)
        main.pack(fill=tk.BOTH, expand=True, padx=22, pady=10)

        # ── Groq API key ──────────────────────────────────────────────────────
        self._section(main, "GROQ API KEY  —  Free · console.groq.com")
        kf = tk.Frame(main, bg=self.C_BG); kf.pack(fill=tk.X, pady=2)
        tk.Label(kf, text="Key:", width=6, anchor="w",
                 bg=self.C_BG).pack(side=tk.LEFT)
        tk.Entry(kf, textvariable=self.groq_api_key, show="*",
                 width=60, font=("Arial", 9)).pack(side=tk.LEFT, padx=4)

        # ── Step 1: Template ──────────────────────────────────────────────────
        self._section(main, "STEP 1 — UPLOAD TEMPLATE  (your correctly-formatted .docx)")
        tf = tk.Frame(main, bg=self.C_BG); tf.pack(fill=tk.X, pady=2)
        tk.Entry(tf, textvariable=self.template_path, width=54,
                 font=("Arial", 9)).pack(side=tk.LEFT)
        self._btn(tf, "Browse", self._browse_template).pack(side=tk.LEFT, padx=6)
        self._btn(tf, "Read Template →", self._read_template,
                  color="#8e44ad").pack(side=tk.LEFT)

        # Rules preview card
        self.rules_card = tk.LabelFrame(
            main, text="  Extracted Rules  ",
            bg=self.C_CARD, fg="#555", font=("Arial", 8),
            padx=10, pady=6,
        )
        self.rules_card.pack(fill=tk.X, pady=4)
        self.rules_label = tk.Label(
            self.rules_card,
            text="Upload a template and click  Read Template →  to extract rules.",
            fg="#999", bg=self.C_CARD, font=("Arial", 8), justify="left",
            wraplength=580,
        )
        self.rules_label.pack(anchor="w")

        # ── Step 2: Page settings ─────────────────────────────────────────────
        self._section(main, "STEP 2 — PAGE SETTINGS  (not read from template)")
        pg = tk.LabelFrame(main, bg=self.C_BG, padx=12, pady=8)
        pg.pack(fill=tk.X, pady=2)

        r1 = tk.Frame(pg, bg=self.C_BG); r1.pack(fill=tk.X, pady=3)
        tk.Label(r1, text="Page Size:", width=16, anchor="w",
                 bg=self.C_BG).pack(side=tk.LEFT)
        ttk.Combobox(r1, textvariable=self.page_size,
                     values=PAGE_SIZE_OPTIONS, width=10,
                     state="readonly").pack(side=tk.LEFT)
        tk.Label(r1, text="   Border:", bg=self.C_BG).pack(side=tk.LEFT, padx=(12,0))
        ttk.Combobox(r1, textvariable=self.page_border_style,
                     values=PAGE_BORDER_OPTIONS, width=10,
                     state="readonly").pack(side=tk.LEFT)
        tk.Label(r1, text="  Colour:", bg=self.C_BG).pack(side=tk.LEFT, padx=(8,0))
        tk.Entry(r1, textvariable=self.page_border_color, width=8,
                 font=("Arial", 9)).pack(side=tk.LEFT, padx=4)

        r2 = tk.Frame(pg, bg=self.C_BG); r2.pack(fill=tk.X, pady=3)
        tk.Label(r2, text="Border Size (cm):", width=16,
                 anchor="w", bg=self.C_BG).pack(side=tk.LEFT)
        for label, var in [("Top",   self.page_border_top),
                            ("Bot",  self.page_border_bottom),
                            ("Left", self.page_border_left),
                            ("Right",self.page_border_right)]:
            tk.Label(r2, text=f"{label}:", bg=self.C_BG).pack(side=tk.LEFT)
            tk.Spinbox(r2, from_=0.1, to=5.0, increment=0.1,
                       textvariable=var, width=5,
                       format="%.1f", font=("Arial",9)).pack(side=tk.LEFT, padx=(0,8))

        # ── Step 3: Input file ────────────────────────────────────────────────
        self._section(main, "STEP 3 — UPLOAD INPUT FILE  (document to format)")
        inf = tk.Frame(main, bg=self.C_BG); inf.pack(fill=tk.X, pady=2)
        tk.Entry(inf, textvariable=self.input_path, width=54,
                 font=("Arial", 9)).pack(side=tk.LEFT)
        self._btn(inf, "Browse", self._browse_input).pack(side=tk.LEFT, padx=6)

        # ── Extra Groq features ───────────────────────────────────────────────
        self._section(main, "EXTRA GROQ FEATURES")
        gf = tk.LabelFrame(main, bg=self.C_BG, padx=10, pady=6)
        gf.pack(fill=tk.X, pady=2)
        tk.Checkbutton(
            gf,
            text="Run-level classification  "
                 "(bold topic headings inside unit paragraphs — slower)",
            variable=self.use_run_classify, bg=self.C_BG,
        ).pack(anchor="w")
        tk.Checkbutton(
            gf,
            text="Clean reference entries  "
                 "(fixes '7thedition' → '7th edition')",
            variable=self.use_ref_cleanup, bg=self.C_BG,
        ).pack(anchor="w", pady=(4,0))

        # ── Output file ───────────────────────────────────────────────────────
        self._section(main, "OUTPUT FILE")
        of = tk.Frame(main, bg=self.C_BG); of.pack(fill=tk.X, pady=2)
        tk.Entry(of, textvariable=self.output_path, width=54,
                 font=("Arial", 9)).pack(side=tk.LEFT)
        self._btn(of, "Browse", self._browse_output).pack(side=tk.LEFT, padx=6)

        # ── Format button ─────────────────────────────────────────────────────
        tk.Button(
            main, text="FORMAT DOCUMENT",
            command=self._run_formatter,
            bg=self.C_RUN, fg="white",
            font=("Arial", 12, "bold"),
            relief=tk.FLAT, padx=28, pady=9, cursor="hand2",
        ).pack(pady=10)

        # ── Status log ────────────────────────────────────────────────────────
        self._section(main, "STATUS LOG")
        self.log = tk.Text(
            main, height=6, state=tk.DISABLED,
            bg="#1e1e1e", fg="#d4d4d4",
            font=("Courier", 9), relief=tk.FLAT,
        )
        self.log.pack(fill=tk.X)
        self._log("Ready — upload a template to begin.")

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _section(self, parent, text: str):
        f = tk.Frame(parent, bg=self.C_BG)
        f.pack(fill=tk.X, pady=(10, 2))
        tk.Label(f, text=text, font=("Arial", 9, "bold"),
                 fg=self.C_LABEL, bg=self.C_BG).pack(side=tk.LEFT)
        tk.Frame(f, bg="#bdc3c7", height=1).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=6)

    def _btn(self, parent, label, cmd, color=None) -> tk.Button:
        return tk.Button(
            parent, text=label, command=cmd,
            bg=color or self.C_BTN, fg="white",
            relief=tk.FLAT, padx=10, cursor="hand2",
        )

    def _log(self, msg: str):
        self.log.config(state=tk.NORMAL)
        self.log.insert(tk.END, f"  {msg}\n")
        self.log.see(tk.END)
        self.log.config(state=tk.DISABLED)

    # ── File dialogs ──────────────────────────────────────────────────────────

    def _browse_template(self):
        p = filedialog.askopenfilename(
            title="Select Template Document",
            filetypes=[("Word Documents","*.docx")],
        )
        if p:
            self.template_path.set(p)
            self._log(f"Template selected: {os.path.basename(p)}")

    def _browse_input(self):
        p = filedialog.askopenfilename(
            title="Select Input Document",
            filetypes=[("Word Documents","*.docx")],
        )
        if p:
            self.input_path.set(p)
            base = os.path.splitext(p)[0]
            self.output_path.set(base + "_formatted.docx")
            self._log(f"Input selected: {os.path.basename(p)}")

    def _browse_output(self):
        p = filedialog.asksaveasfilename(
            title="Save Formatted Document",
            defaultextension=".docx",
            filetypes=[("Word Documents","*.docx")],
        )
        if p:
            self.output_path.set(p)

    # ── Read template ─────────────────────────────────────────────────────────

    def _read_template(self):
        if not self.template_path.get():
            messagebox.showerror("No Template", "Please select a template file first.")
            return
        if not os.path.isfile(self.template_path.get()):
            messagebox.showerror("Not Found", "Template file does not exist.")
            return
        if not self.groq_api_key.get().strip():
            messagebox.showerror("No API Key", "Please enter your Groq API key first.")
            return
        threading.Thread(target=self._do_read_template, daemon=True).start()

    def _do_read_template(self):
        self._log("─" * 50)
        self._log("Reading template …")
        settings = extract_rules_from_template(
            self.template_path.get(),
            self.groq_api_key.get().strip(),
        )
        self._extracted_settings = settings
        self._log("✔  Rules extracted from template:")
        self._log(f"   Heading : {settings['unit_heading_font']} {settings['unit_heading_size']}pt  bold={settings['bold_headings']}")
        self._log(f"   Body    : {settings['body_font']} {settings['body_size']}pt  align={settings['alignment']}")
        self._log(f"   Ref     : {settings['reference_font']} {settings['reference_size']}pt")
        self._log(f"   Author  : italic={settings['author_italic']}  Title: bold={settings['title_bold']}")
        self._log(f"   Hours   : bold={settings['bold_hours']}  italic={settings['italic_hours']}  right={settings['right_align_hours']}")
        self._log(f"   Table   : {settings['table_header_font']} header / {settings['table_body_font']} body")

        # Update the rules card in the UI
        summary = (
            f"Heading: {settings['unit_heading_font']} {settings['unit_heading_size']}pt bold\n"
            f"Body: {settings['body_font']} {settings['body_size']}pt {settings['alignment']}\n"
            f"References: {settings['reference_font']} {settings['reference_size']}pt  |  "
            f"Author italic={settings['author_italic']}  Title bold={settings['title_bold']}\n"
            f"Lecture Hours: bold={settings['bold_hours']}  italic={settings['italic_hours']}  "
            f"right-aligned={settings['right_align_hours']}\n"
            f"Table header: {settings['table_header_font']} {settings['table_header_size']}pt  |  "
            f"Cell borders: {settings['cell_border_style']}"
        )
        self.rules_label.config(text=summary, fg="#222", bg=self.C_CARD)

    # ── Format ────────────────────────────────────────────────────────────────

    def _run_formatter(self):
        if not self.input_path.get() or not os.path.isfile(self.input_path.get()):
            messagebox.showerror("No Input", "Please select a valid input file.")
            return
        if not self.output_path.get():
            messagebox.showerror("No Output", "Please set an output file path.")
            return
        if self._extracted_settings is None:
            messagebox.showerror("No Rules", "Please read a template first.")
            return
        threading.Thread(target=self._do_format, daemon=True).start()

    def _do_format(self):
        self._log("─" * 50)
        api_key = self.groq_api_key.get().strip()

        # Merge extracted settings with manual page settings
        settings = dict(self._extracted_settings)
        settings["page_size"]          = self.page_size.get()
        settings["page_border_style"]  = self.page_border_style.get().lower()
        settings["page_border_color"]  = self.page_border_color.get().lstrip("#")
        settings["page_border_top"]    = self.page_border_top.get()
        settings["page_border_bottom"] = self.page_border_bottom.get()
        settings["page_border_left"]   = self.page_border_left.get()
        settings["page_border_right"]  = self.page_border_right.get()

        # Reference cleanup
        cleaned_refs = None
        if self.use_ref_cleanup.get() and api_key:
            self._log("Cleaning reference entries …")
            try:
                refs = extract_references(self.input_path.get())
                if refs:
                    cleaned_refs = clean_references_with_groq(refs, api_key)
                    self._log(f"✔  {len(refs)} references cleaned.")
            except Exception as exc:
                self._log(f"⚠  Cleanup error: {exc}")

        # Format
        self._log("Applying formatting …")
        try:
            format_document(
                self.input_path.get(),
                self.output_path.get(),
                settings,
                cleaned_refs,
                groq_api_key     = api_key or None,
                use_run_classify = self.use_run_classify.get() and bool(api_key),
            )
            self._log(f"✅  Done!  →  {os.path.basename(self.output_path.get())}")
            messagebox.showinfo("Success",
                f"Formatted document saved!\n\n{self.output_path.get()}")
        except Exception as exc:
            self._log(f"❌  Error: {exc}")
            messagebox.showerror("Error", str(exc))


if __name__ == "__main__":
    root = tk.Tk()
    FormatterApp(root)
    root.mainloop()
