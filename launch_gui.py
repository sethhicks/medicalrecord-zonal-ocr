"""
Medical Record OCR — GUI Launcher
----------------------------------
Double-click this file to open the GUI.
Select a PDF to process and view live output.
"""

import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, font as tkfont

# ---------------------------------------------------------------------------
# Colour palette
# ---------------------------------------------------------------------------
BG          = "#1A1D27"   # dark navy
PANEL       = "#22263A"   # slightly lighter panel
ACCENT      = "#4A9EFF"   # blue accent
ACCENT_DARK = "#2D6FCC"
SUCCESS     = "#3DD68C"
WARNING     = "#FFB84D"
ERROR       = "#FF5F6D"
TEXT        = "#E8EAF2"
TEXT_DIM    = "#7B82A0"
BORDER      = "#2E3450"

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
SCRIPT_PATH = os.path.join(SCRIPT_DIR, "extract_medical_records.py")
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "medical_records_output.xlsx")


class MedicalOCRApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Medical Record OCR")
        self.resizable(True, True)
        self.minsize(700, 520)
        self.configure(bg=BG)

        self._pdf_path   = tk.StringVar(value="")
        self._processing = False
        self._process    = None

        self._build_ui()
        self._center_window(750, 580)

    # ── Layout ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Title bar
        title_bar = tk.Frame(self, bg=PANEL, pady=14)
        title_bar.pack(fill="x")

        tk.Label(
            title_bar, text="Medical Record OCR",
            bg=PANEL, fg=TEXT,
            font=("Segoe UI", 18, "bold"),
        ).pack(side="left", padx=24)

        tk.Label(
            title_bar, text="CMS-1500  ·  UB-04",
            bg=PANEL, fg=TEXT_DIM,
            font=("Segoe UI", 10),
        ).pack(side="left", padx=0)

        sep = tk.Frame(self, bg=BORDER, height=1)
        sep.pack(fill="x")

        # File selection row
        file_frame = tk.Frame(self, bg=BG, pady=18, padx=24)
        file_frame.pack(fill="x")

        tk.Label(
            file_frame, text="PDF File",
            bg=BG, fg=TEXT_DIM,
            font=("Segoe UI", 9, "bold"),
        ).pack(anchor="w")

        input_row = tk.Frame(file_frame, bg=BG)
        input_row.pack(fill="x", pady=(6, 0))

        self._path_entry = tk.Entry(
            input_row,
            textvariable=self._pdf_path,
            bg=PANEL, fg=TEXT,
            insertbackground=TEXT,
            relief="flat",
            font=("Segoe UI", 10),
            state="readonly",
        )
        self._path_entry.pack(side="left", fill="x", expand=True, ipady=8, ipadx=8)

        self._browse_btn = self._make_button(
            input_row, "Browse…", self._browse,
            bg=PANEL, fg=TEXT, hover=BORDER,
            padx=16, pady=0,
        )
        self._browse_btn.pack(side="left", padx=(8, 0))

        # Process button
        btn_frame = tk.Frame(self, bg=BG, padx=24)
        btn_frame.pack(fill="x")

        self._run_btn = self._make_button(
            btn_frame, "▶  Process PDF", self._start_processing,
            bg=ACCENT, fg="white", hover=ACCENT_DARK,
            padx=24, pady=10, font_size=11,
        )
        self._run_btn.pack(side="left")
        self._run_btn.configure(state="disabled")

        self._open_btn = self._make_button(
            btn_frame, "📂  Open Excel", self._open_excel,
            bg=PANEL, fg=TEXT_DIM, hover=BORDER,
            padx=16, pady=10,
        )
        self._open_btn.pack(side="left", padx=(10, 0))
        self._open_btn.configure(state="disabled")

        # Status label
        self._status_var = tk.StringVar(value="Select a PDF file to begin.")
        tk.Label(
            self, textvariable=self._status_var,
            bg=BG, fg=TEXT_DIM,
            font=("Segoe UI", 9),
            anchor="w",
        ).pack(fill="x", padx=24, pady=(10, 4))

        # Output console
        console_frame = tk.Frame(self, bg=BG, padx=24, pady=0)
        console_frame.pack(fill="both", expand=True, pady=(0, 14))

        self._console = tk.Text(
            console_frame,
            bg="#0E1019", fg=TEXT,
            insertbackground=TEXT,
            relief="flat",
            font=("Consolas", 9),
            state="disabled",
            wrap="word",
            padx=12, pady=10,
        )
        self._console.pack(side="left", fill="both", expand=True)

        scrollbar = tk.Scrollbar(console_frame, command=self._console.yview, bg=PANEL)
        scrollbar.pack(side="right", fill="y")
        self._console.configure(yscrollcommand=scrollbar.set)

        # Tag colours for console
        self._console.tag_configure("ok",      foreground=SUCCESS)
        self._console.tag_configure("warn",    foreground=WARNING)
        self._console.tag_configure("err",     foreground=ERROR)
        self._console.tag_configure("dim",     foreground=TEXT_DIM)
        self._console.tag_configure("accent",  foreground=ACCENT)
        self._console.tag_configure("normal",  foreground=TEXT)

    # ── Helpers ──────────────────────────────────────────────────────────────

    def _make_button(self, parent, text, command, bg, fg, hover,
                     padx=12, pady=6, font_size=10):
        btn = tk.Label(
            parent, text=text,
            bg=bg, fg=fg,
            font=("Segoe UI", font_size),
            padx=padx, pady=pady,
            cursor="hand2",
        )
        btn.bind("<Button-1>", lambda e: command() if str(btn["state"]) != "disabled" else None)
        btn.bind("<Enter>",    lambda e: btn.configure(bg=hover) if str(btn["state"]) != "disabled" else None)
        btn.bind("<Leave>",    lambda e: btn.configure(bg=bg))
        btn._default_bg = bg
        btn._hover_bg   = hover
        return btn

    def _center_window(self, w, h):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def _log(self, text, tag="normal"):
        self._console.configure(state="normal")
        # Colour-code common patterns automatically
        if tag == "normal":
            if text.startswith("✅") or "Saved" in text:
                tag = "ok"
            elif text.startswith("⚠") or "LOW CONFIDENCE" in text or "WARNING" in text:
                tag = "warn"
            elif text.startswith("❌") or "Error" in text:
                tag = "err"
            elif text.startswith("  [debug]"):
                tag = "dim"
            elif text.startswith("Reading:") or "page(s) found" in text:
                tag = "accent"
        self._console.insert("end", text + "\n", tag)
        self._console.see("end")
        self._console.configure(state="disabled")

    def _clear_console(self):
        self._console.configure(state="normal")
        self._console.delete("1.0", "end")
        self._console.configure(state="disabled")

    def _set_status(self, text, colour=TEXT_DIM):
        self._status_var.set(text)

    # ── Actions ──────────────────────────────────────────────────────────────

    def _browse(self):
        path = filedialog.askopenfilename(
            title="Select a PDF file",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialdir=os.path.join(SCRIPT_DIR, "pdfs") if os.path.isdir(os.path.join(SCRIPT_DIR, "pdfs")) else SCRIPT_DIR,
        )
        if path:
            self._pdf_path.set(path)
            self._run_btn.configure(state="normal", bg=ACCENT)
            self._open_btn.configure(state="disabled", fg=TEXT_DIM)
            self._set_status(f"Ready — {os.path.basename(path)}")
            self._clear_console()

    def _start_processing(self):
        if self._processing:
            return
        path = self._pdf_path.get()
        if not path or not os.path.isfile(path):
            self._set_status("Please select a valid PDF file.", ERROR)
            return

        self._processing = True
        self._run_btn.configure(state="disabled", bg=PANEL)
        self._open_btn.configure(state="disabled", fg=TEXT_DIM)
        self._browse_btn.configure(state="disabled", bg=PANEL, fg=TEXT_DIM)
        self._clear_console()
        self._set_status("Processing…")

        thread = threading.Thread(target=self._run_script, args=(path,), daemon=True)
        thread.start()

    def _run_script(self, pdf_path):
        try:
            cmd = [sys.executable, "-X", "utf8", "-u", SCRIPT_PATH, pdf_path]
            self._process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                bufsize=1,
                cwd=SCRIPT_DIR,
            )
            for line in self._process.stdout:
                line = line.rstrip()
                if line:
                    self.after(0, self._log, line)

            self._process.wait()
            rc = self._process.returncode

            if rc == 0:
                self.after(0, self._on_success)
            else:
                self.after(0, self._on_failure, f"Script exited with code {rc}.")

        except Exception as e:
            self.after(0, self._on_failure, str(e))

    def _on_success(self):
        self._processing = False
        self._run_btn.configure(state="normal", bg=ACCENT)
        self._browse_btn.configure(state="normal", bg=PANEL, fg=TEXT)
        self._set_status("Done — Excel file saved.")
        if os.path.exists(OUTPUT_FILE):
            self._open_btn.configure(state="normal", bg=SUCCESS, fg="white")

    def _on_failure(self, msg):
        self._processing = False
        self._run_btn.configure(state="normal", bg=ACCENT)
        self._browse_btn.configure(state="normal", bg=PANEL, fg=TEXT)
        self._set_status(f"Failed: {msg}")
        self._log(f"❌  {msg}", "err")

    def _open_excel(self):
        if not os.path.exists(OUTPUT_FILE):
            self._set_status("Excel file not found.", ERROR)
            return
        try:
            if sys.platform == "win32":
                os.startfile(OUTPUT_FILE)
            elif sys.platform == "darwin":
                subprocess.call(["open", OUTPUT_FILE])
            else:
                subprocess.call(["xdg-open", OUTPUT_FILE])
        except Exception as e:
            self._set_status(f"Could not open file: {e}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app = MedicalOCRApp()
    app.mainloop()