#!/usr/bin/env python3
"""
ChatGPT Canvas Markdown -> Markdown / DOCX GUI

This app is designed for Markdown copied from ChatGPT Canvas or saved from Canvas.
Version 3 makes math normalization idempotent, so already-normalized $...$ math is not corrupted by a second pass.
It can optionally normalize the common Canvas-copy math form:

    [
    E = mc^2
    ]

into Pandoc-compatible display math:

    $$
    E = mc^2
    $$

It also converts simple inline math copied as parentheses, e.g. (q), (D(q)),
(C'(q)), (\beta), into $...$.

The app does NOT try to recover math content that has already been lost. It only
adds explicit Markdown math delimiters around TeX/math-looking text.

Requires:
    - Python 3 with Tkinter
    - Pandoc on PATH or at /opt/homebrew/bin/pandoc or /usr/local/bin/pandoc
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import tempfile
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

APP_TITLE = "ChatGPT Canvas Markdown to DOCX v3"

PANDOC_FORMAT = "markdown+tex_math_dollars+tex_math_single_backslash"

# A conservative signal that a string contains TeX/math rather than ordinary prose.
MATH_SIGNAL_RE = re.compile(
    r"""
    \\[A-Za-z]+      |  # TeX commands: \int, \frac, \sum, \boxed, \text...
    [_^]             |  # subscript/superscript
    [=<>≤≥≈]          |  # equations/inequalities
    \{ | \}           # TeX grouping
    """,
    re.VERBOSE,
)

# Tiny variables copied as (q), (p), (q^*), (P_i), (\beta), etc.
FUNCTION_EXPR_RE = re.compile(r"^[A-Za-z][A-Za-z0-9']*\([^)]{1,80}\)\s*[.,;:]?\s*$")

SIMPLE_VAR_RE = re.compile(
    r"""
    ^\s*
    (?:
        \\[A-Za-z]+                                  # \beta
        |
        [A-Za-z][A-Za-z0-9]*                          # q, p, CM, CMgS
        (?:'*)                                        # C'
        (?:_\{?[A-Za-z0-9*]+\}?)?                    # _i or _{ij}
        (?:\^\{?(?:[A-Za-z0-9*]+|\\[A-Za-z]+)\}?)?  # ^*, ^2, ^\top
        (?:\([A-Za-z0-9*+\-_'^{}\\]+\))?             # (q), (q^*)
    )
    \s*$
    """,
    re.VERBOSE,
)

# Parentheses candidates within prose. Kept short to avoid grabbing whole sentences.
PAREN_CANDIDATE_RE = re.compile(r"\(([^()\n]{1,80}(?:\([^()\n]{1,40}\)[^()\n]{0,40})?)\)")


def looks_like_math(text: str) -> bool:
    s = text.strip()
    
    normalized = s.strip().rstrip('.,;:')
    return bool(s and (MATH_SIGNAL_RE.search(s) or SIMPLE_VAR_RE.fullmatch(normalized) or FUNCTION_EXPR_RE.fullmatch(s)))


def normalize_bare_bracket_display_math(text: str) -> str:
    """Convert bare bracket display math blocks into $$ blocks."""
    lines = text.splitlines()
    out: list[str] = []
    i = 0
    in_code = False

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if re.match(r"^\s*(```|~~~)", line):
            in_code = not in_code
            out.append(line)
            i += 1
            continue

        if in_code:
            out.append(line)
            i += 1
            continue

        # Multi-line Canvas-copy display math:
        # [
        # ...
        # ]
        if stripped == "[":
            body: list[str] = []
            j = i + 1
            while j < len(lines) and lines[j].strip() != "]":
                body.append(lines[j])
                j += 1

            if j < len(lines):
                body_text = "\n".join(body).strip()
                if looks_like_math(body_text):
                    out.append("$$")
                    out.extend(body)
                    out.append("$$")
                    i = j + 1
                    continue

        # One-line variant: [ E = mc^2 ]
        m = re.fullmatch(r"\s*\[\s*(.+?)\s*\]\s*", line)
        if m and looks_like_math(m.group(1)):
            out.append("$$")
            out.append(m.group(1))
            out.append("$$")
            i += 1
            continue

        out.append(line)
        i += 1

    return "\n".join(out)


def split_outside_inline_dollar_math(line: str) -> list[tuple[str, bool]]:
    """Split one Markdown line into (segment, is_math) parts for unescaped $...$ spans.

    This makes normalization idempotent: if a previous pass has already produced
    $D(q)$, the parenthesized q inside that math span is not processed again.
    """
    parts: list[tuple[str, bool]] = []
    i = 0
    start = 0
    in_math = False

    while i < len(line):
        ch = line[i]
        escaped = i > 0 and line[i - 1] == "\\"
        if ch == "$" and not escaped:
            # Keep display delimiters or empty dollar pairs untouched here. The
            # caller handles lines that are exactly $$ as display blocks.
            if i > start:
                parts.append((line[start:i], in_math))
            parts.append(("$", True))
            in_math = not in_math
            i += 1
            start = i
        else:
            i += 1

    if start < len(line):
        parts.append((line[start:], in_math))

    if not parts:
        return [(line, False)]

    # Merge delimiter tokens into their adjacent math spans, so "$D(q)$" stays
    # as one protected math-ish region during replacement.
    merged: list[tuple[str, bool]] = []
    buffer = ""
    math_mode = False
    collecting_math = False

    for segment, is_math_token_or_inside in parts:
        if segment == "$" and is_math_token_or_inside:
            if not collecting_math:
                if buffer:
                    merged.append((buffer, False))
                buffer = "$"
                collecting_math = True
                math_mode = True
            else:
                buffer += "$"
                merged.append((buffer, True))
                buffer = ""
                collecting_math = False
                math_mode = False
        else:
            if collecting_math or math_mode or is_math_token_or_inside:
                buffer += segment
            else:
                buffer += segment

    if buffer:
        merged.append((buffer, collecting_math or math_mode))

    return merged


def normalize_inline_parenthesized_math(text: str) -> str:
    """Convert simple inline math copied as (q), (D(q)), (C'(q)), etc. into $...$.

    Important: this skips existing $...$ spans, so repeated normalization does
    not turn $D(q)$ into $D$q$$.
    """
    lines = text.splitlines()
    out: list[str] = []
    in_code = False
    in_display_math = False

    def replacement(match: re.Match[str]) -> str:
        inner = match.group(1).strip()
        # Avoid converting ordinary words or long prose.
        if SIMPLE_VAR_RE.fullmatch(inner) or MATH_SIGNAL_RE.search(inner):
            return f"${inner}$"
        return match.group(0)

    def process_text_segment(segment: str) -> str:
        return PAREN_CANDIDATE_RE.sub(replacement, segment)

    for line in lines:
        if re.match(r"^\s*(```|~~~)", line):
            in_code = not in_code
            out.append(line)
            continue

        if not in_code and line.strip() == "$$":
            in_display_math = not in_display_math
            out.append(line)
            continue

        if in_code or in_display_math:
            out.append(line)
        else:
            rebuilt: list[str] = []
            for segment, is_math in split_outside_inline_dollar_math(line):
                rebuilt.append(segment if is_math else process_text_segment(segment))
            out.append("".join(rebuilt))

    return "\n".join(out)


def normalize_canvas_markdown(text: str) -> str:
    """Normalize the math delimiters commonly produced by Canvas copy."""
    text = normalize_bare_bracket_display_math(text)
    text = normalize_inline_parenthesized_math(text)
    return text


def find_pandoc() -> str | None:
    env_path = os.environ.get("PANDOC_PATH")
    candidates = [
        env_path,
        shutil.which("pandoc"),
        "/opt/homebrew/bin/pandoc",
        "/usr/local/bin/pandoc",
        "/usr/bin/pandoc",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists() and os.access(candidate, os.X_OK):
            return str(candidate)
    return None


class MarkdownToDocxApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1160x780")
        self.minsize(900, 600)

        self.current_md_path: Path | None = None
        self.reference_docx_path: Path | None = None
        self.pandoc_path_var = tk.StringVar(value=find_pandoc() or "")
        self.normalize_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Ready")

        self._build_ui()
        self._update_status_pandoc()

    def _build_ui(self) -> None:
        toolbar = tk.Frame(self, padx=8, pady=6)
        toolbar.pack(side=tk.TOP, fill=tk.X)

        tk.Button(toolbar, text="Load .md…", command=self.load_markdown).pack(side=tk.LEFT, padx=3)
        tk.Button(toolbar, text="Paste", command=self.paste_clipboard).pack(side=tk.LEFT, padx=3)
        tk.Button(toolbar, text="Clear", command=self.clear_text).pack(side=tk.LEFT, padx=3)
        tk.Button(toolbar, text="Normalize preview", command=self.normalize_preview).pack(side=tk.LEFT, padx=3)
        tk.Button(toolbar, text="Save .md…", command=self.save_markdown).pack(side=tk.LEFT, padx=3)
        tk.Button(toolbar, text="Export .docx…", command=self.export_docx).pack(side=tk.LEFT, padx=3)

        tk.Checkbutton(
            toolbar,
            text="Normalize Canvas-copy math before saving/exporting",
            variable=self.normalize_var,
        ).pack(side=tk.LEFT, padx=12)

        pandoc_frame = tk.Frame(self, padx=8, pady=2)
        pandoc_frame.pack(side=tk.TOP, fill=tk.X)
        tk.Label(pandoc_frame, text="Pandoc:").pack(side=tk.LEFT)
        tk.Entry(pandoc_frame, textvariable=self.pandoc_path_var, width=68).pack(side=tk.LEFT, padx=5)
        tk.Button(pandoc_frame, text="Choose Pandoc…", command=self.choose_pandoc).pack(side=tk.LEFT, padx=3)
        tk.Button(pandoc_frame, text="Choose reference.docx…", command=self.choose_reference_docx).pack(side=tk.LEFT, padx=3)
        tk.Button(pandoc_frame, text="Clear reference", command=self.clear_reference_docx).pack(side=tk.LEFT, padx=3)

        help_text = (
            "Paste or load Markdown. If it came from ChatGPT Canvas copy, keep normalization enabled: "
            "bare [ ... ] math blocks become $$...$$, and simple inline (q) math becomes $q$."
        )
        tk.Label(self, text=help_text, anchor="w", justify="left", padx=8).pack(side=tk.TOP, fill=tk.X)

        self.text = ScrolledText(self, wrap=tk.WORD, undo=True, font=("Menlo", 12))
        self.text.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=8, pady=8)

        status = tk.Label(self, textvariable=self.status_var, anchor="w", padx=8, relief=tk.SUNKEN)
        status.pack(side=tk.BOTTOM, fill=tk.X)

    def _update_status_pandoc(self) -> None:
        path = self.pandoc_path_var.get().strip()
        if path and Path(path).exists():
            self.status_var.set(f"Pandoc found: {path}")
        else:
            self.status_var.set("Pandoc not found. Choose Pandoc executable or set PANDOC_PATH.")

    def get_source_text(self) -> str:
        return self.text.get("1.0", tk.END).rstrip("\n")

    def get_output_markdown(self) -> str:
        source = self.get_source_text()
        if self.normalize_var.get():
            return normalize_canvas_markdown(source)
        return source

    def load_markdown(self) -> None:
        path = filedialog.askopenfilename(
            title="Open Markdown file",
            filetypes=[("Markdown files", "*.md *.markdown *.txt"), ("All files", "*.*")],
        )
        if not path:
            return
        p = Path(path)
        try:
            content = p.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            content = p.read_text(encoding="utf-8-sig")
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Could not read file:\n{exc}")
            return

        self.text.delete("1.0", tk.END)
        self.text.insert("1.0", content)
        self.current_md_path = p
        self.status_var.set(f"Loaded: {p}")

    def paste_clipboard(self) -> None:
        try:
            content = self.clipboard_get()
        except tk.TclError:
            messagebox.showwarning(APP_TITLE, "Clipboard is empty or does not contain text.")
            return
        self.text.insert(tk.INSERT, content)
        self.status_var.set("Pasted clipboard text")

    def clear_text(self) -> None:
        if messagebox.askyesno(APP_TITLE, "Clear the editor contents?"):
            self.text.delete("1.0", tk.END)
            self.current_md_path = None
            self.status_var.set("Cleared")

    def normalize_preview(self) -> None:
        normalized = normalize_canvas_markdown(self.get_source_text())
        self.text.delete("1.0", tk.END)
        self.text.insert("1.0", normalized)
        self.status_var.set("Normalized Canvas-copy math in editor")

    def save_markdown(self) -> None:
        initial = self.current_md_path.name if self.current_md_path else "document.md"
        path = filedialog.asksaveasfilename(
            title="Save Markdown",
            defaultextension=".md",
            initialfile=initial,
            filetypes=[("Markdown files", "*.md"), ("Text files", "*.txt"), ("All files", "*.*")],
        )
        if not path:
            return
        p = Path(path)
        try:
            p.write_text(self.get_output_markdown() + "\n", encoding="utf-8")
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Could not save Markdown:\n{exc}")
            return
        self.current_md_path = p
        self.status_var.set(f"Saved Markdown: {p}")

    def choose_pandoc(self) -> None:
        path = filedialog.askopenfilename(title="Choose pandoc executable", filetypes=[("Executable", "*"), ("All files", "*.*")])
        if path:
            self.pandoc_path_var.set(path)
            self._update_status_pandoc()

    def choose_reference_docx(self) -> None:
        path = filedialog.askopenfilename(
            title="Choose Pandoc reference.docx",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
        )
        if path:
            self.reference_docx_path = Path(path)
            self.status_var.set(f"Reference DOCX selected: {path}")

    def clear_reference_docx(self) -> None:
        self.reference_docx_path = None
        self.status_var.set("Reference DOCX cleared")

    def export_docx(self) -> None:
        pandoc = self.pandoc_path_var.get().strip() or find_pandoc()
        if not pandoc or not Path(pandoc).exists():
            messagebox.showerror(
                APP_TITLE,
                "Pandoc was not found. Choose the Pandoc executable.\n\n"
                "On Apple Silicon Homebrew this is often:\n/opt/homebrew/bin/pandoc",
            )
            return
        self.pandoc_path_var.set(pandoc)

        out_path = filedialog.asksaveasfilename(
            title="Export DOCX",
            defaultextension=".docx",
            initialfile="document.docx",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
        )
        if not out_path:
            return

        markdown = self.get_output_markdown()
        if not markdown.strip():
            messagebox.showwarning(APP_TITLE, "There is no Markdown content to export.")
            return

        with tempfile.TemporaryDirectory(prefix="canvas_md_to_docx_") as tmpdir:
            tmp_md = Path(tmpdir) / "input.md"
            tmp_md.write_text(markdown + "\n", encoding="utf-8")

            cmd = [
                pandoc,
                "-f",
                PANDOC_FORMAT,
                str(tmp_md),
                "-o",
                str(out_path),
            ]
            if self.reference_docx_path:
                cmd.insert(-2, f"--reference-doc={self.reference_docx_path}")

            try:
                completed = subprocess.run(cmd, text=True, capture_output=True, check=False)
            except Exception as exc:
                messagebox.showerror(APP_TITLE, f"Could not run Pandoc:\n{exc}")
                return

        if completed.returncode != 0:
            messagebox.showerror(
                APP_TITLE,
                "Pandoc failed.\n\nCommand:\n"
                + " ".join(cmd)
                + "\n\nSTDERR:\n"
                + completed.stderr,
            )
            return

        self.status_var.set(f"Exported DOCX: {out_path}")
        messagebox.showinfo(APP_TITLE, f"DOCX exported successfully:\n{out_path}")


def main() -> None:
    app = MarkdownToDocxApp()
    app.mainloop()


if __name__ == "__main__":
    main()
