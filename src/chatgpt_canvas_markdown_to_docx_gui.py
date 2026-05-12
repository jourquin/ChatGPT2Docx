#!/usr/bin/env python3
"""
ChatGPT Canvas Markdown -> Markdown / DOCX / LaTeX GUI

This app is designed for Markdown copied from ChatGPT, ChatGPT Canvas or saved from Canvas.

It can optionally normalize the common Canvas-copy math form:

    [
    E = mc^2
    ]

into Pandoc-compatible display math:

    $$
    E = mc^2
    $$

It also converts standalone TeX-looking formula lines and bullet-only parenthesized formulas into display math, protects inline code spans, and converts simple inline math copied as parentheses, e.g. (q), (D(q)),
(C'(q)), (\beta), into $...$.

The app does NOT try to recover math content that has already been lost. It only
adds explicit Markdown math delimiters around TeX/math-looking text.

Requires:
    - Python 3 with Tkinter
    - Pandoc on PATH, in a common platform-specific install location,
      or selected manually in the GUI.
"""

from __future__ import annotations

import json
import os
import platform
import re
import shutil
import subprocess
import tempfile
import zipfile
import tkinter as tk
from xml.etree import ElementTree as ET
from pathlib import Path
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

APP_TITLE = "ChatGPT Markdown to DOCX/LaTeX v5.7"

PANDOC_FORMAT = "markdown+tex_math_dollars+tex_math_single_backslash"

APP_ID = "ChatGPT2Docx"
SETTINGS_FILENAME = "settings.json"


def get_settings_path() -> Path:
    """Return the per-user preferences file path.

    macOS:   ~/Library/Application Support/ChatGPT2Docx/settings.json
    Windows: %APPDATA%/ChatGPT2Docx/settings.json
    Linux:   ~/.config/ChatGPT2Docx/settings.json
    """
    system = platform.system()
    home = Path.home()

    if system == "Darwin":
        base = home / "Library" / "Application Support" / APP_ID
    elif system == "Windows":
        appdata = os.environ.get("APPDATA")
        base = Path(appdata) / APP_ID if appdata else home / "AppData" / "Roaming" / APP_ID
    else:
        config_home = os.environ.get("XDG_CONFIG_HOME")
        base = Path(config_home) / APP_ID if config_home else home / ".config" / APP_ID

    return base / SETTINGS_FILENAME


def load_settings() -> dict:
    """Load persisted GUI preferences, returning defaults if unavailable."""
    path = get_settings_path()
    defaults = {
        "pandoc_path": "",
        "reference_docx_path": "",
        "normalize_math": True,
        "shade_code_blocks": True,
    }

    try:
        if path.exists():
            data = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                defaults.update(data)
    except Exception:
        # Preferences should never prevent the app from launching.
        pass

    return defaults


def save_settings(settings: dict) -> None:
    """Persist GUI preferences as JSON."""
    path = get_settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(settings, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


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
                    out.extend(normalize_plain_formula_terms(line) for line in body)
                    out.append("$$")
                    i = j + 1
                    continue

        # One-line variant: [ E = mc^2 ]
        m = re.fullmatch(r"\s*\[\s*(.+?)\s*\]\s*", line)
        if m and looks_like_math(m.group(1)):
            out.append("$$")
            out.append(normalize_plain_formula_terms(m.group(1)))
            out.append("$$")
            i += 1
            continue

        out.append(line)
        i += 1

    return "\n".join(out)


# Strong signals that a full standalone line is intended as display math.
# This is deliberately stricter than looks_like_math(), because it operates on
# whole lines and must avoid wrapping normal prose.
STANDALONE_DISPLAY_MATH_RE = re.compile(
    r"""
    (?:\\(?:frac|sum|int|prod|lim|sqrt|boxed|left|right|begin|end|alpha|beta|gamma|delta|theta|lambda|mu|pi|sigma|top|cdot|times|in|leq|geq|neq|approx|text)\b)
    |
    (?:[A-Za-z][A-Za-z0-9']*(?:_[A-Za-z0-9{}*]+)?\s*=\s*.+(?:\\[A-Za-z]+|[_^{}]))
    """,
    re.VERBOSE,
)


def is_standalone_display_formula(line: str) -> bool:
    r"""Return True if a whole line is probably a display formula.

    This catches regular ChatGPT copy output such as:
        P_i = \frac{e^{V_i}}{\sum_{j \in C} e^{V_j}}

    It intentionally avoids prose, headings, bullets, tables, blockquotes, and
    already-delimited math.
    """
    s = line.strip()
    if not s:
        return False

    # Already math or structural Markdown.
    if s in {"$$", "[", "]"}:
        return False
    if s.startswith(("#", ">", "|", "```", "~~~", "$$", "$", r"\[", r"\(", "- ", "* ", "+ ")):
        return False
    if re.match(r"^\d+[.)]\s+", s):
        return False

    # Avoid ordinary prose and sentences. Formula lines are usually compact and
    # do not end with a colon. A final comma/full stop after a formula is common
    # in LaTeX prose, so do not reject commas/full stops universally.
    if s.endswith(":"):
        return False
    if " " in s and len(s.split()) > 8:
        return False

    return bool(STANDALONE_DISPLAY_MATH_RE.search(s))


def normalize_standalone_display_math_lines(text: str) -> str:
    """Wrap standalone TeX-looking formula lines in $$ blocks.

    This complements normalize_bare_bracket_display_math(). It is for regular
    ChatGPT answers copied with the Copy button, where the rendered display
    formula may become a plain line without any surrounding delimiter.
    """
    lines = text.splitlines()
    out: list[str] = []
    in_code = False
    in_display_math = False

    for line in lines:
        stripped = line.strip()

        if re.match(r"^\s*(```|~~~)", line):
            in_code = not in_code
            out.append(line)
            continue

        if not in_code and stripped == "$$":
            in_display_math = not in_display_math
            out.append(line)
            continue

        if in_code or in_display_math:
            out.append(line)
            continue

        if is_standalone_display_formula(line):
            out.append("$$")
            out.append(normalize_plain_formula_terms(line))
            out.append("$$")
        else:
            out.append(line)

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
                if is_math:
                    rebuilt.append(segment)
                else:
                    for part, protected in split_protected_inline_spans(segment):
                        rebuilt.append(part if protected else process_text_segment(part))
            out.append("".join(rebuilt))

    return "\n".join(out)




# Inline code spans such as `models.logit(V, av, i)` must not be treated as math.
# This regex intentionally supports only ordinary one-backtick spans, which is enough
# for ChatGPT copy output and avoids disturbing fenced code blocks.
INLINE_CODE_SPAN_RE = re.compile(r"`[^`\n]*`")

# A bullet whose whole content is a parenthesized formula is usually a display
# equation copied from rendered ChatGPT output, for example:
#     * ( P_i = \frac{...}{...} )
BULLET_PAREN_MATH_RE = re.compile(r"^\s*(?:[-*+] |\d+[.)]\s+)\(\s*(.+?)\s*\)\s*$")

# A line whose whole content is an inline code span may still be a rendered
# equation copied as code by the Copy button, for example:
#     `denominator = exp(V1) + exp(V2) + exp(V3)`
SOLE_CODE_SPAN_RE = re.compile(r"^\s*`([^`\n]+)`\s*$")


def normalize_exp_variables(expr: str) -> str:
    """Make plain copied expressions slightly more TeX-like.

    This is deliberately small in scope. It improves common copied formula lines
    such as ``denominator = exp(V1) + exp(V2)`` but does not touch real Python
    code blocks or inline code inside prose.
    """
    s = expr.strip()

    # exp(V1) -> \exp(V_1), exp(V_i) -> \exp(V_i)
    def exp_repl(match: re.Match[str]) -> str:
        arg = match.group(1).strip()
        arg = re.sub(r"\b([A-Za-z]+)(\d+)\b", r"\1_\2", arg)
        return rf"\exp({arg})"

    s = re.sub(r"\bexp\(([^()]+)\)", exp_repl, s)

    # denominator = ... -> \text{denominator} = ...
    m = re.match(r"^([A-Za-z][A-Za-z0-9_ ]*)\s*=\s*(.+)$", s)
    if m and "\\" in m.group(2):
        lhs = m.group(1).strip().replace("_", r"\_")
        s = rf"\text{{{lhs}}} = {m.group(2).strip()}"

    return normalize_plain_formula_terms(s)



def normalize_plain_formula_terms(expr: str) -> str:
    r"""Improve copied formula fragments before Pandoc sees them.

    Rendered ChatGPT/Canvas copy can degrade some TeX constructs.  In
    particular, underscores in subscripts may appear as Markdown-like asterisks:

        \hat{q}*{i,m}   ->   \hat{q}_{i,m}
        B^{%}*m         ->   B^{\%}_m

    This function only runs on text that has already been classified as formula
    content, not on ordinary prose or code blocks.
    """
    s = expr

    # Pandoc's math parser treats an unescaped percent as the start of a TeX
    # comment and may then leave the entire display equation as literal $$...$$.
    # Escape bare percentages inside recovered formulas: B^{%} -> B^{\%}.
    s = re.sub(r"(?<!\\)%", r"\\%", s)

    # ChatGPT copied/canvas math sometimes turns subscripts into asterisks.
    # Repair the most common forms without touching spaced multiplication.
    # Examples: q*{i,m} -> q_{i,m}; \hat{q}*{i,m} -> \hat{q}_{i,m}
    s = re.sub(r"(?<=[A-Za-z0-9}\)])\*\{([^{}\n]+)\}", r"_{\1}", s)

    # Example: B^{\%}*m -> B^{\%}_m.  Limit this to a preceding brace so that
    # normal products such as x*y are not rewritten as subscripts.
    s = re.sub(r"(?<=\})\*([A-Za-z][A-Za-z0-9]*)\b", r"_\1", s)

    # avail_j -> \text{avail}_j. Use a function replacement so backslashes are
    # preserved literally and are not interpreted as regex replacement escapes.
    s = re.sub(
        r"(?<!\\text\{)\bavail_([A-Za-z0-9]+)\b",
        lambda m: r"\text{avail}_" + m.group(1),
        s,
    )

    # A few whole-word cases occasionally appear in simple copied equations.
    s = re.sub(r"(?<!\\text\{)\bnavail\b", r"\text{navail}", s)
    s = re.sub(r"(?<![A-Za-z])qtot(?![A-Za-z])", r"q_{\text{tot}}", s)
    return s


def is_formulaish_code_span(content: str) -> bool:
    """True for a whole-line code span that is more formula than code."""
    s = content.strip()
    if not s or "=" not in s:
        return False

    # Avoid converting likely programming assignments with brackets, quotes,
    # method calls, or indexing. The denominator example is intentionally simple.
    if any(token in s for token in ["'", '"', "[", "]", ".", "df", "np."]):
        return False

    return bool(re.search(r"\bexp\([^()]+\)", s) or re.search(r"[A-Za-z]\d", s))


def normalize_sole_code_span_display_math(text: str) -> str:
    """Convert one-line code spans that are really formulas into display math."""
    lines = text.splitlines()
    out: list[str] = []
    in_code = False
    in_display_math = False

    for line in lines:
        stripped = line.strip()
        if re.match(r"^\s*(```|~~~)", line):
            in_code = not in_code
            out.append(line)
            continue
        if not in_code and stripped == "$$":
            in_display_math = not in_display_math
            out.append(line)
            continue
        if in_code or in_display_math:
            out.append(line)
            continue

        m = SOLE_CODE_SPAN_RE.match(line)
        if m and is_formulaish_code_span(m.group(1)):
            out.append("$$")
            out.append(normalize_exp_variables(m.group(1)))
            out.append("$$")
        else:
            out.append(line)

    return "\n".join(out)


def normalize_bullet_parenthesized_display_math(text: str) -> str:
    """Turn bullet-only parenthesized formulas into display math blocks."""
    lines = text.splitlines()
    out: list[str] = []
    in_code = False
    in_display_math = False

    for line in lines:
        stripped = line.strip()
        if re.match(r"^\s*(```|~~~)", line):
            in_code = not in_code
            out.append(line)
            continue
        if not in_code and stripped == "$$":
            in_display_math = not in_display_math
            out.append(line)
            continue
        if in_code or in_display_math:
            out.append(line)
            continue

        m = BULLET_PAREN_MATH_RE.match(line)
        if m:
            inner = m.group(1).strip()
            # Require a strong math signal so ordinary bullets like "* (optional)"
            # are not converted.
            if MATH_SIGNAL_RE.search(inner) or "\\frac" in inner or "\\exp" in inner:
                out.append("$$")
                out.append(normalize_plain_formula_terms(inner))
                out.append("$$")
                continue

        out.append(line)

    return "\n".join(out)


def split_protected_inline_spans(segment: str) -> list[tuple[str, bool]]:
    """Split a text segment into (part, protected) around inline code spans."""
    parts: list[tuple[str, bool]] = []
    pos = 0
    for match in INLINE_CODE_SPAN_RE.finditer(segment):
        if match.start() > pos:
            parts.append((segment[pos:match.start()], False))
        parts.append((match.group(0), True))
        pos = match.end()
    if pos < len(segment):
        parts.append((segment[pos:], False))
    if not parts:
        return [(segment, False)]
    return parts
def normalize_canvas_markdown(text: str) -> str:
    """Normalize math delimiters commonly lost in ChatGPT/Canvas copy."""
    text = normalize_bare_bracket_display_math(text)
    text = normalize_sole_code_span_display_math(text)
    text = normalize_bullet_parenthesized_display_math(text)
    text = normalize_standalone_display_math_lines(text)
    text = normalize_inline_parenthesized_math(text)
    return text


def is_usable_executable(candidate: str | os.PathLike[str] | None) -> bool:
    """Return True if candidate looks like a runnable executable.

    On Unix-like systems, os.access(..., X_OK) is the right test. On Windows,
    executable permissions are less meaningful, so an existing .exe/.bat/.cmd
    file should be accepted as a valid tool path.
    """
    if not candidate:
        return False

    path = Path(candidate).expanduser()
    if not path.exists() or not path.is_file():
        return False

    if platform.system() == "Windows":
        return path.suffix.lower() in {".exe", ".bat", ".cmd", ".com"} or os.access(path, os.X_OK)

    return os.access(path, os.X_OK)


def find_pandoc() -> str | None:
    """Find Pandoc in PATH or in common macOS, Windows, and Linux locations."""
    env_path = os.environ.get("PANDOC_PATH")
    system = platform.system()

    candidates: list[str | None] = [
        env_path,
        shutil.which("pandoc"),
        shutil.which("pandoc.exe"),
    ]

    if system == "Darwin":
        candidates.extend([
            "/opt/homebrew/bin/pandoc",      # Apple Silicon Homebrew
            "/usr/local/bin/pandoc",        # Intel Homebrew
            "/opt/local/bin/pandoc",        # MacPorts
            "/usr/bin/pandoc",
        ])
    elif system == "Windows":
        program_files = os.environ.get("ProgramFiles")
        program_files_x86 = os.environ.get("ProgramFiles(x86)")
        local_appdata = os.environ.get("LOCALAPPDATA")
        chocolatey = os.environ.get("ChocolateyInstall")
        home = Path.home()

        candidates.extend([
            str(Path(program_files) / "Pandoc" / "pandoc.exe") if program_files else None,
            str(Path(program_files_x86) / "Pandoc" / "pandoc.exe") if program_files_x86 else None,
            str(Path(local_appdata) / "Pandoc" / "pandoc.exe") if local_appdata else None,
            str(Path(chocolatey) / "bin" / "pandoc.exe") if chocolatey else None,
            str(home / "scoop" / "shims" / "pandoc.exe"),
            str(home / "scoop" / "apps" / "pandoc" / "current" / "pandoc.exe"),
        ])
    else:
        candidates.extend([
            "/usr/bin/pandoc",
            "/usr/local/bin/pandoc",
            "/snap/bin/pandoc",
            "/opt/pandoc/bin/pandoc",
        ])

    seen: set[str] = set()
    for candidate in candidates:
        if not candidate:
            continue
        normalized = str(Path(candidate).expanduser())
        if normalized in seen:
            continue
        seen.add(normalized)
        if is_usable_executable(normalized):
            return normalized
    return None




def shade_docx_code_blocks(docx_path: str | Path, fill: str = "F2F2F2") -> None:
    """Apply a light gray background to Pandoc fenced code block paragraphs.

    Pandoc maps fenced code block lines to paragraphs with style SourceCode.
    Word does not always display those paragraphs in a gray box by default, so
    this directly adds paragraph shading to those paragraphs in document.xml.
    """
    path = Path(docx_path)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    ET.register_namespace("w", ns["w"])
    ET.register_namespace("m", "http://schemas.openxmlformats.org/officeDocument/2006/math")
    ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

    with tempfile.TemporaryDirectory(prefix="docx_code_shade_") as tmpdir:
        tmp = Path(tmpdir)
        with zipfile.ZipFile(path, "r") as zin:
            zin.extractall(tmp)

        document_xml = tmp / "word" / "document.xml"
        tree = ET.parse(document_xml)
        root = tree.getroot()
        changed = False

        for p in root.findall(".//w:p", ns):
            p_pr = p.find("w:pPr", ns)
            if p_pr is None:
                continue
            p_style = p_pr.find("w:pStyle", ns)
            if p_style is None or p_style.get(f"{{{ns['w']}}}val") != "SourceCode":
                continue

            shd = p_pr.find("w:shd", ns)
            if shd is None:
                shd = ET.SubElement(p_pr, f"{{{ns['w']}}}shd")
            shd.set(f"{{{ns['w']}}}val", "clear")
            shd.set(f"{{{ns['w']}}}color", "auto")
            shd.set(f"{{{ns['w']}}}fill", fill)

            spacing = p_pr.find("w:spacing", ns)
            if spacing is None:
                spacing = ET.SubElement(p_pr, f"{{{ns['w']}}}spacing")
            spacing.set(f"{{{ns['w']}}}before", "0")
            spacing.set(f"{{{ns['w']}}}after", "0")
            changed = True

        if not changed:
            return

        tree.write(document_xml, encoding="UTF-8", xml_declaration=True)

        tmp_docx = path.with_suffix(path.suffix + ".tmp")
        with zipfile.ZipFile(tmp_docx, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for file in tmp.rglob("*"):
                if file.is_file():
                    zout.write(file, file.relative_to(tmp).as_posix())
        tmp_docx.replace(path)


def add_latex_code_shading_args(cmd: list[str]) -> None:
    """Add Pandoc options that make fenced code blocks shaded in LaTeX output.

    Pandoc's default LaTeX highlighting style on many installations is
    ``pygments``, which emits a Shaded environment but defines it as empty.
    The ``tango`` style defines ``shadecolor`` and wraps highlighted code in
    a ``snugshade`` environment, producing a light gray background in the
    generated .tex file and in PDFs compiled from it.

    We use the widely supported ``--highlight-style`` option for compatibility
    with older Pandoc versions. Newer Pandoc versions may describe this as an
    alias for ``--syntax-highlighting``.
    """
    cmd.extend(["--highlight-style=tango"])


class MarkdownToDocxApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1160x780")
        self.minsize(900, 600)

        self.settings = load_settings()

        saved_pandoc = str(self.settings.get("pandoc_path") or "").strip()
        pandoc_initial = saved_pandoc if saved_pandoc else (find_pandoc() or "")

        saved_reference = str(self.settings.get("reference_docx_path") or "").strip()
        reference_path = Path(saved_reference).expanduser() if saved_reference else None

        self.current_md_path: Path | None = None
        # Keep a separate displayed path variable so the saved path is visible
        # at launch, even before the user exports anything.
        self.reference_docx_path: Path | None = reference_path if reference_path and reference_path.exists() else None
        self.pandoc_path_var = tk.StringVar(value=pandoc_initial)
        self.reference_docx_var = tk.StringVar(value=saved_reference)
        self.normalize_var = tk.BooleanVar(value=bool(self.settings.get("normalize_math", True)))
        self.shade_code_var = tk.BooleanVar(value=bool(self.settings.get("shade_code_blocks", True)))
        self.status_var = tk.StringVar(value="Ready")

        self._build_ui()
        self._update_status_pandoc()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def _build_ui(self) -> None:
        # Row 1: primary actions. Keep this row short; options live below so
        # long checkbox labels do not get clipped on smaller windows.
        toolbar = tk.Frame(self, padx=8, pady=6)
        toolbar.pack(side=tk.TOP, fill=tk.X)

        tk.Button(toolbar, text="Load .md…", command=self.load_markdown).pack(side=tk.LEFT, padx=3)
        tk.Button(toolbar, text="Paste", command=self.paste_clipboard).pack(side=tk.LEFT, padx=3)
        tk.Button(toolbar, text="Clear", command=self.clear_text).pack(side=tk.LEFT, padx=3)
        tk.Button(toolbar, text="Normalize preview", command=self.normalize_preview).pack(side=tk.LEFT, padx=14)
        tk.Button(toolbar, text="Save .md…", command=self.save_markdown).pack(side=tk.LEFT, padx=3)
        tk.Button(toolbar, text="Export .docx…", command=self.export_docx).pack(side=tk.LEFT, padx=3)
        tk.Button(toolbar, text="Export .tex…", command=self.export_tex).pack(side=tk.LEFT, padx=3)

        # Row 2: options. Short labels prevent clipping while remaining clear.
        options_frame = tk.LabelFrame(self, text="Options", padx=8, pady=4)
        options_frame.pack(side=tk.TOP, fill=tk.X, padx=8, pady=4)

        tk.Checkbutton(
            options_frame,
            text="Normalize ChatGPT/Canvas copy math",
            variable=self.normalize_var,
        ).pack(side=tk.LEFT, padx=12)

        tk.Checkbutton(
            options_frame,
            text="Shade code blocks (DOCX/LaTeX)",
            variable=self.shade_code_var,
        ).pack(side=tk.LEFT, padx=12)

        # Row 3: Pandoc executable.
        pandoc_frame = tk.Frame(self, padx=8, pady=2)
        pandoc_frame.pack(side=tk.TOP, fill=tk.X)
        tk.Label(pandoc_frame, text="Pandoc:").pack(side=tk.LEFT)
        tk.Entry(pandoc_frame, textvariable=self.pandoc_path_var, width=62).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        tk.Button(pandoc_frame, text="Choose Pandoc…", command=self.choose_pandoc).pack(side=tk.LEFT, padx=3)

        # Row 4: optional Word reference style file. Show the selected path so
        # restored preferences are visible at launch, similarly to the Pandoc path.
        reference_frame = tk.Frame(self, padx=8, pady=2)
        reference_frame.pack(side=tk.TOP, fill=tk.X)
        tk.Label(reference_frame, text="Reference .docx:").pack(side=tk.LEFT)
        tk.Entry(reference_frame, textvariable=self.reference_docx_var, width=62).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        tk.Button(reference_frame, text="Choose reference.docx…", command=self.choose_reference_docx).pack(side=tk.LEFT, padx=3)
        tk.Button(reference_frame, text="Clear reference", command=self.clear_reference_docx).pack(side=tk.LEFT, padx=3)

        help_text = (
            "Paste or load Markdown. Enable normalization for ChatGPT/Canvas copy: "
            "bare [ ... ] blocks, bullet-only formulas, and standalone TeX lines become $$...$$; "
            "inline code spans are protected; simple inline (q) math becomes $q$. "
            "DOCX export can use reference.docx. TeX export creates a standalone LaTeX file; when code shading is enabled, Pandoc uses the tango highlighting style so fenced code blocks get a light background."
        )
        self.help_label = tk.Label(self, text=help_text, anchor="w", justify="left", padx=8, wraplength=1050)
        self.help_label.pack(side=tk.TOP, fill=tk.X)
        self.bind("<Configure>", self._update_help_wraplength)

        self.text = ScrolledText(self, wrap=tk.WORD, undo=True, font=("Menlo", 12))
        self.text.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=8, pady=8)

        status = tk.Label(self, textvariable=self.status_var, anchor="w", padx=8, relief=tk.SUNKEN)
        status.pack(side=tk.BOTTOM, fill=tk.X)

    def _update_help_wraplength(self, event: tk.Event | None = None) -> None:
        # Keep the help text readable and prevent horizontal clipping.
        width = max(400, self.winfo_width() - 40)
        if hasattr(self, "help_label"):
            self.help_label.configure(wraplength=width)

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
            self.save_preferences(silent=True)

    def choose_reference_docx(self) -> None:
        path = filedialog.askopenfilename(
            title="Choose Pandoc reference.docx",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
        )
        if path:
            self.reference_docx_path = Path(path)
            self.reference_docx_var.set(path)
            self.status_var.set(f"Reference DOCX selected: {path}")
            self.save_preferences(silent=True)

    def clear_reference_docx(self) -> None:
        self.reference_docx_path = None
        self.reference_docx_var.set("")
        self.status_var.set("Reference DOCX cleared")
        self.save_preferences(silent=True)

    def get_reference_docx_path(self, *, warn: bool = True) -> Path | None:
        """Return the reference .docx path from the visible entry, if valid."""
        value = self.reference_docx_var.get().strip()
        if not value:
            self.reference_docx_path = None
            return None

        path = Path(value).expanduser()
        if not path.exists():
            self.reference_docx_path = None
            if warn:
                messagebox.showwarning(
                    APP_TITLE,
                    "The selected reference .docx file was not found and will be ignored:\n" + value,
                )
            return None

        self.reference_docx_path = path
        return path

    def save_preferences(self, silent: bool = False) -> None:
        settings = {
            "pandoc_path": self.pandoc_path_var.get().strip(),
            "reference_docx_path": self.reference_docx_var.get().strip(),
            "normalize_math": bool(self.normalize_var.get()),
            "shade_code_blocks": bool(self.shade_code_var.get()),
        }
        try:
            save_settings(settings)
            if not silent:
                self.status_var.set(f"Preferences saved: {get_settings_path()}")
        except Exception as exc:
            if silent:
                # Do not interrupt routine UI actions because preferences failed.
                return
            messagebox.showwarning(APP_TITLE, f"Could not save preferences:\n{exc}")

    def on_close(self) -> None:
        self.save_preferences(silent=True)
        self.destroy()

    def export_docx(self) -> None:
        pandoc = self.pandoc_path_var.get().strip() or find_pandoc()
        if not pandoc or not Path(pandoc).exists():
            messagebox.showerror(
                APP_TITLE,
                "Pandoc was not found. Choose the Pandoc executable.\n\n"
                "Examples:\n"
                "macOS Homebrew: /opt/homebrew/bin/pandoc or /usr/local/bin/pandoc\n"
                "Windows: C:\\Program Files\\Pandoc\\pandoc.exe\n"
                "Linux: /usr/bin/pandoc or /usr/local/bin/pandoc",
            )
            return
        self.pandoc_path_var.set(pandoc)
        self.save_preferences(silent=True)

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
            reference_docx = self.get_reference_docx_path(warn=True)
            if reference_docx:
                cmd.insert(-2, f"--reference-doc={reference_docx}")

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

        if self.shade_code_var.get():
            try:
                shade_docx_code_blocks(out_path)
            except Exception as exc:
                messagebox.showwarning(
                    APP_TITLE,
                    "DOCX was exported, but code-block shading failed:\n" + str(exc),
                )

        self.status_var.set(f"Exported DOCX: {out_path}")
        messagebox.showinfo(APP_TITLE, f"DOCX exported successfully:\n{out_path}")

    def export_tex(self) -> None:
        pandoc = self.pandoc_path_var.get().strip() or find_pandoc()
        if not pandoc or not Path(pandoc).exists():
            messagebox.showerror(
                APP_TITLE,
                "Pandoc was not found. Choose the Pandoc executable.\n\n"
                "Examples:\n"
                "macOS Homebrew: /opt/homebrew/bin/pandoc or /usr/local/bin/pandoc\n"
                "Windows: C:\\Program Files\\Pandoc\\pandoc.exe\n"
                "Linux: /usr/bin/pandoc or /usr/local/bin/pandoc",
            )
            return
        self.pandoc_path_var.set(pandoc)
        self.save_preferences(silent=True)

        out_path = filedialog.asksaveasfilename(
            title="Export LaTeX",
            defaultextension=".tex",
            initialfile="document.tex",
            filetypes=[("LaTeX files", "*.tex"), ("All files", "*.*")],
        )
        if not out_path:
            return

        markdown = self.get_output_markdown()
        if not markdown.strip():
            messagebox.showwarning(APP_TITLE, "There is no Markdown content to export.")
            return

        with tempfile.TemporaryDirectory(prefix="canvas_md_to_tex_") as tmpdir:
            tmp_md = Path(tmpdir) / "input.md"
            tmp_md.write_text(markdown + "\n", encoding="utf-8")

            cmd = [
                pandoc,
                "-f",
                PANDOC_FORMAT,
                "-t",
                "latex",
                "-s",
                str(tmp_md),
                "-o",
                str(out_path),
            ]
            if self.shade_code_var.get():
                add_latex_code_shading_args(cmd)

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

        self.status_var.set(f"Exported LaTeX: {out_path}")
        messagebox.showinfo(APP_TITLE, f"LaTeX exported successfully:\n{out_path}")


def main() -> None:
    app = MarkdownToDocxApp()
    app.mainloop()


if __name__ == "__main__":
    main()
