# ChatGPT Canvas Markdown to DOCX GUI

A small Python/Tkinter desktop application for converting Markdown copied from **ChatGPT Canvas** into:

- a normalized Markdown `.md` file; and/or
- a Word `.docx` file generated with Pandoc.

The app is designed for users who write technical or academic text in ChatGPT Canvas and want to preserve Markdown structure and LaTeX-style mathematical formulas when exporting to Word.

---

## Why this tool exists

Markdown copied from a rendered ChatGPT answer or from Canvas may not always contain math delimiters in a form that Pandoc recognizes directly.

For example, Canvas-copied display math can sometimes appear as:

```markdown
[
D(q^*) = C'(q^*)
]
```

Pandoc does **not** treat bare `[` and `]` as display-math delimiters. It expects something like:

```markdown
$$
D(q^*) = C'(q^*)
$$
```

or:

```markdown
\[
D(q^*) = C'(q^*)
\]
```

The GUI can optionally normalize these Canvas-copy math blocks before saving or exporting.

It can also convert simple inline math copied as plain parentheses, such as:

```markdown
(q), (p), (D(q)), (C'(q)), (\beta)
```

into Pandoc-compatible inline math:

```markdown
$q$, $p$, $D(q)$, $C'(q)$, $\beta$
```

---

## Features

- Paste Markdown directly from ChatGPT Canvas.
- Open an existing `.md` file.
- Save the current content as Markdown.
- Export the content directly to `.docx` using Pandoc.
- Optional normalization of Canvas-copied math syntax.
- Idempotent normalization: running normalization more than once should not corrupt existing `$...$` math.
- Automatic Pandoc detection on common macOS Homebrew paths:
  - `/opt/homebrew/bin/pandoc`
  - `/usr/local/bin/pandoc`
- Manual Pandoc executable selection.
- Optional Pandoc `reference.docx` support for custom Word styles.

---

## Requirements

### 1. Python 3

The application uses Python 3 and the standard `tkinter` GUI library.

Check your Python version:

```bash
python3 --version
```

### 2. Tkinter

Tkinter is included with many Python installations.

On macOS, the Python installer from python.org usually includes Tkinter. Homebrew Python may also work, depending on your setup.

On Debian/Ubuntu Linux, you may need:

```bash
sudo apt install python3-tk
```

### 3. Pandoc

Pandoc must be installed to export `.docx` files.

Check whether Pandoc is available:

```bash
pandoc --version
```

On macOS with Homebrew:

```bash
brew install pandoc
```

On Windows, install Pandoc from the official installer and make sure it is available on your system `PATH`.

---

## Installation

Clone or download this repository, then place the script somewhere convenient:

```bash
git clone https://github.com/YOUR-USER/YOUR-REPO.git
cd YOUR-REPO
```

Or simply download the Python file:

```text
chatgpt_canvas_markdown_to_docx_gui_v3.py
```

No extra Python packages are required.

---

## Running the app

From a terminal:

```bash
python3 chatgpt_canvas_markdown_to_docx_gui.py
```

If your system uses `python` instead of `python3`:

```bash
python chatgpt_canvas_markdown_to_docx_gui.py
```

---

## Basic workflow

1. Open ChatGPT Canvas.
2. Copy the Canvas Markdown content.
3. Launch the GUI.
4. Paste the Markdown into the text area.
5. Leave **Normalize Canvas-copy math before saving/exporting** enabled if the copied text contains formulas like:

   ```markdown
   [
   E = mc^2
   ]
   ```

6. Click **Save .md** to save normalized Markdown.
7. Click **Export .docx** to generate a Word document through Pandoc.

---

## Pandoc conversion used by the app

The app exports DOCX using Pandoc with this input format:

```text
markdown+tex_math_dollars+tex_math_single_backslash
```

Conceptually, the command is equivalent to:

```bash
pandoc -f markdown+tex_math_dollars+tex_math_single_backslash input.md -o output.docx
```

This allows Pandoc to recognize both dollar-style math:

```markdown
Inline math: $P_i$

Display math:

$$
P_i = \frac{e^{V_i}}{\sum_{j \in C} e^{V_j}}
$$
```

and single-backslash LaTeX delimiters:

```markdown
Inline math: \(P_i\)

Display math:

\[
P_i = \frac{e^{V_i}}{\sum_{j \in C} e^{V_j}}
\]
```

---

## Using a reference DOCX

Pandoc can use a `reference.docx` file to control the formatting of the generated Word document.

A reference DOCX does **not** provide content. It provides Word styles, such as:

- `Normal`
- `Heading 1`
- `Heading 2`
- table styles
- bullet and numbered list styles
- margins
- fonts
- line spacing

To generate Pandoc's default reference file:

```bash
pandoc --print-default-data-file reference.docx > reference.docx
```

Open `reference.docx` in Microsoft Word, modify the styles, save it, and then select it in the GUI before exporting.

The export is then conceptually equivalent to:

```bash
pandoc input.md --reference-doc=reference.docx -o output.docx
```

---

## What the normalization does

When enabled, the normalization step converts common Canvas-copy math forms into Pandoc-compatible math.

### Display math

Input:

```markdown
[
B(q)=\int_0^q D(x)\,dx
]
```

Output:

```markdown
$$
B(q)=\int_0^q D(x)\,dx
$$
```

### Inline math

Input:

```markdown
The optimal traffic level is (q^*) and the price is (p^*).
```

Output:

```markdown
The optimal traffic level is $q^*$ and the price is $p^*$.
```

Input:

```markdown
The inverse demand is (D(q)) and marginal cost is (C'(q)).
```

Output:

```markdown
The inverse demand is $D(q)$ and marginal cost is $C'(q)$.
```

---

## What the normalization does not do

The app does **not** reconstruct formulas whose content has already been lost.

For example, if a failed copy/export produced:

```text
P_i =
```

instead of:

```markdown
P_i = \frac{e^{V_i}}{\sum_{j \in C} e^{V_j}}
```

then the app cannot infer the missing formula. It only adds correct Markdown math delimiters around math-like content that is still present.

The normalization is intentionally conservative. Some unusual mathematical expressions may need manual adjustment.

---

## Recommended prompts for ChatGPT

For best results, ask ChatGPT to write documents in Canvas using explicit Pandoc-friendly math delimiters.

Example prompt:

```text
Write this in a Canvas as Markdown. Use $...$ for inline math and $$...$$ for display math so that the saved Markdown can be converted locally with Pandoc to DOCX.
```

Or:

```text
Write this in a Canvas as Markdown. Use \(...\) for inline math and \[...\] for display math so that Pandoc can convert it to DOCX.
```

If you are copying from a normal rendered ChatGPT answer rather than Canvas, ask for raw Markdown in a fenced code block:

```text
Give me the answer as raw Markdown in a fenced code block. Use $...$ for inline math and $$...$$ for display math. Do not render the equations.
```

---

## Troubleshooting

### The app cannot find Pandoc

If Pandoc works in your terminal but not in the GUI, it may be a `PATH` issue, especially on macOS.

Check where Pandoc is installed:

```bash
which pandoc
```

Common locations are:

```text
/opt/homebrew/bin/pandoc
/usr/local/bin/pandoc
```

The app checks those paths automatically. If it still does not find Pandoc, use the GUI button to manually choose the Pandoc executable.

You can also launch the app from the terminal with:

```bash
PANDOC_PATH=/opt/homebrew/bin/pandoc python3 chatgpt_canvas_markdown_to_docx_gui_v3.py
```

### The DOCX still shows raw dollar signs

This usually means one of the following:

- the Markdown was already partially corrupted before import;
- math delimiters were nested incorrectly;
- normalization was disabled even though Canvas-copy math needed it;
- an expression was too unusual for the conservative normalizer.

Try saving the normalized Markdown first, inspect the `.md`, then export again.

### The DOCX has equations, but the formatting is not what I want

Use a Pandoc `reference.docx` file and customize Word styles such as `Normal`, `Heading 1`, and `Heading 2`.

### Accented characters are wrong

Make sure the Markdown file is saved as UTF-8.

---

## Limitations

- The app focuses on Markdown-to-DOCX conversion through Pandoc.
- It is not a full Markdown editor.
- It does not preview rendered Markdown or rendered equations.
- It does not repair formulas whose TeX content has already been deleted.
- The inline math normalization is heuristic and may require manual correction in complex cases.

---

## Development note: yes, ChatGPT helped

This script was developed with help from ChatGPT itself, which feels a bit like asking a photocopier to design a better photocopier. The goal was simple: take Markdown copied from ChatGPT Canvas, stop mathematical formulas from escaping into the wilderness, and produce a cleaner path to Word documents via Pandoc.

ChatGPT helped sketch the first version, debug the macOS Pandoc path issue, improve the math-normalization logic, and document the workflow. Human supervision, testing, and common sense were still required, especially when the early versions cheerfully produced artifacts such as `q$$`.

In other words: this is a small tool made by a human, with assistance from an AI, to solve a problem created by copying text from an AI. Perfectly normal modern software development.

## Suggested file names

For clarity, a typical project can use:

```text
input.md                 # original Markdown copied or saved from Canvas
input_normalized.md      # Markdown after normalization
output.docx              # Word document exported with Pandoc
reference.docx           # optional Word style template
```

---

## License

This project is released under the **MIT License**.

You are free to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the software, provided that the MIT License notice is included in copies or substantial portions of the software.



