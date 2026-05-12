# ChatGPT Markdown to DOCX/LaTeX GUI

A small Python/Tkinter desktop application that converts Markdown copied from ChatGPT into clean export files.

It is especially useful when a ChatGPT answer contains **LaTeX-style mathematical formulas** that look correct on screen but lose their delimiters when copied.

The app can:

- paste or load Markdown;
- normalize common ChatGPT/Canvas-copy math glitches;
- save corrected Markdown as `.md`;
- export directly to Word `.docx` through Pandoc;
- export to standalone LaTeX `.tex` through Pandoc;
- optionally shade fenced code blocks in DOCX and LaTeX outputs;
- repair additional formula-copy artifacts such as `\hat{q}*{i,m}` and `B^{%}*m`;
- remember your Pandoc path, `reference.docx` path, and checkbox preferences between sessions;
- run on macOS, Windows, and Linux, provided Python/Tkinter and Pandoc are installed.

---

## Why this tool exists

ChatGPT can display formulas beautifully in the browser, but the text copied with the regular **Copy** button is not always valid Markdown math.

For example, a rendered formula may be copied as:

```markdown
[
P_i = \frac{e^{V_i}}{\sum_{j \in C} e^{V_j}}
]
```

or even as a standalone line:

```markdown
P_i = \frac{e^{V_i}}{\sum_{j \in C} e^{V_j}}
```

Some Canvas copies can also degrade subscripts or percent signs inside formulas, for example:

```markdown
B_m=\sum_i(\hat{q}*{i,m}-q*{i,m}), \qquad B^{%}*m=...
```

Pandoc does not automatically know that these are display equations or that `*{i,m}` was meant to be a subscript. It expects explicit math delimiters such as:

```markdown
$$
P_i = \frac{e^{V_i}}{\sum_{j \in C} e^{V_j}}
$$
```

This app provides a practical normalization step before exporting.

---

## Features

### Markdown input

- Paste Markdown from the clipboard.
- Load a `.md` file.
- Edit the content in a simple text area.
- Save the current content as `.md`.

### Math normalization

When enabled, the app attempts to repair common copy artifacts from ChatGPT and ChatGPT Canvas.

It can convert bare bracket display math:

```markdown
[
D(q^*) = C'(q^*)
]
```

into:

```markdown
$$
D(q^*) = C'(q^*)
$$
```

It can also convert standalone TeX-looking formula lines:

```markdown
P_i = \frac{\exp(V_i)}{\sum_j \text{avail}_j \exp(V_j)}
```

into:

```markdown
$$
P_i = \frac{\exp(V_i)}{\sum_j \text{avail}_j \exp(V_j)}
$$
```

It can convert simple inline math copied with parentheses:

```markdown
Here, (P_i) is the probability, (x_i) is a vector, and (\beta) is estimated.
```

into:

```markdown
Here, $P_i$ is the probability, $x_i$ is a vector, and $\beta$ is estimated.
```

It also handles bullet-only formulas such as:

```markdown
* ( P_i = \frac{\exp(V_i)}{\sum_j \text{avail}_j \exp(V_j)} )
```

by turning them into display equations.

The latest heuristic also repairs two common Canvas-copy artifacts inside recovered formulas:

- copied subscripts written as `*{...}` instead of `_{...}`;
- unescaped percent signs in math, such as `B^{%}`, which can break TeX parsing.

For example:

```markdown
B_m=\sum_i(\hat{q}*{i,m}-q*{i,m}), \qquad B^{%}*m=...
```

becomes:

```markdown
B_m=\sum_i(\hat{q}_{i,m}-q_{i,m}), \qquad B^{\%}_m=...
```

### Code protection

The normalizer protects:

- fenced code blocks;
- inline code spans such as `` `models.logit(V, av, i)` ``.

This avoids converting Python function calls or code parentheses into math.

### DOCX export

The app exports to `.docx` using Pandoc.

Conceptually, it runs:

```bash
pandoc -f markdown+tex_math_dollars+tex_math_single_backslash input.md -o output.docx
```

It can also use an optional Pandoc `reference.docx` file to control Word styles.

### LaTeX export

The app can export the same normalized Markdown to a standalone `.tex` file.

Conceptually, it runs:

```bash
pandoc -f markdown+tex_math_dollars+tex_math_single_backslash \
  -t latex \
  -s input.md \
  -o output.tex
```

### Code-block shading

The option **Shade code blocks (DOCX/LaTeX)** applies to both export formats.

For DOCX, the app post-processes Pandoc’s generated `.docx` and adds a light gray background to paragraphs using the `SourceCode` style.

For LaTeX, the app adds:

```bash
--highlight-style=tango
```

This makes Pandoc generate shaded environments for fenced code blocks, so code appears with a light background when the `.tex` file is compiled to PDF.

Code shading works best when fenced code blocks include a language name:

````markdown
```python
df['avail1'] = df['cost1'].notna().astype(int)
```
````

### Saved preferences

The app remembers your settings between sessions:

- Pandoc executable path;
- `reference.docx` path;
- **Normalize ChatGPT/Canvas copy math** checkbox state;
- **Shade code blocks (DOCX/LaTeX)** checkbox state.

The preferences are saved when the app closes. The Pandoc and `reference.docx` paths are also saved immediately when you choose or clear them.

Settings are stored in a small JSON file:

| Platform | Settings file |
|---|---|
| macOS | `~/Library/Application Support/ChatGPT2Docx/settings.json` |
| Windows | `%APPDATA%\ChatGPT2Docx\settings.json` |
| Linux | `~/.config/ChatGPT2Docx/settings.json` |

---

## Requirements

### Python 3

The app is written in Python and uses the standard `tkinter` GUI toolkit.

Check your Python version:

```bash
python3 --version
```

or on Windows:

```powershell
py -3 --version
```

### Tkinter

Tkinter is included with many Python installations.

On macOS, the Python installer from python.org usually includes Tkinter. On Windows, the Python installer from python.org normally includes Tkinter as well.

On Debian/Ubuntu, you may need:

```bash
sudo apt install python3-tk
```

On Fedora:

```bash
sudo dnf install python3-tkinter
```

On Arch Linux:

```bash
sudo pacman -S tk
```

### Pandoc

Pandoc is required for `.docx` and `.tex` export.

Check whether Pandoc is available:

```bash
pandoc --version
```

On macOS with Homebrew:

```bash
brew install pandoc
```

On Debian/Ubuntu:

```bash
sudo apt install pandoc
```

On Windows, install Pandoc from the official installer or a package manager, then use **Choose Pandoc…** in the app if it is not detected automatically.

The app automatically checks common locations:

```text
macOS:
  /opt/homebrew/bin/pandoc
  /usr/local/bin/pandoc

Windows:
  C:\Program Files\Pandoc\pandoc.exe
  C:\Program Files (x86)\Pandoc\pandoc.exe

Linux:
  /usr/bin/pandoc
  /usr/local/bin/pandoc
  /snap/bin/pandoc
  /opt/pandoc/bin/pandoc
```

You can always manually select the Pandoc executable from the GUI.

---

## Running the app

### macOS and Linux

From a terminal:

```bash
python3 chatgpt_markdown_to_docx_gui.py
```

### Windows

From PowerShell or Command Prompt:

```powershell
py -3 chatgpt_markdown_to_docx_gui.py
```

or, if Python is directly on your `PATH`:

```powershell
python chatgpt_markdown_to_docx_gui.py
```

---

## Basic workflow

1. Copy a ChatGPT answer or ChatGPT Canvas content.
2. Launch the app.
3. Click **Paste**, or load an existing `.md` file with **Load .md…**.
4. Keep **Normalize ChatGPT/Canvas copy math** enabled if the copied text contains formula artifacts.
5. Optionally click **Normalize preview** to inspect the corrected Markdown.
6. Click **Save .md…** to save normalized Markdown.
7. Click **Export .docx…** to create a Word document.
8. Click **Export .tex…** to create a standalone LaTeX file.

---

## GUI controls

### Main buttons

| Button | Purpose |
|---|---|
| **Load .md…** | Open an existing Markdown file. |
| **Paste** | Paste clipboard content into the editor. |
| **Clear** | Clear the editor. |
| **Normalize preview** | Replace the editor content with normalized Markdown so you can inspect it. |
| **Save .md…** | Save the current content as Markdown. |
| **Export .docx…** | Export to Word through Pandoc. |
| **Export .tex…** | Export to standalone LaTeX through Pandoc. |

### Options

| Option | Purpose |
|---|---|
| **Normalize ChatGPT/Canvas copy math** | Repairs common copied-math patterns before saving/exporting. |
| **Shade code blocks (DOCX/LaTeX)** | Adds light background shading to fenced code blocks in DOCX and LaTeX outputs. |

### Pandoc and reference DOCX controls

| Control | Purpose |
|---|---|
| **Pandoc:** path field | Shows the Pandoc executable currently used by the app. |
| **Choose Pandoc…** | Manually select the Pandoc executable. |
| **Reference .docx:** path field | Shows the selected Word style template, if any. |
| **Choose reference.docx…** | Select a Word style template for DOCX export. |
| **Clear reference** | Remove the selected reference DOCX. |

---

## Using a reference DOCX

A Pandoc `reference.docx` is a Word file that controls the styling of the generated `.docx`.

It does **not** provide content. It provides styles such as:

- `Normal`;
- `Heading 1`;
- `Heading 2`;
- list styles;
- table styles;
- code styles;
- page margins;
- fonts.

You can generate Pandoc’s default reference file with:

```bash
pandoc --print-default-data-file reference.docx > reference.docx
```

Then open it in Word, modify the styles, save it, and select it in the app with **Choose reference.docx…**.

The selected reference path is displayed in the GUI and restored when the app launches. If the saved reference file no longer exists, the app warns you at DOCX export time and ignores the missing reference file.

The reference DOCX applies only to `.docx` export, not `.tex` export.

---

## Examples

### Example 1: copied display formula

Input copied from ChatGPT:

```markdown
[
V_i = \beta^\top x_i
]
```

Normalized output:

```markdown
$$
V_i = \beta^\top x_i
$$
```

### Example 2: copied standalone formula line

Input:

```markdown
P_i = \frac{e^{V_i}}{\sum_{j \in C} e^{V_j}}
```

Normalized output:

```markdown
$$
P_i = \frac{e^{V_i}}{\sum_{j \in C} e^{V_j}}
$$
```

### Example 3: copied inline variables

Input:

```markdown
Here, (P_i) is the probability of choosing alternative (i).
```

Normalized output:

```markdown
Here, $P_i$ is the probability of choosing alternative $i$.
```

### Example 4: code is protected

Input:

```markdown
The model is computed with `models.logit(V, av, i)`.
```

Output remains:

```markdown
The model is computed with `models.logit(V, av, i)`.
```

### Example 5: copied subscripts and percent signs

Some Canvas copies may turn subscripts into star patterns and leave percent signs unescaped:

```markdown
[
B_m=\sum_i(\hat{q}*{i,m}-q*{i,m}), \qquad B^{%}*m=\frac{\sum_i(\hat{q}*{i,m}-q_{i,m})}{\sum_i q_{i,m}}.
]
```

Normalized output:

```markdown
$$
B_m=\sum_i(\hat{q}_{i,m}-q_{i,m}), \qquad B^{\%}_m=\frac{\sum_i(\hat{q}_{i,m}-q_{i,m})}{\sum_i q_{i,m}}.
$$
```

This is important because an unescaped `%` in TeX math can make the rest of the formula behave like a comment, which may leave visible `$$` delimiters or prevent Word math conversion.

---

## Limitations

This is a heuristic normalizer, not a full mathematical parser.

It can fix many common copy artifacts when the TeX content is still present, including missing display delimiters, simple parenthesized inline math, bullet-only formulas, copied subscript patterns such as `*{i,m}`, and unescaped percent signs in recovered math. It cannot recover formulas if the copy process has already deleted the formula content.

For maximum reliability, ask ChatGPT to produce raw Markdown in a fenced code block:

```text
Give the answer as raw Markdown in a fenced code block. Use $...$ for inline math and $$...$$ for display math.
```

For long documents, ChatGPT Canvas plus this tool is usually more comfortable: you can visually inspect formulas in Canvas and then normalize/export locally.

---

## Troubleshooting

### Pandoc works in Terminal, but the app does not find it

This often happens on macOS because GUI-launched Python apps do not always inherit the same shell `PATH` as Terminal.

Use **Choose Pandoc…** and select the executable manually, for example:

```text
/opt/homebrew/bin/pandoc
```

On Windows, select:

```text
C:\Program Files\Pandoc\pandoc.exe
```

if that is where Pandoc is installed.

### The app does not start on Linux because Tkinter is missing

Install the Tkinter package for your distribution. For Debian/Ubuntu:

```bash
sudo apt install python3-tk
```

### The app starts, but export fails

Check that the path shown in the **Pandoc:** field points to a valid Pandoc executable. You can also click **Choose Pandoc…** and select it manually.

### The wrong Pandoc or reference DOCX path keeps coming back

The app restores its saved preferences at launch. To reset them, close the app and delete the settings file:

| Platform | File to delete |
|---|---|
| macOS | `~/Library/Application Support/ChatGPT2Docx/settings.json` |
| Windows | `%APPDATA%\ChatGPT2Docx\settings.json` |
| Linux | `~/.config/ChatGPT2Docx/settings.json` |

### The DOCX contains visible `$$`

This usually means the Markdown was normalized twice by an older version, or the input contains malformed delimiters. Use v5.7 or later and try starting again from the original copied text.

### Some formulas still do not render

Check the normalized Markdown with **Normalize preview**. The formula should be surrounded by either:

```markdown
$$
...
$$
```

or:

```markdown
\[
...
\]
```

If the formula content itself has disappeared, the tool cannot reconstruct it.

### A formula contains `*{...}` or a percent sign and still fails

Use the latest version of the script and run **Normalize preview** again from the original copied text. The current heuristic repairs common cases such as:

```markdown
\hat{q}*{i,m} -> \hat{q}_{i,m}
B^{%}*m -> B^{\%}_m
```

If the pattern is more complex, manually edit the normalized Markdown so subscripts use `_` and literal percent signs inside formulas are escaped as `\%`.

### Code blocks are not shaded in DOCX

Make sure **Shade code blocks (DOCX/LaTeX)** is enabled. Also prefer fenced code blocks:

````markdown
```python
print("hello")
```
````

instead of indented code blocks or plain paragraphs.

### Code blocks are not shaded in LaTeX/PDF

The `.tex` export uses Pandoc highlighting style `tango` when shading is enabled. To see the background, compile the `.tex` file to PDF with a normal LaTeX engine such as `pdflatex`, `xelatex`, or `lualatex`.

---

## Development note: yes, ChatGPT helped

This script was developed with help from ChatGPT itself, which feels a bit like asking a photocopier to design a better photocopier. The goal was simple: take Markdown copied from ChatGPT Canvas, stop mathematical formulas from escaping into the wilderness, and produce a cleaner path to Word documents via Pandoc.

ChatGPT helped sketch the first version, debug the macOS Pandoc path issue, improve the math-normalization logic, add DOCX and LaTeX export options, support code-block shading, document the workflow, and refine oddball Canvas-copy cases such as `\hat{q}*{i,m}` and `B^{%}*m`. Human supervision, testing, and common sense were still required, especially when the early versions cheerfully produced artifacts such as `q$$`.

In other words: this is a small tool made by a human, with assistance from an AI, to solve a problem created by copying text from an AI. Perfectly normal modern software development.

---

## License

This project is released under the **MIT License**.

See the [`LICENSE`](LICENSE) file for details.
