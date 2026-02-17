# Document Processing API

Converts Word (.docx) files to HTML with support for equations, tables, images, footnotes, and shapes.

## Output Formats

- **LaTeX + MathJax** (`latex_html`): Equations converted to LaTeX, rendered as native MathML via MathJax 4
- **MathML** (`mathml_html`): Equations converted directly to MathML at the backend, no JavaScript required

## Output Files

Each conversion produces:
- `{filename}.html` — Full standalone HTML file with styles and MathJax script (no images — team inserts manually)
- `{filename}_body.txt` — Body-only HTML containing just the `<div id="mathjax-content">` wrapper with content and footnotes. No images, no scripts. Uses `.txt` extension for easy pasting into SharePoint.
- `{filename}_equations.docx` — (LaTeX mode only) Intermediate Word file with converted equations

## SharePoint Scripts

For SharePoint, the `_body.txt` provides content and these scripts handle rendering:

### `sharepoint-mathjax-loader.js` (.txt copy for email)
MathJax 4 loader with native MathML output. Upload to SharePoint as a Script Editor or Content Editor web part.

- Detects SharePoint edit mode (URL params, `contenteditable`, `MSOLayout_InDesignMode`, `DisplayModeName`) and skips MathJax when editing
- Scopes processing to `#mathjax-content` div
- Ignores SharePoint UI classes: `sp-*`, `od-*`, `canvasTextArea`
- **Note:** `ms-*` classes are NOT ignored because SharePoint wraps article content in `<div class="ms-rtestate-field">` which must be processed
- Uses MathJax 4 `startup.js` with custom `renderActions` to output native `<math>` elements (text-selectable, copy-pasteable)

### `mathjax-copy-menu.js` (.txt copy for email)
Hover-to-copy buttons for equations. Hover over any equation to see:

- **نسخ LaTeX** — Copies the original TeX source with `\(...\)` or `\[...\]` delimiters
- **نسخ MathML** — Copies the MathML markup directly from the DOM `<math>` element

The script is automatically embedded inline in generated HTML files. Also works as a standalone script after MathJax loads.

### MathJax 4 Native MathML

The scripts use MathJax 4 (`startup.js`) configured to render LaTeX as native browser MathML instead of the old CHTML format. This means:
- Equations are rendered as native `<math>` elements
- Users can select and copy equation text naturally (Ctrl+A, Ctrl+C)
- When pasted into Word, equations are converted to Word's native equation format
- No custom fonts or CSS required — the browser handles rendering
