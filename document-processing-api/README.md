# Document Processing API

Converts Word (.docx) files to HTML with support for equations, tables, images, footnotes, and shapes.

## Output Formats

- **LaTeX + MathJax** (`latex_html`): Equations rendered as LaTeX with MathJax JavaScript
- **MathML** (`mathml_html`): Equations rendered as native MathML, no JavaScript required

## Output Files

Each conversion produces:
- `{filename}.html` — Full standalone HTML file with `<html>`, `<head>`, styles, and scripts
- `{filename}_body.txt` — Body-only HTML containing just the `<div id="mathjax-content">` wrapper with content and footnotes. Uses `.txt` extension for easy copy-pasting into SharePoint or other CMS pages.
- `{filename}_equations.docx` — (LaTeX mode only) Intermediate Word file with converted equations

## SharePoint MathJax Loader Scripts

When using the LaTeX + MathJax output format in SharePoint, the `_body.txt` file provides the content without any `<script>` tags. MathJax rendering is handled separately by uploading one of the loader scripts to SharePoint:

- `sharepoint-mathjax-loader.js` — Original loader script
- `sharepoint-mathjax-loader-2.js` — Compact variant with improved SharePoint edit-mode detection (checks `MSOLayout_InDesignMode`, `DisplayModeName`, URL params, and `contenteditable` ancestors)

These scripts detect SharePoint edit mode and skip MathJax initialization when the page is being edited, preventing interference with the SharePoint editor. They scope MathJax processing to the `#mathjax-content` div and ignore SharePoint UI classes (`sp-*`, `ms-*`, `od-*`).

### MathJax Configuration Notes

- `startup.elements` must be under `startup` (not `options`) in MathJax 3 — placing it under `options` causes an "Invalid option" error.
- `options.enableMenu: false` disables MathJax's built-in right-click context menu.

## Equation Copy Menu

`mathjax-copy-menu.js` — Adds hover-to-copy buttons to MathJax-rendered equations. Hover over any equation to see a popup with two copy modes:

- **LaTeX** (`نسخ LaTeX`) — Copies the original TeX source with `\(...\)` or `\[...\]` delimiters
- **MathML** (`نسخ MathML`) — Copies the MathML markup via MathJax's internal serializer

The script is automatically embedded inline in generated HTML files (LaTeX mode). It can also be included as a standalone file after MathJax loads:
```html
<script src="mathjax-copy-menu.js"></script>
```

### How it works

1. Waits for MathJax to finish typesetting (retries up to 30 seconds)
2. Attaches hover handlers to all `mjx-container` elements
3. On hover, shows a floating button bar above the equation
4. On click, copies the equation in the selected format to the clipboard
5. Shows Arabic feedback: "تم النسخ!" (Copied!) on success

The script skips initialization in SharePoint edit mode and uses `Array.from()` on MathJax's iterable math list for compatibility with different MathJax 3.x builds.
