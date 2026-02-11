# Technical Specification - Document Processing API

## Overview

This system converts Microsoft Word documents (.docx) to HTML with full equation support. It handles the complete document structure including headings, tables, footnotes, images, lists, and mathematical equations.

---

## 1. Conversion Modes

### 1.1 LaTeX + MathJax (Default)

**Pipeline:**
```
DOCX → Extract ZIP → Find Equations → Convert to LaTeX → Replace in XML
     → Repack DOCX → Load Modified DOCX → Convert to HTML → Add MathJax
```

**Output:** HTML file with LaTeX equations rendered by MathJax JavaScript library.

**Files Involved:**
- `enhanced_zip_converter.py` - Finds and replaces equations with LaTeX
- `doc_processor/omml_2_latex.py` - OMML to LaTeX conversion
- `word_to_html_full.py` - HTML generation with MathJax

### 1.2 MathML (No JavaScript)

**Pipeline:**
```
DOCX → Extract ZIP → Load Resources → Convert to HTML
     → Convert Equations to MathML inline → Output HTML
```

**Output:** HTML file with native MathML elements (no JavaScript required).

**Files Involved:**
- `word_to_html_full.py` - Direct conversion
- `doc_processor/omml_to_mathml.py` - OMML to MathML conversion

---

## 2. Word Document Structure

A .docx file is a ZIP archive containing:

```
document.docx (ZIP archive)
├── [Content_Types].xml
├── _rels/
│   └── .rels
├── word/
│   ├── document.xml      ← Main document content
│   ├── footnotes.xml     ← Footnote definitions
│   ├── styles.xml        ← Style definitions
│   ├── numbering.xml     ← List numbering definitions
│   ├── _rels/
│   │   └── document.xml.rels  ← Relationships (images, etc.)
│   └── media/
│       ├── image1.png
│       ├── image2.jpg
│       └── ...
└── docProps/
    ├── app.xml
    └── core.xml
```

---

## 3. Equation Locations

Equations appear in **5 different locations** within the XML:

| Location | XPath | Description |
|----------|-------|-------------|
| Main body | `//m:oMath` | Direct equations in paragraphs |
| mc:Choice | `//mc:Choice//m:oMath` | Modern shapes (Word 2010+) |
| mc:Fallback | `//mc:Fallback//m:oMath` | Legacy compatibility copies |
| VML textbox | `//v:textbox//m:oMath` | VML shape content |
| Textbox | `//w:txbxContent//m:oMath` | Standard textbox content |

### Display vs Inline Detection

```python
def is_display_equation(omath_element):
    # Display if wrapped in m:oMathPara
    parent = omath_element.getparent()
    if parent is not None and parent.tag.endswith('oMathPara'):
        return True
    return False
```

- **Display equation:** Inside `m:oMathPara` → rendered as block
- **Inline equation:** Direct `m:oMath` → rendered inline

---

## 4. OMML Element Mapping

### 4.1 To LaTeX

| OMML Element | LaTeX Output | Example |
|-------------|--------------|---------|
| `m:f` (fraction) | `\frac{num}{den}` | `\frac{a}{b}` |
| `m:rad` (radical) | `\sqrt{x}` or `\sqrt[n]{x}` | `\sqrt{x}` |
| `m:sSup` (superscript) | `base^{sup}` | `x^{2}` |
| `m:sSub` (subscript) | `base_{sub}` | `x_{n}` |
| `m:sSubSup` (both) | `base_{sub}^{sup}` | `x_{n}^{2}` |
| `m:nary` (integral/sum) | `\int`, `\sum`, `\prod` | `\int_{0}^{\infty}` |
| `m:d` (delimiters) | `\left( \right)` | `\left( x \right)` |
| `m:m` (matrix) | `\begin{matrix}...\end{matrix}` | Matrix layout |
| `m:acc` (accent) | `\hat{x}`, `\tilde{x}` | Accented symbols |
| `m:func` (function) | `\sin`, `\cos`, `\lim` | Function names |
| `m:limLow` (limit) | `\lim_{x \to 0}` | Limits |
| `m:eqArr` (array) | `\begin{cases}...\end{cases}` | Piecewise |

### 4.2 To MathML

| OMML Element | MathML Output |
|-------------|---------------|
| `m:f` | `<mfrac><mrow>...</mrow><mrow>...</mrow></mfrac>` |
| `m:rad` | `<msqrt>...</msqrt>` or `<mroot>...</mroot>` |
| `m:sSup` | `<msup><mi>x</mi><mn>2</mn></msup>` |
| `m:sSub` | `<msub><mi>x</mi><mi>n</mi></msub>` |
| `m:nary` | `<munderover><mo>∫</mo>...</munderover>` |
| `m:d` | `<mrow><mo>(</mo>...<mo>)</mo></mrow>` |
| `m:m` | `<mtable><mtr><mtd>...</mtd></mtr></mtable>` |
| `m:r` (text) | `<mi>`, `<mn>`, or `<mo>` based on content type |

---

## 5. HTML Output Format

### 5.1 Structure

```html
<!DOCTYPE html>
<html dir="rtl">
<head>
  <meta charset="UTF-8">
  <title>Document Title</title>
  <!-- MathJax script (LaTeX mode only) -->
</head>
<body>
  <!-- Document content -->
  <h1>Heading 1</h1>
  <p>Paragraph with <strong>bold</strong> and \(x^2\) inline equation.</p>

  <!-- Display equation -->
  <p>\[\frac{a}{b}\]</p>

  <!-- Table -->
  <table>
    <tbody>
      <tr>
        <td width="312">Cell content</td>
        <td colspan="2" width="624">Merged cell</td>
      </tr>
    </tbody>
  </table>

  <!-- Footnote reference -->
  <a href="#_ftn1" name="_ftnref1">[1]</a>

  <!-- Footnotes section -->
  <div>
    <p><a href="#_ftnref1" name="_ftn1">[1]</a> Footnote text.</p>
  </div>
</body>
</html>
```

### 5.2 Footnote Naming Convention

Following wordhtml.com format:

- Reference in body: `<a href="#_ftn1" name="_ftnref1">[1]</a>`
- Definition at bottom: `<a href="#_ftnref1" name="_ftn1">[1]</a>`

### 5.3 Table Attributes

- `<tbody>` wrapper always included
- `width` attribute on `<td>` (pixels)
- `colspan` for merged cells

### 5.4 List Continuation

Lists interrupted by other content continue with proper numbering:

```html
<ol>
  <li>Item 1</li>
  <li>Item 2</li>
</ol>
<p>Some text in between</p>
<ol start="3">
  <li>Item 3</li>
  <li>Item 4</li>
</ol>
```

---

## 6. Configuration Options

### ConversionConfig

```python
@dataclass
class ConversionConfig:
    output_format: str = "latex_html"     # "latex_html" or "mathml_html"
    inline_prefix: str = ""               # Equation marker prefix
    inline_suffix: str = ""               # Equation marker suffix
    display_prefix: str = ""              # Display equation prefix
    display_suffix: str = ""              # Display equation suffix
    convert_shapes_to_svg: bool = False   # Convert Word shapes to SVG
    include_images: bool = True           # Include images in output
    include_mathjax: bool = True          # Include MathJax library
    rtl_direction: bool = True            # RTL text direction
```

---

## 7. API Reference

### POST /api/process

Upload and process documents.

**Request:**
```
Content-Type: multipart/form-data

files: File(s) to process
processor_type: "word_to_html"
conversion_config: JSON string with options
```

**Response:**
```json
{
  "job_id": "abc123",
  "status": "processing"
}
```

### GET /api/status/{job_id}

Check processing status.

**Response:**
```json
{
  "status": "completed",
  "results": [
    {
      "filename": "document.html",
      "size": 12345,
      "type": "text/html"
    }
  ]
}
```

### GET /api/download/{job_id}/{index}

Download specific result file.

### GET /api/download/{job_id}

Download all results as ZIP.

---

## 8. Error Handling

### Common Issues

| Issue | Cause | Solution |
|-------|-------|----------|
| Equations not rendering | MathJax not loaded | Check script inclusion |
| Missing images | Image extraction failed | Check media/ folder |
| Broken footnote links | ID mismatch | Verify _ftn/_ftnref naming |
| List numbers wrong | numId tracking issue | Check list_counters logic |
| RTL text reversed | Missing dir attribute | Ensure dir="rtl" on html |

### Debug Logging

Enable debug output by checking console logs. Key log points:

- `[App]` - Frontend state changes
- `[JobStatus]` - Job polling
- `[ResultDownload]` - Download operations

---

## 9. Dependencies

### Backend (Python)

```
fastapi>=0.68.0
uvicorn>=0.15.0
python-multipart>=0.0.5
lxml>=4.6.3
aiofiles>=0.7.0
```

### Frontend (Node.js)

```
vue@3
vite
tailwindcss
axios
```

### System

- Python 3.9+
- Node.js 16+
- Docker (for deployment)

---

## 10. Deployment Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                    Google Cloud Run                          │
│                                                              │
│  ┌────────────────────────────────────────────────────────┐ │
│  │                 Docker Container                         │ │
│  │                                                          │ │
│  │  ┌─────────────┐    ┌─────────────┐                     │ │
│  │  │   Nginx     │───>│   FastAPI   │                     │ │
│  │  │  (static)   │    │  (backend)  │                     │ │
│  │  └─────────────┘    └─────────────┘                     │ │
│  │        │                   │                             │ │
│  │        v                   v                             │ │
│  │   Vue.js App          Python                            │ │
│  │   (frontend)          processing                        │ │
│  │                                                          │ │
│  └────────────────────────────────────────────────────────┘ │
│                                                              │
└─────────────────────────────────────────────────────────────┘
```

### Resources

- Memory: 1Gi
- CPU: 1
- Timeout: 300s
- Max instances: 10

---

## 11. Symbol Reference

### Greek Letters

| Symbol | LaTeX | MathML |
|--------|-------|--------|
| α | `\alpha` | `<mi>α</mi>` |
| β | `\beta` | `<mi>β</mi>` |
| γ | `\gamma` | `<mi>γ</mi>` |
| Δ | `\Delta` | `<mi>Δ</mi>` |
| ∞ | `\infty` | `<mo>∞</mo>` |

### Operators

| Symbol | LaTeX | MathML |
|--------|-------|--------|
| ∫ | `\int` | `<mo>∫</mo>` |
| ∑ | `\sum` | `<mo>∑</mo>` |
| ∏ | `\prod` | `<mo>∏</mo>` |
| ± | `\pm` | `<mo>±</mo>` |
| ≤ | `\leq` | `<mo>≤</mo>` |
| ≥ | `\geq` | `<mo>≥</mo>` |
| ≠ | `\neq` | `<mo>≠</mo>` |

---

## 12. Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0 | Initial | LaTeX + MathJax conversion |
| 1.1 | Update | Added MathML mode |
| 1.2 | Update | Fixed list continuation, table formatting |
| 1.3 | Current | Simplified UI, footnote improvements |
