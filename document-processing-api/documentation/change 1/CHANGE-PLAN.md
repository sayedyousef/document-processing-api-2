# Change Plan 1: MathML Native HTML Conversion (No JavaScript)

## Overview

Add a new conversion mode that generates **pure HTML with MathML** for equations instead of LaTeX + MathJax JavaScript. This produces self-contained HTML that renders math natively in all modern browsers without any JavaScript dependency.

The HTML output will follow **wordhtml.com** naming conventions for tags, footnotes, headings, and elements.

---

## Table of Contents

1. [Requirements Summary](#1-requirements-summary)
2. [Current Architecture Analysis](#2-current-architecture-analysis)
3. [Proposed Architecture](#3-proposed-architecture)
4. [Detailed Changes](#4-detailed-changes)
5. [File-by-File Change Plan](#5-file-by-file-change-plan)
6. [MathML Conversion Details](#6-mathml-conversion-details)
7. [HTML Output Format (wordhtml.com Style)](#7-html-output-format-wordhtmlcom-style)
8. [Image Handling](#8-image-handling)
9. [Table Handling](#9-table-handling)
10. [Section & Content Preservation](#10-section--content-preservation)
11. [Testing Strategy](#11-testing-strategy)
12. [Migration & Backward Compatibility](#12-migration--backward-compatibility)

---

## 1. Requirements Summary

| # | Requirement | Details |
|---|-------------|---------|
| R1 | **MathML equations** | Convert OMML to MathML `<math>` tags instead of LaTeX text. No JavaScript needed for rendering. |
| R2 | **Direct DOCX-to-HTML** | Skip intermediate Word document generation. Parse DOCX XML directly to HTML. |
| R3 | **wordhtml.com HTML format** | Match footnote naming (`_ftn1`/`_ftnref1`), heading tags, table structure, and element conventions from the reference file. |
| R4 | **Image handling** | Extract images to a separate `images/` folder. Reference them with relative paths. |
| R5 | **Table preservation** | Preserve table structure including `width`, `colspan`, nested tables. |
| R6 | **Section preservation** | Handle Word sections properly. Avoid content trimming during conversion. Process `w:sectPr` elements. |
| R7 | **Old equation format support** | Keep existing equation detection logic for OMML in all locations (body, shapes, textboxes, mc:Choice/Fallback). |
| R8 | **Preserve current framework** | Add as a new option only. All existing functionality remains. Use dependency injection / strategy pattern. Make the new mode the default. |

---

## 2. Current Architecture Analysis

### Current Conversion Flow

```
DOCX Upload
    |
    v
[1] EnhancedZipConverter.process_document()
    - Extracts DOCX as ZIP
    - Finds ALL m:oMath elements (body, mc:Choice, mc:Fallback, VML, textboxes)
    - Converts OMML -> LaTeX via DirectOmmlToLatex.parse()
    - Replaces equations in XML with LaTeX text runs (w:r elements)
    - Saves modified DOCX (intermediate Word file)
    |
    v
[2] FullWordToHTMLConverter.convert()
    - Extracts the modified DOCX
    - Loads relationships, styles, numbering, footnotes, images
    - Converts document body to HTML
    - Equations are already LaTeX text in the XML
    - Generates HTML with MathJax <script> tags
    - Outputs: HTML file + images/ folder + modified .docx
```

### Key Files & Their Roles

| File | Role | Lines |
|------|------|-------|
| `main.py` | API entry point, routes, job processing | 629 |
| `word_to_html_full.py` | `ConversionConfig` + `FullWordToHTMLConverter` (XML->HTML) | 1036 |
| `enhanced_zip_converter.py` | `EnhancedZipConverter` (DOCX ZIP equation replacement) | 403 |
| `doc_processor/omml_2_latex.py` | `DirectOmmlToLatex` (OMML->LaTeX parser) | 817 |
| `frontend/src/App.vue` | UI with conversion options | 228 |

### Current Equation Detection (MUST PRESERVE)

The `EnhancedZipConverter` finds equations in 5 locations:
1. **Main body** - `//m:oMath` directly in `w:body`
2. **mc:Choice** - Modern shapes (Word 2010+) in `wps:txbx`
3. **mc:Fallback** - Legacy compatibility copies in `v:textbox`
4. **VML** - Older VML shapes
5. **Textbox** - `w:txbxContent` containers

This detection logic is correct and complete. We keep it identical.

### Current Equation Type Detection (MUST PRESERVE)

Display vs. inline is determined by checking if `m:oMath` has an `m:oMathPara` ancestor:
- **Display equation**: `m:oMath` inside `m:oMathPara` (block-level)
- **Inline equation**: `m:oMath` directly in `w:p` or `w:r` (inline)

This logic is correct. We keep it identical.

---

## 3. Proposed Architecture

### Strategy Pattern for Equation Conversion

```
                          IEquationConverter (interface)
                         /                              \
            DirectOmmlToLatex                   OmmlToMathMLConverter
            (existing, unchanged)                     (NEW)
                    |                                     |
        Produces LaTeX string                  Produces MathML <math> HTML
        e.g. "\frac{a}{b}"                    e.g. "<math><mfrac>..."
```

### New Conversion Flow (MathML mode)

```
DOCX Upload
    |
    v
[1] FullWordToHTMLConverter.convert()  (ENHANCED)
    - Extracts DOCX as ZIP
    - Loads relationships, styles, numbering, footnotes, images
    - Parses document body XML
    - When encountering m:oMath elements IN-PLACE:
        -> Calls OmmlToMathMLConverter.convert(omml_element)
        -> Gets back MathML HTML string
        -> Embeds directly in HTML output
    - NO intermediate modified .docx needed
    - NO MathJax script needed
    - Outputs: HTML file + images/ folder
```

### Key Architectural Decisions

1. **No intermediate DOCX**: In MathML mode, we parse the ORIGINAL DOCX directly. No equation pre-processing step. Equations are converted to MathML on-the-fly during HTML generation.

2. **Strategy injection**: `FullWordToHTMLConverter` receives an `equation_converter` strategy. If `output_format == "mathml_html"`, it uses `OmmlToMathMLConverter`. If `output_format == "latex_html"`, it uses the existing two-step flow (EnhancedZipConverter + LaTeX text).

3. **Backward compatible**: The existing LaTeX/MathJax flow remains fully functional. The new MathML mode is an additional option (and the new default).

4. **wordhtml.com conventions**: The HTML generator is enhanced to produce footnote anchors, table attributes, and heading structure matching the reference format.

---

## 4. Detailed Changes

### 4.1 New File: `backend/doc_processor/omml_to_mathml.py`

**Purpose**: Convert OMML XML elements to MathML HTML strings.

**Class**: `OmmlToMathMLConverter`

**Approach**: OMML and MathML share the same underlying mathematical structure. The conversion maps OMML elements to their MathML equivalents:

| OMML Element | MathML Equivalent | Description |
|-------------|-------------------|-------------|
| `m:oMath` | `<math>` | Math container |
| `m:oMathPara` | `<math display="block">` | Display (block) math |
| `m:r` / `m:t` | `<mi>`, `<mn>`, `<mo>` | Identifiers, numbers, operators |
| `m:f` (fraction) | `<mfrac>` | Fraction |
| `m:num` | First child of `<mfrac>` | Numerator |
| `m:den` | Second child of `<mfrac>` | Denominator |
| `m:rad` (radical) | `<msqrt>` or `<mroot>` | Square root / nth root |
| `m:sSup` | `<msup>` | Superscript |
| `m:sSub` | `<msub>` | Subscript |
| `m:sSubSup` | `<msubsup>` | Sub+superscript |
| `m:nary` | `<munderover>` + `<mo>` | Integral, sum, product |
| `m:d` (delimiters) | `<mrow><mo>(</mo>...<mo>)</mo></mrow>` | Parentheses, brackets |
| `m:m` (matrix) | `<mtable><mtr><mtd>` | Matrix |
| `m:acc` (accent) | `<mover>` | Hat, tilde, bar, etc. |
| `m:func` | `<mrow>` with function name | sin, cos, lim, etc. |
| `m:limLow` | `<munder>` | Limit with subscript |
| `m:eqArr` | `<mtable>` | Equation array / piecewise |

**Key Methods**:

```python
class OmmlToMathMLConverter:
    def convert(self, omml_element, is_display=False) -> str:
        """Convert m:oMath element to MathML HTML string.

        Args:
            omml_element: lxml element (m:oMath or m:oMathPara)
            is_display: True for block equations, False for inline

        Returns:
            MathML HTML string like '<math xmlns="...">...</math>'
        """

    def _parse_element(self, elem) -> str:
        """Recursively parse OMML element to MathML."""
        # Dispatch to specific handlers based on element tag

    def _parse_fraction(self, elem) -> str:
        """m:f -> <mfrac>"""

    def _parse_radical(self, elem) -> str:
        """m:rad -> <msqrt> or <mroot>"""

    def _parse_superscript(self, elem) -> str:
        """m:sSup -> <msup>"""

    def _parse_subscript(self, elem) -> str:
        """m:sSub -> <msub>"""

    def _parse_subsup(self, elem) -> str:
        """m:sSubSup -> <msubsup>"""

    def _parse_nary(self, elem) -> str:
        """m:nary -> integral/sum with <munderover>"""

    def _parse_delimiter(self, elem) -> str:
        """m:d -> <mrow> with <mo> delimiters, or <mtable> for matrices"""

    def _parse_matrix(self, elem) -> str:
        """m:m -> <mtable><mtr><mtd>"""

    def _parse_run(self, elem) -> str:
        """m:r -> <mi>/<mn>/<mo> based on content type"""

    def _parse_accent(self, elem) -> str:
        """m:acc -> <mover>"""

    def _parse_function(self, elem) -> str:
        """m:func -> function name + argument"""

    def _parse_limit_lower(self, elem) -> str:
        """m:limLow -> <munder>"""

    def _parse_equation_array(self, elem) -> str:
        """m:eqArr -> <mtable> for piecewise/aligned"""

    def _classify_text(self, text) -> str:
        """Classify text as identifier (<mi>), number (<mn>), or operator (<mo>)"""

    def _convert_symbol(self, char) -> str:
        """Convert Unicode math symbol to MathML entity/character"""
```

**Symbol Handling**: Reuse the existing `MATH_SYMBOLS` mapping from `omml_2_latex.py` but convert to MathML `<mo>` operators instead of LaTeX commands. Unicode math symbols map directly to MathML since MathML uses Unicode natively.

**Double-struck / Blackboard Bold**: Map via `<mi mathvariant="double-struck">R</mi>` instead of `\mathbb{R}`.

### 4.2 Modified File: `backend/word_to_html_full.py`

#### Changes to `ConversionConfig`

Add new field:

```python
@dataclass
class ConversionConfig:
    # ... existing fields ...

    # NEW: Output format for equations
    # "mathml_html" = MathML (no JS, native browser rendering) - NEW DEFAULT
    # "latex_html" = LaTeX + MathJax (existing behavior)
    output_format: str = "mathml_html"
```

#### Changes to `FullWordToHTMLConverter`

**Initialization**: Based on `output_format`, select the equation conversion strategy:

```python
def __init__(self, config: ConversionConfig = None):
    self.config = config or ConversionConfig()
    self.svg_converter = ShapeToSVGConverter()

    # Strategy: select equation converter based on output_format
    if self.config.output_format == "mathml_html":
        from doc_processor.omml_to_mathml import OmmlToMathMLConverter
        self.equation_converter = OmmlToMathMLConverter()
    else:
        self.equation_converter = None  # LaTeX mode uses pre-processed DOCX

    # ... rest unchanged ...
```

**Modified `convert()` method**: Two paths based on `output_format`:

```python
def convert(self, input_path, output_path=None, output_dir=None):
    # ... setup same as before ...

    if self.config.output_format == "mathml_html":
        # NEW PATH: Direct conversion, no intermediate DOCX
        return self._convert_mathml_mode(input_path, output_path, output_dir)
    else:
        # EXISTING PATH: Two-step (equation pre-processing + HTML generation)
        return self._convert_latex_mode(input_path, output_path, output_dir)
```

The existing `convert()` logic moves to `_convert_latex_mode()` (unchanged). The new `_convert_mathml_mode()`:

```python
def _convert_mathml_mode(self, input_path, output_path, output_dir):
    """Direct DOCX to HTML with MathML - no intermediate Word file"""

    temp_dir = Path(f"temp_full_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    try:
        # Step 1: Extract ORIGINAL DOCX (no pre-processing)
        extract_dir = temp_dir / "extracted"
        with zipfile.ZipFile(input_path, 'r') as z:
            z.extractall(extract_dir)

        # Step 2: Load resources (same as before)
        self.output_dir = output_dir
        self._load_relationships(extract_dir)
        self._load_styles(extract_dir)
        self._load_numbering(extract_dir)
        self._load_footnotes_wordhtml(extract_dir)  # NEW: wordhtml.com format
        self._extract_images(extract_dir, output_dir)

        # Step 3: Convert document (OMML equations converted inline to MathML)
        doc_xml = extract_dir / "word" / "document.xml"
        with open(doc_xml, 'rb') as f:
            doc_root = etree.fromstring(f.read())

        html_content = self._convert_body(doc_root)  # Enhanced to handle m:oMath

        # Step 4: Generate HTML (no MathJax, wordhtml.com format)
        full_html = self._generate_html_wordhtml(html_content, input_path.stem)

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(full_html)

        return {'success': True, 'output_path': str(output_path)}
    finally:
        if temp_dir.exists():
            shutil.rmtree(temp_dir)
```

**Enhanced `_convert_paragraph_content()`**: Handle `m:oMath` and `m:oMathPara` elements inline:

```python
def _convert_paragraph_content(self, p_elem):
    ns = self.namespaces
    parts = []

    for child in p_elem:
        tag = child.tag.split('}')[-1]

        if tag == 'r':
            parts.append(self._convert_run(child))
        elif tag == 'hyperlink':
            parts.append(self._convert_hyperlink(child))
        elif tag == 'drawing':
            parts.append(self._convert_drawing(child))
        elif tag == 'oMath' and self.equation_converter:
            # NEW: Convert equation inline to MathML
            parts.append(self.equation_converter.convert(child, is_display=False))
        elif tag == 'oMathPara' and self.equation_converter:
            # NEW: Convert display equation inline to MathML
            omath = child.find('m:oMath', namespaces=ns)
            if omath is not None:
                parts.append(self.equation_converter.convert(omath, is_display=True))
        elif tag in ['pPr', 'bookmarkStart', 'bookmarkEnd']:
            continue
        else:
            text = self._extract_text(child)
            if text:
                parts.append(self._escape(text))

    return ''.join(parts)
```

**New `_load_footnotes_wordhtml()`**: Load footnotes with wordhtml.com anchor naming:

```python
def _load_footnotes_wordhtml(self, extract_dir):
    """Load footnotes using wordhtml.com naming convention (_ftn/_ftnref)"""
    fn_path = extract_dir / "word" / "footnotes.xml"
    if fn_path.exists():
        with open(fn_path, 'rb') as f:
            root = etree.fromstring(f.read())
        ns = {'w': self.namespaces['w']}
        for fn in root.xpath('//w:footnote', namespaces=ns):
            fn_id = fn.get(f'{{{ns["w"]}}}id')
            if fn_id and fn_id not in ['0', '-1']:
                # Parse full footnote content (may contain formatting)
                self.footnotes[fn_id] = self._convert_footnote_content(fn)
```

**New `_generate_html_wordhtml()`**: Generate clean HTML matching wordhtml.com format (no MathJax, no custom CSS classes):

```python
def _generate_html_wordhtml(self, content, title):
    """Generate HTML in wordhtml.com format - clean, no JavaScript"""
    # No MathJax script at all
    # Minimal CSS or inline styles only
    # RTL support via dir attribute
    # Footnotes using _ftn/_ftnref naming
```

**Modified `_convert_run()`**: In MathML mode, handle `m:oMath` elements that appear inside runs:

```python
def _convert_run(self, r_elem):
    ns = self.namespaces
    parts = []

    for child in r_elem:
        tag = child.tag.split('}')[-1]

        if tag == 't':
            parts.append(self._escape(child.text or ''))
        elif tag == 'drawing':
            parts.append(self._convert_drawing(child))
        elif tag == 'pict':
            parts.append(self._convert_pict(child))
        elif tag == 'AlternateContent':
            # Handle mc:AlternateContent (shapes, equations in textboxes)
            drawing = child.xpath('.//w:drawing', namespaces=ns)
            pict = child.xpath('.//w:pict', namespaces=ns)
            if drawing:
                parts.append(self._convert_drawing(drawing[0]))
            elif pict:
                parts.append(self._convert_pict(pict[0]))
        elif tag == 'footnoteReference':
            fn_id = child.get(f'{{{ns["w"]}}}id')
            # wordhtml.com format: <a href="#_ftn1" name="_ftnref1">[1]</a>
            if self.config.output_format == "mathml_html":
                parts.append(f'<a href="#_ftn{fn_id}" name="_ftnref{fn_id}">[{fn_id}]</a>')
            else:
                parts.append(f'<sup><a id="fnref{fn_id}" href="#fn{fn_id}">[{fn_id}]</a></sup>')
        elif tag == 'br':
            parts.append('<br>')

    # ... bold/italic wrapping same as before ...
```

### 4.3 Modified File: `backend/main.py`

**Changes in `process_job()`**:

```python
async def process_job(job_id, file_paths, processor_type, output_dir, config_dict=None):
    # ...

    if processor_type == "word_to_html":
        from word_to_html_full import FullWordToHTMLConverter, ConversionConfig

        conv_config = ConversionConfig(
            # ... existing fields ...
            output_format=config_dict.get('output_format', 'mathml_html') if config_dict else 'mathml_html',
        )

        converter = FullWordToHTMLConverter(conv_config)
        result = converter.convert(file_path, output_dir=output_dir)

        # NOTE: In mathml_html mode, no intermediate .docx is generated
        # Only HTML + images/ folder
```

The `output_format` config value controls the strategy. No `if/else` spaghetti in `main.py` -- the decision is encapsulated inside `FullWordToHTMLConverter`.

### 4.4 Modified File: `frontend/src/App.vue`

**Add output format selector** (above existing settings):

```html
<!-- Output Format -->
<div class="mb-4">
  <label class="block text-sm font-medium text-gray-700 mb-2">Output Format</label>
  <div class="flex space-x-2">
    <button
      @click="conversionConfig.output_format = 'mathml_html'"
      :class="[..., conversionConfig.output_format === 'mathml_html' ? 'active' : '']"
    >MathML (No JS) (Recommended)</button>
    <button
      @click="conversionConfig.output_format = 'latex_html'"
      :class="[..., conversionConfig.output_format === 'latex_html' ? 'active' : '']"
    >LaTeX + MathJax</button>
  </div>
</div>
```

**Conditionally show/hide settings**: When `output_format == 'mathml_html'`:
- Hide "Equation Marker Style" (not applicable)
- Hide "Include MathJax library" (not applicable)
- Keep "Convert shapes to SVG", "Include images in HTML", "RTL direction"

**Add to `conversionConfig` reactive object**:

```javascript
const conversionConfig = reactive({
  // ... existing fields ...
  output_format: 'mathml_html'  // NEW: default to MathML
})
```

---

## 5. File-by-File Change Plan

### New Files

| File | Description |
|------|-------------|
| `backend/doc_processor/omml_to_mathml.py` | `OmmlToMathMLConverter` class - OMML to MathML conversion |

### Modified Files

| File | Changes |
|------|---------|
| `backend/word_to_html_full.py` | Add `output_format` to `ConversionConfig`. Add `_convert_mathml_mode()`, `_load_footnotes_wordhtml()`, `_generate_html_wordhtml()`. Enhance `_convert_paragraph_content()` and `_convert_run()` to handle inline OMML. Enhance `_convert_body()` to handle `m:oMathPara` at block level. Improve `_convert_table()` with width attributes. Add section break handling. |
| `backend/main.py` | Pass `output_format` from config_dict to `ConversionConfig`. Adjust ZIP output logic for MathML mode (no .docx output). |
| `frontend/src/App.vue` | Add output format toggle (MathML/LaTeX). Conditionally show/hide equation marker settings. Default to MathML. |

### Unchanged Files (explicitly preserved)

| File | Reason |
|------|--------|
| `doc_processor/omml_2_latex.py` | Untouched. Used only in LaTeX mode. |
| `enhanced_zip_converter.py` | Untouched. Used only in LaTeX mode. |
| `doc_processor/zip_equation_replacer.py` | Untouched. Used for `latex_equations` processor type. |
| `doc_processor/main_word_com_equation_replacer.py` | Untouched. Windows-only COM approach. |
| All frontend components | `FileUploader.vue`, `JobStatus.vue`, `ResultDownload.vue` - no changes. |
| `core/config.py`, `core/logger.py` | No changes needed. |

---

## 6. MathML Conversion Details

### 6.1 OMML to MathML Mapping (Complete)

#### Simple Elements

```
OMML: <m:r><m:t>x</m:t></m:r>
MathML: <mi>x</mi>          (identifier/variable)

OMML: <m:r><m:t>123</m:t></m:r>
MathML: <mn>123</mn>        (number)

OMML: <m:r><m:t>+</m:t></m:r>
MathML: <mo>+</mo>          (operator)
```

#### Fraction

```
OMML: <m:f>
        <m:num><m:r><m:t>a</m:t></m:r></m:num>
        <m:den><m:r><m:t>b</m:t></m:r></m:den>
      </m:f>

MathML: <mfrac>
          <mrow><mi>a</mi></mrow>
          <mrow><mi>b</mi></mrow>
        </mfrac>
```

#### Superscript / Subscript

```
OMML: <m:sSup>
        <m:e><m:r><m:t>x</m:t></m:r></m:e>
        <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
      </m:sSup>

MathML: <msup>
          <mi>x</mi>
          <mn>2</mn>
        </msup>
```

#### Integral with Limits

```
OMML: <m:nary>
        <m:naryPr><m:chr m:val="∫"/></m:naryPr>
        <m:sub><m:r><m:t>0</m:t></m:r></m:sub>
        <m:sup><m:r><m:t>∞</m:t></m:r></m:sup>
        <m:e>...</m:e>
      </m:nary>

MathML: <munderover>
          <mo>&#x222B;</mo>
          <mn>0</mn>
          <mi>&#x221E;</mi>
        </munderover>
        <mrow>...</mrow>
```

#### Matrix

```
OMML: <m:d>
        <m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr>
        <m:e>
          <m:m>
            <m:mr>
              <m:e><m:r><m:t>a</m:t></m:r></m:e>
              <m:e><m:r><m:t>b</m:t></m:r></m:e>
            </m:mr>
            <m:mr>
              <m:e><m:r><m:t>c</m:t></m:r></m:e>
              <m:e><m:r><m:t>d</m:t></m:r></m:e>
            </m:mr>
          </m:m>
        </m:e>
      </m:d>

MathML: <mrow>
          <mo>(</mo>
          <mtable>
            <mtr>
              <mtd><mi>a</mi></mtd>
              <mtd><mi>b</mi></mtd>
            </mtr>
            <mtr>
              <mtd><mi>c</mi></mtd>
              <mtd><mi>d</mi></mtd>
            </mtr>
          </mtable>
          <mo>)</mo>
        </mrow>
```

#### Square Root

```
OMML: <m:rad>
        <m:radPr><m:degHide m:val="1"/></m:radPr>
        <m:e><m:r><m:t>x</m:t></m:r></m:e>
      </m:rad>

MathML: <msqrt><mi>x</mi></msqrt>

OMML: <m:rad>
        <m:deg><m:r><m:t>3</m:t></m:r></m:deg>
        <m:e><m:r><m:t>x</m:t></m:r></m:e>
      </m:rad>

MathML: <mroot><mi>x</mi><mn>3</mn></mroot>
```

#### Accents (hat, tilde, bar, vec)

```
OMML: <m:acc>
        <m:accPr><m:chr m:val="̂"/></m:accPr>
        <m:e><m:r><m:t>x</m:t></m:r></m:e>
      </m:acc>

MathML: <mover>
          <mi>x</mi>
          <mo>&#x302;</mo>
        </mover>
```

#### Delimiters (parentheses, brackets)

```
OMML: <m:d>
        <m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr>
        <m:e>...</m:e>
      </m:d>

MathML: <mrow>
          <mo>[</mo>
          ...
          <mo>]</mo>
        </mrow>
```

### 6.2 Symbol Handling

Reuse the same Unicode symbols from `MATH_SYMBOLS` in `omml_2_latex.py`. In MathML, Unicode symbols are used directly inside `<mo>` tags:

```python
# LaTeX mode:  '≠' -> r'\neq '
# MathML mode: '≠' -> '<mo>≠</mo>'    (Unicode is native in MathML)
```

For special symbols:
- Greek letters: `<mi>α</mi>`, `<mi>β</mi>`, etc. (Unicode directly)
- Blackboard bold: `<mi mathvariant="double-struck">R</mi>`
- Function names: `<mi mathvariant="normal">sin</mi>` or `<mo>sin</mo>`

### 6.3 Display vs. Inline

```html
<!-- Inline equation -->
<math xmlns="http://www.w3.org/1998/Math/MathML">
  <mi>x</mi><mo>=</mo><mn>5</mn>
</math>

<!-- Display (block) equation -->
<math xmlns="http://www.w3.org/1998/Math/MathML" display="block">
  <mfrac><mi>a</mi><mi>b</mi></mfrac>
</math>
```

---

## 7. HTML Output Format (wordhtml.com Style)

### 7.1 Reference Analysis

From the sample file `التشابه (جاهزة للنشر) - Copy.html`, the wordhtml.com format uses:

#### Headings
```html
<h1>التشابه <strong>Similarity</strong></h1>
<h2>تشابه المضلعات</h2>
<h3>حالات تشابه المثلثات</h3>
```
- Direct `<h1>` through `<h6>` tags
- No CSS classes on headings
- Bold text inside headings uses `<strong>`

#### Paragraphs
```html
<p>Paragraph text here.</p>
<p>&nbsp;</p>                        <!-- Empty paragraph as spacing -->
```
- Standard `<p>` tags
- `&nbsp;` for empty paragraphs (preserves Word's empty paragraphs as spacing)

#### Footnote References (in body text)
```html
<a href="#_ftn1" name="_ftnref1">[1]</a>
```
- Uses `name` attribute (not `id`)
- Naming convention: `_ftnref{N}` for reference, `_ftn{N}` for target
- No `<sup>` wrapper (just inline `<a>`)

#### Footnote Definitions (at bottom)
```html
<p><a href="#_ftnref1" name="_ftn1">[1]</a> Footnote text with <strong>bold</strong> and <em>italic</em>.</p>
```
- Bidirectional linking back to `_ftnref{N}`
- Full formatting preserved in footnote text
- No special container div (just `<p>` elements)

#### Tables
```html
<table>
<tbody>
<tr>
<td width="312">&nbsp;</td>
<td width="312">&nbsp;</td>
</tr>
<tr>
<td colspan="2" width="623">&nbsp;</td>
</tr>
</tbody>
</table>
```
- Always has `<tbody>`
- `width` attribute on `<td>` (in pixels, from Word column widths)
- `colspan` preserved
- No CSS classes or border attributes
- Empty cells contain `&nbsp;`

#### Lists
```html
<ul>
<li>List item text.</li>
<li>Another item.</li>
</ul>
```
- Standard `<ul><li>` and `<ol><li>`
- No CSS classes

#### Bold / Italic / Superscript
```html
<strong>bold text</strong>
<em>italic text</em>
<sup>th</sup>
```

#### Images
Images in the reference file appear as `&nbsp;` in cells (stripped by wordhtml.com). Our implementation will use:
```html
<img src="images/image1.png" alt="">
```

### 7.2 HTML Template for MathML Mode

```html
<!DOCTYPE html>
<html dir="rtl">
<head>
<meta charset="UTF-8">
<title>{title}</title>
</head>
<body>
{content}
{footnotes}
</body>
</html>
```

Minimal HTML. No `<style>` block, no `<script>` block. Clean like wordhtml.com output.

Optionally, if `include_styles` is True, add a minimal `<style>` block for readability (tables, spacing), but keep it simple.

---

## 8. Image Handling

### Current Behavior (preserved)
- Images are extracted from `word/media/` to `output_dir/images/`
- Referenced via relative path `images/image1.png`

### Enhanced Behavior
- Same extraction to `images/` subfolder
- In MathML mode, images are referenced the same way
- Table cells that contain images (common in the reference document) properly embed `<img>` tags
- Images in shapes/drawings handled via existing `_convert_drawing()` logic

### Implementation Detail
```python
def _extract_images(self, extract_dir, output_dir):
    """Extract all images to images/ subfolder"""
    media_dir = extract_dir / "word" / "media"
    if not media_dir.exists():
        return
    images_dir = output_dir / "images"
    images_dir.mkdir(exist_ok=True)
    for img in media_dir.iterdir():
        shutil.copy2(img, images_dir / img.name)
        self.images[img.name] = f"images/{img.name}"
```
No change needed. This already works correctly.

---

## 9. Table Handling

### Enhanced Table Conversion

The current `_convert_table()` is basic. Enhance it to match wordhtml.com:

```python
def _convert_table(self, tbl_elem):
    ns = self.namespaces
    rows = []

    for tr in tbl_elem.xpath('./w:tr', namespaces=ns):
        cells = []
        for tc in tr.xpath('./w:tc', namespaces=ns):
            # Get cell properties
            tc_pr = tc.find('w:tcPr', namespaces=ns)
            width = ''
            colspan = ''

            if tc_pr is not None:
                # Width
                tc_w = tc_pr.find('w:tcW', namespaces=ns)
                if tc_w is not None:
                    w_val = tc_w.get(f'{{{ns["w"]}}}w', '')
                    w_type = tc_w.get(f'{{{ns["w"]}}}type', 'dxa')
                    if w_val and w_type == 'dxa':
                        # Convert twips to approximate pixels (1 twip = 1/1440 inch, 96 ppi)
                        px = int(int(w_val) * 96 / 1440)
                        width = f' width="{px}"'

                # Colspan (gridSpan)
                grid_span = tc_pr.find('w:gridSpan', namespaces=ns)
                if grid_span is not None:
                    span_val = grid_span.get(f'{{{ns["w"]}}}val', '')
                    if span_val and int(span_val) > 1:
                        colspan = f' colspan="{span_val}"'

            # Convert cell content
            content = []
            for p in tc.xpath('./w:p', namespaces=ns):
                p_content = self._convert_paragraph_content(p)
                if p_content.strip():
                    content.append(f'<p>{p_content}</p>')

            cell_html = '\n'.join(content) if content else '&nbsp;'
            cells.append((cell_html, width, colspan))

        rows.append(cells)

    # Generate HTML
    html = ['<table>', '<tbody>']
    for row in rows:
        html.append('<tr>')
        for cell_html, width, colspan in row:
            html.append(f'<td{colspan}{width}>{cell_html}</td>')
        html.append('</tr>')
    html.append('</tbody>')
    html.append('</table>')
    return '\n'.join(html)
```

---

## 10. Section & Content Preservation

### Problem
Word documents have `w:sectPr` (section properties) elements that define page layout, columns, and section breaks. The current code skips `sectPr` entirely, which can cause content after section breaks to be lost.

### Solution

1. **Don't skip sectPr entirely** - Process elements that come after section breaks.
2. **Handle section breaks as `<hr>` or `<div>` separators** when meaningful.
3. **Ensure all `w:p` and `w:tbl` elements in `w:body` are processed**, regardless of which section they belong to.

```python
def _convert_body(self, doc_root):
    ns = self.namespaces
    body = doc_root.xpath('//w:body', namespaces=ns)[0]

    parts = []
    list_items = []
    current_list = None

    for child in body:
        tag = child.tag.split('}')[-1]

        if tag == 'p':
            # ... existing paragraph/list handling ...
            # ALSO check for section breaks inside paragraphs
            sect_pr = child.find('w:pPr/w:sectPr', namespaces=ns)
            if sect_pr is not None:
                # Section break inside paragraph - still process the paragraph
                # and add a separator
                parts.append(self._convert_paragraph(child))
                parts.append('<hr class="section-break">')
                continue

            # ... rest of existing logic ...

        elif tag == 'tbl':
            # ... existing table handling ...

        elif tag == 'sectPr':
            # Final section properties - safe to skip
            continue

        else:
            # Process any other element type to avoid content loss
            text = self._extract_text(child)
            if text.strip():
                parts.append(f'<p>{self._escape(text)}</p>')

    # ... rest unchanged ...
```

### Content Trimming Fix

The current code processes all children of `w:body`, so content should not be trimmed. The main risk is:
1. **Equations in textboxes** being lost when `convert_shapes_to_svg` is False (currently returns `''` for shapes). In MathML mode, we should extract text and equations from textboxes even when not converting shapes to SVG.
2. **Empty paragraphs** being dropped. wordhtml.com preserves them as `<p>&nbsp;</p>`.

Fix empty paragraph handling:
```python
def _convert_paragraph(self, p_elem):
    # ... existing style/heading detection ...

    content = self._convert_paragraph_content(p_elem)
    if not content.strip():
        # wordhtml.com preserves empty paragraphs
        if self.config.output_format == "mathml_html":
            return '<p>&nbsp;</p>'
        return ''

    if heading_level:
        return f'<h{heading_level}>{content}</h{heading_level}>'
    return f'<p>{content}</p>'
```

---

## 11. Testing Strategy

### Test Documents

Use existing test documents:
- `التشابه (جاهزة للنشر) - Copy.docx` - Has equations, footnotes, tables, images
- Compare output with `التشابه (جاهزة للنشر) - Copy.html` (wordhtml.com reference)

### Test Cases

| Test | What to Verify |
|------|---------------|
| Basic MathML | Simple inline equation renders correctly |
| Display equation | Block equation with `display="block"` |
| Fraction | `<mfrac>` output correct |
| Matrix | `<mtable>` with rows and cells |
| Integral with limits | `<munderover>` with operator |
| Greek symbols | Unicode renders in `<mi>` |
| Subscript/superscript | `<msub>`, `<msup>`, `<msubsup>` |
| Nested equations | Fraction inside radical inside superscript |
| Equations in shapes | Detected and converted from mc:Choice/Fallback |
| Footnotes | `_ftn`/`_ftnref` naming, bidirectional links |
| Tables | Width, colspan preserved |
| Images | Extracted to images/ folder, referenced correctly |
| Headings | h1-h6 without CSS classes |
| Empty paragraphs | Preserved as `<p>&nbsp;</p>` |
| Section breaks | Content after breaks not lost |
| RTL text | `dir="rtl"` on `<html>` |
| LaTeX mode | Existing functionality unchanged |
| Browser rendering | Open HTML in Chrome, Firefox, Edge - math displays |

### Validation Script

Create `backend/test_mathml_conversion.py`:
```python
"""Test MathML conversion vs reference HTML"""
# 1. Convert test DOCX with MathML mode
# 2. Compare structure with wordhtml.com reference
# 3. Validate all MathML elements are well-formed
# 4. Check footnote links are correct
# 5. Verify no JavaScript in output
```

---

## 12. Migration & Backward Compatibility

### What Changes for Existing Users

| Aspect | Before | After |
|--------|--------|-------|
| Default mode | LaTeX + MathJax | MathML (native) |
| JavaScript needed | Yes (MathJax) | No (MathML native) |
| Output files | HTML + images + .docx | HTML + images (no .docx in MathML mode) |
| Equation format | LaTeX text in HTML | `<math>` MathML elements in HTML |
| Browser support | All (via MathJax) | All modern (Chrome 109+, Firefox, Safari, Edge) |

### Backward Compatibility Guarantee

- **LaTeX mode still works**: Select "LaTeX + MathJax" in UI to use the old flow
- **API compatible**: Same endpoint, same `conversion_config` JSON. New `output_format` field defaults to `mathml_html`, but `latex_html` still works
- **No code deleted**: `DirectOmmlToLatex`, `EnhancedZipConverter`, MathJax generation - all preserved
- **Existing tests still pass**: No changes to existing code paths

### Default Change

The default `output_format` changes from (implied) LaTeX to `mathml_html`. This is the only breaking change. Users who relied on the default LaTeX output need to explicitly set `output_format: "latex_html"`.

---

## Implementation Order

### Phase 1: Core MathML Converter
1. Create `backend/doc_processor/omml_to_mathml.py` with `OmmlToMathMLConverter`
2. Test with sample OMML elements extracted from test documents
3. Verify MathML output renders in browsers

### Phase 2: HTML Generator Enhancement
4. Add `output_format` to `ConversionConfig`
5. Add `_convert_mathml_mode()` to `FullWordToHTMLConverter`
6. Enhance `_convert_paragraph_content()` for inline OMML handling
7. Add wordhtml.com-style footnotes, tables, sections
8. Test with `التشابه` document, compare with reference HTML

### Phase 3: Integration
9. Update `main.py` to pass `output_format` from config
10. Update `App.vue` with output format toggle
11. End-to-end test: upload DOCX, verify HTML output

### Phase 4: Polish
12. Handle edge cases (nested equations, empty cells, special symbols)
13. Test with multiple documents
14. Verify LaTeX mode still works unchanged
15. Cross-browser testing (Chrome, Firefox, Edge, Safari)

---

## Summary of Changes

| Category | Files Changed | Lines Added (est.) | Lines Modified (est.) |
|----------|--------------|-------------------|---------------------|
| New converter | `omml_to_mathml.py` | ~400 | 0 |
| HTML generator | `word_to_html_full.py` | ~200 | ~50 |
| API routing | `main.py` | ~5 | ~10 |
| Frontend UI | `App.vue` | ~30 | ~5 |
| Test script | `test_mathml_conversion.py` | ~100 | 0 |
| **Total** | **5 files** | **~735** | **~65** |

All existing functionality is preserved. The change is additive (new converter, new code paths) with a single config switch to control which path is used.
