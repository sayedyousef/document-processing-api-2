# Document Processing API — Architecture Diagrams

## 1. What a Word Document Actually Is

```
  document.docx  (looks like a simple file)
        |
        v
  ┌─────────────────────────────────────┐
  │          Actually a ZIP archive      │
  │                                      │
  │  word/                               │
  │   ├── document.xml     (content)     │
  │   ├── footnotes.xml    (footnotes)   │
  │   ├── styles.xml       (formatting)  │
  │   ├── numbering.xml    (lists)       │
  │   ├── media/                         │
  │   │    ├── image1.png                │
  │   │    ├── image2.jpg                │
  │   │    └── ...                       │
  │   └── _rels/                         │
  │        └── document.xml.rels         │
  │                                      │
  │  [Content_Types].xml                 │
  └─────────────────────────────────────┘
```

## 2. Where Equations Hide Inside the XML

Equations are NOT in one place. They appear in **5 different locations**:

```
  document.xml
  │
  ├── w:body
  │    ├── w:p (paragraph)
  │    │    └── m:oMath  ◄─────────────────── [1] MAIN BODY
  │    │
  │    ├── w:p (paragraph with shape)
  │    │    └── mc:AlternateContent
  │    │         ├── mc:Choice  (modern Word 2010+)
  │    │         │    └── wps:txbx
  │    │         │         └── m:oMath  ◄───── [2] MODERN SHAPES
  │    │         │
  │    │         └── mc:Fallback  (legacy copy)
  │    │              └── v:textbox
  │    │                   └── m:oMath  ◄───── [3] LEGACY FALLBACK
  │    │
  │    ├── w:p (paragraph with VML)
  │    │    └── w:pict
  │    │         └── v:shape
  │    │              └── v:textbox
  │    │                   └── m:oMath  ◄───── [4] VML TEXTBOX
  │    │
  │    └── w:p (paragraph with textbox)
  │         └── w:txbxContent
  │              └── m:oMath  ◄────────────── [5] OTHER TEXTBOX
  │
  └── ALL of these must be found and converted
```

## 3. What a Single Equation Looks Like in XML

A simple fraction like **a/b** is this in Word's XML:

```xml
<m:oMath>
  <m:f>                          ← fraction
    <m:num>                      ← numerator
      <m:r>                      ← run
        <m:t>a</m:t>             ← text "a"
      </m:r>
    </m:num>
    <m:den>                      ← denominator
      <m:r>
        <m:t>b</m:t>             ← text "b"
      </m:r>
    </m:den>
  </m:f>
</m:oMath>
```

A real equation has **dozens of nested elements** — fractions inside radicals inside superscripts inside matrices.

## 4. The Current System (LaTeX + MathJax)

```
 ┌──────────┐
 │ Word     │
 │ .docx    │
 └────┬─────┘
      │
      v
 ┌──────────────────────────────────────────────┐
 │  STEP 1: Equation Pre-Processing              │
 │  (EnhancedZipConverter)                        │
 │                                                │
 │  ┌─────────┐    ┌──────────────┐               │
 │  │ Extract  │───>│ Find ALL     │               │
 │  │ ZIP      │    │ equations    │               │
 │  └─────────┘    │ (5 locations)│               │
 │                  └──────┬───────┘               │
 │                         │                       │
 │                         v                       │
 │                  ┌──────────────┐               │
 │                  │ Parse each   │               │
 │                  │ equation     │               │
 │                  │ (20+ element │               │
 │                  │  handlers)   │               │
 │                  └──────┬───────┘               │
 │                         │                       │
 │                         v                       │
 │                  ┌──────────────┐               │
 │                  │ Convert to   │               │
 │                  │ LaTeX text   │───> \frac{a}{b}│
 │                  └──────┬───────┘               │
 │                         │                       │
 │                         v                       │
 │                  ┌──────────────┐               │
 │                  │ Replace in   │               │
 │                  │ XML & repack │               │
 │                  └──────┬───────┘               │
 │                         │                       │
 └─────────────────────────┼───────────────────────┘
                           │
                      Modified .docx
                      (LaTeX text)
                           │
                           v
 ┌──────────────────────────────────────────────┐
 │  STEP 2: HTML Generation                      │
 │  (FullWordToHTMLConverter)                     │
 │                                                │
 │  ┌─────────┐    ┌──────────────┐               │
 │  │ Extract  │───>│ Load:        │               │
 │  │ ZIP      │    │ • styles     │               │
 │  │ again    │    │ • footnotes  │               │
 │  └─────────┘    │ • images     │               │
 │                  │ • numbering  │               │
 │                  └──────┬───────┘               │
 │                         │                       │
 │                         v                       │
 │                  ┌──────────────┐               │
 │                  │ Convert to   │               │
 │                  │ HTML tags    │               │
 │                  └──────┬───────┘               │
 │                         │                       │
 │                         v                       │
 │                  ┌──────────────┐               │
 │                  │ Add MathJax  │               │
 │                  │ JavaScript   │               │
 │                  └──────┬───────┘               │
 │                         │                       │
 └─────────────────────────┼───────────────────────┘
                           │
                           v
                    ┌──────────────┐
                    │  HTML file   │
                    │  + images/   │
                    │  + .docx     │
                    │              │
                    │  ⚠ Requires  │
                    │  JavaScript  │
                    │  to render   │
                    │  equations   │
                    └──────────────┘
```

## 5. The NEW System (MathML — No JavaScript)

```
 ┌──────────┐
 │ Word     │
 │ .docx    │
 └────┬─────┘
      │
      v
 ┌──────────────────────────────────────────────┐
 │  SINGLE STEP: Direct Conversion               │
 │  (Enhanced FullWordToHTMLConverter)             │
 │                                                │
 │  ┌─────────┐    ┌──────────────┐               │
 │  │ Extract  │───>│ Load:        │               │
 │  │ ZIP      │    │ • styles     │               │
 │  └─────────┘    │ • footnotes  │               │
 │                  │ • images     │               │
 │                  │ • numbering  │               │
 │                  └──────┬───────┘               │
 │                         │                       │
 │                         v                       │
 │        ┌────────────────────────────────┐       │
 │        │  Convert document body to HTML  │       │
 │        │                                 │       │
 │        │  For each element:              │       │
 │        │  ├── paragraph ──> <p>          │       │
 │        │  ├── heading ────> <h1>-<h6>    │       │
 │        │  ├── table ─────> <table>       │       │
 │        │  ├── image ─────> <img>         │       │
 │        │  ├── footnote ──> <a>           │       │
 │        │  │                              │       │
 │        │  └── EQUATION ──> ┌───────────┐ │       │
 │        │     (m:oMath)     │ NEW       │ │       │
 │        │                   │ OMML to   │ │       │
 │        │                   │ MathML    │ │       │
 │        │                   │ converter │ │       │
 │        │                   │ (20+ new  │ │       │
 │        │                   │ handlers) │ │       │
 │        │                   └─────┬─────┘ │       │
 │        │                         │       │       │
 │        │                   <math>        │       │
 │        │                    <mfrac>      │       │
 │        │                     <mi>a</mi>  │       │
 │        │                     <mi>b</mi>  │       │
 │        │                    </mfrac>     │       │
 │        │                   </math>       │       │
 │        └────────────────────────────────┘       │
 │                         │                       │
 └─────────────────────────┼───────────────────────┘
                           │
                           v
                    ┌──────────────┐
                    │  HTML file   │
                    │  + images/   │
                    │              │
                    │  ✓ No JS     │
                    │  ✓ Copy-     │
                    │    pasteable │
                    │  ✓ Native    │
                    │    rendering │
                    └──────────────┘
```

## 6. The Equation Converter — What "20+ Handlers" Means

Each math element type needs its own conversion logic:

```
  OMML Element          What It Represents         Converter Output
  ─────────────         ──────────────────         ────────────────
  m:f                   Fraction  (a/b)            <mfrac>
  m:rad                 Square root  (√x)          <msqrt> / <mroot>
  m:sSup                Superscript  (x²)          <msup>
  m:sSub                Subscript  (xₙ)            <msub>
  m:sSubSup             Both  (xₙ²)               <msubsup>
  m:nary                Integral/Sum  (∫∑∏)        <munderover>
  m:d                   Brackets  ((x))            <mrow> + <mo>
  m:m                   Matrix                     <mtable>
  m:acc                 Accent  (x̂)                <mover>
  m:func                Function  (sin, cos)       <mi> + <mrow>
  m:limLow              Limit  (lim)               <munder>
  m:eqArr               Piecewise / aligned        <mtable>
  m:r                   Text / symbols             <mi>/<mn>/<mo>
  m:bar                 Overbar                    <mover> + <mo>
  m:borderBox           Box around equation        <menclose>
  m:groupChr            Group character            <munder>/<mover>
  m:phant               Phantom (spacing)          <mphantom>
  m:sPre                Pre-sub/superscript        <mmultiscripts>
  ...                   ...                        ...

  + Symbol mapping (100+ Unicode math symbols)
  + Display vs. inline detection
  + Nested combinations of all of the above
```

## 7. Side-by-Side: What Changes

```
  ┌─────────────────────────┬─────────────────────────┐
  │   CURRENT (LaTeX)       │   NEW (MathML)           │
  ├─────────────────────────┼─────────────────────────┤
  │                         │                          │
  │  2-step pipeline        │  1-step pipeline         │
  │                         │                          │
  │  OMML ──> LaTeX text    │  OMML ──> MathML tags    │
  │  (existing converter)   │  (NEW converter needed)  │
  │                         │                          │
  │  Intermediate .docx     │  Direct to HTML          │
  │  generated              │  (no intermediate file)  │
  │                         │                          │
  │  Requires MathJax JS    │  No JavaScript at all    │
  │  to render equations    │  Browser renders natively│
  │                         │                          │
  │  Not copy-pasteable     │  Copy-paste works        │
  │  (JS must run first)    │  immediately             │
  │                         │                          │
  │  ~3000 lines of code    │  ~3700 lines of code     │
  │  (PRESERVED as-is)      │  (+735 new lines)        │
  │                         │                          │
  │  Still available as     │  New default option      │
  │  "LaTeX + MathJax"      │  "MathML (No JS)"        │
  │  option in UI           │                          │
  │                         │                          │
  └─────────────────────────┴─────────────────────────┘
```

## 8. Full System Overview

```
  ┌──────────────────────────────────────────────────────────┐
  │                    FRONTEND (Vue.js)                       │
  │                                                           │
  │   ┌─────────────┐  ┌──────────┐  ┌───────────────────┐   │
  │   │ File Upload  │  │ Settings │  │ Output Format     │   │
  │   │ (.docx)      │  │ Panel    │  │ ○ MathML (new)    │   │
  │   └──────┬───────┘  └────┬─────┘  │ ○ LaTeX (current) │   │
  │          │               │        └─────────┬─────────┘   │
  │          └───────┬───────┘                  │             │
  │                  │                          │             │
  └──────────────────┼──────────────────────────┼─────────────┘
                     │   POST /api/process      │
                     v                          v
  ┌──────────────────────────────────────────────────────────┐
  │                    BACKEND (FastAPI)                       │
  │                                                           │
  │            ┌─────────────────────────┐                    │
  │            │  Which output_format?   │                    │
  │            └────────┬────────────────┘                    │
  │                     │                                     │
  │          ┌──────────┴──────────┐                          │
  │          │                     │                          │
  │          v                     v                          │
  │   ┌─────────────┐      ┌─────────────┐                   │
  │   │  "mathml"   │      │  "latex"    │                   │
  │   │             │      │             │                   │
  │   │ OmmlToMath  │      │ OmmlToLaTeX │ (existing)        │
  │   │ MLConverter │      │ + MathJax   │                   │
  │   │   (NEW)     │      │             │                   │
  │   └──────┬──────┘      └──────┬──────┘                   │
  │          │                     │                          │
  │          v                     v                          │
  │   ┌─────────────┐      ┌─────────────┐                   │
  │   │ Clean HTML  │      │ HTML + JS   │                   │
  │   │ No scripts  │      │ + MathJax   │                   │
  │   │ + images/   │      │ + images/   │                   │
  │   │             │      │ + .docx     │                   │
  │   └─────────────┘      └─────────────┘                   │
  │                                                           │
  └───────────────────────────────────────────────────────────┘
```
