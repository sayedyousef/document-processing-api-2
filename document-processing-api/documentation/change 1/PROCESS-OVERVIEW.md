# Document Processing — What's Really Happening

## What You See

```
    📄 Word File    ───>    🌐 HTML Page

    "Just convert it, how hard can it be?"
```

## What's Actually Happening

```
    📄 Word File
         │
         │  Step 1: Open the box
         │
         ▼
    ┌─────────────────────────────────────────────┐
    │                                              │
    │   A Word file is actually an archive         │
    │   containing 10+ separate files inside it.   │
    │                                              │
    │   Think of it like a box with:               │
    │   • The actual text                          │
    │   • The formatting rules                     │
    │   • All the images (as separate files)       │
    │   • All the footnotes (as a separate file)   │
    │   • Table structures                         │
    │   • Mathematical equations                   │
    │   • Shapes and drawings                      │
    │   • List numbering rules                     │
    │   • Page layout settings                     │
    │                                              │
    │   Each of these needs to be read separately  │
    │   and understood in relation to the others.  │
    │                                              │
    └──────────────────┬───────────────────────────┘
                       │
                       │  Step 2: Find the equations
                       ▼
    ┌─────────────────────────────────────────────┐
    │                                              │
    │   Equations are scattered across 5 different │
    │   locations inside the file. Some are inside │
    │   shapes, some are copies for compatibility  │
    │   with older versions of Word.               │
    │                                              │
    │   A document with 150 equations might        │
    │   actually contain 250+ equation fragments   │
    │   when you count all the hidden copies.      │
    │                                              │
    │   Each one must be found and identified.     │
    │                                              │
    └──────────────────┬───────────────────────────┘
                       │
                       │  Step 3: Translate every equation
                       ▼
    ┌─────────────────────────────────────────────┐
    │                                              │
    │   Each equation is a tree of nested pieces.  │
    │   For example, one equation might contain:   │
    │                                              │
    │   "the integral from 0 to infinity of        │
    │    a fraction whose numerator is the         │
    │    square root of x-squared plus y-squared   │
    │    and whose denominator is sigma times      │
    │    the sine of theta"                        │
    │                                              │
    │   That's: an integral, containing a          │
    │   fraction, containing a square root,        │
    │   containing a superscript, plus Greek       │
    │   letters, plus a function name — all        │
    │   nested inside each other.                  │
    │                                              │
    │   There are 20+ different piece types.       │
    │   Each needs its own translation rule.       │
    │   Multiply that by 150 equations.            │
    │                                              │
    └──────────────────┬───────────────────────────┘
                       │
                       │  Step 4: Convert everything else
                       ▼
    ┌─────────────────────────────────────────────┐
    │                                              │
    │   Equations are only part of the job.        │
    │   The system also handles:                   │
    │                                              │
    │   • Headings (6 levels)                      │
    │   • Tables (merged cells, widths, nesting)   │
    │   • Footnotes (with clickable links          │
    │     that go both directions)                 │
    │   • Images (extracted to a folder)           │
    │   • Bold, italic, superscript text           │
    │   • Numbered and bullet lists                │
    │   • Shapes and drawings                      │
    │   • Hyperlinks                               │
    │   • Arabic right-to-left text                │
    │   • Section breaks                           │
    │                                              │
    │   All of this must work together correctly.  │
    │                                              │
    └──────────────────┬───────────────────────────┘
                       │
                       │  Step 5: Assemble the page
                       ▼
                  🌐 HTML Page + 📁 Images folder


    ═══════════════════════════════════════════

    This was built ONCE for the current system.

    The new change requires building the
    equation translation engine A SECOND TIME
    with completely different output rules.

    ═══════════════════════════════════════════
```

---

## Don't Existing Tools Already Do This?

**Honest answer: partially, but none cover our full set of requirements.**

There are tools that handle PARTS of this problem. Here's what they can and cannot do:

```
    ┌────────────────────────────────────────────────────────────────┐
    │                                                                │
    │                   WHAT WE NEED (all together)                  │
    │                                                                │
    │   ✓ Convert equations (150+ per document)                     │
    │   ✓ Convert full document structure (tables, footnotes, etc.) │
    │   ✓ Arabic right-to-left text support                         │
    │   ✓ Clean MathML output (no JavaScript)                       │
    │   ✓ Specific HTML format (wordhtml.com conventions)           │
    │   ✓ Automated batch processing                                │
    │   ✓ Free / no per-seat licensing costs                        │
    │                                                                │
    └────────────────────────────────────────────────────────────────┘
```

### Tool 1: Pandoc (free, open source, code-callable)

**The closest alternative.** Pandoc is a command-line tool that CAN convert Word to HTML with MathML equations. It can be called from code and handles batch processing.

```
    What Pandoc CAN do                What Pandoc CANNOT do well
    ─────────────────────             ───────────────────────────
    ✓ Convert equations to MathML     ✗ RTL + footnotes are broken
    ✓ Handle basic tables               (footnote numbers appear in
    ✓ Handle basic footnotes              wrong position in Arabic)
    ✓ Handle images                   ✗ Equations get repositioned
    ✓ Batch processing                  (moved to end of paragraph
    ✓ Free                               instead of staying inline)
                                      ✗ No wordhtml.com format
                                      ✗ Table annotations lost
                                      ✗ No control over HTML style
```

Pandoc's RTL + footnote bug and equation positioning issues are **documented and open** on their GitHub. For Arabic academic documents with 150 equations, these are not minor issues — they break the output.

### Tool 2: Aspose.Words (commercial API, code-callable)

**A commercial product** that can convert Word to HTML. It has a MathML output mode.

```
    What Aspose CAN do                What Aspose CANNOT do well
    ─────────────────────             ───────────────────────────
    ✓ Convert to HTML                 ✗ Costs $1,199+ per developer
    ✓ MathML output mode              ✗ Had equation-to-image issues
    ✓ Handle tables, footnotes           for years (logged 2016,
    ✓ Professional support               described as "more complex
                                         than initially estimated")
                                      ✗ Some MathML rendering bugs
                                        (special math fonts, notations)
                                      ✗ No wordhtml.com format
                                      ✗ Ongoing licensing cost
```

Aspose could work for some use cases, but it's expensive and has had its own documented struggles with equation conversion.

### Tool 3: Equation-only converters (free, code-callable)

Libraries like `omml2mathml` (Ruby), `scienceai/omml2mathml` (JavaScript), and Microsoft's XSLT stylesheet.

```
    What they CAN do                  What they CANNOT do
    ─────────────────────             ───────────────────────────
    ✓ Convert equations only          ✗ No document handling at all
    ✓ Can be called from code         ✗ No footnotes, tables, images
    ✓ Free                            ✗ No RTL support
                                      ✗ No HTML generation
                                      ✗ Microsoft's XSLT has known bugs
                                      ✗ You still need to build
                                        everything else around them
```

### Tool 4: wordhtml.com and similar online converters

```
    What they CAN do                  What they CANNOT do
    ─────────────────────             ───────────────────────────
    ✓ Convert basic documents         ✗ Equations are STRIPPED entirely
    ✓ Nice HTML output                  (they disappear from output)
    ✓ Easy to use                     ✗ Manual only (no API)
                                      ✗ Cannot process 150 equations
                                      ✗ Not automatable
```

---

## Why Our Custom Solution Exists

```
    ┌────────────────────────────────────────────────────────────┐
    │                                                            │
    │  The problem is not that these tools are bad.              │
    │  The problem is that NO SINGLE TOOL handles the            │
    │  COMBINATION of all our requirements:                      │
    │                                                            │
    │    Accurate equations                                      │
    │    + Full document structure                                │
    │    + Arabic right-to-left text                              │
    │    + Clean MathML (no JavaScript)                           │
    │    + Specific output format                                 │
    │    + Automated processing                                   │
    │    + No licensing costs                                     │
    │                                                            │
    │  Pandoc comes closest but breaks on RTL + equations.        │
    │  Aspose works but costs $1,199+ per developer.              │
    │  Equation-only tools need an entire system built around     │
    │  them — which is essentially what we built.                 │
    │                                                            │
    └────────────────────────────────────────────────────────────┘
```

---

## The Scale of Work

```
    WHAT WAS ALREADY BUILT:

        10+ source code files
        ~3,000 lines of code
        20+ equation element handlers
        100+ mathematical symbol mappings
        Full web interface
        Cloud deployment pipeline


    WHAT THE NEW CHANGE ADDS:

        1 new source file
        ~735 new lines of code
        20+ NEW equation handlers (different format)
        4 existing files modified
        1,080-line technical specification
        All existing features still work
```
