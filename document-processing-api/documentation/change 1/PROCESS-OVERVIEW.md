# Document Processing — What's Really Happening

## What You See

```
    📄 Word File    ───────────>    🌐 HTML Page

    "Just convert it"
```

## What's Actually Happening

```
    📄 Word File
         │
         │  A Word file is NOT a simple document.
         │  It's an archive containing 10+ XML files,
         │  images, styles, footnotes, and relationships.
         │
         ▼
    ┌─────────────────────────────────────────────┐
    │                                             │
    │   📦 Unpack the archive                     │
    │                                             │
    │   Inside we find:                           │
    │   • The document text (as XML)              │
    │   • Formatting rules                        │
    │   • Images (as separate files)              │
    │   • Footnotes (as separate file)            │
    │   • Table structures                        │
    │   • Mathematical equations                  │
    │   • Shapes and drawings                     │
    │   • Numbered list rules                     │
    │   • Section & page layout settings          │
    │                                             │
    └──────────────────┬──────────────────────────┘
                       │
                       ▼
    ┌─────────────────────────────────────────────┐
    │                                             │
    │   🔍 Find ALL mathematical equations        │
    │                                             │
    │   Equations are NOT in one place.           │
    │   They are scattered across:                │
    │                                             │
    │   • Normal paragraphs                       │
    │   • Inside shapes and text boxes            │
    │   • Inside legacy compatibility sections    │
    │   • Inside drawing objects                  │
    │   • Some are duplicated for compatibility   │
    │                                             │
    │   Each must be found and identified.        │
    │                                             │
    └──────────────────┬──────────────────────────┘
                       │
                       ▼
    ┌─────────────────────────────────────────────┐
    │                                             │
    │   🧮 Convert EACH equation                  │
    │                                             │
    │   Every equation is a tree of nested        │
    │   elements. A single equation can contain:  │
    │                                             │
    │   • Fractions (with numerator/denominator)  │
    │   • Matrices (rows and columns of values)   │
    │   • Integrals, sums, products               │
    │   • Square roots (with optional degree)     │
    │   • Superscripts and subscripts             │
    │   • Greek letters and special symbols       │
    │   • Brackets, parentheses, braces           │
    │   • Accents (hat, bar, tilde, arrow)        │
    │   • Limits and function names               │
    │   • ALL of the above nested inside          │
    │     each other in any combination           │
    │                                             │
    │   Each type needs dedicated handling.        │
    │   There are 20+ different element types.    │
    │                                             │
    └──────────────────┬──────────────────────────┘
                       │
                       ▼
    ┌─────────────────────────────────────────────┐
    │                                             │
    │   📝 Convert all document elements          │
    │                                             │
    │   Besides equations, we must handle:        │
    │                                             │
    │   • Headings (detect levels 1-6)            │
    │   • Paragraphs (preserve formatting)        │
    │   • Bold, italic, underline text            │
    │   • Tables (widths, merged cells, nesting)  │
    │   • Numbered and bullet lists               │
    │   • Footnotes (with clickable links)        │
    │   • Images (extract and reference)          │
    │   • Shapes and drawings                     │
    │   • Hyperlinks                              │
    │   • Right-to-left Arabic text               │
    │   • Section breaks                          │
    │   • Empty paragraphs and spacing            │
    │                                             │
    └──────────────────┬──────────────────────────┘
                       │
                       ▼
    ┌─────────────────────────────────────────────┐
    │                                             │
    │   🏗️ Assemble the final HTML page           │
    │                                             │
    │   • Combine all converted elements          │
    │   • Add proper document structure           │
    │   • Link footnotes bidirectionally          │
    │   • Reference extracted images              │
    │   • Ensure right-to-left text works         │
    │                                             │
    └──────────────────┬──────────────────────────┘
                       │
                       ▼
                 🌐 HTML Page
                 + 📁 Images folder


    This entire process was built TWICE:

    ✅ First time: LaTeX output (requires JavaScript to display)
    ✅ Second time: MathML output (works without JavaScript)

    Both share the document processing, but each requires
    its own equation converter with 20+ element handlers.
```


## Why There's No "Just Use an Existing Tool"

```
    ┌─────────────────────────────────────────────┐
    │                                             │
    │   What existing tools CAN do:               │
    │                                             │
    │   ✓ Convert equations only (not full doc)   │
    │   ✓ Convert simple documents (no equations) │
    │   ✓ Convert with known bugs and limitations │
    │                                             │
    ├─────────────────────────────────────────────┤
    │                                             │
    │   What NO existing tool does:               │
    │                                             │
    │   ✗ Full document + equations + footnotes   │
    │     + tables + images + RTL Arabic text     │
    │     + shapes — all in one pipeline          │
    │                                             │
    │   ✗ Find equations in ALL 5 locations       │
    │     inside Word's XML structure             │
    │                                             │
    │   ✗ Produce clean, copy-pasteable HTML      │
    │     with no JavaScript dependency           │
    │                                             │
    │   ✗ Match specific output format            │
    │     (wordhtml.com conventions)              │
    │                                             │
    │   ✗ Handle Arabic right-to-left text        │
    │     alongside mathematical equations        │
    │                                             │
    └─────────────────────────────────────────────┘


    Here's what's available and why it's not enough:

    ┌──────────────────┬──────────────────────────┐
    │ Tool             │ What's missing            │
    ├──────────────────┼──────────────────────────┤
    │                  │                           │
    │ Microsoft's XSLT │ Only converts equations.  │
    │ (omml2mml.xsl)  │ Known bugs. Does not      │
    │                  │ handle full documents.    │
    │                  │                           │
    ├──────────────────┼──────────────────────────┤
    │                  │                           │
    │ Pandoc           │ Documented issues with    │
    │                  │ equation accuracy.         │
    │                  │ Moves inline equations.   │
    │                  │ Loses equation numbers.   │
    │                  │ No custom HTML format.    │
    │                  │                           │
    ├──────────────────┼──────────────────────────┤
    │                  │                           │
    │ wordhtml.com     │ Strips all equations      │
    │                  │ entirely. They disappear  │
    │                  │ from the output.          │
    │                  │                           │
    ├──────────────────┼──────────────────────────┤
    │                  │                           │
    │ MathType         │ Manual one-by-one copy.   │
    │                  │ Not automated. Not a      │
    │                  │ pipeline. Costs $$.       │
    │                  │                           │
    ├──────────────────┼──────────────────────────┤
    │                  │                           │
    │ omml2mathml      │ Equation-only converter.  │
    │ (open source)    │ No document handling.     │
    │                  │ No footnotes, tables,     │
    │                  │ images, or RTL support.   │
    │                  │                           │
    └──────────────────┴──────────────────────────┘

    CONCLUSION: A custom solution is the only way
    to meet all requirements together.
```


## The Scale of Work — Simple Numbers

```
    WHAT WAS BUILT (existing system):

    ┌─────────────────────────────────────────────┐
    │                                             │
    │   📁 10+ source code files                  │
    │   📝 ~3,000 lines of code                   │
    │   🧮 20+ equation element handlers          │
    │   🔣 100+ mathematical symbol mappings      │
    │   📋 3 processor types                      │
    │   🖥️ Full web interface (upload/download)   │
    │   🐳 Docker deployment configuration        │
    │   ☁️ Google Cloud deployment pipeline       │
    │                                             │
    └─────────────────────────────────────────────┘


    WHAT THE NEW CHANGE ADDS:

    ┌─────────────────────────────────────────────┐
    │                                             │
    │   📄 1 new source file (equation converter) │
    │   📝 ~735 new lines of code                 │
    │   ✏️ ~65 modified lines in 4 existing files │
    │   🧮 20+ NEW equation element handlers      │
    │     (different output format = different    │
    │      conversion logic for each one)         │
    │   📋 1,080-line technical specification     │
    │   🔀 Full backward compatibility            │
    │     (nothing breaks, old mode still works)  │
    │                                             │
    └─────────────────────────────────────────────┘
```
