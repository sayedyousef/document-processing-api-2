# Documentation - Document Processing API

## Quick Reference

| Document | Purpose |
|----------|---------|
| [../../README.md](../../README.md) | Main project README - setup, deployment, usage |
| [TECHNICAL-SPECIFICATION.md](TECHNICAL-SPECIFICATION.md) | Technical details of conversion system |
| [change 1/](change%201/) | MathML feature documentation (implemented) |

---

## System Overview

The Document Processing API converts Word documents (.docx) to HTML with full equation support.

### Conversion Modes

| Mode | Status | Description |
|------|--------|-------------|
| **LaTeX + MathJax** | Default | Converts equations to LaTeX, renders with MathJax |
| **MathML (No JS)** | Available | Direct MathML output, native browser rendering |

### Key Features

- **Equation Conversion**: 150+ equations per document supported
- **Full Document Structure**: Headings, tables, footnotes, images, lists
- **RTL Support**: Arabic and Hebrew text
- **Cloud Deployment**: Google Cloud Run with auto-scaling

---

## Architecture

```
Word Document (.docx)
       │
       ▼
┌─────────────────────────────────────┐
│  1. Extract ZIP archive             │
│     - document.xml (content)        │
│     - footnotes.xml                 │
│     - styles.xml                    │
│     - numbering.xml                 │
│     - media/ (images)               │
└─────────────────────────────────────┘
       │
       ▼
┌─────────────────────────────────────┐
│  2. Parse & Convert                 │
│     - Find equations (5 locations)  │
│     - Convert OMML to LaTeX/MathML  │
│     - Convert structure to HTML     │
│     - Extract images                │
└─────────────────────────────────────┘
       │
       ▼
┌─────────────────────────────────────┐
│  3. Output                          │
│     - HTML file                     │
│     - images/ folder                │
└─────────────────────────────────────┘
```

---

## Files & Their Roles

### Backend

| File | Purpose | Lines |
|------|---------|-------|
| `main.py` | FastAPI endpoints, job management | ~630 |
| `word_to_html_full.py` | Main converter (LaTeX + MathML modes) | ~1200 |
| `enhanced_zip_converter.py` | OMML to LaTeX pre-processor | ~400 |
| `doc_processor/omml_2_latex.py` | OMML to LaTeX parser | ~820 |
| `doc_processor/omml_to_mathml.py` | OMML to MathML converter | ~720 |

### Frontend

| File | Purpose |
|------|---------|
| `App.vue` | Main application, settings |
| `FileUploader.vue` | Drag-drop file upload |
| `JobStatus.vue` | Processing status display |
| `ResultDownload.vue` | Download results, MathJax script |

---

## Deployment

### Production URL

https://document-processing-api-788366675655.us-central1.run.app/

### Deploy Command

```bash
gcloud builds submit --config=cloudbuild.yaml
```

See [../../README.md](../../README.md) for full deployment instructions.

---

## Chat History Archive

Historical development conversations preserved for reference:

- `chat_history/thread 1/` - Initial setup
- `chat_history/thread 2/` - VML textbox discovery
- `chat_history/thread 3/` - Final implementation

These contain detailed technical discussions but may be outdated.
