# ⚠️ IMPORTANT: These Assets are for the PUBLISHING SITE

## NOT for Our System!

The files in this folder are **NOT** part of our document processing system. They are assets that should be deployed on the **publishing site** where articles are displayed as HTML.

---

## 📍 Where These Files Belong

### Our System (Document Processing API):
- **Purpose**: Preprocesses Word documents
- **Output**: Modified .docx files with LaTeX markers
- **Location**: Backend processing server
- **No frontend needed**

### Publishing Site (Article Display):
- **Purpose**: Displays articles to readers
- **Input**: HTML with equation markers
- **Location**: Publishing website/platform
- **Uses these JavaScript files**

---

## 🎯 Clear Separation

```
OUR SYSTEM                           PUBLISHING SITE
─────────────────                    ─────────────────

Backend Only                         Frontend Website
│                                    │
├── Process .docx                    ├── Display articles
├── Convert OMML → LaTeX             ├── Process markers
├── Add markers                      ├── Render equations
└── Output .docx                     └── Use equation_processor.js

         ↓                                    ↑
         │                                    │
    Modified .docx ──────────────────────────┘
    (with markers)      (via any HTML converter)
```

---

## 📦 Files in This Folder

### equation_processor.js
- **What**: JavaScript that converts equation markers to HTML elements
- **Where to use**: On the publishing/article display website
- **When it runs**: After HTML is loaded in the browser
- **Purpose**: Converts `MATHSTARTINLINE...\)MATHENDINLINE` to proper HTML

---

## 🚀 How to Deploy on Publishing Site

1. **Upload equation_processor.js** to the publishing site's assets folder

2. **Include in article HTML template**:
```html
<!DOCTYPE html>
<html>
<head>
    <title>Article</title>

    <!-- MathJax for equation rendering -->
    <script src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js"></script>
</head>
<body>
    <!-- Article content with markers -->
    ${article_html_content}

    <!-- Equation processor (converts markers to HTML) -->
    <script src="/assets/equation_processor.js"></script>
</body>
</html>
```

3. **The JavaScript automatically**:
   - Finds all equation markers
   - Converts them to span/div elements
   - Triggers MathJax rendering

---

## ❌ Common Mistakes

### Wrong:
- Putting this JavaScript in our document processing API
- Including it in the Word document
- Running it on the server

### Right:
- Deploy on the publishing website
- Include in article HTML pages
- Let it run in readers' browsers

---

## 📝 Summary

**Our system** = Backend utility that preprocesses Word documents

**Publishing site** = Where these JavaScript files are actually used

The JavaScript is the **final step** that happens in the **reader's browser**, not in our processing system!