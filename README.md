# Document Processing API

A web application that converts Word documents (.docx) to HTML with LaTeX equation support using MathJax.

## Features

- **Word to HTML Conversion** - Full document structure preserved
- **LaTeX Equations** - OMML equations converted to LaTeX with MathJax rendering
- **Complete Document Support:**
  - Headings (6 levels)
  - Tables (with width, colspan, merged cells)
  - Footnotes (clickable bidirectional links)
  - Images (extracted to folder)
  - Numbered and bullet lists (continuation support)
  - Bold, italic, superscript text
  - RTL (Arabic/Hebrew) text support
  - Section breaks

## Live Demo

**Production URL:** https://document-processing-api-788366675655.us-central1.run.app/

## Tech Stack

- **Frontend:** Vue.js 3 + Vite + TailwindCSS
- **Backend:** Python FastAPI
- **Deployment:** Google Cloud Run with Cloud Build

## Project Structure

```
document-processing-api-2/
├── document-processing-api/
│   ├── backend/
│   │   ├── main.py                    # FastAPI endpoints
│   │   ├── word_to_html_full.py       # Main converter (LaTeX + MathML modes)
│   │   ├── enhanced_zip_converter.py  # OMML to LaTeX converter
│   │   └── doc_processor/
│   │       ├── omml_2_latex.py        # OMML parsing
│   │       └── omml_to_mathml.py      # OMML to MathML converter
│   └── frontend/
│       └── src/
│           ├── App.vue                # Main application
│           └── components/
│               ├── FileUploader.vue
│               ├── JobStatus.vue
│               └── ResultDownload.vue
├── Dockerfile                         # Multi-stage Docker build
├── cloudbuild.yaml                    # Cloud Build configuration
└── README.md
```

---

## Local Development

### Backend

```bash
cd document-processing-api/backend
pip install -r requirements.txt
python main.py
```

Backend runs on **http://localhost:8000**

### Frontend

```bash
cd document-processing-api/frontend
npm install
npm run dev
```

Frontend runs on **http://localhost:5173**

### Running Both Together

1. Start backend in one terminal
2. Start frontend in another terminal
3. Open http://localhost:5173 in browser

---

## Cloud Deployment

### Project Info

- **Project ID:** `topic-project-408412`
- **Region:** `us-central1`
- **Service:** `document-processing-api`

### Prerequisites

1. Install Google Cloud CLI: https://cloud.google.com/sdk/docs/install
2. Authenticate:
   ```bash
   gcloud auth login
   gcloud config set project topic-project-408412
   ```

### Deploy with Cloud Build

```bash
# From project root directory
gcloud builds submit --config=cloudbuild.yaml
```

This will:
1. Build the Docker image
2. Push to Google Container Registry
3. Deploy to Cloud Run

### Manual Deployment (Alternative)

```bash
# Build and push image
gcloud builds submit --tag gcr.io/topic-project-408412/document-processing-api:latest

# Deploy to Cloud Run
gcloud run deploy document-processing-api \
  --image gcr.io/topic-project-408412/document-processing-api:latest \
  --region us-central1 \
  --platform managed \
  --allow-unauthenticated \
  --memory 1Gi \
  --cpu 1 \
  --timeout 300
```

### View Deployment Status

```bash
# Check service status
gcloud run services describe document-processing-api --region us-central1

# View logs
gcloud run services logs read document-processing-api --region us-central1 --limit 50
```

---

## API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/api/process` | POST | Upload and process documents |
| `/api/status/{job_id}` | GET | Check processing status |
| `/api/download/{job_id}/{index}` | GET | Download specific result file |
| `/api/download/{job_id}` | GET | Download all results as ZIP |
| `/api/health` | GET | Health check |

---

## Conversion Options

| Option | Description | Default |
|--------|-------------|---------|
| **Convert shapes to SVG** | Convert Word shapes to SVG elements | Off |
| **Include images** | Embed images in HTML output | On |
| **Include MathJax** | Add MathJax library script | On |
| **RTL direction** | Right-to-left text direction | On |

---

## Output Files

When processing completes, you receive:
- **HTML file** - Complete document with embedded equations
- **images/** folder - Extracted images (if any)

The HTML includes MathJax script for rendering LaTeX equations. Simply open the HTML file in any modern browser.

### MathJax Script (for custom HTML)

```html
<script>
  window.MathJax = {
    tex: {
      inlineMath: [['\\(', '\\)']],
      displayMath: [['\\[', '\\]']]
    },
    svg: { fontCache: 'global' }
  };
</script>
<script src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-chtml.js" async></script>
```

---

## Troubleshooting

### Cloud Build Fails

1. Check authentication:
   ```bash
   gcloud auth application-default set-quota-project topic-project-408412
   ```

2. Verify project permissions:
   ```bash
   gcloud projects get-iam-policy topic-project-408412
   ```

### Local Backend Won't Start

1. Ensure Python 3.9+ installed
2. Install all requirements:
   ```bash
   pip install -r requirements.txt
   ```

### Equations Not Rendering

- Ensure MathJax script is included in HTML
- Check browser console for JavaScript errors
- Verify LaTeX syntax in output

---

## License

MIT
