# Document Processing API

A web application that converts Word documents (.docx) to HTML with LaTeX equation support.

## Features

- Convert Word documents to HTML
- Automatic OMML to LaTeX equation conversion
- Configurable equation markers (None, MATHSTART/END, or Custom)
- MathJax integration for equation rendering
- RTL (Arabic/Hebrew) text support
- SVG shape conversion

## Live Demo

**Production URL:** https://document-processing-api-788366675655.us-central1.run.app/

## Tech Stack

- **Frontend:** Vue.js 3 + Vite + TailwindCSS
- **Backend:** Python FastAPI
- **Deployment:** Google Cloud Run with CI/CD via Cloud Build

## Project Structure

```
document-processing-api-2/
├── document-processing-api/
│   ├── backend/           # FastAPI backend
│   │   ├── main.py        # Main API endpoints
│   │   ├── word_to_html_full.py    # Word to HTML converter
│   │   ├── enhanced_zip_converter.py  # OMML to LaTeX converter
│   │   └── doc_processor/  # Processing utilities
│   └── frontend/          # Vue.js frontend
│       └── src/
│           ├── App.vue    # Main application
│           └── components/
├── Dockerfile             # Multi-stage Docker build
├── cloudbuild.yaml        # CI/CD configuration
├── deploy.bat / deploy.sh # Manual deployment scripts
└── README.md
```

## Local Development

### Backend
```bash
cd document-processing-api/backend
pip install -r requirements.txt
python main.py
```
Backend runs on http://localhost:8000

### Frontend
```bash
cd document-processing-api/frontend
npm install
npm run dev
```
Frontend runs on http://localhost:5173

## Deployment

### Automatic (CI/CD)
Push to `main` branch triggers automatic deployment via Cloud Build:
```bash
git add . && git commit -m "your message" && git push
```

### Manual
```bash
# Windows
deploy.bat topic-project-408412 us-central1

# Linux/Mac
./deploy.sh topic-project-408412 us-central1
```

## API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/api/process` | POST | Upload and process documents |
| `/api/status/{job_id}` | GET | Check processing status |
| `/api/download/{job_id}` | GET | Download results as ZIP |
| `/api/health` | GET | Health check |

## Conversion Options

- **Equation Marker Style:** None, MATHSTART/END, or Custom prefix/suffix
- **Convert shapes to SVG:** Enable/disable SVG conversion
- **Include images:** Embed images in HTML
- **Include MathJax:** Add MathJax library for equation rendering
- **RTL direction:** Right-to-left text support

## Platform Detection

- **Windows:** Uses Word COM for equation processing (local development)
- **Linux/Cloud Run:** Uses ZIP-based processing (production)

## License

MIT
