"""
Flask API Server for Word to HTML Converter
============================================

This server provides the backend API for the upload page.
It receives configuration from the frontend and runs the converter.
"""

import os
import sys
import io
import json
import tempfile
import shutil
from pathlib import Path
from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename

# Import the converter
from word_to_html_full import FullWordToHTMLConverter, ConversionConfig

app = Flask(__name__)
CORS(app)  # Enable CORS for frontend

# Configuration
UPLOAD_FOLDER = Path(tempfile.gettempdir()) / 'word_converter_uploads'
OUTPUT_FOLDER = Path(tempfile.gettempdir()) / 'word_converter_output'
ALLOWED_EXTENSIONS = {'docx'}

UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)

app.config['UPLOAD_FOLDER'] = str(UPLOAD_FOLDER)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    """Serve the upload page"""
    return send_file('upload_page.html')


@app.route('/api/convert', methods=['POST'])
def convert_document():
    """
    Convert a Word document to HTML with the specified configuration.

    Expects:
    - file: The .docx file (multipart/form-data)
    - config: JSON string with conversion settings

    Returns:
    - JSON with success status and download URL
    """

    # Check if file was uploaded
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'No file uploaded'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'success': False, 'error': 'No file selected'}), 400

    if not allowed_file(file.filename):
        return jsonify({'success': False, 'error': 'Invalid file type. Only .docx files are allowed'}), 400

    # Get configuration from request
    config_json = request.form.get('config', '{}')
    try:
        config_data = json.loads(config_json)
    except json.JSONDecodeError:
        config_data = {}

    # Create ConversionConfig from request data
    config = ConversionConfig(
        convert_shapes_to_svg=config_data.get('convert_shapes_to_svg', True),
        include_images=config_data.get('include_images', True),
        inline_prefix=config_data.get('inline_prefix', 'MATHSTARTINLINE'),
        inline_suffix=config_data.get('inline_suffix', 'MATHENDINLINE'),
        display_prefix=config_data.get('display_prefix', 'MATHSTARTDISPLAY'),
        display_suffix=config_data.get('display_suffix', 'MATHENDDISPLAY'),
        include_mathjax=config_data.get('include_mathjax', True),
        rtl_direction=config_data.get('rtl_direction', True)
    )

    # Save uploaded file
    filename = secure_filename(file.filename)
    # Keep original name for Arabic filenames
    original_name = file.filename
    input_path = UPLOAD_FOLDER / filename
    file.save(str(input_path))

    # Create output directory for this conversion
    output_id = f"{Path(filename).stem}_{os.urandom(4).hex()}"
    output_dir = OUTPUT_FOLDER / output_id
    output_dir.mkdir(exist_ok=True)

    try:
        # Run conversion
        converter = FullWordToHTMLConverter(config)
        result = converter.convert(input_path, output_dir=output_dir)

        if result.get('success'):
            output_path = Path(result['output_path'])

            return jsonify({
                'success': True,
                'message': 'Conversion successful',
                'output_id': output_id,
                'filename': output_path.name,
                'download_url': f'/api/download/{output_id}/{output_path.name}',
                'download_zip_url': f'/api/download-zip/{output_id}'
            })
        else:
            return jsonify({
                'success': False,
                'error': result.get('error', 'Unknown error during conversion')
            }), 500

    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500

    finally:
        # Clean up uploaded file
        if input_path.exists():
            input_path.unlink()


@app.route('/api/download/<output_id>/<filename>')
def download_file(output_id, filename):
    """Download a converted HTML file"""
    output_dir = OUTPUT_FOLDER / output_id

    if not output_dir.exists():
        return jsonify({'error': 'File not found'}), 404

    return send_from_directory(str(output_dir), filename, as_attachment=True)


@app.route('/api/download-zip/<output_id>')
def download_zip(output_id):
    """Download all output files as a ZIP"""
    output_dir = OUTPUT_FOLDER / output_id

    if not output_dir.exists():
        return jsonify({'error': 'Output not found'}), 404

    # Create ZIP file
    zip_path = OUTPUT_FOLDER / f"{output_id}.zip"
    shutil.make_archive(str(zip_path.with_suffix('')), 'zip', str(output_dir))

    return send_file(str(zip_path), as_attachment=True, download_name=f"{output_id}.zip")


@app.route('/api/preview/<output_id>/<filename>')
def preview_file(output_id, filename):
    """Preview a converted HTML file in browser"""
    output_dir = OUTPUT_FOLDER / output_id

    if not output_dir.exists():
        return jsonify({'error': 'File not found'}), 404

    return send_from_directory(str(output_dir), filename)


@app.route('/api/preview/<output_id>/images/<image_name>')
def preview_image(output_id, image_name):
    """Serve images for preview"""
    images_dir = OUTPUT_FOLDER / output_id / 'images'

    if not images_dir.exists():
        return jsonify({'error': 'Image not found'}), 404

    return send_from_directory(str(images_dir), image_name)


if __name__ == '__main__':
    print("=" * 60)
    print("Word to HTML Converter API Server")
    print("=" * 60)
    print(f"Upload folder: {UPLOAD_FOLDER}")
    print(f"Output folder: {OUTPUT_FOLDER}")
    print("=" * 60)
    print("Starting server on http://localhost:5000")
    print("Open http://localhost:5000 in your browser")
    print("=" * 60)

    app.run(debug=True, host='0.0.0.0', port=5000)
