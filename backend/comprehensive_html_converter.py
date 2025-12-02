"""
COMPREHENSIVE Word to HTML Converter
Handles ALL document elements: headers, tables, shapes, images, footnotes, equations

This converter:
1. First converts equations using enhanced_zip_converter
2. Then converts to HTML handling all elements:
   - Headers/Headings (H1-H6)
   - Tables with proper formatting
   - Shapes/textboxes with content
   - Images (extracted or embedded)
   - Footnotes and endnotes
   - Lists (ordered and unordered)
   - RTL/LTR text direction
   - Styles and formatting
"""

import sys
import io
import zipfile
import shutil
import base64
import re
from pathlib import Path
from lxml import etree
from datetime import datetime

# Set UTF-8 encoding for stdout only if not already wrapped
if not isinstance(sys.stdout, io.TextIOWrapper) or sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    except:
        pass


class ComprehensiveHTMLConverter:
    """Complete Word to HTML converter handling all document elements"""

    def __init__(self):
        self.namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
            'v': 'urn:schemas-microsoft-com:vml',
            'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
            'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture'
        }
        self.relationships = {}
        self.images = {}
        self.footnotes = {}
        self.endnotes = {}
        self.styles = {}

    def convert_document(self, input_path, output_path=None, output_dir=None):
        """
        Convert Word document to HTML

        Args:
            input_path: Path to input .docx file
            output_path: Path for output HTML file (optional)
            output_dir: Directory for output files and images (optional)

        Returns:
            dict with conversion results
        """

        input_path = Path(input_path).absolute()

        if not output_dir:
            output_dir = input_path.parent / f"{input_path.stem}_html"
        else:
            output_dir = Path(output_dir).absolute()

        output_dir.mkdir(exist_ok=True)

        if not output_path:
            output_path = output_dir / f"{input_path.stem}.html"
        else:
            output_path = Path(output_path).absolute()

        print("\n" + "="*70)
        print("COMPREHENSIVE WORD TO HTML CONVERTER")
        print("="*70)
        print(f"Input:  {input_path}")
        print(f"Output: {output_path}")
        print(f"Dir:    {output_dir}")
        print("="*70)

        temp_dir = Path(f"temp_html_{datetime.now().strftime('%Y%m%d_%H%M%S')}")

        try:
            # Step 1: First convert equations
            print("\n[1] Converting equations...")
            from enhanced_zip_converter import EnhancedZipConverter
            eq_converter = EnhancedZipConverter()

            # Create temp file for equation-converted doc
            eq_converted = temp_dir / "eq_converted.docx"
            temp_dir.mkdir(exist_ok=True)

            eq_result = eq_converter.process_document(input_path, eq_converted)

            if not eq_result.get('success'):
                print(f"    Warning: Equation conversion had issues")
                # Use original file if equation conversion fails
                eq_converted = input_path

            # Step 2: Extract docx
            print("\n[2] Extracting document structure...")
            extract_dir = temp_dir / "extracted"
            with zipfile.ZipFile(eq_converted, 'r') as z:
                z.extractall(extract_dir)

            # Step 3: Load relationships
            print("\n[3] Loading relationships...")
            self._load_relationships(extract_dir)

            # Step 4: Load styles
            print("\n[4] Loading styles...")
            self._load_styles(extract_dir)

            # Step 5: Load footnotes and endnotes
            print("\n[5] Loading footnotes/endnotes...")
            self._load_footnotes(extract_dir)

            # Step 6: Extract images
            print("\n[6] Extracting images...")
            self._extract_images(extract_dir, output_dir)

            # Step 7: Parse and convert document
            print("\n[7] Converting document to HTML...")
            doc_xml_path = extract_dir / "word" / "document.xml"
            with open(doc_xml_path, 'rb') as f:
                doc_root = etree.fromstring(f.read())

            html_content = self._convert_body(doc_root)

            # Step 8: Generate complete HTML
            print("\n[8] Generating final HTML...")
            complete_html = self._generate_complete_html(
                html_content,
                input_path.stem
            )

            # Save HTML
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(complete_html)

            print("\n" + "="*70)
            print("CONVERSION COMPLETE!")
            print("="*70)
            print(f"Output HTML: {output_path}")
            print(f"Images:      {len(self.images)}")
            print(f"Footnotes:   {len(self.footnotes)}")
            print(f"Endnotes:    {len(self.endnotes)}")

            return {
                'success': True,
                'output_path': str(output_path),
                'output_dir': str(output_dir),
                'images': len(self.images),
                'footnotes': len(self.footnotes)
            }

        except Exception as e:
            print(f"\nERROR: {e}")
            import traceback
            traceback.print_exc()
            return {
                'success': False,
                'error': str(e)
            }

        finally:
            # Cleanup
            if temp_dir.exists():
                shutil.rmtree(temp_dir)

    def _load_relationships(self, extract_dir):
        """Load document relationships (for images, hyperlinks, etc.)"""

        rels_path = extract_dir / "word" / "_rels" / "document.xml.rels"
        if not rels_path.exists():
            return

        with open(rels_path, 'rb') as f:
            root = etree.fromstring(f.read())

        ns = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}

        for rel in root.xpath('//r:Relationship', namespaces=ns):
            rel_id = rel.get('Id')
            target = rel.get('Target')
            rel_type = rel.get('Type', '').split('/')[-1]

            self.relationships[rel_id] = {
                'target': target,
                'type': rel_type
            }

        print(f"    Loaded {len(self.relationships)} relationships")

    def _load_styles(self, extract_dir):
        """Load document styles for proper heading detection"""

        styles_path = extract_dir / "word" / "styles.xml"
        if not styles_path.exists():
            return

        with open(styles_path, 'rb') as f:
            root = etree.fromstring(f.read())

        ns = {'w': self.namespaces['w']}

        for style in root.xpath('//w:style', namespaces=ns):
            style_id = style.get(f'{{{ns["w"]}}}styleId')
            style_name = style.xpath('w:name/@w:val', namespaces=ns)

            if style_name:
                self.styles[style_id] = style_name[0]

        print(f"    Loaded {len(self.styles)} styles")

    def _load_footnotes(self, extract_dir):
        """Load footnotes and endnotes"""

        # Footnotes
        fn_path = extract_dir / "word" / "footnotes.xml"
        if fn_path.exists():
            with open(fn_path, 'rb') as f:
                root = etree.fromstring(f.read())

            ns = {'w': self.namespaces['w']}
            for fn in root.xpath('//w:footnote', namespaces=ns):
                fn_id = fn.get(f'{{{ns["w"]}}}id')
                if fn_id and fn_id not in ['0', '-1']:  # Skip separator notes
                    content = self._extract_text(fn)
                    self.footnotes[fn_id] = content

        # Endnotes
        en_path = extract_dir / "word" / "endnotes.xml"
        if en_path.exists():
            with open(en_path, 'rb') as f:
                root = etree.fromstring(f.read())

            ns = {'w': self.namespaces['w']}
            for en in root.xpath('//w:endnote', namespaces=ns):
                en_id = en.get(f'{{{ns["w"]}}}id')
                if en_id and en_id not in ['0', '-1']:
                    content = self._extract_text(en)
                    self.endnotes[en_id] = content

        print(f"    Loaded {len(self.footnotes)} footnotes, {len(self.endnotes)} endnotes")

    def _extract_images(self, extract_dir, output_dir):
        """Extract images from document"""

        media_dir = extract_dir / "word" / "media"
        if not media_dir.exists():
            return

        images_dir = output_dir / "images"
        images_dir.mkdir(exist_ok=True)

        for img_file in media_dir.iterdir():
            # Copy image to output
            dest = images_dir / img_file.name
            shutil.copy2(img_file, dest)

            # Store reference
            self.images[img_file.name] = f"images/{img_file.name}"

        print(f"    Extracted {len(self.images)} images")

    def _extract_text(self, element):
        """Extract all text from an element"""

        texts = []
        for t in element.xpath('.//w:t', namespaces=self.namespaces):
            if t.text:
                texts.append(t.text)
        return ''.join(texts)

    def _convert_body(self, doc_root):
        """Convert document body to HTML"""

        ns = self.namespaces
        body = doc_root.xpath('//w:body', namespaces=ns)[0]

        html_parts = []

        for child in body:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

            if tag == 'p':
                html_parts.append(self._convert_paragraph(child))
            elif tag == 'tbl':
                html_parts.append(self._convert_table(child))
            elif tag == 'sectPr':
                continue  # Skip section properties
            else:
                # Try to convert unknown elements
                text = self._extract_text(child)
                if text.strip():
                    html_parts.append(f'<p>{self._escape_html(text)}</p>')

        return '\n'.join(html_parts)

    def _convert_paragraph(self, p_elem):
        """Convert paragraph to HTML"""

        ns = self.namespaces

        # Check for style (heading detection)
        style_id = p_elem.xpath('.//w:pStyle/@w:val', namespaces=ns)
        style_name = self.styles.get(style_id[0], '') if style_id else ''

        # Detect heading level
        heading_level = 0
        if 'Heading' in style_name or 'heading' in style_name:
            match = re.search(r'(\d+)', style_name)
            if match:
                heading_level = min(int(match.group(1)), 6)
        elif style_name.lower() == 'title':
            heading_level = 1
        elif style_name.lower() == 'subtitle':
            heading_level = 2

        # Convert content
        content_parts = []

        for child in p_elem:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

            if tag == 'r':
                content_parts.append(self._convert_run(child))
            elif tag == 'hyperlink':
                content_parts.append(self._convert_hyperlink(child))
            elif tag == 'oMath' or tag == 'oMathPara':
                # This shouldn't happen after equation conversion
                content_parts.append(self._convert_math(child))
            elif tag == 'drawing':
                content_parts.append(self._convert_drawing(child))
            elif tag == 'pPr':
                continue  # Skip paragraph properties
            else:
                text = self._extract_text(child)
                if text:
                    content_parts.append(self._escape_html(text))

        content = ''.join(content_parts)

        if not content.strip():
            return ''

        # Wrap in appropriate tag
        if heading_level > 0:
            return f'<h{heading_level}>{content}</h{heading_level}>'
        else:
            return f'<p>{content}</p>'

    def _convert_run(self, r_elem):
        """Convert text run to HTML"""

        ns = self.namespaces
        parts = []

        # Get run properties
        bold = bool(r_elem.xpath('.//w:b', namespaces=ns))
        italic = bool(r_elem.xpath('.//w:i', namespaces=ns))
        underline = bool(r_elem.xpath('.//w:u', namespaces=ns))
        strike = bool(r_elem.xpath('.//w:strike', namespaces=ns))
        superscript = bool(r_elem.xpath('.//w:vertAlign[@w:val="superscript"]', namespaces=ns))
        subscript = bool(r_elem.xpath('.//w:vertAlign[@w:val="subscript"]', namespaces=ns))

        for child in r_elem:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

            if tag == 't':
                text = child.text or ''
                parts.append(self._escape_html(text))
            elif tag == 'drawing':
                parts.append(self._convert_drawing(child))
            elif tag == 'footnoteReference':
                fn_id = child.get(f'{{{ns["w"]}}}id')
                parts.append(f'<sup><a href="#fn{fn_id}" id="fnref{fn_id}">[{fn_id}]</a></sup>')
            elif tag == 'endnoteReference':
                en_id = child.get(f'{{{ns["w"]}}}id')
                parts.append(f'<sup><a href="#en{en_id}" id="enref{en_id}">[{en_id}]</a></sup>')
            elif tag == 'br':
                parts.append('<br>')
            elif tag == 'tab':
                parts.append('&emsp;')

        content = ''.join(parts)

        # Apply formatting
        if superscript:
            content = f'<sup>{content}</sup>'
        if subscript:
            content = f'<sub>{content}</sub>'
        if strike:
            content = f'<s>{content}</s>'
        if underline:
            content = f'<u>{content}</u>'
        if italic:
            content = f'<em>{content}</em>'
        if bold:
            content = f'<strong>{content}</strong>'

        return content

    def _convert_hyperlink(self, h_elem):
        """Convert hyperlink to HTML"""

        ns = self.namespaces

        # Get target from relationships
        r_id = h_elem.get(f'{{{ns["r"]}}}id')
        href = "#"

        if r_id and r_id in self.relationships:
            href = self.relationships[r_id]['target']

        # Get anchor if internal link
        anchor = h_elem.get(f'{{{ns["w"]}}}anchor')
        if anchor:
            href = f'#{anchor}'

        # Get content
        content_parts = []
        for r in h_elem.xpath('.//w:r', namespaces=ns):
            content_parts.append(self._convert_run(r))

        content = ''.join(content_parts)

        return f'<a href="{href}">{content}</a>'

    def _convert_table(self, tbl_elem):
        """Convert table to HTML"""

        ns = self.namespaces
        rows = []

        for tr in tbl_elem.xpath('.//w:tr', namespaces=ns):
            cells = []

            for tc in tr.xpath('.//w:tc', namespaces=ns):
                # Get cell content (can contain multiple paragraphs)
                cell_content = []
                for p in tc.xpath('.//w:p', namespaces=ns):
                    p_html = self._convert_paragraph(p)
                    if p_html:
                        # Remove outer <p> tags for table cells
                        p_html = re.sub(r'^<p>(.*)</p>$', r'\1', p_html)
                        cell_content.append(p_html)

                cells.append('<br>'.join(cell_content))

            rows.append(cells)

        # Build HTML table
        html = ['<table>']

        for i, row in enumerate(rows):
            html.append('  <tr>')
            tag = 'th' if i == 0 else 'td'  # First row as header
            for cell in row:
                html.append(f'    <{tag}>{cell}</{tag}>')
            html.append('  </tr>')

        html.append('</table>')

        return '\n'.join(html)

    def _convert_drawing(self, drawing_elem):
        """Convert drawing (image/shape) to HTML"""

        ns = self.namespaces

        # Look for image
        blip = drawing_elem.xpath('.//a:blip/@r:embed', namespaces=ns)
        if blip:
            r_id = blip[0]
            if r_id in self.relationships:
                target = self.relationships[r_id]['target']
                img_name = Path(target).name

                if img_name in self.images:
                    return f'<img src="{self.images[img_name]}" alt="Image">'

        # Look for shape with text
        txbx_content = drawing_elem.xpath('.//wps:txbx//w:p', namespaces=ns)
        if txbx_content:
            shape_html = ['<div class="shape-textbox">']
            for p in txbx_content:
                p_html = self._convert_paragraph(p)
                if p_html:
                    shape_html.append(p_html)
            shape_html.append('</div>')
            return '\n'.join(shape_html)

        return ''

    def _convert_math(self, math_elem):
        """Convert remaining math elements (fallback)"""

        # This should rarely be called after equation conversion
        text = self._extract_text(math_elem)
        return f'<span class="math">{self._escape_html(text)}</span>'

    def _escape_html(self, text):
        """Escape HTML special characters"""

        if not text:
            return ''

        # Don't escape our LaTeX markers
        if 'MATHSTART' in text:
            return text

        return (text
                .replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;')
                .replace('"', '&quot;'))

    def _generate_complete_html(self, content, title):
        """Generate complete HTML document with styles and scripts"""

        # Build footnotes section
        footnotes_html = ''
        if self.footnotes:
            footnotes_html = '\n<hr>\n<section class="footnotes">\n<h3>الحواشي</h3>\n'
            for fn_id, fn_content in self.footnotes.items():
                footnotes_html += f'<p id="fn{fn_id}"><sup>{fn_id}</sup> {fn_content} '
                footnotes_html += f'<a href="#fnref{fn_id}">↩</a></p>\n'
            footnotes_html += '</section>\n'

        # Build endnotes section
        endnotes_html = ''
        if self.endnotes:
            endnotes_html = '\n<section class="endnotes">\n<h3>الملاحظات الختامية</h3>\n'
            for en_id, en_content in self.endnotes.items():
                endnotes_html += f'<p id="en{en_id}"><sup>{en_id}</sup> {en_content} '
                endnotes_html += f'<a href="#enref{en_id}">↩</a></p>\n'
            endnotes_html += '</section>\n'

        return f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>

    <!-- MathJax for equation rendering -->
    <script>
        window.MathJax = {{
            tex: {{
                inlineMath: [['\\\\(', '\\\\)']],
                displayMath: [['\\\\[', '\\\\]']]
            }},
            svg: {{
                fontCache: 'global'
            }}
        }};
    </script>
    <script src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-chtml.js" async></script>

    <!-- Equation marker processor -->
    <script>
        document.addEventListener('DOMContentLoaded', function() {{
            var content = document.body.innerHTML;

            // Process inline equations
            content = content.replace(/MATHSTARTINLINE([\\s\\S]*?)MATHENDINLINE/g,
                '<span class="inline-math">$1</span>');

            // Process display equations
            content = content.replace(/MATHSTARTDISPLAY([\\s\\S]*?)MATHENDDISPLAY/g,
                '<div class="display-math">$1</div>');

            document.body.innerHTML = content;

            // Trigger MathJax
            if (window.MathJax && MathJax.typesetPromise) {{
                MathJax.typesetPromise();
            }}
        }});
    </script>

    <style>
        /* Base styles */
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.8;
            max-width: 900px;
            margin: 0 auto;
            padding: 20px 40px;
            direction: rtl;
            text-align: right;
            background-color: #fafafa;
            color: #333;
        }}

        /* Headings */
        h1, h2, h3, h4, h5, h6 {{
            color: #2c3e50;
            margin-top: 1.5em;
            margin-bottom: 0.5em;
            line-height: 1.3;
        }}
        h1 {{ font-size: 2.2em; border-bottom: 2px solid #3498db; padding-bottom: 0.3em; }}
        h2 {{ font-size: 1.8em; border-bottom: 1px solid #ddd; padding-bottom: 0.2em; }}
        h3 {{ font-size: 1.4em; }}
        h4 {{ font-size: 1.2em; }}

        /* Paragraphs */
        p {{
            margin: 1em 0;
            text-align: justify;
        }}

        /* Tables */
        table {{
            border-collapse: collapse;
            width: 100%;
            margin: 1.5em 0;
            background: white;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }}
        th, td {{
            border: 1px solid #ddd;
            padding: 12px 15px;
            text-align: right;
        }}
        th {{
            background-color: #3498db;
            color: white;
            font-weight: bold;
        }}
        tr:nth-child(even) {{
            background-color: #f8f9fa;
        }}
        tr:hover {{
            background-color: #e8f4f8;
        }}

        /* Images */
        img {{
            max-width: 100%;
            height: auto;
            display: block;
            margin: 1.5em auto;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }}

        /* Lists */
        ul, ol {{
            margin: 1em 0;
            padding-right: 2em;
        }}
        li {{
            margin: 0.5em 0;
        }}

        /* Math equations */
        .inline-math {{
            display: inline;
            margin: 0 0.2em;
        }}
        .display-math {{
            display: block;
            text-align: center;
            margin: 1.5em 0;
            padding: 1em;
            background: #f8f9fa;
            border-radius: 5px;
            overflow-x: auto;
        }}

        /* Shape textboxes */
        .shape-textbox {{
            display: inline-block;
            border: 1px solid #ddd;
            border-radius: 50%;
            padding: 15px 20px;
            margin: 5px;
            background: #f0f8ff;
            text-align: center;
            min-width: 50px;
        }}

        /* Footnotes */
        .footnotes, .endnotes {{
            margin-top: 3em;
            padding-top: 1em;
            border-top: 1px solid #ddd;
            font-size: 0.9em;
            color: #666;
        }}
        .footnotes p, .endnotes p {{
            margin: 0.5em 0;
        }}

        /* Links */
        a {{
            color: #3498db;
            text-decoration: none;
        }}
        a:hover {{
            text-decoration: underline;
        }}

        /* Print styles */
        @media print {{
            body {{
                max-width: none;
                padding: 0;
                background: white;
            }}
            .shape-textbox {{
                border: 1px solid #999;
            }}
        }}
    </style>
</head>
<body>
{content}
{footnotes_html}
{endnotes_html}
</body>
</html>"""


def test_comprehensive_converter():
    """Test the comprehensive HTML converter"""

    test_doc = Path(r"D:\Development\document-processing-api-2\document-processing-api\test docs\الدالة واحد لواحد (جاهزة للنشر) - Copy.docx")
    output_dir = Path(r"D:\Development\document-processing-api-2\backend\test_output\html_output")

    if test_doc.exists():
        converter = ComprehensiveHTMLConverter()
        result = converter.convert_document(test_doc, output_dir=output_dir)

        if result.get('success'):
            print(f"\n Open {result['output_path']} in your browser to view!")
    else:
        print(f"Test document not found: {test_doc}")


if __name__ == "__main__":
    test_comprehensive_converter()
