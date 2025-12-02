"""
COMPREHENSIVE Word to HTML Converter V2
Handles ALL document elements: headers, tables, shapes, images, footnotes, equations, bullets

FIXES in V2:
1. Proper bullet/numbered list handling
2. Better shape/drawing handling (ovals with content)
3. Info box styling
4. Better table cell backgrounds
5. Improved CSS that actually applies
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


class ComprehensiveHTMLConverterV2:
    """Complete Word to HTML converter V2 - handles all document elements"""

    def __init__(self):
        self.namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
            'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
            'v': 'urn:schemas-microsoft-com:vml',
            'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
            'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
            'o': 'urn:schemas-microsoft-com:office:office'
        }
        self.relationships = {}
        self.images = {}
        self.footnotes = {}
        self.endnotes = {}
        self.styles = {}
        self.numbering = {}  # For bullet/numbered lists
        self.current_list_id = None
        self.list_stack = []

    def convert_document(self, input_path, output_path=None, output_dir=None):
        """Convert Word document to HTML"""

        input_path = Path(input_path).absolute()

        if not output_dir:
            output_dir = input_path.parent / f"{input_path.stem}_html_v2"
        else:
            output_dir = Path(output_dir).absolute()

        output_dir.mkdir(exist_ok=True)

        if not output_path:
            output_path = output_dir / f"{input_path.stem}.html"
        else:
            output_path = Path(output_path).absolute()

        print("\n" + "="*70)
        print("COMPREHENSIVE WORD TO HTML CONVERTER V2")
        print("="*70)
        print(f"Input:  {input_path}")
        print(f"Output: {output_path}")
        print("="*70)

        temp_dir = Path(f"temp_html_v2_{datetime.now().strftime('%Y%m%d_%H%M%S')}")

        try:
            # Step 1: First convert equations
            print("\n[1] Converting equations...")
            from enhanced_zip_converter import EnhancedZipConverter
            eq_converter = EnhancedZipConverter()

            eq_converted = temp_dir / "eq_converted.docx"
            temp_dir.mkdir(exist_ok=True)

            eq_result = eq_converter.process_document(input_path, eq_converted)

            if not eq_result.get('success'):
                print(f"    Warning: Equation conversion had issues, using original")
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

            # Step 5: Load numbering (bullets/lists)
            print("\n[5] Loading numbering definitions...")
            self._load_numbering(extract_dir)

            # Step 6: Load footnotes and endnotes
            print("\n[6] Loading footnotes/endnotes...")
            self._load_footnotes(extract_dir)

            # Step 7: Extract images
            print("\n[7] Extracting images...")
            self._extract_images(extract_dir, output_dir)

            # Step 8: Parse and convert document
            print("\n[8] Converting document to HTML...")
            doc_xml_path = extract_dir / "word" / "document.xml"
            with open(doc_xml_path, 'rb') as f:
                doc_root = etree.fromstring(f.read())

            html_content = self._convert_body(doc_root)

            # Step 9: Generate complete HTML
            print("\n[9] Generating final HTML...")
            complete_html = self._generate_complete_html(html_content, input_path.stem)

            # Save HTML
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(complete_html)

            print("\n" + "="*70)
            print("CONVERSION COMPLETE!")
            print("="*70)
            print(f"Output HTML: {output_path}")
            print(f"Images:      {len(self.images)}")
            print(f"Footnotes:   {len(self.footnotes)}")

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
            return {'success': False, 'error': str(e)}

        finally:
            if temp_dir.exists():
                shutil.rmtree(temp_dir)

    def _load_relationships(self, extract_dir):
        """Load document relationships"""
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
            self.relationships[rel_id] = {'target': target, 'type': rel_type}

        print(f"    Loaded {len(self.relationships)} relationships")

    def _load_styles(self, extract_dir):
        """Load document styles"""
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

    def _load_numbering(self, extract_dir):
        """Load numbering definitions for bullets and lists"""
        num_path = extract_dir / "word" / "numbering.xml"
        if not num_path.exists():
            return

        with open(num_path, 'rb') as f:
            root = etree.fromstring(f.read())

        ns = {'w': self.namespaces['w']}

        # Load abstract numbering definitions
        abstract_nums = {}
        for abstract in root.xpath('//w:abstractNum', namespaces=ns):
            abstract_id = abstract.get(f'{{{ns["w"]}}}abstractNumId')
            levels = {}
            for lvl in abstract.xpath('.//w:lvl', namespaces=ns):
                lvl_id = lvl.get(f'{{{ns["w"]}}}ilvl')
                num_fmt = lvl.xpath('.//w:numFmt/@w:val', namespaces=ns)
                levels[lvl_id] = num_fmt[0] if num_fmt else 'bullet'
            abstract_nums[abstract_id] = levels

        # Load numbering instances
        for num in root.xpath('//w:num', namespaces=ns):
            num_id = num.get(f'{{{ns["w"]}}}numId')
            abstract_ref = num.xpath('.//w:abstractNumId/@w:val', namespaces=ns)
            if abstract_ref and abstract_ref[0] in abstract_nums:
                self.numbering[num_id] = abstract_nums[abstract_ref[0]]

        print(f"    Loaded {len(self.numbering)} numbering definitions")

    def _load_footnotes(self, extract_dir):
        """Load footnotes and endnotes"""
        fn_path = extract_dir / "word" / "footnotes.xml"
        if fn_path.exists():
            with open(fn_path, 'rb') as f:
                root = etree.fromstring(f.read())
            ns = {'w': self.namespaces['w']}
            for fn in root.xpath('//w:footnote', namespaces=ns):
                fn_id = fn.get(f'{{{ns["w"]}}}id')
                if fn_id and fn_id not in ['0', '-1']:
                    content = self._extract_text(fn)
                    self.footnotes[fn_id] = content

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
            dest = images_dir / img_file.name
            shutil.copy2(img_file, dest)
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
        current_list = None
        list_items = []

        for child in body:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

            if tag == 'p':
                # Check if this is a list item
                num_pr = child.xpath('.//w:numPr', namespaces=ns)
                if num_pr:
                    # This is a list item
                    num_id = child.xpath('.//w:numId/@w:val', namespaces=ns)
                    ilvl = child.xpath('.//w:ilvl/@w:val', namespaces=ns)

                    list_type = 'ul'  # default to bullet
                    if num_id and num_id[0] in self.numbering:
                        lvl = ilvl[0] if ilvl else '0'
                        fmt = self.numbering[num_id[0]].get(lvl, 'bullet')
                        if fmt in ['decimal', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman']:
                            list_type = 'ol'

                    if current_list != list_type:
                        # Close previous list if any
                        if list_items:
                            html_parts.append(self._wrap_list(list_items, current_list or 'ul'))
                            list_items = []
                        current_list = list_type

                    item_content = self._convert_paragraph_content(child)
                    if item_content.strip():
                        list_items.append(item_content)
                else:
                    # Not a list item - close any open list
                    if list_items:
                        html_parts.append(self._wrap_list(list_items, current_list or 'ul'))
                        list_items = []
                        current_list = None

                    html_parts.append(self._convert_paragraph(child))

            elif tag == 'tbl':
                # Close any open list
                if list_items:
                    html_parts.append(self._wrap_list(list_items, current_list or 'ul'))
                    list_items = []
                    current_list = None

                html_parts.append(self._convert_table(child))

            elif tag == 'sectPr':
                continue

        # Close any remaining list
        if list_items:
            html_parts.append(self._wrap_list(list_items, current_list or 'ul'))

        return '\n'.join(filter(None, html_parts))

    def _wrap_list(self, items, list_type):
        """Wrap list items in appropriate list tag"""
        if not items:
            return ''
        items_html = '\n'.join(f'  <li>{item}</li>' for item in items)
        return f'<{list_type}>\n{items_html}\n</{list_type}>'

    def _convert_paragraph_content(self, p_elem):
        """Convert paragraph content without wrapper tag (for list items)"""
        ns = self.namespaces
        content_parts = []

        for child in p_elem:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

            if tag == 'r':
                content_parts.append(self._convert_run(child))
            elif tag == 'hyperlink':
                content_parts.append(self._convert_hyperlink(child))
            elif tag == 'oMath' or tag == 'oMathPara':
                content_parts.append(self._convert_math(child))
            elif tag == 'drawing':
                content_parts.append(self._convert_drawing(child))
            elif tag == 'pPr' or tag == 'bookmarkStart' or tag == 'bookmarkEnd':
                continue
            else:
                text = self._extract_text(child)
                if text:
                    content_parts.append(self._escape_html(text))

        return ''.join(content_parts)

    def _convert_paragraph(self, p_elem):
        """Convert paragraph to HTML"""
        ns = self.namespaces

        # Check for style
        style_id = p_elem.xpath('.//w:pStyle/@w:val', namespaces=ns)
        style_name = self.styles.get(style_id[0], '') if style_id else ''

        # Check for background color (info boxes)
        shd = p_elem.xpath('.//w:shd/@w:fill', namespaces=ns)
        bg_color = shd[0] if shd and shd[0] != 'auto' else None

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

        content = self._convert_paragraph_content(p_elem)

        if not content.strip():
            return ''

        # Build attributes
        style_attr = ''
        if bg_color:
            style_attr = f' style="background-color: #{bg_color}; padding: 10px; border-radius: 5px;"'

        # Wrap in appropriate tag
        if heading_level > 0:
            return f'<h{heading_level}{style_attr}>{content}</h{heading_level}>'
        else:
            return f'<p{style_attr}>{content}</p>'

    def _convert_run(self, r_elem):
        """Convert text run to HTML"""
        ns = self.namespaces
        parts = []

        # Get run properties
        bold = bool(r_elem.xpath('.//w:b[not(@w:val="false") and not(@w:val="0")]', namespaces=ns)) or \
               bool(r_elem.xpath('.//w:b[not(@w:val)]', namespaces=ns))
        italic = bool(r_elem.xpath('.//w:i[not(@w:val="false") and not(@w:val="0")]', namespaces=ns)) or \
                 bool(r_elem.xpath('.//w:i[not(@w:val)]', namespaces=ns))
        underline = bool(r_elem.xpath('.//w:u[@w:val]', namespaces=ns))
        strike = bool(r_elem.xpath('.//w:strike', namespaces=ns))
        superscript = bool(r_elem.xpath('.//w:vertAlign[@w:val="superscript"]', namespaces=ns))
        subscript = bool(r_elem.xpath('.//w:vertAlign[@w:val="subscript"]', namespaces=ns))

        # Get color
        color = r_elem.xpath('.//w:color/@w:val', namespaces=ns)
        color_style = f'color: #{color[0]};' if color and color[0] != 'auto' else ''

        for child in r_elem:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

            if tag == 't':
                text = child.text or ''
                parts.append(self._escape_html(text))
            elif tag == 'drawing':
                parts.append(self._convert_drawing(child))
            elif tag == 'pict':
                parts.append(self._convert_vml_pict(child))
            elif tag == 'object':
                parts.append(self._convert_object(child))
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
            elif tag == 'sym':
                # Symbol character
                char_code = child.get(f'{{{ns["w"]}}}char')
                if char_code:
                    try:
                        parts.append(chr(int(char_code, 16)))
                    except:
                        parts.append('?')

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
        if color_style:
            content = f'<span style="{color_style}">{content}</span>'

        return content

    def _convert_hyperlink(self, h_elem):
        """Convert hyperlink to HTML"""
        ns = self.namespaces
        r_id = h_elem.get(f'{{{ns["r"]}}}id')
        href = "#"

        if r_id and r_id in self.relationships:
            href = self.relationships[r_id]['target']

        anchor = h_elem.get(f'{{{ns["w"]}}}anchor')
        if anchor:
            href = f'#{anchor}'

        content_parts = []
        for r in h_elem.xpath('.//w:r', namespaces=ns):
            content_parts.append(self._convert_run(r))

        content = ''.join(content_parts)
        return f'<a href="{href}">{content}</a>'

    def _convert_table(self, tbl_elem):
        """Convert table to HTML with better styling"""
        ns = self.namespaces
        rows = []

        # Check for table style/background
        tbl_shd = tbl_elem.xpath('.//w:tblPr/w:shd/@w:fill', namespaces=ns)
        table_bg = tbl_shd[0] if tbl_shd and tbl_shd[0] != 'auto' else None

        for tr in tbl_elem.xpath('./w:tr', namespaces=ns):
            cells = []
            for tc in tr.xpath('./w:tc', namespaces=ns):
                # Get cell properties
                cell_shd = tc.xpath('.//w:shd/@w:fill', namespaces=ns)
                cell_bg = cell_shd[0] if cell_shd and cell_shd[0] != 'auto' else None

                # Get cell content
                cell_content = []
                for p in tc.xpath('./w:p', namespaces=ns):
                    p_html = self._convert_paragraph_content(p)
                    if p_html.strip():
                        cell_content.append(p_html)

                # Check for nested tables
                for nested_tbl in tc.xpath('./w:tbl', namespaces=ns):
                    cell_content.append(self._convert_table(nested_tbl))

                cells.append({
                    'content': '<br>'.join(cell_content) if cell_content else '',
                    'bg': cell_bg
                })

            rows.append(cells)

        # Build HTML table
        table_style = f' style="background-color: #{table_bg};"' if table_bg else ''
        html = [f'<table{table_style}>']

        for i, row in enumerate(rows):
            html.append('  <tr>')
            tag = 'th' if i == 0 else 'td'
            for cell in row:
                cell_style = f' style="background-color: #{cell["bg"]};"' if cell.get('bg') else ''
                html.append(f'    <{tag}{cell_style}>{cell["content"]}</{tag}>')
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
                    return f'<img src="{self.images[img_name]}" alt="Image" class="doc-image">'

        # Look for shape with text (textbox/oval/etc)
        shape_content = []

        # Check for wps:txbx (Word 2010+ shapes)
        txbx_content = drawing_elem.xpath('.//wps:txbx//w:p', namespaces=ns)
        if txbx_content:
            for p in txbx_content:
                p_html = self._convert_paragraph_content(p)
                if p_html.strip():
                    shape_content.append(p_html)

        # Check shape type for styling
        shape_type = 'rectangle'
        sp_auto = drawing_elem.xpath('.//wps:spPr//a:prstGeom/@prst', namespaces=ns)
        if sp_auto:
            shape_type = sp_auto[0]

        if shape_content:
            shape_class = 'shape-oval' if 'ellipse' in shape_type or 'oval' in shape_type else 'shape-box'
            return f'<div class="{shape_class}">{" ".join(shape_content)}</div>'

        return ''

    def _convert_vml_pict(self, pict_elem):
        """Convert VML pict elements (legacy shapes)"""
        ns = self.namespaces

        # Look for shapes with text
        shape_content = []

        # VML textbox content
        txbx = pict_elem.xpath('.//v:textbox//w:p', namespaces=ns)
        if txbx:
            for p in txbx:
                p_html = self._convert_paragraph_content(p)
                if p_html.strip():
                    shape_content.append(p_html)

        # Check shape type
        shape_type = 'box'
        oval = pict_elem.xpath('.//v:oval', namespaces=ns)
        if oval:
            shape_type = 'oval'

        if shape_content:
            shape_class = 'shape-oval' if shape_type == 'oval' else 'shape-box'
            return f'<span class="{shape_class}">{" ".join(shape_content)}</span>'

        # Check for image in VML
        imagedata = pict_elem.xpath('.//v:imagedata/@r:id', namespaces=ns)
        if imagedata:
            r_id = imagedata[0]
            if r_id in self.relationships:
                target = self.relationships[r_id]['target']
                img_name = Path(target).name
                if img_name in self.images:
                    return f'<img src="{self.images[img_name]}" alt="Image" class="doc-image">'

        return ''

    def _convert_object(self, obj_elem):
        """Convert embedded objects"""
        # Usually these are OLE objects, try to extract any image
        ns = self.namespaces
        imagedata = obj_elem.xpath('.//v:imagedata/@r:id', namespaces=ns)
        if imagedata:
            r_id = imagedata[0]
            if r_id in self.relationships:
                target = self.relationships[r_id]['target']
                img_name = Path(target).name
                if img_name in self.images:
                    return f'<img src="{self.images[img_name]}" alt="Object" class="doc-image">'
        return ''

    def _convert_math(self, math_elem):
        """Convert remaining math elements"""
        text = self._extract_text(math_elem)
        return f'<span class="math">{self._escape_html(text)}</span>'

    def _escape_html(self, text):
        """Escape HTML special characters"""
        if not text:
            return ''
        if 'MATHSTART' in text:
            return text
        return (text
                .replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;')
                .replace('"', '&quot;'))

    def _generate_complete_html(self, content, title):
        """Generate complete HTML document"""

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
            endnotes_html = '\n<section class="endnotes">\n<h3>الملاحظات</h3>\n'
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
            svg: {{ fontCache: 'global' }}
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
        /* ========== BASE STYLES ========== */
        * {{
            box-sizing: border-box;
        }}

        body {{
            font-family: 'Segoe UI', 'Arial', sans-serif;
            font-size: 16px;
            line-height: 1.8;
            max-width: 900px;
            margin: 0 auto;
            padding: 30px 40px;
            direction: rtl;
            text-align: right;
            background-color: #ffffff;
            color: #333333;
        }}

        /* ========== HEADINGS ========== */
        h1, h2, h3, h4, h5, h6 {{
            color: #1a365d;
            margin-top: 1.5em;
            margin-bottom: 0.5em;
            line-height: 1.3;
            font-weight: bold;
        }}

        h1 {{
            font-size: 2em;
            border-bottom: 3px solid #3182ce;
            padding-bottom: 0.3em;
            text-align: center;
        }}

        h2 {{
            font-size: 1.6em;
            border-bottom: 2px solid #63b3ed;
            padding-bottom: 0.2em;
            color: #2c5282;
        }}

        h3 {{ font-size: 1.3em; color: #2b6cb0; }}
        h4 {{ font-size: 1.1em; }}

        /* ========== PARAGRAPHS ========== */
        p {{
            margin: 1em 0;
            text-align: justify;
        }}

        /* ========== LISTS ========== */
        ul, ol {{
            margin: 1em 0;
            padding-right: 2em;
            padding-left: 0;
        }}

        li {{
            margin: 0.5em 0;
            line-height: 1.6;
        }}

        ul {{
            list-style-type: disc;
        }}

        ul ul {{
            list-style-type: circle;
        }}

        ol {{
            list-style-type: decimal;
        }}

        /* ========== TABLES ========== */
        table {{
            border-collapse: collapse;
            width: 100%;
            margin: 1.5em 0;
            background: white;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            border-radius: 8px;
            overflow: hidden;
        }}

        th, td {{
            border: 1px solid #e2e8f0;
            padding: 12px 16px;
            text-align: right;
            vertical-align: top;
        }}

        th {{
            background-color: #3182ce;
            color: white;
            font-weight: bold;
        }}

        tr:nth-child(even) {{
            background-color: #f7fafc;
        }}

        tr:hover {{
            background-color: #edf2f7;
        }}

        /* ========== IMAGES ========== */
        img, .doc-image {{
            max-width: 100%;
            height: auto;
            display: block;
            margin: 1.5em auto;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }}

        /* ========== SHAPES (OVALS, BOXES) ========== */
        .shape-oval {{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            min-width: 40px;
            min-height: 40px;
            padding: 8px 15px;
            margin: 3px;
            border: 2px solid #4a5568;
            border-radius: 50%;
            background: #f7fafc;
            text-align: center;
            font-weight: 500;
        }}

        .shape-box {{
            display: inline-block;
            padding: 10px 15px;
            margin: 5px;
            border: 1px solid #cbd5e0;
            border-radius: 5px;
            background: #f7fafc;
        }}

        /* ========== MATH EQUATIONS ========== */
        .inline-math {{
            display: inline;
            margin: 0 0.2em;
        }}

        .display-math {{
            display: block;
            text-align: center;
            margin: 1.5em 0;
            padding: 1em;
            background: #f8fafc;
            border-radius: 8px;
            border-right: 4px solid #3182ce;
            overflow-x: auto;
        }}

        .math {{
            font-style: italic;
        }}

        /* ========== INFO BOXES ========== */
        .info-box, [style*="background-color"] {{
            padding: 15px 20px;
            border-radius: 8px;
            margin: 1em 0;
        }}

        /* ========== LINKS ========== */
        a {{
            color: #3182ce;
            text-decoration: none;
        }}

        a:hover {{
            text-decoration: underline;
            color: #2c5282;
        }}

        /* ========== FOOTNOTES ========== */
        .footnotes, .endnotes {{
            margin-top: 3em;
            padding-top: 1em;
            border-top: 2px solid #e2e8f0;
            font-size: 0.9em;
            color: #4a5568;
        }}

        .footnotes h3, .endnotes h3 {{
            font-size: 1.2em;
            margin-bottom: 1em;
        }}

        .footnotes p, .endnotes p {{
            margin: 0.5em 0;
        }}

        /* ========== PRINT STYLES ========== */
        @media print {{
            body {{
                max-width: none;
                padding: 0;
                background: white;
            }}

            table {{
                box-shadow: none;
            }}

            .shape-oval, .shape-box {{
                border: 1px solid #333;
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


def test_converter_v2():
    """Test the V2 converter"""
    test_doc = Path(r"D:\Development\document-processing-api-2\document-processing-api\test docs\الدالة واحد لواحد (جاهزة للنشر) - Copy.docx")
    output_dir = Path(r"D:\Development\document-processing-api-2\backend\test_output\html_v2_output")

    if test_doc.exists():
        converter = ComprehensiveHTMLConverterV2()
        result = converter.convert_document(test_doc, output_dir=output_dir)

        if result.get('success'):
            print(f"\n Open {result['output_path']} in your browser!")
    else:
        print(f"Test document not found: {test_doc}")


if __name__ == "__main__":
    test_converter_v2()
