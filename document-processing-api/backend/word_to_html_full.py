"""
FULL Word to HTML Converter with Configuration Options
=======================================================

Features:
1. SVG conversion for shapes (ovals, rectangles, lines, groups)
2. Configurable equation prefix/suffix
3. User control panel in generated HTML
4. All standard Word elements (headers, tables, lists, images, footnotes)

Configuration Options:
- convert_shapes_to_svg: True/False
- equation_prefix: '' | 'current' | custom string
- equation_suffix: '' | 'current' | custom string
- include_styles: True/False
- include_mathjax: True/False
"""

import sys
import io
import zipfile
import shutil
import base64
import re
import json
from pathlib import Path
from lxml import etree
from datetime import datetime
from dataclasses import dataclass, field
from typing import Optional, List, Dict

# Set UTF-8 encoding
if not isinstance(sys.stdout, io.TextIOWrapper) or sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    except:
        pass


@dataclass
class ConversionConfig:
    """Configuration options for Word to HTML conversion

    These settings are applied DURING conversion (from upload page),
    NOT embedded in the output HTML.
    """
    # Shape conversion - default False to ignore Word shapes and equations inside them
    convert_shapes_to_svg: bool = False

    # Image settings
    include_images: bool = True  # Include images in output (always extracted to subfolder)

    # Equation settings - prefix/suffix applied during conversion
    inline_prefix: str = 'MATHSTARTINLINE'  # Prefix for inline equations
    inline_suffix: str = 'MATHENDINLINE'    # Suffix for inline equations
    display_prefix: str = 'MATHSTARTDISPLAY'  # Prefix for display equations
    display_suffix: str = 'MATHENDDISPLAY'    # Suffix for display equations

    # Output settings
    include_styles: bool = True
    include_mathjax: bool = True

    # Styling
    rtl_direction: bool = True
    font_family: str = "'Segoe UI', Arial, sans-serif"
    max_width: str = "900px"

    # Output format: "mathml_html" (MathML, no JS) or "latex_html" (LaTeX + MathJax)
    output_format: str = "mathml_html"


class ShapeToSVGConverter:
    """Converts Word shapes to SVG"""

    def __init__(self):
        self.ns = {
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
            'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'v': 'urn:schemas-microsoft-com:vml',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
        }

    def convert_drawing_to_svg(self, drawing_elem, equations_map=None):
        """Convert a drawing element to SVG"""

        # Get dimensions
        extent = drawing_elem.xpath('.//wp:extent', namespaces=self.ns)
        if extent:
            # EMUs to pixels (914400 EMUs = 1 inch = 96 pixels)
            cx = int(extent[0].get('cx', '914400')) / 914400 * 96
            cy = int(extent[0].get('cy', '914400')) / 914400 * 96
        else:
            cx, cy = 200, 150  # Default size

        svg_parts = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{cx}" height="{cy}" viewBox="0 0 {cx} {cy}">']

        # Check for group of shapes
        grp_sp = drawing_elem.xpath('.//wpg:wgp', namespaces=self.ns)
        if grp_sp:
            svg_parts.append(self._convert_group(grp_sp[0], cx, cy))
        else:
            # Single shape
            wsp = drawing_elem.xpath('.//wps:wsp', namespaces=self.ns)
            if wsp:
                svg_parts.append(self._convert_shape(wsp[0], 0, 0, cx, cy))

        svg_parts.append('</svg>')
        return '\n'.join(svg_parts)

    def _convert_group(self, grp_elem, width, height):
        """Convert a group of shapes"""
        parts = []

        # Get all shapes in group
        shapes = grp_elem.xpath('.//wps:wsp', namespaces=self.ns)

        for i, shape in enumerate(shapes):
            # Get shape position within group
            off = shape.xpath('.//a:off', namespaces=self.ns)
            ext = shape.xpath('.//a:ext', namespaces=self.ns)

            x = int(off[0].get('x', '0')) / 914400 * 96 if off else i * 50
            y = int(off[0].get('y', '0')) / 914400 * 96 if off else 0
            w = int(ext[0].get('cx', '914400')) / 914400 * 96 if ext else 50
            h = int(ext[0].get('cy', '914400')) / 914400 * 96 if ext else 50

            parts.append(self._convert_shape(shape, x, y, w, h))

        # Get connectors/lines
        cxn_sps = grp_elem.xpath('.//wps:cxnSp', namespaces=self.ns)
        for cxn in cxn_sps:
            parts.append(self._convert_connector(cxn))

        return '\n'.join(parts)

    def _convert_shape(self, shape_elem, x, y, width, height):
        """Convert a single shape to SVG element"""

        # Get shape type
        prst_geom = shape_elem.xpath('.//a:prstGeom/@prst', namespaces=self.ns)
        shape_type = prst_geom[0] if prst_geom else 'rect'

        # Get fill color
        solid_fill = shape_elem.xpath('.//a:solidFill/a:srgbClr/@val', namespaces=self.ns)
        fill_color = f'#{solid_fill[0]}' if solid_fill else '#f0f0f0'

        # Get outline color
        ln_fill = shape_elem.xpath('.//a:ln/a:solidFill/a:srgbClr/@val', namespaces=self.ns)
        stroke_color = f'#{ln_fill[0]}' if ln_fill else '#333333'

        # Get text content
        text_content = self._extract_shape_text(shape_elem)

        # Check if text contains LaTeX equations
        has_latex = text_content and ('\\(' in text_content or '\\[' in text_content)

        svg = []

        if shape_type in ['ellipse', 'oval']:
            cx = x + width / 2
            cy = y + height / 2
            rx = width / 2
            ry = height / 2
            svg.append(f'<ellipse cx="{cx}" cy="{cy}" rx="{rx}" ry="{ry}" fill="{fill_color}" stroke="{stroke_color}" stroke-width="2"/>')
            if text_content:
                if has_latex:
                    # Use foreignObject for MathJax rendering
                    svg.append(f'<foreignObject x="{x}" y="{y}" width="{width}" height="{height}">')
                    svg.append(f'<div xmlns="http://www.w3.org/1999/xhtml" style="width:100%;height:100%;display:flex;align-items:center;justify-content:center;text-align:center;font-size:12px;">{text_content}</div>')
                    svg.append('</foreignObject>')
                else:
                    svg.append(f'<text x="{cx}" y="{cy}" text-anchor="middle" dominant-baseline="middle" font-size="14" fill="#333">{text_content}</text>')

        elif shape_type in ['rect', 'rectangle', 'roundRect']:
            r = 5 if shape_type == 'roundRect' else 0
            svg.append(f'<rect x="{x}" y="{y}" width="{width}" height="{height}" rx="{r}" fill="{fill_color}" stroke="{stroke_color}" stroke-width="2"/>')
            if text_content:
                tx = x + width / 2
                ty = y + height / 2
                if has_latex:
                    svg.append(f'<foreignObject x="{x}" y="{y}" width="{width}" height="{height}">')
                    svg.append(f'<div xmlns="http://www.w3.org/1999/xhtml" style="width:100%;height:100%;display:flex;align-items:center;justify-content:center;text-align:center;font-size:12px;">{text_content}</div>')
                    svg.append('</foreignObject>')
                else:
                    svg.append(f'<text x="{tx}" y="{ty}" text-anchor="middle" dominant-baseline="middle" font-size="14" fill="#333">{text_content}</text>')

        elif shape_type in ['line', 'straightConnector1']:
            svg.append(f'<line x1="{x}" y1="{y}" x2="{x + width}" y2="{y + height}" stroke="{stroke_color}" stroke-width="2"/>')

        else:
            # Default to rectangle
            svg.append(f'<rect x="{x}" y="{y}" width="{width}" height="{height}" fill="{fill_color}" stroke="{stroke_color}" stroke-width="2"/>')
            if text_content:
                tx = x + width / 2
                ty = y + height / 2
                if has_latex:
                    svg.append(f'<foreignObject x="{x}" y="{y}" width="{width}" height="{height}">')
                    svg.append(f'<div xmlns="http://www.w3.org/1999/xhtml" style="width:100%;height:100%;display:flex;align-items:center;justify-content:center;text-align:center;font-size:12px;">{text_content}</div>')
                    svg.append('</foreignObject>')
                else:
                    svg.append(f'<text x="{tx}" y="{ty}" text-anchor="middle" dominant-baseline="middle" font-size="14" fill="#333">{text_content}</text>')

        return '\n'.join(svg)

    def _convert_connector(self, cxn_elem):
        """Convert connector/line to SVG"""
        # Get start and end points
        stCxn = cxn_elem.xpath('.//a:stCxn', namespaces=self.ns)
        endCxn = cxn_elem.xpath('.//a:endCxn', namespaces=self.ns)

        # Simplified - just draw a line
        return '<line x1="50" y1="50" x2="150" y2="50" stroke="#333" stroke-width="2" marker-end="url(#arrow)"/>'

    def _extract_shape_text(self, shape_elem):
        """Extract text content from shape, including LaTeX equations"""
        texts = []
        for t in shape_elem.xpath('.//w:t/text() | .//a:t/text()', namespaces=self.ns):
            if t and t.strip():
                texts.append(t.strip())

        # Join with spaces, preserving LaTeX markers
        result = ' '.join(texts) if texts else ''

        # Escape HTML entities but preserve LaTeX
        if result and '\\' not in result:
            result = result.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

        return result

    def convert_vml_to_svg(self, pict_elem):
        """Convert VML pict element to SVG"""
        ns = self.ns

        # Get oval
        ovals = pict_elem.xpath('.//v:oval', namespaces=ns)
        rects = pict_elem.xpath('.//v:rect', namespaces=ns)
        lines = pict_elem.xpath('.//v:line', namespaces=ns)

        # Default dimensions
        width, height = 100, 100

        svg_parts = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">']

        for oval in ovals:
            style = oval.get('style', '')
            # Parse style for dimensions
            w = self._parse_vml_dimension(style, 'width', 50)
            h = self._parse_vml_dimension(style, 'height', 50)

            fill = oval.get('fillcolor', '#f0f0f0')
            stroke = oval.get('strokecolor', '#333')

            cx, cy = w/2, h/2
            rx, ry = w/2, h/2

            svg_parts.append(f'<ellipse cx="{cx}" cy="{cy}" rx="{rx}" ry="{ry}" fill="{fill}" stroke="{stroke}" stroke-width="2"/>')

            # Get text
            text = self._extract_vml_text(oval)
            if text:
                svg_parts.append(f'<text x="{cx}" y="{cy}" text-anchor="middle" dominant-baseline="middle" font-size="14">{text}</text>')

        for rect in rects:
            style = rect.get('style', '')
            w = self._parse_vml_dimension(style, 'width', 50)
            h = self._parse_vml_dimension(style, 'height', 50)

            fill = rect.get('fillcolor', '#f0f0f0')
            stroke = rect.get('strokecolor', '#333')

            svg_parts.append(f'<rect x="0" y="0" width="{w}" height="{h}" fill="{fill}" stroke="{stroke}" stroke-width="2"/>')

        svg_parts.append('</svg>')
        return '\n'.join(svg_parts)

    def _parse_vml_dimension(self, style, prop, default):
        """Parse dimension from VML style string"""
        match = re.search(rf'{prop}:(\d+(?:\.\d+)?)(pt|px|in)?', style)
        if match:
            val = float(match.group(1))
            unit = match.group(2) or 'pt'
            if unit == 'pt':
                return val * 1.333  # pt to px
            elif unit == 'in':
                return val * 96
            return val
        return default

    def _extract_vml_text(self, elem):
        """Extract text from VML element"""
        texts = []
        for t in elem.xpath('.//w:t/text()', namespaces=self.ns):
            if t.strip():
                texts.append(t.strip())
        return ' '.join(texts) if texts else ''


class FullWordToHTMLConverter:
    """Full-featured Word to HTML converter with configuration"""

    def __init__(self, config: ConversionConfig = None):
        self.config = config or ConversionConfig()
        self.svg_converter = ShapeToSVGConverter()

        # Strategy: select equation converter based on output_format
        if self.config.output_format == "mathml_html":
            from doc_processor.omml_to_mathml import OmmlToMathMLConverter
            self.equation_converter = OmmlToMathMLConverter()
        else:
            self.equation_converter = None  # LaTeX mode uses pre-processed DOCX

        self.namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
            'v': 'urn:schemas-microsoft-com:vml',
            'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
            'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
            'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture'
        }

        self.relationships = {}
        self.images = {}
        self.footnotes = {}
        self.styles = {}
        self.numbering = {}
        self.output_dir = None  # Set during conversion
        self.svg_counter = 0  # Counter for generated SVG files

    def convert(self, input_path, output_path=None, output_dir=None):
        """Main conversion method - branches based on output_format"""

        input_path = Path(input_path).absolute()

        if not output_dir:
            output_dir = input_path.parent / f"{input_path.stem}_html_full"
        output_dir = Path(output_dir)
        output_dir.mkdir(exist_ok=True)

        if not output_path:
            output_path = output_dir / f"{input_path.stem}.html"
        output_path = Path(output_path)

        print(f"\n{'='*70}")
        print("FULL WORD TO HTML CONVERTER")
        print(f"{'='*70}")
        print(f"Input:  {input_path}")
        print(f"Output: {output_path}")
        print(f"Mode:   {self.config.output_format}")
        print(f"Config: SVG={self.config.convert_shapes_to_svg}, MathJax={self.config.include_mathjax}")
        print(f"{'='*70}")

        # Debug: Show exact output_format value
        print(f"DEBUG: output_format = '{self.config.output_format}' (type: {type(self.config.output_format).__name__})")
        print(f"DEBUG: output_format == 'mathml_html': {self.config.output_format == 'mathml_html'}")
        print(f"DEBUG: output_format == 'latex_html': {self.config.output_format == 'latex_html'}")

        if self.config.output_format == "mathml_html":
            print("DEBUG: Using MathML mode")
            return self._convert_mathml_mode(input_path, output_path, output_dir)
        else:
            print("DEBUG: Using LaTeX mode")
            return self._convert_latex_mode(input_path, output_path, output_dir)

    def _convert_latex_mode(self, input_path, output_path, output_dir):
        """Existing two-step conversion: equation pre-processing + HTML generation"""

        temp_dir = Path(f"temp_full_{datetime.now().strftime('%Y%m%d_%H%M%S')}")

        try:
            # Step 1: Convert equations first
            print("\n[1] Converting equations...")
            print(f"    Markers: inline={self.config.inline_prefix}/{self.config.inline_suffix}, display={self.config.display_prefix}/{self.config.display_suffix}")
            from enhanced_zip_converter import EnhancedZipConverter
            eq_converter = EnhancedZipConverter(
                inline_prefix=self.config.inline_prefix,
                inline_suffix=self.config.inline_suffix,
                display_prefix=self.config.display_prefix,
                display_suffix=self.config.display_suffix
            )

            eq_converted = temp_dir / "eq_converted.docx"
            temp_dir.mkdir(exist_ok=True)

            eq_result = eq_converter.process_document(input_path, eq_converted)

            # Save converted Word document to output folder (user requested this)
            word_output_path = output_dir / f"{input_path.stem}_equations.docx"
            if eq_result.get('success') and eq_converted.exists():
                shutil.copy2(eq_converted, word_output_path)
                print(f"    Word document with equations saved: {word_output_path}")
            else:
                # Copy original if conversion failed
                shutil.copy2(input_path, word_output_path)
                eq_converted = input_path

            # Step 2: Extract document
            print("\n[2] Extracting document...")
            extract_dir = temp_dir / "extracted"
            with zipfile.ZipFile(eq_converted, 'r') as z:
                z.extractall(extract_dir)

            # Step 3: Load resources
            print("\n[3] Loading resources...")
            self.output_dir = output_dir  # Store for SVG saving
            self._load_relationships(extract_dir)
            self._load_styles(extract_dir)
            self._load_numbering(extract_dir)
            self._load_footnotes(extract_dir)
            self._extract_images(extract_dir, output_dir)  # Always extract images to subfolder

            # Step 4: Convert document
            print("\n[4] Converting document...")
            doc_xml = extract_dir / "word" / "document.xml"
            with open(doc_xml, 'rb') as f:
                doc_root = etree.fromstring(f.read())

            html_content = self._convert_body(doc_root)

            # Step 5: Generate HTML
            print("\n[5] Generating HTML...")
            full_html = self._generate_html(html_content, input_path.stem)

            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(full_html)

            # Generate body-only file for SharePoint pasting
            body_output_path = output_dir / f"{input_path.stem}_body.txt"
            body_html = self._generate_body_html(html_content)
            with open(body_output_path, 'w', encoding='utf-8') as f:
                f.write(body_html)

            print(f"\n{'='*70}")
            print("CONVERSION COMPLETE!")
            print(f"HTML Output: {output_path}")
            print(f"Body Output: {body_output_path}")
            print(f"Word Output: {word_output_path}")
            print(f"{'='*70}")

            return {
                'success': True,
                'output_path': str(output_path),
                'body_output_path': str(body_output_path),
                'word_output_path': str(word_output_path)
            }

        except Exception as e:
            print(f"ERROR: {e}")
            import traceback
            traceback.print_exc()
            return {'success': False, 'error': str(e)}

        finally:
            if temp_dir.exists():
                shutil.rmtree(temp_dir)

    def _convert_mathml_mode(self, input_path, output_path, output_dir):
        """Direct DOCX to HTML with MathML - no intermediate Word file"""

        temp_dir = Path(f"temp_mathml_{datetime.now().strftime('%Y%m%d_%H%M%S')}")

        try:
            # Step 1: Extract ORIGINAL DOCX directly (no pre-processing)
            print("\n[1] Extracting original document (MathML mode)...")
            extract_dir = temp_dir / "extracted"
            temp_dir.mkdir(exist_ok=True)
            with zipfile.ZipFile(input_path, 'r') as z:
                z.extractall(extract_dir)

            # Step 2: Load resources
            print("\n[2] Loading resources...")
            self.output_dir = output_dir
            self._load_relationships(extract_dir)
            self._load_styles(extract_dir)
            self._load_numbering(extract_dir)
            self._load_footnotes_wordhtml(extract_dir)
            self._extract_images(extract_dir, output_dir)

            # Step 3: Convert document (OMML equations converted inline to MathML)
            print("\n[3] Converting document with inline MathML...")
            doc_xml = extract_dir / "word" / "document.xml"
            with open(doc_xml, 'rb') as f:
                doc_root = etree.fromstring(f.read())

            html_content = self._convert_body(doc_root)

            # Step 4: Generate clean HTML (no MathJax, wordhtml.com format)
            print("\n[4] Generating HTML (wordhtml.com format, no JavaScript)...")
            full_html = self._generate_html_wordhtml(html_content, input_path.stem)

            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(full_html)

            # Generate body-only file for SharePoint pasting
            body_output_path = output_dir / f"{input_path.stem}_body.txt"
            body_html = self._generate_body_html(html_content)
            with open(body_output_path, 'w', encoding='utf-8') as f:
                f.write(body_html)

            print(f"\n{'='*70}")
            print("CONVERSION COMPLETE (MathML mode)!")
            print(f"HTML Output: {output_path}")
            print(f"Body Output: {body_output_path}")
            print(f"{'='*70}")

            return {
                'success': True,
                'output_path': str(output_path),
                'body_output_path': str(body_output_path)
            }

        except Exception as e:
            print(f"ERROR: {e}")
            import traceback
            traceback.print_exc()
            return {'success': False, 'error': str(e)}

        finally:
            if temp_dir.exists():
                shutil.rmtree(temp_dir)

    def _load_relationships(self, extract_dir):
        rels_path = extract_dir / "word" / "_rels" / "document.xml.rels"
        if not rels_path.exists():
            return
        with open(rels_path, 'rb') as f:
            root = etree.fromstring(f.read())
        ns = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
        for rel in root.xpath('//r:Relationship', namespaces=ns):
            self.relationships[rel.get('Id')] = {
                'target': rel.get('Target'),
                'type': rel.get('Type', '').split('/')[-1]
            }

    def _load_styles(self, extract_dir):
        styles_path = extract_dir / "word" / "styles.xml"
        if not styles_path.exists():
            return
        with open(styles_path, 'rb') as f:
            root = etree.fromstring(f.read())
        ns = {'w': self.namespaces['w']}
        for style in root.xpath('//w:style', namespaces=ns):
            style_id = style.get(f'{{{ns["w"]}}}styleId')
            name = style.xpath('w:name/@w:val', namespaces=ns)
            if name:
                self.styles[style_id] = name[0]

    def _load_numbering(self, extract_dir):
        num_path = extract_dir / "word" / "numbering.xml"
        if not num_path.exists():
            return
        with open(num_path, 'rb') as f:
            root = etree.fromstring(f.read())
        ns = {'w': self.namespaces['w']}

        abstract_nums = {}
        for abstract in root.xpath('//w:abstractNum', namespaces=ns):
            abs_id = abstract.get(f'{{{ns["w"]}}}abstractNumId')
            levels = {}
            for lvl in abstract.xpath('.//w:lvl', namespaces=ns):
                lvl_id = lvl.get(f'{{{ns["w"]}}}ilvl')
                fmt = lvl.xpath('.//w:numFmt/@w:val', namespaces=ns)
                levels[lvl_id] = fmt[0] if fmt else 'bullet'
            abstract_nums[abs_id] = levels

        for num in root.xpath('//w:num', namespaces=ns):
            num_id = num.get(f'{{{ns["w"]}}}numId')
            abs_ref = num.xpath('.//w:abstractNumId/@w:val', namespaces=ns)
            if abs_ref and abs_ref[0] in abstract_nums:
                self.numbering[num_id] = abstract_nums[abs_ref[0]]

    def _load_footnotes(self, extract_dir):
        fn_path = extract_dir / "word" / "footnotes.xml"
        if fn_path.exists():
            with open(fn_path, 'rb') as f:
                root = etree.fromstring(f.read())
            ns = {'w': self.namespaces['w']}
            for fn in root.xpath('//w:footnote', namespaces=ns):
                fn_id = fn.get(f'{{{ns["w"]}}}id')
                if fn_id and fn_id not in ['0', '-1']:
                    # Escape HTML entities in footnote text for safe output
                    raw_text = self._extract_text(fn)
                    self.footnotes[fn_id] = self._escape(raw_text)

    def _load_footnotes_wordhtml(self, extract_dir):
        """Load footnotes with full formatting for wordhtml.com style output"""
        fn_path = extract_dir / "word" / "footnotes.xml"
        if not fn_path.exists():
            return
        with open(fn_path, 'rb') as f:
            root = etree.fromstring(f.read())
        ns = {'w': self.namespaces['w']}
        for fn in root.xpath('//w:footnote', namespaces=ns):
            fn_id = fn.get(f'{{{ns["w"]}}}id')
            if fn_id and fn_id not in ['0', '-1']:
                self.footnotes[fn_id] = self._convert_footnote_content(fn)

    def _convert_footnote_content(self, fn_elem):
        """Convert footnote content with formatting preserved"""
        ns = self.namespaces
        parts = []
        for p in fn_elem.xpath('.//w:p', namespaces=ns):
            p_parts = []
            for child in p:
                tag = child.tag.split('}')[-1]
                if tag == 'r':
                    # Skip footnote reference marker inside footnote itself
                    if child.xpath('.//w:footnoteRef', namespaces=ns):
                        continue
                    p_parts.append(self._convert_run(child))
                elif tag == 'hyperlink':
                    p_parts.append(self._convert_hyperlink(child))
                elif tag == 'oMath' and self.equation_converter:
                    p_parts.append(self.equation_converter.convert(child, is_display=False))
                elif tag == 'oMathPara' and self.equation_converter:
                    omath = child.find('m:oMath', namespaces=ns)
                    if omath is not None:
                        p_parts.append(self.equation_converter.convert(omath, is_display=True))
            content = ''.join(p_parts)
            if content.strip():
                parts.append(content)
        return ' '.join(parts)

    def _extract_images(self, extract_dir, output_dir):
        media_dir = extract_dir / "word" / "media"
        if not media_dir.exists():
            return
        images_dir = output_dir / "images"
        images_dir.mkdir(exist_ok=True)
        for img in media_dir.iterdir():
            shutil.copy2(img, images_dir / img.name)
            self.images[img.name] = f"images/{img.name}"

    def _extract_text(self, elem):
        return ''.join(t.text or '' for t in elem.xpath('.//w:t', namespaces=self.namespaces))

    def _convert_body(self, doc_root):
        ns = self.namespaces
        body = doc_root.xpath('//w:body', namespaces=ns)[0]
        is_mathml = self.config.output_format == "mathml_html"

        parts = []
        list_items = []
        current_list = None
        current_num_id = None  # Track numId to continue lists after interruptions
        list_counters = {}  # Track count for each numId to use <ol start="N">

        for child in body:
            tag = child.tag.split('}')[-1]

            if tag == 'p':
                # Check for section break inside paragraph
                sect_pr = child.find('.//w:pPr/w:sectPr', namespaces=ns) if is_mathml else None

                num_pr = child.xpath('.//w:numPr', namespaces=ns)
                if num_pr:
                    num_id = child.xpath('.//w:numId/@w:val', namespaces=ns)
                    num_id_val = num_id[0] if num_id else None
                    ilvl = child.xpath('.//w:ilvl/@w:val', namespaces=ns)

                    list_type = 'ul'
                    if num_id_val and num_id_val in self.numbering:
                        lvl = ilvl[0] if ilvl else '0'
                        fmt = self.numbering[num_id_val].get(lvl, 'bullet')
                        if fmt in ['decimal', 'lowerLetter', 'upperLetter']:
                            list_type = 'ol'

                    # Check if we're continuing a different list or starting fresh
                    if current_list != list_type or current_num_id != num_id_val:
                        if list_items:
                            parts.append(self._wrap_list(list_items, current_list, current_num_id, list_counters))
                            list_items = []
                        current_list = list_type
                        current_num_id = num_id_val

                    content = self._convert_paragraph_content(child)
                    if content.strip():
                        list_items.append(content)
                        # Track count for this numId
                        if num_id_val:
                            list_counters[num_id_val] = list_counters.get(num_id_val, 0) + 1
                else:
                    if list_items:
                        parts.append(self._wrap_list(list_items, current_list, current_num_id, list_counters))
                        list_items = []
                        current_list = None
                        # Don't reset current_num_id - we might continue the list later
                    parts.append(self._convert_paragraph(child))

                # Add section break separator if present
                if sect_pr is not None:
                    parts.append('<hr>')

            elif tag == 'tbl':
                if list_items:
                    parts.append(self._wrap_list(list_items, current_list, current_num_id, list_counters))
                    list_items = []
                    current_list = None
                parts.append(self._convert_table(child))

            elif tag == 'sectPr':
                # Final section properties - safe to skip
                continue

            else:
                # Process any other element type to avoid content loss
                text = self._extract_text(child)
                if text.strip():
                    parts.append(f'<p>{self._escape(text)}</p>')

        if list_items:
            parts.append(self._wrap_list(list_items, current_list, current_num_id, list_counters))

        return '\n'.join(filter(None, parts))

    def _wrap_list(self, items, list_type, num_id=None, counters=None):
        """Wrap list items in <ol> or <ul> with proper numbering continuation.

        Args:
            items: List of HTML content for each list item
            list_type: 'ol' or 'ul'
            num_id: Word's numId for this list (used to track continuation)
            counters: Dict tracking total count for each numId
        """
        if not items:
            return ''

        html = '\n'.join(f'  <li>{item}</li>' for item in items)
        tag = list_type or 'ul'

        # For ordered lists, calculate start number to continue from previous items
        if tag == 'ol' and num_id and counters:
            total_count = counters.get(num_id, len(items))
            start_num = total_count - len(items) + 1
            if start_num > 1:
                # Continue numbering from where we left off
                return f'<ol start="{start_num}">\n{html}\n</ol>'

        return f'<{tag}>\n{html}\n</{tag}>'

    def _convert_paragraph_content(self, p_elem):
        ns = self.namespaces
        parts = []

        for child in p_elem:
            tag = child.tag.split('}')[-1]

            if tag == 'r':
                parts.append(self._convert_run(child))
            elif tag == 'hyperlink':
                parts.append(self._convert_hyperlink(child))
            elif tag == 'drawing':
                parts.append(self._convert_drawing(child))
            elif tag == 'oMath' and self.equation_converter:
                # MathML mode: convert equation inline
                parts.append(self.equation_converter.convert(child, is_display=False))
            elif tag == 'oMathPara' and self.equation_converter:
                # MathML mode: convert display equation
                omath = child.find('m:oMath', namespaces=ns)
                if omath is not None:
                    parts.append(self.equation_converter.convert(omath, is_display=True))
                else:
                    parts.append(self.equation_converter.convert(child, is_display=True))
            elif tag in ['pPr', 'bookmarkStart', 'bookmarkEnd']:
                continue
            else:
                text = self._extract_text(child)
                if text:
                    parts.append(self._escape(text))

        return ''.join(parts)

    def _convert_paragraph(self, p_elem):
        ns = self.namespaces

        style_id = p_elem.xpath('.//w:pStyle/@w:val', namespaces=ns)
        style_name = self.styles.get(style_id[0], '') if style_id else ''

        # Also get the style ID itself (e.g., "Heading1")
        actual_style_id = style_id[0] if style_id else ''

        heading_level = 0
        # Check both style name (e.g., "heading 1") and style ID (e.g., "Heading1")
        if 'heading' in style_name.lower() or 'Heading' in actual_style_id:
            # Try to extract number from style name first, then style ID
            match = re.search(r'(\d+)', style_name) or re.search(r'(\d+)', actual_style_id)
            if match:
                heading_level = min(int(match.group(1)), 6)

        content = self._convert_paragraph_content(p_elem)
        if not content.strip():
            # wordhtml.com preserves empty paragraphs as spacing
            if self.config.output_format == "mathml_html":
                return '<p>&nbsp;</p>'
            return ''

        if heading_level:
            return f'<h{heading_level}>{content}</h{heading_level}>'
        return f'<p>{content}</p>'

    def _convert_run(self, r_elem):
        ns = self.namespaces
        parts = []

        bold = bool(r_elem.xpath('.//w:b[not(@w:val="false")]', namespaces=ns))
        italic = bool(r_elem.xpath('.//w:i[not(@w:val="false")]', namespaces=ns))
        superscript = bool(r_elem.xpath('.//w:vertAlign[@w:val="superscript"]', namespaces=ns))
        subscript_text = bool(r_elem.xpath('.//w:vertAlign[@w:val="subscript"]', namespaces=ns))

        for child in r_elem:
            tag = child.tag.split('}')[-1]

            if tag == 't':
                parts.append(self._escape(child.text or ''))
            elif tag == 'drawing':
                parts.append(self._convert_drawing(child))
            elif tag == 'pict':
                parts.append(self._convert_pict(child))
            elif tag == 'AlternateContent':
                # Handle mc:AlternateContent - shapes, equations in textboxes
                # In MathML mode, also look for equations inside textboxes
                if self.equation_converter:
                    omath_elems = child.xpath('.//m:oMath', namespaces=ns)
                    if omath_elems:
                        for omath in omath_elems:
                            parts.append(self.equation_converter.convert(omath, is_display=False))
                        continue
                drawing = child.xpath('.//w:drawing', namespaces=ns)
                pict = child.xpath('.//w:pict', namespaces=ns)
                if drawing:
                    parts.append(self._convert_drawing(drawing[0]))
                elif pict:
                    parts.append(self._convert_pict(pict[0]))
            elif tag == 'footnoteReference':
                fn_id = child.get(f'{{{ns["w"]}}}id')
                # wordhtml.com format for both modes: <a href="#_ftn1" name="_ftnref1">[1]</a>
                parts.append(f'<a href="#_ftn{fn_id}" name="_ftnref{fn_id}">[{fn_id}]</a>')
            elif tag == 'br':
                parts.append('<br>')

        content = ''.join(parts)
        if italic:
            content = f'<i>{content}</i>'
        if bold:
            content = f'<b>{content}</b>'
        if superscript:
            content = f'<sup>{content}</sup>'
        if subscript_text:
            content = f'<sub>{content}</sub>'

        return content

    def _convert_hyperlink(self, h_elem):
        ns = self.namespaces
        r_id = h_elem.get(f'{{{ns["r"]}}}id')
        href = self.relationships.get(r_id, {}).get('target', '#')
        content = ''.join(self._convert_run(r) for r in h_elem.xpath('.//w:r', namespaces=ns))
        return f'<a href="{href}">{content}</a>'

    def _convert_table(self, tbl_elem):
        """Convert table to wordhtml.com format with tbody, width, and colspan."""
        ns = self.namespaces
        rows = []

        for tr in tbl_elem.xpath('./w:tr', namespaces=ns):
            cells = []
            for tc in tr.xpath('./w:tc', namespaces=ns):
                # Get cell properties for wordhtml.com format (both modes)
                width_attr = ''
                colspan_attr = ''

                tc_pr = tc.find('w:tcPr', namespaces=ns)
                if tc_pr is not None:
                    # Width: convert twips to pixels
                    tc_w = tc_pr.find('w:tcW', namespaces=ns)
                    if tc_w is not None:
                        w_val = tc_w.get(f'{{{ns["w"]}}}w', '')
                        w_type = tc_w.get(f'{{{ns["w"]}}}type', 'dxa')
                        if w_val and w_type == 'dxa':
                            px = int(int(w_val) * 96 / 1440)
                            width_attr = f' width="{px}"'

                    # Colspan (gridSpan)
                    grid_span = tc_pr.find('w:gridSpan', namespaces=ns)
                    if grid_span is not None:
                        span_val = grid_span.get(f'{{{ns["w"]}}}val', '')
                        if span_val and int(span_val) > 1:
                            colspan_attr = f' colspan="{span_val}"'

                # Convert cell content
                content = []
                for p in tc.xpath('./w:p', namespaces=ns):
                    p_content = self._convert_paragraph_content(p)
                    if p_content.strip():
                        content.append(p_content)

                cell_html = '<br>'.join(content) if content else '&nbsp;'
                cells.append((cell_html, width_attr, colspan_attr))
            rows.append(cells)

        # wordhtml.com format for both modes: <table><tbody><tr><td width="..." colspan="...">
        html = ['<table>', '<tbody>']
        for row in rows:
            html.append('<tr>')
            for cell_html, width_attr, colspan_attr in row:
                html.append(f'<td{colspan_attr}{width_attr}>{cell_html}</td>')
            html.append('</tr>')
        html.append('</tbody>')
        html.append('</table>')

        return '\n'.join(html)

    def _convert_drawing(self, drawing_elem):
        ns = self.namespaces

        # Check for image
        blip = drawing_elem.xpath('.//a:blip/@r:embed', namespaces=ns)
        if blip:
            r_id = blip[0]
            if r_id in self.relationships:
                target = self.relationships[r_id]['target']
                img_name = Path(target).name
                if img_name in self.images:
                    # Only include image tag if include_images is True
                    if self.config.include_images:
                        return f'<img src="{self.images[img_name]}" alt="Image" class="doc-image">'
                    else:
                        return ''  # Skip image but file is still extracted

        # Check for shapes (wsp = single shape, wpg = group of shapes)
        wsp = drawing_elem.xpath('.//wps:wsp', namespaces=ns)
        wpg = drawing_elem.xpath('.//wpg:wgp', namespaces=ns)

        if wsp or wpg:
            # This is a Word shape (not an image)
            if self.config.convert_shapes_to_svg:
                # Convert shape to SVG
                svg = self.svg_converter.convert_drawing_to_svg(drawing_elem)
                svg_filename = self._save_svg(svg)
                return f'<img src="{svg_filename}" alt="Shape" class="shape-svg">'
            else:
                # Shapes disabled - completely ignore shape and its content
                return ''

        return ''

    def _save_svg(self, svg_content):
        """Save SVG content to file and return relative path"""
        self.svg_counter += 1
        svg_filename = f"shape_{self.svg_counter}.svg"

        # Ensure images subfolder exists
        images_dir = self.output_dir / "images"
        images_dir.mkdir(exist_ok=True)

        # Save SVG file
        svg_path = images_dir / svg_filename
        with open(svg_path, 'w', encoding='utf-8') as f:
            f.write(svg_content)

        return f"images/{svg_filename}"

    def _convert_pict(self, pict_elem):
        """Convert VML pict elements"""
        ns = self.namespaces

        if self.config.convert_shapes_to_svg:
            # Check for shapes
            ovals = pict_elem.xpath('.//v:oval', namespaces=ns)
            rects = pict_elem.xpath('.//v:rect', namespaces=ns)

            if ovals or rects:
                svg = self.svg_converter.convert_vml_to_svg(pict_elem)
                # Save SVG to file
                svg_filename = self._save_svg(svg)
                return f'<img src="{svg_filename}" alt="Shape" class="shape-svg">'

        # Fallback: extract text
        texts = []
        for t in pict_elem.xpath('.//w:t/text()', namespaces=ns):
            if t.strip():
                texts.append(t.strip())

        if texts:
            return f'<span class="shape-text">{" ".join(texts)}</span>'

        return ''

    def _escape(self, text):
        if not text:
            return ''
        # Don't escape text containing math markers (they contain LaTeX)
        if '\\(' in text or '\\[' in text:
            return text
        # Don't escape MathML content
        if '<math ' in text or '<math>' in text:
            return text
        return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

    def _generate_body_html(self, content):
        """Generate body-only HTML for SharePoint pasting.

        Contains just the <div id="mathjax-content"> wrapper with content and
        footnotes. No DOCTYPE, html, head, style, script, or body tags.
        Image tags are removed (SharePoint team inserts images manually).
        Equations are wrapped with semantic HTML classes:
          inline  -> <span class="inline-math">\(...\)</span>
          display -> <span class="display-math">\[...\]</span>
        """
        import re
        # Remove image tags - SharePoint team inserts images manually
        body = re.sub(r'<img\s[^>]*>', '', content)

        # Strip any remaining equation markers (legacy)
        for marker in [self.config.inline_prefix, self.config.inline_suffix,
                       self.config.display_prefix, self.config.display_suffix]:
            if marker:
                body = body.replace(marker, '')

        # Collapse double spaces in equations (prevents invisible Unicode in MathML)
        body = re.sub(r'  +', ' ', body)

        # Wrap equations with semantic HTML classes
        body = re.sub(r'(\\\(.+?\\\))', r'<span class="inline-math">\1</span>', body)
        body = re.sub(r'(\\\[.+?\\\])', r'<span class="display-math">\1</span>', body, flags=re.DOTALL)

        footnotes_html = ''
        if self.footnotes:
            footnotes_parts = []
            for fn_id, fn_content in self.footnotes.items():
                footnotes_parts.append(
                    f'<p><a href="#_ftnref{fn_id}" name="_ftn{fn_id}">[{fn_id}]</a> {fn_content}</p>'
                )
            footnotes_html = '\n'.join(footnotes_parts)

        return f'''<div id="mathjax-content">
{body}
{footnotes_html}
</div>'''

    def _generate_html_wordhtml(self, content, title):
        """Generate HTML in wordhtml.com format - clean, no JavaScript"""
        import re

        config = self.config

        # Remove image tags - images are not included in output
        content = re.sub(r'<img\s[^>]*>', '', content)

        # Strip any remaining equation markers (legacy)
        for marker in [config.inline_prefix, config.inline_suffix, config.display_prefix, config.display_suffix]:
            if marker:
                content = content.replace(marker, '')

        # Collapse double spaces in equations (prevents invisible Unicode in MathML)
        content = re.sub(r'  +', ' ', content)

        # Wrap equations with semantic HTML classes
        content = re.sub(r'(\\\(.+?\\\))', r'<span class="inline-math">\1</span>', content)
        content = re.sub(r'(\\\[.+?\\\])', r'<span class="display-math">\1</span>', content, flags=re.DOTALL)

        direction = 'rtl' if config.rtl_direction else 'ltr'

        # Minimal styles for readability (optional)
        styles = ''
        if config.include_styles:
            styles = '''
    <style>
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            line-height: 1.8;
            max-width: 900px;
            margin: 0 auto;
            padding: 30px 40px;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin: 1em 0;
        }
        td, th {
            border: 1px solid #ddd;
            padding: 8px;
        }
        img {
            max-width: 100%;
        }
    </style>'''

        # Footnotes HTML with wordhtml.com naming (_ftn/_ftnref)
        footnotes_html = ''
        if self.footnotes:
            footnotes_parts = []
            for fn_id, fn_content in self.footnotes.items():
                footnotes_parts.append(
                    f'<p><a href="#_ftnref{fn_id}" name="_ftn{fn_id}">[{fn_id}]</a> {fn_content}</p>'
                )
            footnotes_html = '\n'.join(footnotes_parts)

        return f'''<!DOCTYPE html>
<html dir="{direction}">
<head>
<meta charset="UTF-8">
<title>{title}</title>{styles}
</head>
<body>
<div id="mathjax-content">
{content}
{footnotes_html}
</div>
</body>
</html>'''

    def _generate_html(self, content, title):
        """Generate clean HTML output - NO config panel (settings applied during conversion)"""
        import re

        config = self.config

        # Remove image tags - images are not included in output
        content = re.sub(r'<img\s[^>]*>', '', content)

        # Strip any remaining equation markers (legacy)
        for marker in [config.inline_prefix, config.inline_suffix, config.display_prefix, config.display_suffix]:
            if marker:
                content = content.replace(marker, '')

        # Collapse double spaces in equations (prevents invisible Unicode in MathML)
        content = re.sub(r'  +', ' ', content)

        # Wrap equations with semantic HTML classes
        # inline  \(...\) -> <span class="inline-math">\(...\)</span>
        # display \[...\] -> <span class="display-math">\[...\]</span>
        content = re.sub(r'(\\\(.+?\\\))', r'<span class="inline-math">\1</span>', content)
        content = re.sub(r'(\\\[.+?\\\])', r'<span class="display-math">\1</span>', content, flags=re.DOTALL)

        # MathJax script
        mathjax = ''
        if config.include_mathjax:
            mathjax = f'''
    <style>
        mjx-container {{ display: inline; }}
        mjx-container[display="block"] {{ display: block; text-align: center; margin: 1em 0; }}
    </style>
    <script>
        // SharePoint edit mode detection - skip MathJax in edit mode
        (function() {{
            var isEditMode = (
                window.location.search.indexOf('Mode=Edit') !== -1 ||
                window.location.search.indexOf('mode=edit') !== -1 ||
                document.querySelector('.sp-pageLayout-editMode') !== null ||
                document.querySelector('#spPageCanvasContent [contenteditable="true"]') !== null
            );
            if (isEditMode) return;

            // Check contenteditable ancestor of mathjax-content
            var el = document.getElementById('mathjax-content');
            if (el) {{
                var parent = el.parentElement;
                while (parent) {{
                    if (parent.getAttribute && parent.getAttribute('contenteditable') === 'true') return;
                    parent = parent.parentElement;
                }}
            }}

            window.MathJax = {{
                loader: {{load: ['input/tex']}},
                tex: {{
                    inlineMath: [['\\\\(', '\\\\)']],
                    displayMath: [['\\\\[', '\\\\]']]
                }},
                options: {{
                    renderActions: {{
                        assistiveMml: [],
                        typeset: [150,
                            function(doc) {{ for (var math of doc.math) MathJax.config.renderMathML(math, doc); }},
                            function(math, doc) {{ MathJax.config.renderMathML(math, doc); }}
                        ]
                    }}
                }},
                startup: {{
                    elements: ['#mathjax-content'],
                    pageReady: function() {{
                        return MathJax.startup.document.render();
                    }}
                }},
                renderMathML: function(math, doc) {{
                    math.typesetRoot = document.createElement('mjx-container');
                    var mml = MathJax.startup.toMML(math.root);
                    // Strip invisible operators, zero-width chars, bidi marks, BOM
                    mml = mml.replace(/[\\u2060-\\u2064\\u200B-\\u200F\\u061C\\u202A-\\u202C\\u2066-\\u2069\\uFEFF]/g, '');
                    mml = mml.replace(/&#x(206[0-9a-f]|200[b-f]|061c|202[a-c]|feff);/gi, '');
                    mml = mml.replace(/<mo[^>]*>\\s*<\\/mo>/g, '');
                    mml = mml.replace(/ data-[a-z-]+="[^"]*"/g, '');
                    // Collapse <msup><mi></mi><mo>X</mo></msup> → <mo>X</mo>
                    mml = mml.replace(/<msup>\\s*<mi\\s*\\/?>\\s*(<\\/mi>)?\\s*(<mo[^>]*>[^<]*<\\/mo>)\\s*<\\/msup>/g, '$2');
                    math.typesetRoot.innerHTML = mml;
                    if (math.display) math.typesetRoot.setAttribute('display', 'block');
                }}
            }};
        }})();
        // Clean clipboard: browser native MathML adds invisible chars on copy
        (function() {{
            var R = /[\\u2060-\\u2064\\u200B-\\u200F\\u061C\\u202A-\\u202C\\u2066-\\u2069\\uFEFF]/g;
            document.addEventListener('copy', function(e) {{
                var sel = window.getSelection();
                if (!sel || !sel.rangeCount || !sel.toString()) return;
                var range = sel.getRangeAt(0);
                var els = document.querySelectorAll('mjx-container, math');
                var hit = false;
                for (var i = 0; i < els.length; i++) {{
                    if (range.intersectsNode(els[i])) {{ hit = true; break; }}
                }}
                if (!hit) return;
                e.clipboardData.setData('text/plain', sel.toString().replace(R, ''));
                var d = document.createElement('div');
                d.appendChild(range.cloneContents());
                e.clipboardData.setData('text/html', d.innerHTML.replace(R, ''));
                e.preventDefault();
            }});
        }})();
    </script>
    <script defer src="https://cdn.jsdelivr.net/npm/mathjax@4/startup.js"></script>
'''

        # CSS
        direction = 'rtl' if config.rtl_direction else 'ltr'

        styles = f'''
    <style>
        * {{ box-sizing: border-box; }}

        body {{
            font-family: {config.font_family};
            font-size: 16px;
            line-height: 1.8;
            max-width: {config.max_width};
            margin: 0 auto;
            padding: 30px 40px;
            direction: {direction};
            background: #fff;
            color: #333;
        }}

        /* Headings */
        h1, h2, h3, h4 {{
            color: #1a365d;
            margin-top: 1.5em;
        }}
        h1 {{ font-size: 2em; border-bottom: 3px solid #3182ce; padding-bottom: 0.3em; text-align: center; }}
        h2 {{ font-size: 1.6em; border-bottom: 2px solid #63b3ed; color: #2c5282; }}

        /* Tables */
        table {{
            border-collapse: collapse;
            width: 100%;
            margin: 1.5em 0;
        }}
        th, td {{
            border: 1px solid #ddd;
            padding: 12px;
            text-align: right;
        }}
        th {{ background: #3182ce; color: white; }}
        tr:nth-child(even) {{ background: #f7fafc; }}

        /* Lists */
        ul, ol {{
            margin: 1em 0;
            padding-right: 2em;
        }}
        li {{ margin: 0.5em 0; }}

        /* Images */
        .doc-image {{
            max-width: 100%;
            display: block;
            margin: 1.5em auto;
        }}

        /* Shapes */
        .shape-svg {{
            display: inline-block;
            margin: 5px;
            vertical-align: middle;
        }}
        .shape-text {{
            display: inline-block;
            padding: 5px 10px;
            border: 1px solid #ccc;
            border-radius: 50%;
            background: #f7f7f7;
            margin: 3px;
        }}

        /* Math */
        .inline-math {{ display: inline; margin: 0 0.2em; }}
        .display-math {{
            display: block;
            text-align: center;
            margin: 1.5em 0;
            padding: 1em;
            background: #f8fafc;
            border-radius: 8px;
            border-right: 4px solid #3182ce;
        }}

        /* Footnote references in text */
        sup a {{
            color: #3182ce;
            text-decoration: none;
            font-weight: bold;
        }}
        sup a:hover {{
            text-decoration: underline;
        }}

        /* Print */
        @media print {{
            body {{ max-width: none; padding: 0; }}
        }}
    </style>
''' if config.include_styles else ''

        # Footnotes HTML with wordhtml.com naming (_ftn/_ftnref)
        footnotes_html = ''
        if self.footnotes:
            footnotes_parts = []
            for fn_id, fn_content in self.footnotes.items():
                footnotes_parts.append(
                    f'<p><a href="#_ftnref{fn_id}" name="_ftn{fn_id}">[{fn_id}]</a> {fn_content}</p>'
                )
            footnotes_html = '\n'.join(footnotes_parts)

        # Equation copy menu script (inline) - read from external JS file
        copy_menu_script = ''
        if config.include_mathjax:
            copy_menu_js = Path(__file__).resolve().parent.parent / 'mathjax-copy-menu.js'
            try:
                js_content = copy_menu_js.read_text(encoding='utf-8')
                copy_menu_script = f'\n    <script>\n{js_content}\n    </script>'
            except FileNotFoundError:
                print(f"    Warning: {copy_menu_js} not found, skipping copy menu")

        return f'''<!DOCTYPE html>
<html lang="ar" dir="{direction}">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
{mathjax}
{styles}
</head>
<body>
<div id="mathjax-content">
{content}
{footnotes_html}
</div>
{copy_menu_script}
</body>
</html>'''


def test_full_converter():
    """Test the full converter on ALL documents in test folder"""

    config = ConversionConfig(
        convert_shapes_to_svg=True,
        include_images=True,
        include_mathjax=True,
        inline_prefix='MATHSTARTINLINE',
        inline_suffix='MATHENDINLINE',
        display_prefix='MATHSTARTDISPLAY',
        display_suffix='MATHENDDISPLAY'
    )

    converter = FullWordToHTMLConverter(config)

    test_folder = Path(r"D:\Development\document-processing-api-2\document-processing-api\test docs")
    output_dir = Path(r"D:\Development\document-processing-api-2\backend\test_output\html_full_output")
    output_dir.mkdir(exist_ok=True)

    # Find all .docx files (excluding already converted ones)
    docx_files = [f for f in test_folder.glob("*.docx")
                  if not f.name.endswith('_latex_equations.docx')
                  and not f.name.endswith('_standalone.docx')
                  and not f.name.endswith('_converted.docx')]

    print(f"\n{'#'*70}")
    print(f"# TESTING ALL {len(docx_files)} DOCUMENTS IN TEST FOLDER")
    print(f"{'#'*70}")

    results = []
    for i, doc in enumerate(docx_files, 1):
        print(f"\n{'='*70}")
        print(f"# Document {i}/{len(docx_files)}: {doc.name}")
        print(f"{'='*70}")

        result = converter.convert(doc, output_dir=output_dir)
        results.append({
            'doc': doc.name,
            'success': result.get('success', False),
            'output': result.get('output_path', ''),
            'error': result.get('error', '')
        })

        # Reset converter state for next document
        converter = FullWordToHTMLConverter(config)

    # Summary
    print(f"\n{'#'*70}")
    print("# CONVERSION SUMMARY")
    print(f"{'#'*70}")

    success_count = sum(1 for r in results if r['success'])
    print(f"\nTotal documents: {len(results)}")
    print(f"Successful:      {success_count}")
    print(f"Failed:          {len(results) - success_count}")

    print(f"\nResults:")
    for r in results:
        status = "✅" if r['success'] else "❌"
        print(f"  {status} {r['doc']}")
        if r['error']:
            print(f"      Error: {r['error']}")

    print(f"\nOutput folder: {output_dir}")


if __name__ == "__main__":
    test_full_converter()
