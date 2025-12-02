"""
ENHANCED ZIP-based Equation Converter
Handles ALL equation types including shapes/textboxes WITHOUT Word COM

Key insight: Equations in shapes appear in TWO places in XML:
1. mc:Choice - Modern version (Word 2010+) in wps:txbx
2. mc:Fallback - Legacy version in v:textbox

We need to:
1. Process equations in mc:Choice (the "real" ones)
2. ALSO process equations in mc:Fallback (for compatibility)
3. Process regular body equations

This achieves 100% conversion like Word COM, but without requiring Word!
"""

import sys
import io
import zipfile
import shutil
from pathlib import Path
from lxml import etree
from datetime import datetime

# Set UTF-8 encoding for stdout only if not already wrapped
if not isinstance(sys.stdout, io.TextIOWrapper) or sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    except:
        pass


class EnhancedZipConverter:
    """Enhanced ZIP converter that handles ALL equation types"""

    def __init__(self, inline_prefix='MATHSTARTINLINE', inline_suffix='MATHENDINLINE',
                 display_prefix='MATHSTARTDISPLAY', display_suffix='MATHENDDISPLAY'):
        # Store marker configuration
        self.inline_prefix = inline_prefix
        self.inline_suffix = inline_suffix
        self.display_prefix = display_prefix
        self.display_suffix = display_suffix

        self.namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
            'v': 'urn:schemas-microsoft-com:vml',
            'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
            'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
            'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }

        # Import LaTeX converter
        try:
            # Add current backend folder to path
            backend_dir = Path(__file__).parent
            if str(backend_dir) not in sys.path:
                sys.path.insert(0, str(backend_dir))
            from doc_processor.omml_2_latex import DirectOmmlToLatex
            self.latex_converter = DirectOmmlToLatex()
            print("LaTeX converter loaded successfully")
        except ImportError as e:
            print(f"Warning: Could not import LaTeX converter: {e}")
            self.latex_converter = None

    def _extract_equation_text(self, omml_element):
        """Extract text content from OMML equation"""
        texts = []
        for text_elem in omml_element.xpath('.//m:t', namespaces=self.namespaces):
            if text_elem.text:
                texts.append(text_elem.text)
        return ''.join(texts)

    def _convert_to_latex(self, omml_element):
        """Convert OMML element to LaTeX"""
        if self.latex_converter:
            try:
                result = self.latex_converter.parse(omml_element)
                if result and result.strip():
                    return result
                else:
                    print(f"    Warning: LaTeX converter returned empty for equation")
            except Exception as e:
                print(f"    Warning: LaTeX conversion error: {e}")
        # Fallback: just extract text
        fallback = self._extract_equation_text(omml_element)
        if fallback:
            print(f"    Using fallback text extraction: {fallback[:50]}...")
        return fallback

    def _create_latex_run(self, latex_text, is_display=False):
        """Create a w:r element containing LaTeX text with markers"""

        w_ns = '{' + self.namespaces['w'] + '}'

        # Create run
        run = etree.Element(w_ns + 'r')

        # Add text element
        text = etree.SubElement(run, w_ns + 't')
        text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

        # Add markers (use configured prefix/suffix)
        if is_display:
            prefix = self.display_prefix
            suffix = self.display_suffix
            text.text = f' {prefix}\\[{latex_text}\\]{suffix} ' if prefix or suffix else f' \\[{latex_text}\\] '
        else:
            prefix = self.inline_prefix
            suffix = self.inline_suffix
            text.text = f' {prefix}\\({latex_text}\\){suffix} ' if prefix or suffix else f' \\({latex_text}\\) '

        return run

    def _get_equation_location(self, eq):
        """Determine where an equation is located"""

        # Check ancestors
        is_in_choice = bool(eq.xpath('ancestor::mc:Choice', namespaces=self.namespaces))
        is_in_fallback = bool(eq.xpath('ancestor::mc:Fallback', namespaces=self.namespaces))
        is_in_vml = bool(eq.xpath('ancestor::v:textbox', namespaces=self.namespaces))
        is_in_wps = bool(eq.xpath('ancestor::wps:txbx', namespaces=self.namespaces))
        is_in_txbx = bool(eq.xpath('ancestor::w:txbxContent', namespaces=self.namespaces))

        if is_in_choice:
            return 'mc_choice'
        elif is_in_fallback:
            return 'mc_fallback'
        elif is_in_vml:
            return 'vml'
        elif is_in_wps or is_in_txbx:
            return 'textbox'
        else:
            return 'main_body'

    def _replace_equation(self, eq, latex_text):
        """Replace a single equation with LaTeX text run"""

        parent = eq.getparent()
        if parent is None:
            return False

        parent_tag = parent.tag.split('}')[-1] if '}' in parent.tag else parent.tag

        # Determine if display or inline
        is_display = len(latex_text) > 50

        # Check if equation is in oMathPara (display equation)
        omath_para = eq.xpath('ancestor::m:oMathPara', namespaces=self.namespaces)
        if omath_para:
            is_display = True

        # Create replacement
        latex_run = self._create_latex_run(latex_text, is_display)

        try:
            # Get position of equation in parent
            index = list(parent).index(eq)

            # Remove equation
            parent.remove(eq)

            # Insert LaTeX run at same position
            parent.insert(index, latex_run)

            return True

        except Exception as e:
            print(f"    Error replacing: {e}")
            return False

    def process_document(self, input_path, output_path=None):
        """
        Process document and convert ALL equations

        Args:
            input_path: Path to input .docx file
            output_path: Path for output file (optional)

        Returns:
            dict with conversion results
        """

        input_path = Path(input_path).absolute()

        if not output_path:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_path = input_path.parent / f"{input_path.stem}_zip_converted_{timestamp}.docx"
        else:
            output_path = Path(output_path).absolute()

        print("\n" + "="*70)
        print("ENHANCED ZIP EQUATION CONVERTER")
        print("="*70)
        print(f"Input:  {input_path}")
        print(f"Output: {output_path}")
        print("="*70)

        # Create temp directory
        temp_dir = Path(f"temp_zip_{datetime.now().strftime('%Y%m%d_%H%M%S')}")

        try:
            # Extract docx
            print("\n[1] Extracting document...")
            with zipfile.ZipFile(input_path, 'r') as z:
                z.extractall(temp_dir)

            # Read document.xml
            doc_xml_path = temp_dir / "word" / "document.xml"
            with open(doc_xml_path, 'rb') as f:
                content = f.read()

            root = etree.fromstring(content)

            # Find ALL equations
            print("\n[2] Analyzing equations...")
            all_equations = root.xpath('//m:oMath', namespaces=self.namespaces)
            print(f"    Total m:oMath elements: {len(all_equations)}")

            # Categorize equations
            categories = {
                'main_body': [],
                'mc_choice': [],
                'mc_fallback': [],
                'vml': [],
                'textbox': []
            }

            for eq in all_equations:
                location = self._get_equation_location(eq)
                categories[location].append(eq)

            print(f"\n    Equation locations:")
            print(f"      Main body:     {len(categories['main_body'])}")
            print(f"      mc:Choice:     {len(categories['mc_choice'])}")
            print(f"      mc:Fallback:   {len(categories['mc_fallback'])}")
            print(f"      VML:           {len(categories['vml'])}")
            print(f"      Other textbox: {len(categories['textbox'])}")

            # Calculate unique equations
            unique_count = len(categories['main_body']) + len(categories['mc_choice'])
            duplicate_count = len(categories['mc_fallback']) + len(categories['vml'])

            print(f"\n    UNIQUE equations:    {unique_count}")
            print(f"    Fallback duplicates: {duplicate_count}")

            # Process ALL equations (including duplicates for compatibility)
            print("\n[3] Converting equations...")

            replaced_count = 0
            failed_count = 0

            # Process in order: main_body, mc_choice, mc_fallback, vml, textbox
            # Process in REVERSE order within each category to maintain positions

            for category in ['main_body', 'mc_choice', 'mc_fallback', 'vml', 'textbox']:
                equations = categories[category]
                if not equations:
                    continue

                print(f"\n    Processing {category}: {len(equations)} equations")

                # Reverse to process from end to start
                for eq in reversed(equations):
                    try:
                        # Get LaTeX
                        latex = self._convert_to_latex(eq)

                        # Replace
                        if self._replace_equation(eq, latex):
                            replaced_count += 1
                        else:
                            failed_count += 1

                    except Exception as e:
                        print(f"      Error: {e}")
                        failed_count += 1

            # Save modified XML
            print("\n[4] Saving modified document...")
            with open(doc_xml_path, 'wb') as f:
                f.write(etree.tostring(root, encoding='UTF-8', xml_declaration=True))

            # Repackage docx
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as z:
                for file_path in temp_dir.rglob('*'):
                    if file_path.is_file():
                        arcname = str(file_path.relative_to(temp_dir))
                        z.write(file_path, arcname)

            # Results
            print("\n" + "="*70)
            print("CONVERSION COMPLETE!")
            print("="*70)
            print(f"Total equations found:  {len(all_equations)}")
            print(f"Unique equations:       {unique_count}")
            print(f"Fallback copies:        {duplicate_count}")
            print(f"Successfully replaced:  {replaced_count}")
            print(f"Failed:                 {failed_count}")
            print(f"Success rate:           {replaced_count/len(all_equations)*100:.1f}%")
            print(f"\nOutput: {output_path}")

            return {
                'success': True,
                'output_path': str(output_path),
                'total_equations': len(all_equations),
                'unique_equations': unique_count,
                'replaced': replaced_count,
                'failed': failed_count
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


def verify_conversion(converted_path):
    """Verify conversion by counting remaining OMML and markers"""

    print("\n" + "="*70)
    print("VERIFYING CONVERSION")
    print("="*70)

    with zipfile.ZipFile(converted_path, 'r') as z:
        with z.open('word/document.xml') as f:
            content = f.read()
            root = etree.fromstring(content)

            ns = {'m': 'http://schemas.openxmlformats.org/officeDocument/2006/math'}

            # Count remaining OMML
            remaining = root.xpath('//m:oMath', namespaces=ns)

            # Get all text content
            text_content = etree.tostring(root, encoding='unicode')

            # Count markers
            inline_count = text_content.count('MATHSTARTINLINE')
            display_count = text_content.count('MATHSTARTDISPLAY')

    print(f"Remaining OMML equations: {len(remaining)}")
    print(f"LaTeX markers created:")
    print(f"  Inline:  {inline_count}")
    print(f"  Display: {display_count}")
    print(f"  Total:   {inline_count + display_count}")

    if len(remaining) == 0:
        print("\n SUCCESS: All equations converted!")
    else:
        print(f"\n WARNING: {len(remaining)} equations remain unconverted")

    return {
        'remaining': len(remaining),
        'markers': inline_count + display_count
    }


if __name__ == "__main__":
    # Test with document that has shape equations (144 total, 107 unique)
    test_doc1 = Path(r"D:\Development\document-processing-api-2\document-processing-api\test docs\الدالة واحد لواحد (جاهزة للنشر) - Copy.docx")

    # Test with document without shapes (89 equations)
    test_doc2 = Path(r"D:\Development\document-processing-api-2\document-processing-api\test docs\التشابه (جاهزة للنشر) - Copy.docx")

    output_dir = Path(r"D:\Development\document-processing-api-2\backend\test_output")
    output_dir.mkdir(exist_ok=True)

    converter = EnhancedZipConverter()

    # Test document 1 (with shapes)
    if test_doc1.exists():
        print("\n" + "#"*70)
        print("# TEST 1: Document with shapes (144 equations, 107 unique)")
        print("#"*70)

        output1 = output_dir / "doc1_zip_converted.docx"
        result1 = converter.process_document(test_doc1, output1)

        if result1.get('success'):
            verify_conversion(output1)

    # Test document 2 (no shapes)
    if test_doc2.exists():
        print("\n" + "#"*70)
        print("# TEST 2: Document without shapes (89 equations)")
        print("#"*70)

        output2 = output_dir / "doc2_zip_converted.docx"
        result2 = converter.process_document(test_doc2, output2)

        if result2.get('success'):
            verify_conversion(output2)
