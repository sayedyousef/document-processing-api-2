"""
Standalone ZIP Implementation for Word Document Equation Conversion
No Word COM required - pure Python implementation
"""

import sys
import io
import zipfile
import shutil
import re
import json
from pathlib import Path
from lxml import etree
from datetime import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


class StandaloneZipConverter:
    """Convert Word equations using only ZIP manipulation"""

    def __init__(self):
        self.namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
            'v': 'urn:schemas-microsoft-com:vml',
            'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
        }
        self.equation_count = 0
        self.equations_replaced = 0
        self.equations_in_vml = 0

    def extract_docx(self, docx_path):
        """Extract docx to temporary directory"""
        temp_dir = Path(f"temp_extract_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
        with zipfile.ZipFile(docx_path, 'r') as z:
            z.extractall(temp_dir)
        return temp_dir

    def repackage_docx(self, temp_dir, output_path):
        """Repackage extracted files back to docx"""
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as z:
            for root, dirs, files in temp_dir.walk():
                for file in files:
                    file_path = root / file
                    arcname = str(file_path.relative_to(temp_dir))
                    z.write(file_path, arcname)

    def convert_omml_to_latex(self, omml_element):
        """Convert OMML equation to LaTeX
        This is a simplified converter - in production you'd use a full OMML->LaTeX library
        """
        # For now, return a placeholder that indicates the equation was found
        # In a real implementation, you'd parse the OMML tree and generate LaTeX

        # Check if it's a simple equation we can handle
        equation_text = self.extract_equation_text(omml_element)

        # Default to inline math
        return f"\\({equation_text}\\)"

    def extract_equation_text(self, omml_element):
        """Extract text from OMML equation"""
        # This is a simplified extraction - gets all text nodes
        texts = []
        for text_elem in omml_element.xpath('.//m:t', namespaces=self.namespaces):
            if text_elem.text:
                texts.append(text_elem.text)

        return ' '.join(texts) if texts else 'equation'

    def is_in_vml(self, element):
        """Check if element is inside VML textbox"""
        # Check for VML textbox ancestors
        vml_ancestor = element.xpath('ancestor::v:textbox', namespaces=self.namespaces)
        fallback_ancestor = element.xpath('ancestor::mc:Fallback', namespaces=self.namespaces)
        txbx_ancestor = element.xpath('ancestor::w:txbxContent', namespaces=self.namespaces)

        return bool(vml_ancestor or fallback_ancestor or txbx_ancestor)

    def create_text_run_with_latex(self, latex_text):
        """Create a text run element with LaTeX"""
        # Create paragraph with text run containing LaTeX
        w_ns = self.namespaces['w']

        # Create run element
        run = etree.Element(f"{{{w_ns}}}r")

        # Add run properties if needed
        run_props = etree.SubElement(run, f"{{{w_ns}}}rPr")

        # Add text element
        text = etree.SubElement(run, f"{{{w_ns}}}t")
        text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

        # Add LaTeX markers for HTML conversion
        if latex_text.startswith('\\('):
            text.text = f"MATHSTARTINLINE{latex_text}MATHENDINLINE"
        elif latex_text.startswith('\\['):
            text.text = f"MATHSTARTDISPLAY{latex_text}MATHENDDISPLAY"
        else:
            text.text = latex_text

        return run

    def process_document_xml(self, xml_path, convert_vml=False):
        """Process document.xml to replace equations

        Args:
            xml_path: Path to document.xml
            convert_vml: If True, attempt to convert VML textbox equations (WARNING: may break document!)
        """
        with open(xml_path, 'rb') as f:
            xml_content = f.read()

        # Parse XML
        root = etree.fromstring(xml_content)

        # Find all OMML equations
        equations = root.xpath('//m:oMath', namespaces=self.namespaces)
        self.equation_count = len(equations)

        print(f"Found {self.equation_count} OMML equations")
        if convert_vml:
            print("⚠️ WARNING: VML conversion may break document structure!")
        print(f"VML conversion mode: {'ENABLED (RISKY!)' if convert_vml else 'DISABLED (SAFE)'}")

        replacements = []
        vml_replacements = []
        vml_skipped = 0

        for i, eq in enumerate(equations, 1):
            # Check if in VML
            in_vml = self.is_in_vml(eq)

            if in_vml:
                self.equations_in_vml += 1
                if not convert_vml:
                    print(f"  Equation {i}: Inside VML textbox (skipping to preserve document)")
                    vml_skipped += 1
                    continue
                else:
                    print(f"  Equation {i}: Inside VML textbox (RISKY conversion attempt)")

            # Get parent element
            parent = eq.getparent()
            if parent is None:
                continue

            # Convert to LaTeX
            latex = self.convert_omml_to_latex(eq)

            # Store replacement info
            if in_vml:
                vml_replacements.append({
                    'equation': eq,
                    'parent': parent,
                    'latex': latex,
                    'index': i
                })
            else:
                replacements.append({
                    'equation': eq,
                    'parent': parent,
                    'latex': latex,
                    'index': i
                })

        # Perform regular replacements first
        print(f"\nProcessing {len(replacements)} regular equations...")
        for repl in replacements:
            eq = repl['equation']
            parent = repl['parent']
            latex = repl['latex']

            try:
                # Create text run with LaTeX
                text_run = self.create_text_run_with_latex(latex)

                # Find position of equation in parent
                index = list(parent).index(eq)

                # Remove equation
                parent.remove(eq)

                # Insert text run at same position
                parent.insert(index, text_run)

                self.equations_replaced += 1
                print(f"  Equation {repl['index']}: Replaced with LaTeX")

            except Exception as e:
                print(f"  Equation {repl['index']}: Failed to replace - {e}")

        # Attempt VML replacements if enabled
        if convert_vml and vml_replacements:
            print(f"\nAttempting to convert {len(vml_replacements)} VML textbox equations...")
            vml_converted = 0

            for repl in vml_replacements:
                eq = repl['equation']
                parent = repl['parent']
                latex = repl['latex']

                try:
                    # For VML equations, we need to be more careful
                    # Try to create a text run with LaTeX
                    text_run = self.create_text_run_with_latex(latex)

                    # Find position of equation in parent
                    index = list(parent).index(eq)

                    # Remove equation
                    parent.remove(eq)

                    # Insert text run at same position
                    parent.insert(index, text_run)

                    self.equations_replaced += 1
                    vml_converted += 1
                    print(f"  Equation {repl['index']}: VML equation replaced with LaTeX!")

                except Exception as e:
                    print(f"  Equation {repl['index']}: Failed to replace VML equation - {e}")

            print(f"\nVML conversion results: {vml_converted}/{len(vml_replacements)} successfully converted")

        # Save modified XML
        with open(xml_path, 'wb') as f:
            f.write(etree.tostring(root, encoding='UTF-8', xml_declaration=True))

        return {
            'equations_found': self.equation_count,
            'equations_replaced': self.equations_replaced,
            'equations_in_vml': self.equations_in_vml,
            'vml_attempted': len(vml_replacements) if convert_vml else 0
        }

    def convert_document(self, input_path, output_path, convert_vml=False):
        """Main conversion method

        Args:
            input_path: Input Word document path
            output_path: Output Word document path
            convert_vml: If True, attempt to convert VML textbox equations
        """
        print(f"\n{'='*60}")
        print("STANDALONE ZIP CONVERSION")
        if convert_vml:
            print("MODE: FULL CONVERSION (Including VML textboxes)")
        else:
            print("MODE: REGULAR CONVERSION (Skipping VML textboxes)")
        print(f"{'='*60}")
        print(f"Input: {input_path}")
        print(f"Output: {output_path}")

        temp_dir = None

        try:
            # Extract docx
            print("\n📦 Extracting document...")
            temp_dir = self.extract_docx(input_path)

            # Process document.xml
            doc_xml_path = temp_dir / "word" / "document.xml"
            if not doc_xml_path.exists():
                raise FileNotFoundError("document.xml not found in docx")

            print("\n🔧 Processing equations...")
            results = self.process_document_xml(doc_xml_path, convert_vml=convert_vml)

            # Repackage docx
            print("\n📦 Repackaging document...")
            self.repackage_docx(temp_dir, output_path)

            print(f"\n✅ Conversion complete!")
            print(f"  Equations found: {results['equations_found']}")
            print(f"  Equations replaced: {results['equations_replaced']}")
            print(f"  Equations in VML: {results['equations_in_vml']}")
            if convert_vml and results.get('vml_attempted', 0) > 0:
                print(f"  VML equations attempted: {results['vml_attempted']}")
                success_rate = (results['equations_replaced'] / results['equations_found'] * 100) if results['equations_found'] > 0 else 0
                print(f"  Overall success rate: {success_rate:.1f}%")
            print(f"  Output saved to: {output_path}")

            return results

        except Exception as e:
            print(f"\n❌ Conversion failed: {e}")
            import traceback
            traceback.print_exc()
            return {'error': str(e)}

        finally:
            # Cleanup temp directory
            if temp_dir and temp_dir.exists():
                shutil.rmtree(temp_dir)


class AdvancedOmmlToLatexConverter:
    """Advanced OMML to LaTeX converter with more equation support"""

    def __init__(self):
        self.namespaces = {
            'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math'
        }

    def convert(self, omml_element):
        """Convert OMML element to LaTeX"""
        # Check for fraction
        frac = omml_element.find('.//m:f', self.namespaces)
        if frac is not None:
            return self.convert_fraction(frac)

        # Check for superscript
        sup = omml_element.find('.//m:sSup', self.namespaces)
        if sup is not None:
            return self.convert_superscript(sup)

        # Check for subscript
        sub = omml_element.find('.//m:sSub', self.namespaces)
        if sub is not None:
            return self.convert_subscript(sub)

        # Check for square root
        rad = omml_element.find('.//m:rad', self.namespaces)
        if rad is not None:
            return self.convert_radical(rad)

        # Default: extract all text
        texts = []
        for text_elem in omml_element.xpath('.//m:t', namespaces=self.namespaces):
            if text_elem.text:
                texts.append(text_elem.text)

        text = ' '.join(texts) if texts else 'equation'
        return f"\\({text}\\)"

    def convert_fraction(self, frac_elem):
        """Convert fraction element"""
        num = frac_elem.find('.//m:num', self.namespaces)
        den = frac_elem.find('.//m:den', self.namespaces)

        num_text = self.extract_text(num) if num is not None else '?'
        den_text = self.extract_text(den) if den is not None else '?'

        return f"\\(\\frac{{{num_text}}}{{{den_text}}}\\)"

    def convert_superscript(self, sup_elem):
        """Convert superscript element"""
        base = sup_elem.find('.//m:e', self.namespaces)
        sup = sup_elem.find('.//m:sup', self.namespaces)

        base_text = self.extract_text(base) if base is not None else ''
        sup_text = self.extract_text(sup) if sup is not None else ''

        return f"\\({base_text}^{{{sup_text}}}\\)"

    def convert_subscript(self, sub_elem):
        """Convert subscript element"""
        base = sub_elem.find('.//m:e', self.namespaces)
        sub = sub_elem.find('.//m:sub', self.namespaces)

        base_text = self.extract_text(base) if base is not None else ''
        sub_text = self.extract_text(sub) if sub is not None else ''

        return f"\\({base_text}_{{{sub_text}}}\\)"

    def convert_radical(self, rad_elem):
        """Convert radical (root) element"""
        deg = rad_elem.find('.//m:deg', self.namespaces)
        rad = rad_elem.find('.//m:e', self.namespaces)

        rad_text = self.extract_text(rad) if rad is not None else ''

        if deg is not None:
            deg_text = self.extract_text(deg)
            return f"\\(\\sqrt[{deg_text}]{{{rad_text}}}\\)"
        else:
            return f"\\(\\sqrt{{{rad_text}}}\\)"

    def extract_text(self, element):
        """Extract all text from element"""
        if element is None:
            return ''

        texts = []
        for text_elem in element.xpath('.//m:t', namespaces=self.namespaces):
            if text_elem.text:
                texts.append(text_elem.text)

        return ''.join(texts)


def test_standalone_converter():
    """Test the standalone converter"""
    print("="*60)
    print("TESTING STANDALONE ZIP CONVERTER")
    print("="*60)

    # Test with default test docs
    test_dir = Path(r"D:\Development\document-processing-api-2\document-processing-api\test docs")
    output_dir = Path("test_standalone_output")
    output_dir.mkdir(exist_ok=True)

    converter = StandaloneZipConverter()

    # Test with first Arabic document
    test_file = test_dir / "التشابه (جاهزة للنشر) - Copy.docx"
    if test_file.exists():
        output_file = output_dir / f"{test_file.stem}_standalone.docx"
        results = converter.convert_document(test_file, output_file)

        # Save results
        results_file = output_dir / "standalone_results.json"
        with open(results_file, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, ensure_ascii=False)

        print(f"\n✓ Results saved to: {results_file}")
    else:
        print(f"❌ Test file not found: {test_file}")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        # Process specific file
        input_file = Path(sys.argv[1])
        if input_file.exists():
            output_file = input_file.parent / f"{input_file.stem}_standalone.docx"
            converter = StandaloneZipConverter()
            converter.convert_document(input_file, output_file)
        else:
            print(f"❌ File not found: {input_file}")
    else:
        # Run test
        test_standalone_converter()