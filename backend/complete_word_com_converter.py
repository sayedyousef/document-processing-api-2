"""
COMPLETE Word COM Equation Converter
Handles ALL equation types including shapes/textboxes

This converter can access:
1. Main document equations (via doc.OMaths)
2. Shape/textbox equations (via Shape.TextFrame.TextRange.OMaths)

Key insight: XML may show more equations than actually exist because
mc:Fallback contains DUPLICATE copies for older Word versions.
Word COM accesses the mc:Choice (modern) versions.
"""

import sys
import io
import win32com.client
import pythoncom
import zipfile
from pathlib import Path
from lxml import etree
from datetime import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


class CompleteWordCOMConverter:
    """Complete Word COM converter that handles ALL equation types"""

    def __init__(self):
        self.word = None
        self.doc = None
        self.all_equations = []

    def _start_word(self):
        """Initialize Word application"""
        pythoncom.CoInitialize()
        self.word = win32com.client.Dispatch("Word.Application")
        self.word.Visible = False
        self.word.DisplayAlerts = False
        self.word.ScreenUpdating = False

    def _cleanup(self):
        """Clean up Word application"""
        try:
            if self.doc:
                self.doc.Close(SaveChanges=False)
        except:
            pass
        try:
            if self.word:
                self.word.Quit()
        except:
            pass
        try:
            pythoncom.CoUninitialize()
        except:
            pass

    def _collect_all_equations(self):
        """Collect ALL equations from document including shapes"""

        print("\n" + "="*60)
        print("COLLECTING ALL EQUATIONS")
        print("="*60)

        self.all_equations = []

        # ============================================
        # STEP 1: Main document equations
        # ============================================
        print("\n[1] Main document equations...")
        try:
            main_count = self.doc.OMaths.Count
            print(f"    Found: {main_count}")

            for i in range(1, main_count + 1):
                eq = self.doc.OMaths.Item(i)
                self.all_equations.append({
                    'source': 'main',
                    'index': i,
                    'object': eq,
                    'range': eq.Range,
                    'text': eq.Range.Text or "",
                    'start': eq.Range.Start
                })
        except Exception as e:
            print(f"    Error: {e}")

        # ============================================
        # STEP 2: Shape/textbox equations
        # ============================================
        print("\n[2] Shape/textbox equations...")
        shape_eq_count = 0

        try:
            shape_count = self.doc.Shapes.Count
            print(f"    Total shapes: {shape_count}")

            for i in range(1, shape_count + 1):
                shape = self.doc.Shapes.Item(i)

                try:
                    if shape.TextFrame.HasText:
                        text_range = shape.TextFrame.TextRange
                        eq_count = text_range.OMaths.Count

                        if eq_count > 0:
                            print(f"    Shape '{shape.Name}': {eq_count} equations")

                            for j in range(1, eq_count + 1):
                                eq = text_range.OMaths.Item(j)
                                shape_eq_count += 1

                                self.all_equations.append({
                                    'source': 'shape',
                                    'shape_name': shape.Name,
                                    'index': j,
                                    'object': eq,
                                    'range': eq.Range,
                                    'text': eq.Range.Text or "",
                                    'start': eq.Range.Start,
                                    'shape_obj': shape
                                })
                except:
                    pass  # Shape doesn't have text frame

            print(f"    Total shape equations: {shape_eq_count}")

        except Exception as e:
            print(f"    Error: {e}")

        # Summary
        main_eqs = len([e for e in self.all_equations if e['source'] == 'main'])
        shape_eqs = len([e for e in self.all_equations if e['source'] == 'shape'])

        print("\n" + "-"*40)
        print(f"TOTAL EQUATIONS FOUND: {len(self.all_equations)}")
        print(f"  Main document: {main_eqs}")
        print(f"  In shapes:     {shape_eqs}")

        return len(self.all_equations)

    def _get_latex_from_xml(self, docx_path):
        """Extract LaTeX conversions from XML (for better LaTeX quality)"""

        print("\n" + "="*60)
        print("EXTRACTING LATEX FROM XML")
        print("="*60)

        # Import the OMML to LaTeX converter
        try:
            sys.path.insert(0, str(Path(__file__).parent.parent / "document-processing-api" / "backend"))
            from doc_processor.omml_2_latex import DirectOmmlToLatex
            latex_converter = DirectOmmlToLatex()
        except ImportError:
            print("Warning: Could not import LaTeX converter, using text fallback")
            latex_converter = None

        latex_map = {}

        with zipfile.ZipFile(docx_path, 'r') as z:
            with z.open('word/document.xml') as f:
                content = f.read()
                root = etree.fromstring(content)

                ns = {
                    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
                    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006'
                }

                # Get equations NOT in mc:Fallback (avoid duplicates)
                # Main body equations
                main_eqs = root.xpath('//m:oMath[not(ancestor::mc:Fallback)]', namespaces=ns)

                print(f"Found {len(main_eqs)} unique equations in XML")

                for i, eq in enumerate(main_eqs):
                    # Get text for matching
                    texts = eq.xpath('.//m:t/text()', namespaces=ns)
                    eq_text = ''.join(texts)

                    # Convert to LaTeX
                    if latex_converter:
                        try:
                            latex = latex_converter.parse(eq)
                        except:
                            latex = eq_text
                    else:
                        latex = eq_text

                    latex_map[i] = {
                        'text': eq_text,
                        'latex': latex
                    }

                    if i < 5:  # Show first 5
                        print(f"  {i+1}. '{eq_text[:20]}' -> '{latex[:30]}...'")

        print(f"\nExtracted {len(latex_map)} LaTeX conversions")
        return latex_map

    def _replace_all_equations(self, latex_map):
        """Replace all equations with LaTeX markers"""

        print("\n" + "="*60)
        print("REPLACING EQUATIONS WITH LATEX MARKERS")
        print("="*60)

        replaced_count = 0
        failed_count = 0

        # Separate main and shape equations
        main_eqs = [e for e in self.all_equations if e['source'] == 'main']
        shape_eqs = [e for e in self.all_equations if e['source'] == 'shape']

        # Sort by position descending (process from end to start)
        main_eqs.sort(key=lambda x: x['start'], reverse=True)

        # Group shape equations by shape, then sort within each shape
        shape_groups = {}
        for eq in shape_eqs:
            shape_name = eq['shape_name']
            if shape_name not in shape_groups:
                shape_groups[shape_name] = []
            shape_groups[shape_name].append(eq)

        for shape_name in shape_groups:
            shape_groups[shape_name].sort(key=lambda x: x['start'], reverse=True)

        # ============================================
        # Process shape equations first
        # ============================================
        print("\n[1] Processing shape equations...")

        for shape_name, eqs in shape_groups.items():
            print(f"    Shape '{shape_name}': {len(eqs)} equations")

            for eq_data in eqs:
                try:
                    eq_range = eq_data['range']
                    eq_text = eq_data['text'].strip()

                    # Get LaTeX (use text as fallback)
                    latex_text = eq_text or "equation"

                    # Create marked text
                    is_inline = len(latex_text) < 30
                    if is_inline:
                        marked_text = f' MATHSTARTINLINE\\({latex_text}\\)MATHENDINLINE '
                    else:
                        marked_text = f' MATHSTARTDISPLAY\\[{latex_text}\\]MATHENDDISPLAY '

                    # Replace
                    eq_range.Delete()
                    eq_range.InsertAfter(marked_text)

                    replaced_count += 1

                except Exception as e:
                    print(f"      Failed: {e}")
                    failed_count += 1

        # ============================================
        # Process main document equations
        # ============================================
        print("\n[2] Processing main document equations...")

        for i, eq_data in enumerate(main_eqs):
            try:
                eq_range = eq_data['range']
                eq_text = eq_data['text'].strip()

                # Try to get better LaTeX from map
                # Match by original index (before sorting)
                orig_idx = eq_data['index'] - 1
                if orig_idx in latex_map:
                    latex_text = latex_map[orig_idx]['latex']
                else:
                    latex_text = eq_text or "equation"

                # Create marked text
                is_inline = len(latex_text) < 50
                if is_inline:
                    marked_text = f' MATHSTARTINLINE\\({latex_text}\\)MATHENDINLINE '
                else:
                    marked_text = f' MATHSTARTDISPLAY\\[{latex_text}\\]MATHENDDISPLAY '

                # Replace
                eq_range.Delete()
                eq_range.InsertAfter(marked_text)

                replaced_count += 1

            except Exception as e:
                print(f"    Failed eq {eq_data['index']}: {e}")
                failed_count += 1

        print("\n" + "-"*40)
        print(f"REPLACEMENT RESULTS:")
        print(f"  Successfully replaced: {replaced_count}")
        print(f"  Failed:                {failed_count}")

        return replaced_count, failed_count

    def convert_document(self, input_path, output_path=None):
        """
        Main entry point - convert all equations in document

        Args:
            input_path: Path to input .docx file
            output_path: Path for output file (optional)

        Returns:
            dict with conversion results
        """

        input_path = Path(input_path).absolute()

        if not output_path:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_path = input_path.parent / f"{input_path.stem}_converted_{timestamp}.docx"
        else:
            output_path = Path(output_path).absolute()

        print("\n" + "="*70)
        print("COMPLETE WORD COM EQUATION CONVERTER")
        print("="*70)
        print(f"Input:  {input_path}")
        print(f"Output: {output_path}")
        print("="*70)

        try:
            # Start Word
            print("\nStarting Word application...")
            self._start_word()

            # Open document
            print("Opening document...")
            self.doc = self.word.Documents.Open(str(input_path))

            # Check for tracked changes
            if self.doc.TrackRevisions or self.doc.Revisions.Count > 0:
                print("\n WARNING: Document has tracked changes!")
                print("Please accept all changes before processing.")
                return {
                    'error': 'Document has tracked changes',
                    'success': False
                }

            # Step 1: Extract LaTeX from XML
            latex_map = self._get_latex_from_xml(input_path)

            # Step 2: Collect all equations via COM
            total_found = self._collect_all_equations()

            if total_found == 0:
                print("\nNo equations found in document.")
                return {
                    'success': True,
                    'equations_found': 0,
                    'equations_replaced': 0
                }

            # Step 3: Replace all equations
            replaced, failed = self._replace_all_equations(latex_map)

            # Step 4: Save document
            print(f"\nSaving document to: {output_path}")
            self.doc.SaveAs2(str(output_path))

            # Results
            print("\n" + "="*70)
            print("CONVERSION COMPLETE!")
            print("="*70)
            print(f"Equations found:    {total_found}")
            print(f"Equations replaced: {replaced}")
            print(f"Failed:             {failed}")
            print(f"Success rate:       {replaced/total_found*100:.1f}%")
            print(f"\nOutput: {output_path}")

            return {
                'success': True,
                'output_path': str(output_path),
                'equations_found': total_found,
                'equations_replaced': replaced,
                'equations_failed': failed
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
            self._cleanup()


def verify_conversion(original_path, converted_path):
    """Verify conversion by checking for remaining OMML and counting markers"""

    print("\n" + "="*70)
    print("VERIFYING CONVERSION")
    print("="*70)

    # Check converted document
    with zipfile.ZipFile(converted_path, 'r') as z:
        with z.open('word/document.xml') as f:
            content = f.read()
            root = etree.fromstring(content)

            ns = {'m': 'http://schemas.openxmlformats.org/officeDocument/2006/math'}

            # Count remaining OMML
            remaining = root.xpath('//m:oMath', namespaces=ns)

            # Count markers in text
            text_content = etree.tostring(root, encoding='unicode')
            inline_markers = text_content.count('MATHSTARTINLINE')
            display_markers = text_content.count('MATHSTARTDISPLAY')

    print(f"Remaining OMML equations: {len(remaining)}")
    print(f"LaTeX markers added:")
    print(f"  Inline:  {inline_markers}")
    print(f"  Display: {display_markers}")
    print(f"  Total:   {inline_markers + display_markers}")

    if len(remaining) == 0:
        print("\n SUCCESS: All equations converted!")
    else:
        print(f"\n WARNING: {len(remaining)} equations still remain as OMML")

    return {
        'remaining_omml': len(remaining),
        'inline_markers': inline_markers,
        'display_markers': display_markers,
        'total_markers': inline_markers + display_markers
    }


if __name__ == "__main__":
    # Test with document that has shape equations
    test_doc = Path(r"D:\Development\document-processing-api-2\document-processing-api\test docs\الدالة واحد لواحد (جاهزة للنشر) - Copy.docx")
    output_doc = Path(r"D:\Development\document-processing-api-2\backend\test_output\complete_converted.docx")

    # Create output directory
    output_doc.parent.mkdir(exist_ok=True)

    if test_doc.exists():
        converter = CompleteWordCOMConverter()
        result = converter.convert_document(test_doc, output_doc)

        if result.get('success'):
            # Verify the conversion
            verify_result = verify_conversion(test_doc, output_doc)

            print("\n" + "="*70)
            print("FINAL SUMMARY")
            print("="*70)
            print(f"Original document: {test_doc.name}")
            print(f"Equations found:   {result['equations_found']}")
            print(f"Equations replaced:{result['equations_replaced']}")
            print(f"Markers created:   {verify_result['total_markers']}")
            print(f"Remaining OMML:    {verify_result['remaining_omml']}")
    else:
        print(f"Test document not found: {test_doc}")
