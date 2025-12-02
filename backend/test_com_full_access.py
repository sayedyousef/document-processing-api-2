"""
Test Word COM FULL access to ALL equations including VML textboxes
Uses StoryRanges and Shapes.TextFrame to access textbox equations
"""

import sys
import io
import win32com.client
import pythoncom
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


def get_all_equations_via_com(docx_path):
    """Get ALL equations using COM - including textbox equations"""

    print(f"\n{'='*70}")
    print("COLLECTING ALL EQUATIONS VIA WORD COM")
    print(f"{'='*70}")
    print(f"Document: {docx_path}\n")

    pythoncom.CoInitialize()
    word = None
    doc = None

    try:
        # Start Word
        print("Starting Word application...")
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False

        # Open document
        print("Opening document...")
        doc = word.Documents.Open(str(docx_path))
        print("Document opened successfully\n")

        all_equations = []

        # ============================================
        # STEP 1: Get equations from main document
        # ============================================
        print("="*50)
        print("STEP 1: Main Document Equations")
        print("="*50)

        try:
            main_count = doc.OMaths.Count
            print(f"Found {main_count} equations in main document")

            for i in range(1, main_count + 1):
                eq = doc.OMaths.Item(i)
                eq_data = {
                    'source': 'main_document',
                    'index': i,
                    'range_start': eq.Range.Start,
                    'range_end': eq.Range.End,
                    'text': eq.Range.Text or "",
                    'equation_object': eq,
                    'parent_range': doc.Content
                }
                all_equations.append(eq_data)

        except Exception as e:
            print(f"Error: {e}")

        # ============================================
        # STEP 2: Get equations from Shapes (textboxes)
        # ============================================
        print("\n" + "="*50)
        print("STEP 2: Shape/Textbox Equations")
        print("="*50)

        try:
            shape_count = doc.Shapes.Count
            print(f"Found {shape_count} shapes in document")

            shapes_with_equations = 0
            for i in range(1, shape_count + 1):
                shape = doc.Shapes.Item(i)

                try:
                    if shape.TextFrame.HasText:
                        text_range = shape.TextFrame.TextRange
                        eq_count = text_range.OMaths.Count

                        if eq_count > 0:
                            shapes_with_equations += 1
                            print(f"\n  Shape '{shape.Name}': {eq_count} equations")

                            for j in range(1, eq_count + 1):
                                eq = text_range.OMaths.Item(j)
                                eq_data = {
                                    'source': f'shape_{shape.Name}',
                                    'index': j,
                                    'range_start': eq.Range.Start,
                                    'range_end': eq.Range.End,
                                    'text': eq.Range.Text or "",
                                    'equation_object': eq,
                                    'parent_range': text_range,
                                    'shape': shape
                                }
                                all_equations.append(eq_data)
                                print(f"    Equation {j}: {eq.Range.Text[:30] if eq.Range.Text else '[empty]'}...")

                except Exception as e:
                    pass  # Shape doesn't have text frame

            print(f"\nTotal shapes with equations: {shapes_with_equations}")

        except Exception as e:
            print(f"Error: {e}")

        # ============================================
        # SUMMARY
        # ============================================
        print("\n" + "="*70)
        print("EQUATION COLLECTION SUMMARY")
        print("="*70)

        main_eqs = [e for e in all_equations if e['source'] == 'main_document']
        shape_eqs = [e for e in all_equations if e['source'].startswith('shape_')]

        print(f"Main document equations: {len(main_eqs)}")
        print(f"Shape/textbox equations: {len(shape_eqs)}")
        print(f"TOTAL EQUATIONS:         {len(all_equations)}")

        return all_equations, doc, word

    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()

        if doc:
            try:
                doc.Close(SaveChanges=False)
            except:
                pass
        if word:
            try:
                word.Quit()
            except:
                pass
        pythoncom.CoUninitialize()
        return None, None, None


def replace_all_equations(all_equations, doc, output_path):
    """Replace ALL equations with LaTeX markers"""

    print(f"\n{'='*70}")
    print("REPLACING ALL EQUATIONS WITH LATEX MARKERS")
    print(f"{'='*70}\n")

    # Import LaTeX converter
    sys.path.insert(0, str(Path(__file__).parent.parent / "document-processing-api" / "backend"))
    from doc_processor.omml_2_latex import DirectOmmlToLatex

    latex_converter = DirectOmmlToLatex()

    replaced_count = 0
    failed_count = 0

    # Process equations in REVERSE order (to maintain positions)
    # First sort by source (shapes first, then main), then by position descending

    # Separate main and shape equations
    main_eqs = [e for e in all_equations if e['source'] == 'main_document']
    shape_eqs = [e for e in all_equations if e['source'].startswith('shape_')]

    # Sort each group by position descending
    main_eqs.sort(key=lambda x: x['range_start'], reverse=True)
    shape_eqs.sort(key=lambda x: (x['source'], x['range_start']), reverse=True)

    # Process shape equations first (they're independent of main document)
    print("Processing shape/textbox equations...")
    for eq_data in shape_eqs:
        try:
            eq = eq_data['equation_object']
            eq_range = eq.Range

            # Get LaTeX (simplified - just use text for now)
            latex_text = eq_data['text'].strip() or "equation"

            # Determine if inline or display
            is_inline = len(latex_text) < 30

            if is_inline:
                marked_text = f' MATHSTARTINLINE\\({latex_text}\\)MATHENDINLINE '
            else:
                marked_text = f' MATHSTARTDISPLAY\\[{latex_text}\\]MATHENDDISPLAY '

            # Replace
            eq_range.Delete()
            eq_range.InsertAfter(marked_text)

            replaced_count += 1
            print(f"  Replaced: {eq_data['source']} eq {eq_data['index']}")

        except Exception as e:
            print(f"  Failed: {eq_data['source']} eq {eq_data['index']} - {e}")
            failed_count += 1

    # Process main document equations
    print("\nProcessing main document equations...")
    for eq_data in main_eqs:
        try:
            eq = eq_data['equation_object']
            eq_range = eq.Range

            # Get LaTeX
            latex_text = eq_data['text'].strip() or "equation"

            # Determine if inline or display
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
            print(f"  Failed: main eq {eq_data['index']} - {e}")
            failed_count += 1

    print(f"\n{'='*50}")
    print(f"REPLACEMENT RESULTS")
    print(f"{'='*50}")
    print(f"Successfully replaced: {replaced_count}")
    print(f"Failed:                {failed_count}")
    print(f"Total:                 {len(all_equations)}")

    # Save document
    print(f"\nSaving to: {output_path}")
    doc.SaveAs2(str(output_path))
    print("Document saved!")

    return replaced_count, failed_count


def test_full_conversion():
    """Test full conversion including textbox equations"""

    test_doc = Path(r"D:\Development\document-processing-api-2\document-processing-api\test docs\الدالة واحد لواحد (جاهزة للنشر) - Copy.docx")
    output_doc = Path(r"D:\Development\document-processing-api-2\backend\output_all_144_equations.docx")

    if not test_doc.exists():
        print(f"Test document not found: {test_doc}")
        return

    # Step 1: Collect all equations
    all_equations, doc, word = get_all_equations_via_com(test_doc)

    if not all_equations or not doc:
        print("Failed to collect equations")
        return

    try:
        # Step 2: Replace all equations
        replaced, failed = replace_all_equations(all_equations, doc, output_doc)

        print(f"\n{'='*70}")
        print("FINAL RESULTS")
        print(f"{'='*70}")
        print(f"Total equations in document: 144 (expected)")
        print(f"Total equations found:       {len(all_equations)}")
        print(f"Successfully replaced:       {replaced}")
        print(f"Failed:                      {failed}")

        if len(all_equations) == 144 and replaced == 144:
            print("\n SUCCESS! All 144 equations converted!")
        elif len(all_equations) >= 144:
            print(f"\n Found all equations! {replaced}/144 replaced")
        else:
            print(f"\n PARTIAL: Only found {len(all_equations)}/144 equations")

    finally:
        # Cleanup
        if doc:
            try:
                doc.Close(SaveChanges=False)
            except:
                pass
        if word:
            try:
                word.Quit()
            except:
                pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    test_full_conversion()
