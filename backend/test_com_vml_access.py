"""
Test Word COM access to VML textbox equations
Investigates different methods to access equations in textboxes/shapes
"""

import sys
import io
import win32com.client
import pythoncom
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def test_com_vml_access(docx_path):
    """Test various COM methods to access VML textbox equations"""

    print(f"\n{'='*70}")
    print("TESTING WORD COM ACCESS TO VML TEXTBOX EQUATIONS")
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

        results = {
            'document_omaths': 0,
            'story_ranges_omaths': 0,
            'shapes_omaths': 0,
            'inline_shapes_omaths': 0,
            'textboxes_omaths': 0,
            'headers_footers_omaths': 0,
            'total_unique': 0
        }

        all_positions = set()

        # ============================================
        # METHOD 1: Direct document.OMaths
        # ============================================
        print("="*50)
        print("METHOD 1: document.OMaths")
        print("="*50)
        try:
            count = doc.OMaths.Count
            results['document_omaths'] = count
            print(f"Found {count} equations via document.OMaths")

            for i in range(1, min(count + 1, 10)):  # Show first 10
                eq = doc.OMaths.Item(i)
                pos = eq.Range.Start
                all_positions.add(pos)
                text = eq.Range.Text[:50] if eq.Range.Text else "[empty]"
                print(f"  Equation {i}: pos={pos}, text={text}...")
        except Exception as e:
            print(f"Error: {e}")

        # ============================================
        # METHOD 2: StoryRanges (includes textboxes!)
        # ============================================
        print("\n" + "="*50)
        print("METHOD 2: StoryRanges (may include textboxes)")
        print("="*50)
        try:
            story_count = 0
            story_types = {
                1: "wdMainTextStory",
                2: "wdFootnotesStory",
                3: "wdEndnotesStory",
                4: "wdCommentsStory",
                5: "wdTextFrameStory",  # This is for textboxes!
                6: "wdEvenPagesHeaderStory",
                7: "wdPrimaryHeaderStory",
                8: "wdEvenPagesFooterStory",
                9: "wdPrimaryFooterStory",
                10: "wdFirstPageHeaderStory",
                11: "wdFirstPageFooterStory"
            }

            for story in doc.StoryRanges:
                story_type = story.StoryType
                story_name = story_types.get(story_type, f"Unknown({story_type})")

                current_story = story
                while current_story:
                    try:
                        eq_count = current_story.OMaths.Count
                        if eq_count > 0:
                            print(f"\n  {story_name}: {eq_count} equations")
                            story_count += eq_count

                            for i in range(1, min(eq_count + 1, 5)):
                                eq = current_story.OMaths.Item(i)
                                pos = eq.Range.Start
                                all_positions.add(pos)
                                text = eq.Range.Text[:40] if eq.Range.Text else "[empty]"
                                print(f"    Equation {i}: pos={pos}, text={text}...")
                    except:
                        pass

                    try:
                        current_story = current_story.NextStoryRange
                    except:
                        break

            results['story_ranges_omaths'] = story_count
            print(f"\nTotal via StoryRanges: {story_count}")
        except Exception as e:
            print(f"Error: {e}")

        # ============================================
        # METHOD 3: Shapes (VML shapes including textboxes)
        # ============================================
        print("\n" + "="*50)
        print("METHOD 3: Shapes collection")
        print("="*50)
        try:
            shape_count = doc.Shapes.Count
            print(f"Found {shape_count} shapes in document")

            shapes_with_equations = 0
            for i in range(1, shape_count + 1):
                shape = doc.Shapes.Item(i)
                shape_type = shape.Type
                shape_name = shape.Name

                # Check if shape has text frame
                try:
                    if shape.TextFrame.HasText:
                        text_range = shape.TextFrame.TextRange
                        eq_count = text_range.OMaths.Count

                        if eq_count > 0:
                            shapes_with_equations += eq_count
                            print(f"\n  Shape '{shape_name}' (type={shape_type}): {eq_count} equations")

                            for j in range(1, min(eq_count + 1, 3)):
                                eq = text_range.OMaths.Item(j)
                                pos = eq.Range.Start
                                all_positions.add(pos)
                                text = eq.Range.Text[:40] if eq.Range.Text else "[empty]"
                                print(f"    Equation {j}: pos={pos}, text={text}...")
                except:
                    pass

            results['shapes_omaths'] = shapes_with_equations
            print(f"\nTotal equations in Shapes: {shapes_with_equations}")
        except Exception as e:
            print(f"Error: {e}")

        # ============================================
        # METHOD 4: InlineShapes
        # ============================================
        print("\n" + "="*50)
        print("METHOD 4: InlineShapes collection")
        print("="*50)
        try:
            inline_count = doc.InlineShapes.Count
            print(f"Found {inline_count} inline shapes")

            inline_with_equations = 0
            for i in range(1, inline_count + 1):
                ishape = doc.InlineShapes.Item(i)

                try:
                    if ishape.HasChart:
                        continue

                    # Try to get text frame
                    try:
                        text_range = ishape.Range
                        eq_count = text_range.OMaths.Count
                        if eq_count > 0:
                            inline_with_equations += eq_count
                            print(f"  InlineShape {i}: {eq_count} equations")
                    except:
                        pass
                except:
                    pass

            results['inline_shapes_omaths'] = inline_with_equations
            print(f"Total equations in InlineShapes: {inline_with_equations}")
        except Exception as e:
            print(f"Error: {e}")

        # ============================================
        # METHOD 5: Find equations via Range.Find
        # ============================================
        print("\n" + "="*50)
        print("METHOD 5: Using Range.Find for equations")
        print("="*50)
        try:
            # Try to find all equation fields
            rng = doc.Content
            find_count = 0

            # Reset find
            rng.Find.ClearFormatting()

            # This might find equation placeholders
            print("Searching for equation patterns...")

        except Exception as e:
            print(f"Error: {e}")

        # ============================================
        # METHOD 6: ContentControls
        # ============================================
        print("\n" + "="*50)
        print("METHOD 6: ContentControls")
        print("="*50)
        try:
            cc_count = doc.ContentControls.Count
            print(f"Found {cc_count} content controls")
        except Exception as e:
            print(f"Error: {e}")

        # ============================================
        # SUMMARY
        # ============================================
        results['total_unique'] = len(all_positions)

        print("\n" + "="*70)
        print("SUMMARY")
        print("="*70)
        print(f"document.OMaths:     {results['document_omaths']}")
        print(f"StoryRanges:         {results['story_ranges_omaths']}")
        print(f"Shapes (textboxes):  {results['shapes_omaths']}")
        print(f"InlineShapes:        {results['inline_shapes_omaths']}")
        print(f"")
        print(f"TOTAL UNIQUE POSITIONS: {results['total_unique']}")
        print("="*70)

        return results

    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
        return None

    finally:
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


def compare_with_xml_count(docx_path):
    """Compare COM access with XML equation count"""

    import zipfile
    from lxml import etree

    print(f"\n{'='*70}")
    print("XML EQUATION COUNT (for comparison)")
    print(f"{'='*70}")

    with zipfile.ZipFile(docx_path, 'r') as z:
        with z.open('word/document.xml') as f:
            content = f.read()
            root = etree.fromstring(content)

            ns = {
                'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
                'v': 'urn:schemas-microsoft-com:vml',
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006'
            }

            all_equations = root.xpath('//m:oMath', namespaces=ns)
            print(f"Total OMML equations in XML: {len(all_equations)}")

            # Categorize
            regular = 0
            in_vml = 0
            in_txbx = 0
            in_fallback = 0

            for eq in all_equations:
                has_vml = bool(eq.xpath('ancestor::v:*', namespaces=ns))
                has_txbx = bool(eq.xpath('ancestor::w:txbxContent', namespaces=ns))
                has_fallback = bool(eq.xpath('ancestor::mc:Fallback', namespaces=ns))

                if has_vml or has_txbx or has_fallback:
                    if has_vml:
                        in_vml += 1
                    elif has_txbx:
                        in_txbx += 1
                    else:
                        in_fallback += 1
                else:
                    regular += 1

            print(f"\nBreakdown:")
            print(f"  Regular equations:     {regular}")
            print(f"  In VML elements:       {in_vml}")
            print(f"  In w:txbxContent:      {in_txbx}")
            print(f"  In mc:Fallback:        {in_fallback}")
            print(f"  ---")
            print(f"  Total in textboxes:    {in_vml + in_txbx + in_fallback}")

            return {
                'total': len(all_equations),
                'regular': regular,
                'in_textboxes': in_vml + in_txbx + in_fallback
            }


if __name__ == "__main__":
    # Test with the document that has VML textbox equations
    test_doc = Path(r"D:\Development\document-processing-api-2\document-processing-api\test docs\الدالة واحد لواحد (جاهزة للنشر) - Copy.docx")

    if test_doc.exists():
        # First get XML count
        xml_results = compare_with_xml_count(test_doc)

        # Then test COM access
        com_results = test_com_vml_access(test_doc)

        if com_results and xml_results:
            print(f"\n{'='*70}")
            print("FINAL COMPARISON")
            print(f"{'='*70}")
            print(f"XML Total:          {xml_results['total']}")
            print(f"XML Regular:        {xml_results['regular']}")
            print(f"XML in Textboxes:   {xml_results['in_textboxes']}")
            print(f"")
            print(f"COM Total Found:    {com_results['total_unique']}")
            print(f"")

            if com_results['total_unique'] >= xml_results['total']:
                print("SUCCESS: COM can access ALL equations!")
            elif com_results['total_unique'] > xml_results['regular']:
                missing = xml_results['total'] - com_results['total_unique']
                print(f"PARTIAL: COM found more than regular, but missing {missing}")
            else:
                missing = xml_results['total'] - com_results['total_unique']
                print(f"ISSUE: COM cannot access {missing} textbox equations")
    else:
        print(f"Test document not found: {test_doc}")
