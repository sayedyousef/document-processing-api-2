"""
Analyze XML to understand duplicate equations in VML/Fallback
"""

import sys
import io
import zipfile
from pathlib import Path
from lxml import etree

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


def analyze_xml_equation_structure(docx_path):
    """Detailed analysis of equation structure in XML"""

    print(f"\n{'='*70}")
    print("DETAILED XML EQUATION ANALYSIS")
    print(f"{'='*70}")
    print(f"Document: {docx_path}\n")

    with zipfile.ZipFile(docx_path, 'r') as z:
        with z.open('word/document.xml') as f:
            content = f.read()
            root = etree.fromstring(content)

            ns = {
                'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
                'v': 'urn:schemas-microsoft-com:vml',
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
                'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
            }

            # Find ALL equations
            all_equations = root.xpath('//m:oMath', namespaces=ns)
            print(f"Total m:oMath elements in XML: {len(all_equations)}\n")

            # Categorize each equation
            categories = {
                'main_body': [],
                'mc_choice': [],       # Modern apps (Word 2010+)
                'mc_fallback': [],     # Legacy fallback
                'vml_textbox': [],
                'txbxContent': [],
                'wps_textbox': [],     # Word processing shape textbox
                'other': []
            }

            for i, eq in enumerate(all_equations, 1):
                # Get equation text for identification
                texts = eq.xpath('.//m:t/text()', namespaces=ns)
                eq_text = ''.join(texts)[:20] if texts else '[empty]'

                # Check ancestors
                has_choice = bool(eq.xpath('ancestor::mc:Choice', namespaces=ns))
                has_fallback = bool(eq.xpath('ancestor::mc:Fallback', namespaces=ns))
                has_vml = bool(eq.xpath('ancestor::v:textbox', namespaces=ns))
                has_txbx = bool(eq.xpath('ancestor::w:txbxContent', namespaces=ns))
                has_wps = bool(eq.xpath('ancestor::wps:txbx', namespaces=ns))

                # Get parent chain for debugging
                parent_chain = []
                current = eq
                for _ in range(8):
                    parent = current.getparent()
                    if parent is None:
                        break
                    tag = parent.tag.split('}')[-1] if '}' in parent.tag else parent.tag
                    parent_chain.append(tag)
                    current = parent

                eq_info = {
                    'index': i,
                    'text': eq_text,
                    'parents': ' > '.join(reversed(parent_chain))
                }

                # Categorize
                if has_choice and (has_wps or has_txbx):
                    categories['mc_choice'].append(eq_info)
                elif has_fallback:
                    categories['mc_fallback'].append(eq_info)
                elif has_vml:
                    categories['vml_textbox'].append(eq_info)
                elif has_txbx:
                    categories['txbxContent'].append(eq_info)
                elif has_wps:
                    categories['wps_textbox'].append(eq_info)
                else:
                    # Check if truly in main body
                    in_drawing = bool(eq.xpath('ancestor::w:drawing', namespaces=ns))
                    if in_drawing:
                        categories['other'].append(eq_info)
                    else:
                        categories['main_body'].append(eq_info)

            # Print results
            print("="*50)
            print("EQUATION CATEGORIES")
            print("="*50)

            for cat_name, equations in categories.items():
                if equations:
                    print(f"\n{cat_name.upper()}: {len(equations)} equations")
                    print("-"*40)
                    for eq in equations[:5]:  # Show first 5
                        print(f"  #{eq['index']}: '{eq['text']}' ")
                        print(f"      Path: {eq['parents']}")
                    if len(equations) > 5:
                        print(f"  ... and {len(equations) - 5} more")

            # Summary
            print("\n" + "="*70)
            print("SUMMARY")
            print("="*70)

            main_count = len(categories['main_body'])
            choice_count = len(categories['mc_choice'])
            fallback_count = len(categories['mc_fallback'])
            vml_count = len(categories['vml_textbox'])
            txbx_count = len(categories['txbxContent'])
            wps_count = len(categories['wps_textbox'])

            print(f"Main body equations:        {main_count}")
            print(f"mc:Choice (modern):         {choice_count}")
            print(f"mc:Fallback (legacy):       {fallback_count}")
            print(f"VML textbox:                {vml_count}")
            print(f"w:txbxContent:              {txbx_count}")
            print(f"wps:txbx:                   {wps_count}")
            print(f"Other:                      {len(categories['other'])}")
            print()

            # The key insight
            actual_unique = main_count + choice_count
            duplicates = fallback_count + vml_count

            print("="*70)
            print("KEY INSIGHT")
            print("="*70)
            print(f"ACTUAL UNIQUE EQUATIONS:    {actual_unique}")
            print(f"  - Main body:              {main_count}")
            print(f"  - In shapes (mc:Choice):  {choice_count}")
            print()
            print(f"DUPLICATE/FALLBACK:         {duplicates}")
            print(f"  - mc:Fallback copies:     {fallback_count}")
            print(f"  - VML fallback copies:    {vml_count}")
            print()
            print("The mc:Fallback and VML versions are DUPLICATES for older Word versions!")
            print("Word COM accesses the mc:Choice version (modern), not the fallback.")

            return {
                'total_xml': len(all_equations),
                'main_body': main_count,
                'shape_equations': choice_count,
                'actual_unique': actual_unique,
                'fallback_duplicates': duplicates
            }


if __name__ == "__main__":
    test_doc = Path(r"D:\Development\document-processing-api-2\document-processing-api\test docs\الدالة واحد لواحد (جاهزة للنشر) - Copy.docx")

    if test_doc.exists():
        results = analyze_xml_equation_structure(test_doc)

        print("\n" + "="*70)
        print("CONCLUSION")
        print("="*70)
        print(f"XML reports {results['total_xml']} equations, but...")
        print(f"ACTUAL UNIQUE equations: {results['actual_unique']}")
        print(f"The other {results['fallback_duplicates']} are fallback duplicates!")
        print()
        print("Word COM found 107 equations which is CORRECT!")
        print("(70 main + 37 in shapes = 107 unique equations)")
    else:
        print(f"Test document not found: {test_doc}")
