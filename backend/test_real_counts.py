"""
Test to get REAL equation counts and verify conversion
"""

import sys
import io
import zipfile
from pathlib import Path
from lxml import etree

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def count_equations_accurately(docx_path):
    """Get accurate equation counts"""

    with zipfile.ZipFile(docx_path, 'r') as z:
        with z.open('word/document.xml') as xml_file:
            xml_content = xml_file.read()

    root = etree.fromstring(xml_content)

    namespaces = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
        'v': 'urn:schemas-microsoft-com:vml',
        'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006'
    }

    # Find ALL equations
    all_equations = root.xpath('//m:oMath', namespaces=namespaces)
    print(f"\nTotal OMML equations: {len(all_equations)}")

    # Check different VML contexts
    in_vml_textbox = 0
    in_vml_any = 0
    in_fallback = 0
    in_txbx = 0
    regular = 0

    for eq in all_equations:
        has_vml_textbox = bool(eq.xpath('ancestor::v:textbox', namespaces=namespaces))
        has_vml_any = bool(eq.xpath('ancestor::v:*', namespaces=namespaces))
        has_fallback = bool(eq.xpath('ancestor::mc:Fallback', namespaces=namespaces))
        has_txbx = bool(eq.xpath('ancestor::w:txbxContent', namespaces=namespaces))

        if has_vml_textbox:
            in_vml_textbox += 1
        elif has_vml_any:
            in_vml_any += 1
        elif has_fallback:
            in_fallback += 1
        elif has_txbx:
            in_txbx += 1
        else:
            regular += 1

    print(f"\nEquation locations:")
    print(f"  Regular (safely convertible): {regular}")
    print(f"  In VML textbox: {in_vml_textbox}")
    print(f"  In other VML elements: {in_vml_any}")
    print(f"  In mc:Fallback: {in_fallback}")
    print(f"  In w:txbxContent: {in_txbx}")

    print(f"\nSUMMARY:")
    print(f"  ✅ Can safely convert: {regular}")
    print(f"  ⚠️ Cannot convert (VML/Fallback/txbx): {in_vml_textbox + in_vml_any + in_fallback + in_txbx}")

    return {
        'total': len(all_equations),
        'regular': regular,
        'vml_textbox': in_vml_textbox,
        'vml_other': in_vml_any,
        'fallback': in_fallback,
        'txbx': in_txbx
    }


def test_both_documents():
    """Test both documents to get accurate counts"""

    doc1 = Path(r"D:\Development\document-processing-api-2\document-processing-api\test docs\التشابه (جاهزة للنشر) - Copy.docx")
    doc2 = Path(r"D:\Development\document-processing-api-2\document-processing-api\test docs\الدالة واحد لواحد (جاهزة للنشر) - Copy.docx")

    print("="*60)
    print("ACCURATE EQUATION COUNTS")
    print("="*60)

    if doc1.exists():
        print(f"\n📄 Document 1: {doc1.name}")
        results1 = count_equations_accurately(doc1)

    if doc2.exists():
        print(f"\n📄 Document 2: {doc2.name}")
        results2 = count_equations_accurately(doc2)

    # Test converted files
    print("\n" + "="*60)
    print("CHECKING CONVERTED FILES")
    print("="*60)

    converted1 = Path("safe_output_89_equations.docx")
    converted2 = Path("safe_output_70_of_144_equations.docx")

    if converted1.exists():
        print(f"\n📄 Converted 1: {converted1.name}")
        conv_results1 = count_equations_accurately(converted1)

    if converted2.exists():
        print(f"\n📄 Converted 2: {converted2.name}")
        conv_results2 = count_equations_accurately(converted2)

        print(f"\n✅ Conversion success:")
        print(f"  Original regular equations: {results2['regular']}")
        print(f"  Remaining after conversion: {conv_results2['total']}")
        print(f"  Equations converted: {results2['regular'] - conv_results2['total']}")


if __name__ == "__main__":
    test_both_documents()