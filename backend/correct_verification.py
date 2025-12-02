"""
CORRECT VERIFICATION: Count LaTeX brackets in actual text content
This is the only reliable way to verify equation conversion
"""

import sys
import io
import zipfile
import re
from pathlib import Path
from lxml import etree

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


def extract_all_text_from_docx(docx_path):
    """Extract ALL text content from Word document"""
    with zipfile.ZipFile(docx_path, 'r') as z:
        with z.open('word/document.xml') as xml_file:
            xml_content = xml_file.read()

    # Parse XML
    root = etree.fromstring(xml_content)

    # Extract all text from w:t elements
    namespaces = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    }

    text_elements = root.xpath('//w:t', namespaces=namespaces)
    all_text = []

    for elem in text_elements:
        if elem.text:
            all_text.append(elem.text)

    return ' '.join(all_text)


def count_latex_equations_in_text(text):
    """Count LaTeX equations by counting brackets in text"""

    # Count inline equations: \( ... \)
    inline_open = text.count('\\(')
    inline_close = text.count('\\)')

    # Count display equations: \[ ... \]
    display_open = text.count('\\[')
    display_close = text.count('\\]')

    # Also count with markers
    inline_markers = text.count('MATHSTARTINLINE')
    display_markers = text.count('MATHSTARTDISPLAY')

    results = {
        'inline_equations': inline_open,  # Use opening brackets as count
        'display_equations': display_open,
        'total_equations': inline_open + display_open,
        'inline_brackets': {
            'open': inline_open,
            'close': inline_close,
            'matched': inline_open == inline_close
        },
        'display_brackets': {
            'open': display_open,
            'close': display_close,
            'matched': display_open == display_close
        },
        'markers': {
            'inline': inline_markers,
            'display': display_markers
        }
    }

    return results


def check_omml_equations(docx_path):
    """Check if OMML equations still exist"""
    with zipfile.ZipFile(docx_path, 'r') as z:
        with z.open('word/document.xml') as xml_file:
            xml_content = xml_file.read()

    # Parse XML
    root = etree.fromstring(xml_content)

    namespaces = {
        'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math'
    }

    omml_equations = root.xpath('//m:oMath', namespaces=namespaces)

    return len(omml_equations)


def verify_conversion(original_path, converted_path, expected_count=None):
    """Properly verify equation conversion"""

    print("\n" + "="*60)
    print("CORRECT VERIFICATION METHOD")
    print("="*60)

    # 1. Check original document
    print(f"\n📄 ORIGINAL: {Path(original_path).name}")
    print("-" * 40)

    original_text = extract_all_text_from_docx(original_path)
    original_latex = count_latex_equations_in_text(original_text)
    original_omml = check_omml_equations(original_path)

    print(f"OMML equations: {original_omml}")
    print(f"LaTeX equations in text: {original_latex['total_equations']}")
    print(f"  - Inline \\(...\\): {original_latex['inline_equations']}")
    print(f"  - Display \\[...\\]: {original_latex['display_equations']}")

    # 2. Check converted document
    print(f"\n📄 CONVERTED: {Path(converted_path).name}")
    print("-" * 40)

    converted_text = extract_all_text_from_docx(converted_path)
    converted_latex = count_latex_equations_in_text(converted_text)
    converted_omml = check_omml_equations(converted_path)

    print(f"OMML equations remaining: {converted_omml}")
    print(f"LaTeX equations in text: {converted_latex['total_equations']}")
    print(f"  - Inline \\(...\\): {converted_latex['inline_equations']}")
    print(f"  - Display \\[...\\]: {converted_latex['display_equations']}")

    # Check bracket matching
    if not converted_latex['inline_brackets']['matched']:
        print("⚠️ Warning: Inline brackets don't match!")
    if not converted_latex['display_brackets']['matched']:
        print("⚠️ Warning: Display brackets don't match!")

    # 3. Show markers
    if converted_latex['markers']['inline'] or converted_latex['markers']['display']:
        print(f"\n📍 Markers found:")
        print(f"  - MATHSTARTINLINE: {converted_latex['markers']['inline']}")
        print(f"  - MATHSTARTDISPLAY: {converted_latex['markers']['display']}")

    # 4. Verification results
    print(f"\n✅ VERIFICATION RESULTS:")
    print("-" * 40)

    success = True

    # Check OMML was removed
    if converted_omml == 0:
        print(f"✅ All OMML equations removed (was {original_omml}, now 0)")
    else:
        print(f"❌ OMML equations still present: {converted_omml}")
        success = False

    # Check LaTeX was added
    if converted_latex['total_equations'] > 0:
        print(f"✅ LaTeX equations added: {converted_latex['total_equations']}")
    else:
        print(f"❌ No LaTeX equations found in text!")
        success = False

    # Check against expected count if provided
    if expected_count:
        if converted_latex['total_equations'] == expected_count:
            print(f"✅ Equation count matches expected: {expected_count}")
        else:
            print(f"⚠️ Equation count mismatch: found {converted_latex['total_equations']}, expected {expected_count}")
            success = False

    # 5. Show sample text to prove conversion
    print(f"\n📝 SAMPLE TEXT (first 500 chars):")
    print("-" * 40)
    sample = converted_text[:500]

    # Highlight equations in sample
    sample = sample.replace('\\(', '【\\(').replace('\\)', '\\)】')
    sample = sample.replace('\\[', '【\\[').replace('\\]', '\\]】')
    print(sample)

    if '\\(' in converted_text or '\\[' in converted_text:
        print("\n✅ LaTeX equations found in plain text - conversion successful!")
    else:
        print("\n❌ No LaTeX equations found in text - conversion may have failed!")

    return {
        'success': success,
        'original_omml': original_omml,
        'converted_omml': converted_omml,
        'latex_equations': converted_latex['total_equations'],
        'text_sample': converted_text[:200]
    }


def main():
    """Test the verification on our converted documents"""

    print("="*60)
    print("TESTING CORRECT VERIFICATION METHOD")
    print("="*60)

    # Test document paths
    original = Path(r"D:\Development\document-processing-api-2\document-processing-api\test docs\الدالة واحد لواحد (جاهزة للنشر) - Copy.docx")
    converted = Path("test_144_all_equations.docx")  # File is in current directory

    if original.exists() and converted.exists():
        result = verify_conversion(original, converted, expected_count=144)

        print("\n" + "="*60)
        print("FINAL VERIFICATION:")
        if result['success']:
            print("✅ CONVERSION VERIFIED SUCCESSFULLY!")
            print(f"   - {result['original_omml']} OMML equations replaced")
            print(f"   - {result['latex_equations']} LaTeX equations added")
        else:
            print("❌ VERIFICATION FAILED")
    else:
        if not original.exists():
            print(f"❌ Original file not found: {original}")
        if not converted.exists():
            print(f"❌ Converted file not found: {converted}")

    # Also test the first document
    print("\n" + "="*60)
    print("TESTING FIRST DOCUMENT")
    print("="*60)

    original1 = Path(r"D:\Development\document-processing-api-2\document-processing-api\test docs\التشابه (جاهزة للنشر) - Copy.docx")
    converted1 = Path("test_standalone_output/التشابه (جاهزة للنشر) - Copy_standalone.docx")

    if original1.exists() and converted1.exists():
        verify_conversion(original1, converted1, expected_count=89)


if __name__ == "__main__":
    main()