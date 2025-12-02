"""
Analyze VML structure to understand why sections are lost
"""

import sys
import io
import zipfile
from pathlib import Path
from lxml import etree

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def analyze_vml_structure(docx_path):
    """Deep analysis of VML textbox structure"""

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

    # Find VML textboxes
    vml_textboxes = root.xpath('//v:textbox', namespaces=namespaces)
    print(f"\n📦 Found {len(vml_textboxes)} VML textboxes")

    # Analyze each textbox
    for i, textbox in enumerate(vml_textboxes, 1):
        print(f"\n--- VML Textbox {i} ---")

        # Find equations in this textbox
        equations = textbox.xpath('.//m:oMath', namespaces=namespaces)
        print(f"  Equations: {len(equations)}")

        # Check text content
        text_runs = textbox.xpath('.//w:t', namespaces=namespaces)
        text_content = ' '.join([t.text for t in text_runs if t.text])
        print(f"  Text preview: {text_content[:100]}...")

        # Check parent structure
        parent = textbox.getparent()
        if parent is not None:
            print(f"  Parent tag: {parent.tag}")
            grandparent = parent.getparent()
            if grandparent is not None:
                print(f"  Grandparent tag: {grandparent.tag}")

    # Find ALL equations
    all_equations = root.xpath('//m:oMath', namespaces=namespaces)
    print(f"\n📊 Total equations: {len(all_equations)}")

    # Categorize equations
    vml_equations = []
    regular_equations = []

    for eq in all_equations:
        if eq.xpath('ancestor::v:textbox', namespaces=namespaces):
            vml_equations.append(eq)
        else:
            regular_equations.append(eq)

    print(f"  Regular: {len(regular_equations)}")
    print(f"  In VML: {len(vml_equations)}")

    # Check document structure
    print(f"\n📄 Document structure analysis:")
    body = root.find('.//w:body', namespaces)
    if body is not None:
        children = list(body)
        print(f"  Body has {len(children)} direct children")

        # Count different types
        paragraphs = body.xpath('./w:p', namespaces=namespaces)
        tables = body.xpath('./w:tbl', namespaces=namespaces)
        print(f"  Paragraphs: {len(paragraphs)}")
        print(f"  Tables: {len(tables)}")

        # Check last elements
        print(f"\n  Last 5 elements in body:")
        for elem in children[-5:]:
            tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
            has_vml = bool(elem.xpath('.//v:*', namespaces=namespaces))
            has_equations = bool(elem.xpath('.//m:oMath', namespaces=namespaces))
            print(f"    - {tag} (VML: {has_vml}, Equations: {has_equations})")


# Test on the document with VML issues
doc_path = Path(r"D:\Development\document-processing-api-2\document-processing-api\test docs\الدالة واحد لواحد (جاهزة للنشر) - Copy.docx")
if doc_path.exists():
    print(f"Analyzing: {doc_path.name}")
    analyze_vml_structure(doc_path)
else:
    print("Document not found!")