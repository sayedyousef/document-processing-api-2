"""
Test if Word can open the converted documents
"""

import sys
import io
import time
from pathlib import Path
import win32com.client
import pythoncom

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


def test_word_open(file_path):
    """Test if Word can open a document without errors"""
    file_path = Path(file_path).resolve()

    if not file_path.exists():
        print(f"❌ File not found: {file_path}")
        return False

    print(f"\n{'='*60}")
    print(f"TESTING WORD OPEN: {file_path.name}")
    print(f"{'='*60}")

    word = None
    doc = None

    try:
        # Initialize COM
        pythoncom.CoInitialize()

        # Create Word application
        print("🔧 Starting Word application...")
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True  # Make Word visible to see if it opens

        # Open document
        print(f"📄 Opening document: {file_path}")
        doc = word.Documents.Open(str(file_path))

        # Wait a moment for document to fully load
        time.sleep(2)

        # Check document properties
        print(f"✅ Document opened successfully!")
        print(f"  Pages: {doc.Range().Information(4)}")  # wdNumberOfPagesInDocument
        print(f"  Words: {doc.Words.Count}")
        print(f"  Paragraphs: {doc.Paragraphs.Count}")

        # Check for equations
        equation_count = 0
        try:
            equation_count = doc.OMaths.Count
            print(f"  OMML Equations: {equation_count}")
        except:
            print(f"  OMML Equations: Unable to count")

        # Check if document has any errors
        if doc.SpellingErrors.Count > 0:
            print(f"  ⚠ Spelling errors: {doc.SpellingErrors.Count}")

        print("\n✅ Word can open and read the document successfully!")

        return True

    except Exception as e:
        print(f"\n❌ Failed to open document: {e}")
        import traceback
        traceback.print_exc()
        return False

    finally:
        # Clean up
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


def test_all_converted_files():
    """Test all converted files in test directories"""
    print("="*60)
    print("TESTING ALL CONVERTED DOCUMENTS WITH WORD")
    print("="*60)

    # Test standalone converted files
    standalone_dir = Path("backend/test_standalone_output")
    if standalone_dir.exists():
        print(f"\n📁 Testing standalone converted files in: {standalone_dir}")
        for docx_file in standalone_dir.glob("*_standalone.docx"):
            test_word_open(docx_file)

    # Test Word COM converted files
    test_analysis_dir = Path("backend/test_analysis")
    if test_analysis_dir.exists():
        # Find most recent test directory
        test_dirs = sorted([d for d in test_analysis_dir.iterdir() if d.is_dir()])
        if test_dirs:
            latest_test = test_dirs[-1]
            print(f"\n📁 Testing Word COM converted files in: {latest_test}")
            for docx_file in latest_test.glob("*_converted.docx"):
                test_word_open(docx_file)

    print("\n✅ All tests complete!")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        # Test specific file
        file_path = sys.argv[1]
        test_word_open(file_path)
    else:
        # Test all converted files
        test_all_converted_files()