"""
Main Testing Application - Converter
Handles Word document conversion and extraction
"""

import sys
import io
import os
import shutil
import zipfile
import json
import time
from pathlib import Path
from datetime import datetime

# Fix encoding
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Add backend path
backend_path = Path(r"D:\Development\document-processing-api-2\document-processing-api\backend")
sys.path.insert(0, str(backend_path))

try:
    from doc_processor.main_word_com_equation_replacer import WordCOMEquationReplacer
except ImportError:
    print("Error: Cannot import WordCOMEquationReplacer. Check module path.")
    sys.exit(1)


class DocumentConverter:
    """Main converter for testing document processing"""

    def __init__(self, output_base_dir="test_analysis"):
        self.output_base = Path(output_base_dir)
        self.output_base.mkdir(exist_ok=True)
        self.test_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.test_dir = self.output_base / self.test_id
        self.test_dir.mkdir(exist_ok=True)
        print(f"✓ Test directory created: {self.test_dir}")

    def copy_as_zip(self, docx_path):
        """Copy Word file as ZIP"""
        docx_path = Path(docx_path)
        if not docx_path.exists():
            raise FileNotFoundError(f"Input file not found: {docx_path}")

        zip_path = self.test_dir / f"{docx_path.stem}.zip"
        shutil.copy2(docx_path, zip_path)
        print(f"✓ Copied as ZIP: {zip_path.name}")
        return zip_path

    def extract_docx(self, zip_path, suffix=""):
        """Extract ZIP to analyze structure"""
        extract_dir = self.test_dir / f"{zip_path.stem}{suffix}_extracted"

        with zipfile.ZipFile(zip_path, 'r') as z:
            z.extractall(extract_dir)
            files = z.namelist()
            print(f"✓ Extracted {len(files)} files to: {extract_dir.name}")

        return extract_dir

    def save_pretty_xml(self, extract_dir, suffix=""):
        """Save pretty formatted document.xml"""
        from lxml import etree

        doc_xml_path = extract_dir / "word" / "document.xml"
        if not doc_xml_path.exists():
            print(f"⚠ document.xml not found in {extract_dir}")
            return None

        with open(doc_xml_path, 'r', encoding='utf-8') as f:
            xml_content = f.read()

        # Pretty print
        try:
            root = etree.fromstring(xml_content.encode('utf-8'))
            pretty_xml = etree.tostring(root, pretty_print=True, encoding='unicode')

            pretty_path = self.test_dir / f"{extract_dir.stem.replace('_extracted', '')}_pretty.xml"
            with open(pretty_path, 'w', encoding='utf-8') as f:
                f.write(pretty_xml)
            print(f"✓ Saved pretty XML: {pretty_path.name}")
            return xml_content
        except Exception as e:
            print(f"⚠ Could not prettify XML: {e}")
            return xml_content

    def convert_with_word_com(self, docx_path):
        """Convert using Word COM processor"""
        print(f"\n{'='*60}")
        print("CONVERTING WITH WORD COM")
        print(f"{'='*60}")

        output_path = self.test_dir / f"{docx_path.stem}_converted.docx"

        start_time = time.time()
        processor = WordCOMEquationReplacer()

        try:
            result = processor.process_document(str(docx_path), str(output_path))
            elapsed = time.time() - start_time

            if result is None:
                print("❌ Conversion returned None")
                return None, {'error': 'No result returned'}

            if 'error' in result:
                print(f"❌ Conversion failed: {result['error']}")
                return None, result

            # Safely access results
            equations_found = result.get('equations_found', 0)
            equations_replaced = result.get('equations_replaced', 0)
            equations_inaccessible = result.get('equations_inaccessible', 0)

            print(f"\n✅ Conversion complete in {elapsed:.2f} seconds:")
            print(f"  Equations found: {equations_found}")
            print(f"  Equations replaced: {equations_replaced}")
            print(f"  Equations inaccessible: {equations_inaccessible}")

            # Copy HTML if exists
            if result.get('html_path'):
                html_src = Path(result['html_path'])
                if html_src.exists():
                    html_dst = self.test_dir / html_src.name
                    shutil.copy2(html_src, html_dst)
                    print(f"  HTML copied to: {html_dst.name}")

                    # Check for images folder
                    images_src = html_src.parent / f"{html_src.stem}_files"
                    if images_src.exists():
                        images_dst = self.test_dir / f"{html_src.stem}_files"
                        shutil.copytree(images_src, images_dst, dirs_exist_ok=True)
                        print(f"  Images copied to: {images_dst.name}")

            result['conversion_time'] = elapsed
            return output_path, result

        except Exception as e:
            elapsed = time.time() - start_time
            print(f"❌ Exception during conversion: {e}")
            import traceback
            traceback.print_exc()

            return None, {
                'error': str(e),
                'traceback': traceback.format_exc(),
                'conversion_time': elapsed
            }

    def process_document(self, docx_path):
        """Complete processing pipeline"""
        docx_path = Path(docx_path)

        print(f"\n{'='*60}")
        print(f"PROCESSING: {docx_path.name}")
        print(f"Test ID: {self.test_id}")
        print(f"{'='*60}\n")

        results = {
            'test_id': self.test_id,
            'input_file': str(docx_path),
            'test_dir': str(self.test_dir),
            'steps': {}
        }

        # Step 1: Copy and extract original
        print("\n📋 STEP 1: Extract Original")
        print("-" * 40)
        try:
            original_zip = self.copy_as_zip(docx_path)
            original_extract = self.extract_docx(original_zip, "_original")
            original_xml = self.save_pretty_xml(original_extract, "_original")

            results['steps']['extract_original'] = {
                'success': True,
                'zip_path': str(original_zip),
                'extract_path': str(original_extract)
            }
        except Exception as e:
            print(f"❌ Failed to extract original: {e}")
            results['steps']['extract_original'] = {'success': False, 'error': str(e)}
            return results

        # Step 2: Convert with Word COM
        print("\n📋 STEP 2: Convert with Word COM")
        print("-" * 40)
        converted_path, conversion_result = self.convert_with_word_com(docx_path)
        results['steps']['conversion'] = conversion_result

        if converted_path:
            # Step 3: Extract converted document
            print("\n📋 STEP 3: Extract Converted Document")
            print("-" * 40)
            try:
                converted_zip = self.copy_as_zip(converted_path)
                converted_extract = self.extract_docx(converted_zip, "_converted")
                converted_xml = self.save_pretty_xml(converted_extract, "_converted")

                results['steps']['extract_converted'] = {
                    'success': True,
                    'zip_path': str(converted_zip),
                    'extract_path': str(converted_extract)
                }
            except Exception as e:
                print(f"❌ Failed to extract converted: {e}")
                results['steps']['extract_converted'] = {'success': False, 'error': str(e)}

        # Save results
        results_path = self.test_dir / 'conversion_results.json'
        with open(results_path, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"\n✓ Results saved: {results_path}")

        return results


def main(folder_path=None):
    """Main test runner - can process single folder or default test docs"""

    # Determine which folder to process
    if folder_path:
        test_dir = Path(folder_path)
        if not test_dir.exists():
            print(f"❌ Folder not found: {test_dir}")
            return None
        print(f"\n📁 Processing folder: {test_dir}")
    else:
        # Default test documents
        test_dir = Path(r"D:\Development\document-processing-api-2\document-processing-api\test docs")
        print(f"\n📁 Using default test folder: {test_dir}")

    # Find all .docx files in folder
    test_files = list(test_dir.glob("*.docx"))

    # Filter out temporary files (starting with ~$)
    test_files = [f for f in test_files if not f.name.startswith('~$')]

    if not test_files:
        print(f"❌ No Word documents found in: {test_dir}")
        return None

    print(f"\n📄 Found {len(test_files)} documents to process:")
    for i, f in enumerate(test_files, 1):
        print(f"  {i}. {f.name}")

    # Create converter
    converter = DocumentConverter()

    all_results = []
    success_count = 0
    error_count = 0

    # Process each file
    for i, test_file in enumerate(test_files, 1):
        print(f"\n{'='*60}")
        print(f"Processing {i}/{len(test_files)}: {test_file.name}")
        print(f"{'='*60}")

        try:
            results = converter.process_document(test_file)
            all_results.append(results)

            # Check if conversion was successful
            if results.get('steps', {}).get('conversion', {}).get('equations_replaced', 0) > 0:
                success_count += 1
            else:
                error_count += 1

        except Exception as e:
            print(f"❌ Failed to process {test_file.name}: {e}")
            error_count += 1
            all_results.append({
                'input_file': str(test_file),
                'error': str(e)
            })

    # Save combined results
    combined_path = converter.test_dir / 'all_results.json'
    with open(combined_path, 'w', encoding='utf-8') as f:
        json.dump(all_results, f, indent=2, ensure_ascii=False)

    # Create summary report
    summary = {
        'test_id': converter.test_id,
        'folder_processed': str(test_dir),
        'total_files': len(test_files),
        'successful': success_count,
        'errors': error_count,
        'files': [f.name for f in test_files]
    }

    summary_path = converter.test_dir / 'summary.json'
    with open(summary_path, 'w', encoding='utf-8') as f:
        json.dump(summary, f, indent=2, ensure_ascii=False)

    print(f"\n{'='*60}")
    print(f"✅ ALL TESTS COMPLETE")
    print(f"Files processed: {len(test_files)}")
    print(f"Successful: {success_count}")
    print(f"Errors: {error_count}")
    print(f"Output directory: {converter.test_dir}")
    print(f"{'='*60}")

    return converter.test_dir


if __name__ == "__main__":
    # Check if folder path provided as command line argument
    if len(sys.argv) > 1:
        folder_path = sys.argv[1]
        test_output_dir = main(folder_path)
    else:
        # Use default test folder
        test_output_dir = main()

    if test_output_dir:
        print(f"\nRun analyzer with: python test_analyzer.py {test_output_dir}")