"""
Simple test to convert all 144 equations
"""

from pathlib import Path
from standalone_zip_converter import StandaloneZipConverter

# The document with 144 equations
input_file = Path(r"D:\Development\document-processing-api-2\document-processing-api\test docs\الدالة واحد لواحد (جاهزة للنشر) - Copy.docx")

# Test with VML conversion enabled
converter = StandaloneZipConverter()
output_file = Path("test_144_all_equations.docx")

print("\nATTEMPTING TO CONVERT ALL 144 EQUATIONS (INCLUDING VML)")
print("-" * 60)

results = converter.convert_document(input_file, output_file, convert_vml=True)

print("\n" + "=" * 60)
print("FINAL RESULTS:")
print(f"Total equations found: {results['equations_found']}")
print(f"Equations replaced: {results['equations_replaced']}")
print(f"Equations in VML: {results['equations_in_vml']}")

if results['equations_replaced'] == 144:
    print("\n*** SUCCESS! ALL 144 EQUATIONS CONVERTED! ***")
else:
    print(f"\nConverted {results['equations_replaced']} out of 144 equations")