"""
Testing Application - Analyzer
Analyzes conversion results and generates reports
"""

import sys
import io
import json
import re
from pathlib import Path
from lxml import etree

# Fix encoding
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


class ConversionAnalyzer:
    """Analyzer for document conversion results"""

    def __init__(self, test_dir):
        self.test_dir = Path(test_dir)
        if not self.test_dir.exists():
            raise FileNotFoundError(f"Test directory not found: {test_dir}")

    def analyze_xml_content(self, xml_path, label=""):
        """Analyze XML for equations and replacements"""
        if not xml_path.exists():
            return None

        with open(xml_path, 'r', encoding='utf-8') as f:
            xml_content = f.read()

        analysis = {
            'label': label,
            'total_size': len(xml_content),
            'omml_equations': xml_content.count('<m:oMath'),
            'latex_inline': xml_content.count('MATHSTARTINLINE'),
            'latex_display': xml_content.count('MATHSTARTDISPLAY'),
            'vml_textboxes': xml_content.count('<v:textbox'),
            'mc_fallback': xml_content.count('mc:Fallback')
        }

        # Parse for detailed analysis
        try:
            root = etree.fromstring(xml_content.encode('utf-8'))
            ns = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
                'v': 'urn:schemas-microsoft-com:vml',
                'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006'
            }

            # Find equations in different contexts
            all_equations = root.xpath('//m:oMath', namespaces=ns)
            vml_equations = []
            regular_equations = []

            for i, eq in enumerate(all_equations, 1):
                # Check if in VML
                vml_ancestor = eq.xpath('ancestor::v:textbox', namespaces=ns)
                fallback_ancestor = eq.xpath('ancestor::mc:Fallback', namespaces=ns)
                txbx_ancestor = eq.xpath('ancestor::w:txbxContent', namespaces=ns)

                if vml_ancestor or fallback_ancestor or txbx_ancestor:
                    vml_equations.append(i)
                else:
                    regular_equations.append(i)

            analysis['equations_detail'] = {
                'total': len(all_equations),
                'regular': len(regular_equations),
                'vml': len(vml_equations),
                'vml_indices': vml_equations[:10] if vml_equations else []  # First 10 for reference
            }

        except Exception as e:
            analysis['parse_error'] = str(e)

        return analysis

    def extract_latex_samples(self, xml_path, max_samples=10):
        """Extract sample LaTeX equations"""
        if not xml_path.exists():
            return []

        with open(xml_path, 'r', encoding='utf-8') as f:
            xml_content = f.read()

        samples = []

        # Extract inline equations
        inline_pattern = r'MATHSTARTINLINE\\?\((.*?)\\?\)MATHENDINLINE'
        for match in re.finditer(inline_pattern, xml_content):
            samples.append({
                'type': 'inline',
                'latex': match.group(1).strip()
            })
            if len(samples) >= max_samples:
                break

        # Extract display equations
        display_pattern = r'MATHSTARTDISPLAY\\?\[(.*?)\\?\]MATHENDDISPLAY'
        for match in re.finditer(display_pattern, xml_content):
            samples.append({
                'type': 'display',
                'latex': match.group(1).strip()
            })
            if len(samples) >= max_samples:
                break

        return samples

    def analyze_html_output(self, html_path):
        """Analyze HTML output"""
        if not html_path.exists():
            return None

        with open(html_path, 'r', encoding='utf-8') as f:
            html_content = f.read()

        analysis = {
            'file_size': len(html_content),
            'inline_math': html_content.count('\\('),
            'display_math': html_content.count('\\['),
            'mathjax_present': 'MathJax' in html_content,
            'images': html_content.count('<img')
        }

        # Check for equation classes
        analysis['equation_classes'] = {
            'inlineMath': html_content.count('class="inlineMath"'),
            'Math_box': html_content.count('class="Math_box"')
        }

        return analysis

    def compare_before_after(self, doc_name):
        """Compare original and converted documents"""
        # Paths
        original_xml = self.test_dir / f"{doc_name}_original_extracted" / "word" / "document.xml"
        converted_xml = self.test_dir / f"{doc_name}_converted_extracted" / "word" / "document.xml"

        comparison = {
            'document': doc_name,
            'original': self.analyze_xml_content(original_xml, "ORIGINAL"),
            'converted': self.analyze_xml_content(converted_xml, "CONVERTED")
        }

        # Calculate changes
        if comparison['original'] and comparison['converted']:
            orig = comparison['original']
            conv = comparison['converted']

            comparison['changes'] = {
                'omml_removed': orig['omml_equations'] - conv['omml_equations'],
                'latex_added': conv['latex_inline'] + conv['latex_display'],
                'conversion_rate': 0
            }

            if orig['omml_equations'] > 0:
                comparison['changes']['conversion_rate'] = (
                    comparison['changes']['latex_added'] / orig['omml_equations'] * 100
                )

            # Check VML handling
            if orig.get('equations_detail') and conv.get('equations_detail'):
                orig_vml = orig['equations_detail']['vml']
                conv_vml = conv['equations_detail']['vml']
                comparison['changes']['vml_handling'] = {
                    'original_vml': orig_vml,
                    'converted_vml': conv_vml,
                    'vml_converted': orig_vml - conv_vml if orig_vml > conv_vml else 0
                }

        # Get LaTeX samples
        comparison['latex_samples'] = self.extract_latex_samples(converted_xml, 5)

        return comparison

    def generate_report(self):
        """Generate comprehensive analysis report"""
        report = {
            'test_dir': str(self.test_dir),
            'analyses': []
        }

        # Load conversion results
        results_file = self.test_dir / 'conversion_results.json'
        if results_file.exists():
            with open(results_file, 'r', encoding='utf-8') as f:
                conversion_results = json.load(f)
                report['conversion_results'] = conversion_results

        # Analyze each document - look for both .zip files and _extracted directories
        processed_docs = set()

        # Try to find documents by extracted directories first
        for extract_dir in self.test_dir.glob("*_original_extracted"):
            doc_name = extract_dir.name.replace("_original_extracted", "")
            processed_docs.add(doc_name)

        # Process each found document
        for doc_name in processed_docs:
            print(f"\n📊 Analyzing: {doc_name}")
            print("-" * 40)

            analysis = self.compare_before_after(doc_name)

            # Add HTML analysis
            html_path = self.test_dir / f"{doc_name}_converted.html"
            if html_path.exists():
                analysis['html'] = self.analyze_html_output(html_path)

            report['analyses'].append(analysis)

            # Print summary
            if analysis['original'] and analysis['converted']:
                orig = analysis['original']['omml_equations']
                conv_omml = analysis['converted']['omml_equations']
                conv_latex = analysis['converted']['latex_inline'] + analysis['converted']['latex_display']

                print(f"  Original OMML: {orig}")
                print(f"  Converted OMML: {conv_omml}")
                print(f"  LaTeX markers: {conv_latex}")

                if 'changes' in analysis:
                    rate = analysis['changes']['conversion_rate']
                    print(f"  Conversion rate: {rate:.1f}%")

                    if 'vml_handling' in analysis['changes']:
                        vml = analysis['changes']['vml_handling']
                        if vml['original_vml'] > 0:
                            print(f"  VML equations: {vml['original_vml']}")
                            print(f"  VML converted: {vml['vml_converted']}")

        # Save report
        report_path = self.test_dir / 'analysis_report.json'
        with open(report_path, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False)

        print(f"\n✓ Report saved: {report_path}")

        return report

    def print_summary(self):
        """Print detailed equation counts and types"""
        report = self.generate_report()

        print(f"\n{'='*60}")
        print("DETAILED EQUATION ANALYSIS")
        print(f"{'='*60}")

        for analysis in report['analyses']:
            if analysis['original'] and analysis['converted']:
                doc = analysis['document']

                print(f"\n📄 DOCUMENT: {doc}")
                print("="*50)

                # ORIGINAL FILE DETAILS
                orig = analysis['original']
                print("\n🔵 ORIGINAL FILE:")
                print(f"  Total OMML equations: {orig['omml_equations']}")

                if 'equations_detail' in orig:
                    detail = orig['equations_detail']
                    print(f"    - Regular (accessible): {detail['regular']}")
                    print(f"    - Inside VML textboxes: {detail['vml']}")

                    if detail['vml'] > 0 and detail['vml_indices']:
                        print(f"    - VML equation positions: {detail['vml_indices'][:10]}...")

                print(f"  VML textboxes in document: {orig['vml_textboxes']}")
                print(f"  Fallback sections: {orig['mc_fallback']}")

                # CONVERTED FILE DETAILS
                conv = analysis['converted']
                print("\n🟢 CONVERTED FILE:")
                print(f"  Remaining OMML equations: {conv['omml_equations']}")

                if 'equations_detail' in conv:
                    detail = conv['equations_detail']
                    print(f"    - Still regular: {detail['regular']}")
                    print(f"    - Still in VML: {detail['vml']}")

                print(f"  LaTeX replacements added:")
                print(f"    - Inline (\\(...\\)): {conv['latex_inline']}")
                print(f"    - Display (\\[...\\]): {conv['latex_display']}")
                print(f"    - TOTAL LaTeX: {conv['latex_inline'] + conv['latex_display']}")

                # CHANGES SUMMARY
                if 'changes' in analysis:
                    changes = analysis['changes']
                    print("\n📊 EQUATION CHANGES:")
                    print(f"  OMML removed: {changes['omml_removed']}")
                    print(f"  LaTeX added: {changes['latex_added']}")

                    if 'vml_handling' in changes:
                        vml = changes['vml_handling']
                        print(f"  VML equations:")
                        print(f"    - Original VML: {vml['original_vml']}")
                        print(f"    - Still VML after conversion: {vml['converted_vml']}")
                        print(f"    - VML converted to LaTeX: {vml['vml_converted']}")

                # HTML OUTPUT
                if 'html' in analysis:
                    html = analysis['html']
                    print("\n🌐 HTML OUTPUT:")
                    print(f"  Inline equations (\\(...\\)): {html['inline_math']}")
                    print(f"  Display equations (\\[...\\]): {html['display_math']}")
                    print(f"  TOTAL in HTML: {html['inline_math'] + html['display_math']}")
                    print(f"  Images: {html['images']}")

                    if html['equation_classes']['inlineMath'] > 0 or html['equation_classes']['Math_box'] > 0:
                        print(f"  Equation classes:")
                        print(f"    - class='inlineMath': {html['equation_classes']['inlineMath']}")
                        print(f"    - class='Math_box': {html['equation_classes']['Math_box']}")

                # LATEX SAMPLES
                if analysis['latex_samples']:
                    print("\n📝 SAMPLE LaTeX EQUATIONS:")
                    for i, sample in enumerate(analysis['latex_samples'][:5], 1):
                        latex = sample['latex']
                        print(f"  {i}. [{sample['type']:7s}] {latex}")

                print("\n" + "-"*50)

        # SUMMARY TABLE
        print(f"\n{'='*60}")
        print("SUMMARY TABLE")
        print(f"{'='*60}")
        print(f"{'Document':<30} {'Original OMML':>15} {'LaTeX Added':>15} {'Remaining OMML':>15}")
        print("-"*76)

        for analysis in report['analyses']:
            if analysis['original'] and analysis['converted']:
                doc = analysis['document'][:30]
                orig_omml = analysis['original']['omml_equations']
                latex_total = analysis['converted']['latex_inline'] + analysis['converted']['latex_display']
                remain_omml = analysis['converted']['omml_equations']

                print(f"{doc:<30} {orig_omml:>15} {latex_total:>15} {remain_omml:>15}")

        print(f"\n✅ Analysis complete!")
        print(f"Test directory: {self.test_dir}")


def main():
    """Main analyzer entry point"""
    if len(sys.argv) > 1:
        test_dir = sys.argv[1]
    else:
        # Find most recent test directory
        test_base = Path("test_analysis")
        if test_base.exists():
            test_dirs = sorted([d for d in test_base.iterdir() if d.is_dir()])
            if test_dirs:
                test_dir = test_dirs[-1]
                print(f"Using most recent test: {test_dir}")
            else:
                print("No test directories found")
                return
        else:
            print("No test_analysis directory found")
            return

    try:
        analyzer = ConversionAnalyzer(test_dir)
        analyzer.print_summary()
    except Exception as e:
        print(f"❌ Analysis failed: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()