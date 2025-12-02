"""
Test entire folder with conversion and analysis
Usage: python test_folder.py [folder_path]
"""

import sys
import io
import subprocess
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def test_folder(folder_path=None):
    """Run complete test on folder"""

    print("="*60)
    print("FOLDER TESTING WITH CONVERSION AND ANALYSIS")
    print("="*60)

    # Step 1: Run converter on folder
    print("\n📋 Step 1: Running Converter on Folder...")

    cmd = [sys.executable, "test_converter.py"]
    if folder_path:
        cmd.append(folder_path)

    result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')

    # Extract test directory from output
    test_dir = None
    for line in result.stdout.split('\n'):
        if 'Output directory:' in line:
            test_dir = line.split('Output directory:')[1].strip()
            break

    print(result.stdout)

    if test_dir:
        print(f"\n✓ Conversion complete. Output: {test_dir}")

        # Step 2: Run analyzer
        print("\n📋 Step 2: Running Analyzer...")
        print("="*60)

        result = subprocess.run([sys.executable, "test_analyzer.py", test_dir],
                              capture_output=True, text=True, encoding='utf-8')

        print(result.stdout)

        if result.stderr:
            print("Errors:", result.stderr)

        print("\n✅ Folder testing complete!")
        print(f"Results saved in: {test_dir}")

        # Create a combined HTML report
        create_html_report(test_dir)

    else:
        print("❌ Could not find test directory")
        print(result.stdout)
        if result.stderr:
            print("Errors:", result.stderr)

def create_html_report(test_dir):
    """Create HTML report with all results"""
    import json

    test_dir = Path(test_dir)

    # Load analysis report
    analysis_file = test_dir / 'analysis_report.json'
    if not analysis_file.exists():
        print("❌ No analysis report found")
        return

    with open(analysis_file, 'r', encoding='utf-8') as f:
        report = json.load(f)

    # Create HTML
    html = """<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Conversion Test Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #333; }
        table { border-collapse: collapse; width: 100%; margin: 20px 0; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        .success { color: green; }
        .error { color: red; }
        .warning { color: orange; }
    </style>
</head>
<body>
    <h1>Document Conversion Test Report</h1>
    <p>Test ID: """ + test_dir.name + """</p>

    <h2>Summary</h2>
    <table>
        <tr><th>Document</th><th>Original OMML</th><th>LaTeX Added</th><th>Remaining OMML</th><th>Success Rate</th></tr>
"""

    for analysis in report.get('analyses', []):
        if analysis.get('original') and analysis.get('converted'):
            doc = analysis['document'][:40]
            orig = analysis['original']['omml_equations']
            latex = analysis['converted']['latex_inline'] + analysis['converted']['latex_display']
            remain = analysis['converted']['omml_equations']
            rate = (latex / orig * 100) if orig > 0 else 0

            html += f"""        <tr>
            <td>{doc}</td>
            <td>{orig}</td>
            <td class="{'success' if latex > 0 else 'error'}">{latex}</td>
            <td class="{'success' if remain == 0 else 'warning'}">{remain}</td>
            <td>{rate:.1f}%</td>
        </tr>
"""

    html += """    </table>

    <h2>Detailed Results</h2>
"""

    for analysis in report.get('analyses', []):
        if analysis.get('original') and analysis.get('converted'):
            html += f"""    <h3>{analysis['document']}</h3>
    <ul>
        <li>Original OMML equations: {analysis['original']['omml_equations']}</li>
        <li>LaTeX inline added: {analysis['converted']['latex_inline']}</li>
        <li>LaTeX display added: {analysis['converted']['latex_display']}</li>
        <li>Remaining OMML: {analysis['converted']['omml_equations']}</li>
"""

            if analysis['original'].get('equations_detail'):
                detail = analysis['original']['equations_detail']
                html += f"""        <li>VML equations: {detail.get('vml', 0)}</li>
"""

            html += """    </ul>
"""

    html += """</body>
</html>"""

    # Save HTML report
    html_path = test_dir / 'test_report.html'
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"\n✓ HTML report created: {html_path}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        folder_path = sys.argv[1]
        test_folder(folder_path)
    else:
        # Use default test folder
        test_folder()