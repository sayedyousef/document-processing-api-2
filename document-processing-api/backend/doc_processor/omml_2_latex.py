"""
Complete OMML to LaTeX Parser - All Issues Fixed
Properly handles matrices, integrals, binomials, symbols with spacing
"""

import os
import zipfile
import re
from lxml import etree

# Wingdings / Symbol font → Unicode characters
# Returns Unicode so smart_symbol_convert handles LaTeX conversion
WSYM_UNICODE_MAP = {
    ('Wingdings 2', 'F0CD'): '\u00D7',  # × multiplication sign
    ('Wingdings 2', 'F0CE'): '\u00F7',  # ÷ division sign
    ('Symbol', 'F0B4'): '\u00D7',       # × multiplication sign
    ('Symbol', 'F0B8'): '\u00F7',       # ÷ division sign
    ('Symbol', 'F0B1'): '\u00B1',       # ± plus-minus
    ('Symbol', 'F0B2'): '\u2265',       # ≥ greater-or-equal
    ('Symbol', 'F0A3'): '\u2264',       # ≤ less-or-equal
    ('Symbol', 'F0B9'): '\u2260',       # ≠ not equal
    ('Symbol', 'F0C5'): '\u221E',       # ∞ infinity
    ('Symbol', 'F0D6'): '\u221A',       # √ square root
}


def _resolve_wsym_latex(sym_elem, ns):
    """Resolve w:sym element to a Unicode character for LaTeX processing."""
    font = sym_elem.get(f'{{{ns["w"]}}}font', '')
    char = sym_elem.get(f'{{{ns["w"]}}}char', '')
    return WSYM_UNICODE_MAP.get((font, char), '')


# Symbol mapping
class LatexCommands:
    COMMANDS = {
        # Your existing:
        'sqrt': (1, False),
        'mathbb': (1, True),
        'frac': (2, False),
        'binom': (2, False),
        'neq': (0, True),
        'alpha': (0, True),
        
        # ADD THESE ACCENTS:
        'hat': (1, False),
        'tilde': (1, False),
        'bar': (1, False),
        'dot': (1, False),
        'ddot': (1, False),
        'vec': (1, False),
    }
    
    @staticmethod
    def format(cmd, *args):
        """Generic LaTeX formatter"""
        if cmd not in LatexCommands.COMMANDS:
            return '\\' + cmd  # Unknown command
        
        num_params, needs_space = LatexCommands.COMMANDS[cmd]
        
        result = '\\' + cmd
        
        # Add parameters with braces
        for i in range(num_params):
            if i < len(args):
                result += '{' + str(args[i]) + '}'
        
        # Add space if needed
        if needs_space:
            result += ' '
            
        return result

MATH_SYMBOLS = {
    # Comparison/Relations - ALL need space
    '≠': r'\neq ',
    '≤': r'\leq ',
    '≥': r'\geq ',
    '≈': r'\approx ',
    '≡': r'\equiv ',
    '∼': r'\sim ',
    
    # Set operations - ALL need space
    '∈': r'\in ',
    '∉': r'\notin ',
    '⊂': r'\subset ',
    '⊆': r'\subseteq ',
    '∪': r'\cup ',
    '∩': r'\cap ',
    '∅': r'\emptyset ',
    
    # Logic - ALL need space
    '∧': r'\land ',
    '∨': r'\lor ',
    '¬': r'\neg ',
    '∀': r'\forall ',
    '∃': r'\exists ',
    
    # Arrows - ALL need space
    '→': r'\rightarrow ',
    '←': r'\leftarrow ',
    '↔': r'\leftrightarrow ',
    '⇒': r'\Rightarrow ',
    '⟹': r'\implies ',
    '⟸': r'\impliedby ',
    
    # Greek letters - ALL need space
    'α': r'\alpha ',
    'β': r'\beta ',
    'γ': r'\gamma ',
    'δ': r'\delta ',
    'ε': r'\epsilon ',
    'θ': r'\theta ',
    'λ': r'\lambda ',
    'μ': r'\mu ',
    'π': r'\pi ',
    'σ': r'\sigma ',
    'τ': r'\tau ',
    'φ': r'\phi ',
    'ψ': r'\psi ',
    'ω': r'\omega ',
    'υ': r'\upsilon ',
    'Γ': r'\Gamma ',
    'Δ': r'\Delta ',
    'Σ': r'\Sigma ',
    'Ω': r'\Omega ',
    'ϒ': r'\Upsilon ',
    
    # Other symbols ending with letters - need space
    '∂': r'\partial ',
    '∇': r'\nabla ',
    '∞': r'\infty ',
    '∠': r'\angle ',
    '⊥': r'\perp ',
    '∥': r'\parallel ',
    '…': r'\ldots ',
    '∴': r'\therefore ',
    '∵': r'\because ',
    
    # Binary operators - ALSO need space
    '±': r'\pm ',
    '∓': r'\mp ',
    '×': r'\times ',
    '÷': r'\div ',
    '·': r'\cdot ',
    
    # Big operators - need space
    '∑': r'\sum ',
    '∏': r'\prod ',
    '∫': r'\int ',
    
    # Special cases - NO trailing space
    '√': r'\sqrt',  # Always followed by {
    '°': r'^\circ',  # Superscript notation
    'ⅆ': r'\, d',  # Already has special spacing


    # Blackboard bold letters (mathematical sets)
    'ℝ': r'\mathbb{R} ',  # Real numbers
    'ℂ': r'\mathbb{C} ',  # Complex numbers
    'ℕ': r'\mathbb{N} ',  # Natural numbers
    'ℤ': r'\mathbb{Z} ',  # Integers
    'ℚ': r'\mathbb{Q} ',  # Rational numbers
    'ℍ': r'\mathbb{H} ',  # Quaternions
    '𝔽': r'\mathbb{F} ',  # Field
    '𝕂': r'\mathbb{K} ',  # Field (alternative)
    '𝔸': r'\mathbb{A} ',  # Algebraic numbers
    '𝔹': r'\mathbb{B} ',  # Boolean domain
    '𝕊': r'\mathbb{S} ',  # Sphere
    '𝕋': r'\mathbb{T} ',  # Torus
    '𝕌': r'\mathbb{U} ',  # 
    '𝕍': r'\mathbb{V} ',  # 
    '𝕎': r'\mathbb{W} ',  # 
    '𝕏': r'\mathbb{X} ',  # 
    '𝕐': r'\mathbb{Y} ',  # 
    'ℙ': r'\mathbb{P} ',  # Projective space/Primes
    

}

MATH_SYMBOLS_old = {
    '≠': r'\neq', '≤': r'\leq', '≥': r'\geq', '±': r'\pm', '×': r'\times',
    '÷': r'\div', '·': r'\cdot', '≈': r'\approx', '≡': r'\equiv', '∼': r'\sim',
    '∈': r'\in', '∉': r'\notin', '⊂': r'\subset', '⊆': r'\subseteq',
    '∪': r'\cup', '∩': r'\cap', '∅': r'\emptyset', '∧': r'\land', '∨': r'\lor',
    '¬': r'\neg', '∀': r'\forall', '∃': r'\exists', '→': r'\rightarrow',
    '←': r'\leftarrow', '↔': r'\leftrightarrow', '⇒': r'\Rightarrow',
    'α': r'\alpha', 'β': r'\beta', 'γ': r'\gamma', 'δ': r'\delta', 'ε': r'\epsilon',
    'θ': r'\theta', 'λ': r'\lambda', 'μ': r'\mu', 'π': r'\pi', 'σ': r'\sigma',
    'τ': r'\tau', 'φ': r'\phi', 'ψ': r'\psi', 'ω': r'\omega', 
    'Γ': r'\Gamma', 'Δ': r'\Delta', 'Σ': r'\Sigma', 'Ω': r'\Omega',
    '∂': r'\partial', '∇': r'\nabla', '∑': r'\sum', '∏': r'\prod', '∫': r'\int',
    '∞': r'\infty', '√': r'\sqrt', '∠': r'\angle', '⊥': r'\perp', '∥': r'\parallel',
    '…': r'\ldots', '∴': r'\therefore', '∵': r'\because', '°': r'^\circ',
    'υ': r'\upsilon', 'ϒ': r'\Upsilon',
    'ⅆ': r'\, d',  # Differential d with thin space before it
    '∓': r'\mp',  # Missing minus-plus
    '⟹': r'\implies',
    '⟸': r'\impliedby',  

}

FUNCTION_NAMES = {
    'sin': r'\sin ', 'cos': r'\cos ', 'tan': r'\tan ', 'sec': r'\sec ',
    'csc': r'\csc ', 'cot': r'\cot ', 'arcsin': r'\arcsin ', 'arccos': r'\arccos ',
    'sinh': r'\sinh ', 'cosh': r'\cosh ', 'tanh': r'\tanh ', 'log': r'\log ',
    'ln': r'\ln ', 'exp': r'\exp ', 'lim': r'\lim ', 'sup': r'\sup ', 'inf': r'\inf ',
    'min': r'\min ', 'max': r'\max ', 'det': r'\det ', 'dim': r'\dim ',
}

class DirectOmmlToLatex:
    def __init__(self):
        self.ns = {
            'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        }


    def smart_symbol_convert(self, text):
        """Convert symbols - spacing already handled in mapping"""
        result = []
        i = 0
        while i < len(text):
            found = False
            for symbol, latex in MATH_SYMBOLS.items():
                if text[i:i+len(symbol)] == symbol:
                    result.append(latex)
                    i += len(symbol)
                    found = True
                    break
            if not found:
                result.append(text[i])
                i += 1
        return ''.join(result)


    def smart_symbol_convert_old(self, text):
        """Convert symbols with smart spacing"""
        result = []
        i = 0
        while i < len(text):
            found = False
            for symbol, latex in MATH_SYMBOLS.items():
                if text[i:i+len(symbol)] == symbol:
                    result.append(latex)
                    # General rule: LaTeX commands need space before letters
                    if i + len(symbol) < len(text):
                        next_char = text[i + len(symbol)]
                        # If it's a LaTeX command and next is letter, add space
                        #if latex.startswith('\\') and next_char.isalpha():
                        if latex.startswith('\\') and latex[-1].isalpha() and next_char.isalpha():
    
                            result.append(' ')
                    i += len(symbol)
                    found = True
                    break
            if not found:
                result.append(text[i])
                i += 1
        return ''.join(result)


    
    def convert_function_names(self, text):
        """Convert function names to LaTeX"""
        if text.startswith('\\'):
            return text
        # Strip U+2061 FUNCTION APPLICATION (invisible char inserted by Word)
        text = text.replace('\u2061', '')
        # Sort by length (longest first) so 'arcsin' matches before 'sin', etc.
        sorted_funcs = sorted(FUNCTION_NAMES.items(), key=lambda x: len(x[0]), reverse=True)
        for func, latex_func in sorted_funcs:
            # Match function name at word boundary, NOT followed by more lowercase letters
            # that form a longer known word (e.g. 'inf' should not match inside 'infty')
            # But DO match 'sin' before variable like 'sinx' → '\sin x'
            text = re.sub(r'\b' + re.escape(func) + r'(?![a-z]{2,})',
                         lambda m, lf=latex_func: lf, text)
        return text


    def clean_output(self, latex):
        """Clean LaTeX output carefully"""
        # Skip cleaning for certain patterns INCLUDING mathbb, sqrt, frac
        if any(cmd in latex for cmd in ['\\binom', '\\left', '\\right', '\\begin', '\\mathbb', '\\sqrt', '\\frac']):
            # Only do minimal cleaning for complex structures
            latex = re.sub(r'\s+_', '_', latex)
            latex = re.sub(r'\s+\^', '^', latex)
            # Fix partial derivatives
            latex = re.sub(r'(\\partial)([a-zA-Z])', r'\1 \2', latex)
            # Fix missing braces in fractions
            latex = re.sub(r'\\frac([a-zA-Z0-9])\{', r'\\frac{\1}{', latex)
            return latex
            
        # Regular cleaning for simple content
        # GOOD - you commented out the problematic line!
        #latex = re.sub(r'(?<!\\[a-zA-Z])\{([a-zA-Z0-9])\}', r'\1', latex)
        latex = re.sub(r'\{\{([^}]+)\}\}', r'{\1}', latex)
        latex = re.sub(r'\s+_', '_', latex)
        latex = re.sub(r'\s+\^', '^', latex)
        # Fix partial derivatives
        latex = re.sub(r'(\\partial)([a-zA-Z])', r'\1 \2', latex)
        # Fix missing braces in fractions
        latex = re.sub(r'\\frac([a-zA-Z0-9])\{', r'\\frac{\1}{', latex)
        return latex

    def clean_output_old(self, latex):
        """Clean LaTeX output carefully"""
        # Skip cleaning for certain patterns
        if any(cmd in latex for cmd in ['\\binom', '\\left', '\\right', '\\begin']):
            # Only do minimal cleaning for complex structures
            latex = re.sub(r'\s+_', '_', latex)
            latex = re.sub(r'\s+\^', '^', latex)
            # Fix partial derivatives
            latex = re.sub(r'(\\partial)([a-zA-Z])', r'\1 \2', latex)
            # Fix missing braces in fractions
            latex = re.sub(r'\\frac([a-zA-Z0-9])\{', r'\\frac{\1}{', latex)
            return latex
            
        # Regular cleaning for simple content
        # Don't remove braces from single characters after backslash commands
        #latex = re.sub(r'(?<!\\[a-zA-Z])\{([a-zA-Z0-9])\}', r'\1', latex)
        latex = re.sub(r'\{\{([^}]+)\}\}', r'{\1}', latex)
        latex = re.sub(r'\s+_', '_', latex)
        latex = re.sub(r'\s+\^', '^', latex)
        # Fix partial derivatives
        latex = re.sub(r'(\\partial)([a-zA-Z])', r'\1 \2', latex)
        # Fix missing braces in fractions
        latex = re.sub(r'\\frac([a-zA-Z0-9])\{', r'\\frac{\1}{', latex)
        return latex
    
    def parse(self, elem):
        """Main parsing function"""
        if elem is None:
            return ''
        
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
        handler = getattr(self, f'parse_{tag}', self.parse_default)
        return handler(elem)
    
    def parse_oMath(self, elem):
        latex = ''.join(self.parse(child) for child in elem)
        latex = self.clean_output(latex)
        latex = self.apply_post_processing(latex)
        return latex
    
    def parse_oMathPara(self, elem):
        return ''.join(self.parse(child) for child in elem)

    def parse_r(self, elem):
        """Run element with smart symbol handling"""

        # Check for double-struck (blackboard bold) formatting
        scr_elem = elem.find('.//m:scr', self.ns)
        if scr_elem is not None and scr_elem.get(f'{{{self.ns["m"]}}}val') == 'double-struck':
            # Get the text
            texts = elem.xpath('.//m:t/text()', namespaces=self.ns)
            text = ''.join(texts)
            
            # Convert to \mathbb{} format
            # Handle common blackboard bold letters
            bb_map = {
                'R': r'\mathbb{R} ',
                'C': r'\mathbb{C} ',
                'N': r'\mathbb{N} ',
                'Z': r'\mathbb{Z} ',
                'Q': r'\mathbb{Q} ',
                'H': r'\mathbb{H} ',
                'F': r'\mathbb{F} ',
                'P': r'\mathbb{P} ',
            }
            
            if text in bb_map:
                return bb_map[text]
            else:
                # Generic blackboard bold for any letter
                #return f'\\mathbb{{{text}}} '
                #return '\\mathbb{' + text + '} '
                return LatexCommands.format('mathbb', text)      





        # Get text content (m:t text nodes + w:sym symbol elements)
        parts = []
        for child in elem:
            ctag = child.tag.split('}')[-1]
            if ctag == 't':
                parts.append(child.text or '')
            elif ctag == 'sym':
                sym_char = _resolve_wsym_latex(child, self.ns)
                if sym_char:
                    parts.append(sym_char)
        text = ''.join(parts)

        # Handle minus sign first
        text = text.replace('−', '-')
        

        # ADD THIS: Fix spacing for LaTeX commands that are ALREADY in the text
        # This handles cases where \neq, \in etc are already present without spaces
        text = re.sub(r'(\\[a-zA-Z]+)([a-zA-Z0-9])', lambda m: m.group(1) + (' ' if m.group(1)[-1].isalpha() else '') + m.group(2), text)

        # FIX: Handle differential d (ⅆ) with proper LaTeX spacing
        # Pattern 'rⅆrⅆ' should become 'r \, dr \, d'
        text = re.sub(r'([a-z])ⅆ([a-z])ⅆ', r'\1 \, d\2 \, d', text)

        
        # Handle single differential like 'xⅆ' -> 'x \, d'
        text = re.sub(r'([a-z])ⅆ', r'\1 \, d', text)
        
        # Also handle regular 'd' as differential when it follows a variable
        # This catches cases where 'd' is already regular 'd' not 'ⅆ'
        text = re.sub(r'([a-z])d([a-z])d\b', r'\1 \, d\2 \, d', text)
        text = re.sub(r'([a-z])d([αβγδεζηθικλμνξοπρστυφχψω])', r'\1 \, d\2', text)
        
        # Convert symbols with smart spacing
        text = self.smart_symbol_convert(text)
        
        # FIX: Add space after Greek letters when followed by variables
        # This fixes γz → \gamma z in superscripts
        #text = re.sub(r'(\\gamma|\\alpha|\\beta|\\delta|\\theta|\\sigma)([a-z])', r'\1 \2', text)
        #text = re.sub(r'(\\neq|\\in|\\rightarrow|\\leftarrow|\\implies|\\leq|\\geq)([a-zA-Z])', r'\1 \2', text)
        text = re.sub(r'(\\neq|\\in|\\rightarrow|\\leftarrow|\\implies|\\leq|\\geq)(?![a-z])([a-zA-Z])', r'\1 \2', text)

        # ADD THIS: Final spacing fix for any LaTeX commands we might have missed
        # This is more comprehensive than your current pattern
        #text = re.sub(r'(\\(?:neq|eq|leq|geq|in|notin|subset|subseteq|rightarrow|leftarrow|implies|Rightarrow|forall|exists|pm|mp|times|div|cdot|approx|equiv|sim|alpha|beta|gamma|delta|epsilon|theta|lambda|mu|pi|sigma|tau|phi|psi|omega|Gamma|Delta|Sigma|Omega))([a-zA-Z])', r'\1 \2', text)
        text = re.sub(r'(\\(?:neq|eq|leq|geq|in|notin|subset|subseteq|rightarrow|leftarrow|implies|Rightarrow|forall|exists|pm|mp|times|div|cdot|approx|equiv|sim|alpha|beta|gamma|delta|epsilon|theta|lambda|mu|pi|sigma|tau|phi|psi|omega|Gamma|Delta|Sigma|Omega))(?![a-z])([a-zA-Z])', r'\1 \2', text)
        
        # Your existing Greek letter spacing (keep this)
        text = re.sub(r'(\\gamma|\\alpha|\\beta|\\delta|\\theta|\\sigma)([a-z])', r'\1 \2', text)
        


        # Convert function names
        text = self.convert_function_names(text)
        
        return text


    def parse_f(self, elem):
        """Fraction with proper binomial detection"""
        num_elem = elem.find('.//m:num', self.ns)
        den_elem = elem.find('.//m:den', self.ns)
        
        num = self.parse(num_elem) if num_elem is not None else ''
        den = self.parse(den_elem) if den_elem is not None else ''
        
        # Strip spaces but preserve content
        num = num.strip()
        den = den.strip()
        
        # Special case for 1/2 type fractions in superscripts
        if num in ['1', '2', '3'] and den in ['2', '3', '4']:
            #return f'\\frac{{{num}}}{{{den}}}'
            return LatexCommands.format('frac', num, den)    

        
        # Binomial coefficient detection - only for n and k pattern (common binomial notation)
        if (len(num) == 1 and len(den) == 1 and
            #num == 'n' and den == 'k'):
            #return f'\\binom{{{num}}}{{{den}}}'
            #if (len(num) == 1 and len(den) == 1 and
            num.isalpha() and den.isalpha() and
            ((num == 'n' and den == 'k') or 
            (elem.getparent() is not None and 
            elem.getparent().tag.endswith('d')))):  # Check if inside delimiters
            #return f'\\binom{{{num}}}{{{den}}}'
            return LatexCommands.format('binom', num, den)

        
        # Regular fraction - ensure braces are always present
        #return f'\\frac{{{num}}}{{{den}}}'
        return LatexCommands.format('frac', num, den)    

    
    def parse_sSup(self, elem):
        """Superscript - handle complex nested structures"""
        base_elem = elem.find('m:e', self.ns)
        sup_elem = elem.find('m:sup', self.ns)
        
        base = self.parse(base_elem) if base_elem is not None else ''
        sup = self.parse(sup_elem) if sup_elem is not None else ''
        
        # For complex bracketed expressions, check for duplicate content
        if base.startswith('\\left['):
            # Check if the base already contains nested integrals
            # Count occurrences of key elements
            integral_count = base.count('\\int')
            if integral_count > 2:  # More than expected means duplication
                # Try to extract just the first integral expression
                parts = base.split('\\int')
                if len(parts) > 3:
                    # Reconstruct with just the needed parts
                    base = '\\left[' + '\\int'.join(parts[:3]) + '\\right]'
        
        # Clean the base for simple cases
        if not any(cmd in base for cmd in ['\\binom', '\\left', '\\right', '\\begin']):
            base = self.clean_output(base)
        
        #return f'{base}^{{{sup}}}'
        return base + '^{' + sup + '}'

    
    def parse_sSub(self, elem):
        """Subscript"""
        base_elem = elem.find('m:e', self.ns)
        sub_elem = elem.find('m:sub', self.ns)
        
        base = self.parse(base_elem) if base_elem is not None else ''
        sub = self.parse(sub_elem) if sub_elem is not None else ''
        
        base = self.clean_output(base)
        #return f'{base}_{{{sub}}}'
        return base + '_{' + sub + '}'

    
    def parse_sSubSup(self, elem):
        """Sub and superscript"""
        base_elem = elem.find('m:e', self.ns)
        sub_elem = elem.find('m:sub', self.ns)
        sup_elem = elem.find('m:sup', self.ns)
        
        base = self.parse(base_elem) if base_elem is not None else ''
        sub = self.parse(sub_elem) if sub_elem is not None else ''
        sup = self.parse(sup_elem) if sup_elem is not None else ''
        
        base = self.clean_output(base)
        #return f'{base}_{{{sub}}}^{{{sup}}}'
        return base + '_{' + sub + '}^{' + sup + '}'

    
    def parse_nary(self, elem):
        """N-ary operations"""
        chr_elem = elem.find('.//m:naryPr/m:chr', self.ns)
        
        if chr_elem is not None:
            op_val = chr_elem.get(f'{{{self.ns["m"]}}}val', '∫')
        else:
            op_val = '∫'
        
        operator = self.smart_symbol_convert(op_val)
        
        sub_elem = elem.find('m:sub', self.ns)
        sup_elem = elem.find('m:sup', self.ns)
        expr_elem = elem.find('m:e', self.ns)
        
        result = operator
        if sub_elem is not None:
            #result += f'_{{{self.parse(sub_elem)}}}'
            result += '_{' + self.parse(sub_elem) + '}'

        if sup_elem is not None:
            #result += f'^{{{self.parse(sup_elem)}}}'
            result += '^{' + self.parse(sup_elem) + '}'

        if expr_elem is not None:
            #result += f' {self.parse(expr_elem)}'
            result += ' ' + self.parse(expr_elem)

        
        return result
    
    def parse_rad(self, elem):
        """Radical"""
        deg_elem = elem.find('m:deg', self.ns)
        expr_elem = elem.find('m:e', self.ns)
        
        expr = self.parse(expr_elem) if expr_elem is not None else ''
        
        # Check if degree is hidden
        deg_hide = elem.find('.//m:degHide', self.ns)
        if deg_hide is not None and deg_hide.get(f'{{{self.ns["m"]}}}val') == '1':
            #return f'\\sqrt{{{expr}}}'
            #return f'\\sqrt{{{expr}}}'  # f-string might be eating braces
            return LatexCommands.format('sqrt', expr) 
               


        
        if deg_elem is not None:
            deg_text = self.parse(deg_elem)
            if deg_text and deg_text.strip():
                #return f'\\sqrt[{deg_text}]{{{expr}}}'
                #return LatexCommands.format('sqrt', expr)
                return '\\sqrt[' + deg_text + ']{' + expr + '}'
        

        
        #return f'\\sqrt{{{expr}}}'
        return LatexCommands.format('sqrt', expr)        

    
    def parse_d(self, elem):
        """Delimiters - handle all types properly"""
        beg_chr = elem.find('.//m:begChr', self.ns)
        end_chr = elem.find('.//m:endChr', self.ns)
        
        open_d = '('
        close_d = ')'
        
        if beg_chr is not None:
            open_d = beg_chr.get(f'{{{self.ns["m"]}}}val', '(')
        if end_chr is not None:
            close_d = end_chr.get(f'{{{self.ns["m"]}}}val', ')')
        
        # Get direct e children
        e_children = []
        for child in elem:
            if child.tag.endswith('e'):
                e_children.append(child)
        
        if not e_children:
            return ''
        
        # Check first e child for special structures
        first_e = e_children[0]
        for grandchild in first_e:
            if grandchild.tag.endswith('m'):
                # Matrix
                matrix_type = 'pmatrix'
                if open_d == '[':
                    matrix_type = 'bmatrix'
                elif open_d == '{':
                    matrix_type = 'Bmatrix'
                elif open_d == '|':
                    matrix_type = 'vmatrix'
                return self.parse_matrix(grandchild, matrix_type)
            elif grandchild.tag.endswith('eqArr'):
                content = self.parse(grandchild)
                if open_d == '{' and (not close_d or close_d == ''):
                    return '\\begin{cases} ' + content + ' \\end{cases}'
                # eqArr inside bracket delimiters = column vector
                elif open_d == '[' and close_d == ']':
                    return '\\begin{bmatrix} ' + content + ' \\end{bmatrix}'
                elif open_d == '(' and close_d == ')':
                    return '\\begin{pmatrix} ' + content + ' \\end{pmatrix}'
                elif open_d == '|' and close_d == '|':
                    return '\\begin{vmatrix} ' + content + ' \\end{vmatrix}'
                return content
        
        # Parse the content
        inner = self.parse(first_e)
        
        # Apply delimiters based on type
        if open_d == '(' and close_d == ')':
            #return f'\\left({inner}\\right)'
            return '\\left(' + inner + '\\right)'

        elif open_d == '[' and close_d == ']':
            #return f'\\left[{inner}\\right]'
            return '\\left[' + inner + '\\right]'

        elif open_d == '{' and close_d == '}':
            #return f'\\left\\{{{inner}\\right\\}}'
            return '\\left\\{' + inner + '\\right\\}'

        elif open_d == '|' and close_d == '|':
            #return f'\\left|{inner}\\right|'
            return '\\left|' + inner + '\\right|'

        else:
            return f'{open_d}{inner}{close_d}'
    
    def parse_matrix(self, elem, matrix_type='pmatrix'):
        """Parse matrix elements correctly"""
        rows = []
        
        # Process each row (mr element)
        for child in elem:
            if child.tag.endswith('mr'):
                cols = []
                # Process each cell (e element) in the row
                for cell in child:
                    if cell.tag.endswith('e'):
                        cell_content = self.parse(cell)
                        if cell_content:
                            cols.append(cell_content)
                
                # Only add non-empty rows
                if cols:
                    rows.append(' & '.join(cols))
        
        # Join rows with line breaks
        if rows:
            content = ' \\\\ '.join(rows)
            #return f'\\begin{{{matrix_type}}} {content} \\end{{{matrix_type}}}'
            return '\\begin{' + matrix_type + '} ' + content + ' \\end{' + matrix_type + '}'

        
        return ''
    
    def parse_m(self, elem):
        """Matrix without delimiters"""
        return self.parse_matrix(elem, 'matrix')
    
    def parse_func(self, elem):
        """Functions with proper handling"""
        fname_elem = elem.find('m:fName', self.ns)
        arg_elem = elem.find('m:e', self.ns)
        
        # Handle limit with subscript
        if fname_elem is not None:
            limlower = fname_elem.find('.//m:limLow', self.ns)
            if limlower is not None:
                fname_parsed = self.parse(fname_elem)
                arg_parsed = self.parse(arg_elem) if arg_elem is not None else ''
                if arg_parsed:
                    return f'{fname_parsed} {arg_parsed}'
                return fname_parsed
        
        # Regular functions
        fname = self.parse(fname_elem) if fname_elem is not None else ''
        arg = self.parse(arg_elem) if arg_elem is not None else ''
        
        # Convert function names
        if fname and not fname.startswith('\\'):
            fname = self.convert_function_names(fname)
        # Strip trailing space from FUNCTION_NAMES to prevent double spaces
        fname = fname.rstrip()

        # Limits don't get parentheses
        if fname and 'lim' in fname.lower():
            if arg:
                return f'{fname} {arg}'
            return fname
        
        # Other functions
        if fname and arg:
            # Don't add extra () if arg already has delimiters from parse_d
            stripped = arg.strip()
            if stripped.startswith('\\left') or stripped.startswith('(') or stripped.startswith('['):
                return f'{fname}{arg}'
            return f'{fname}({arg})'
        elif fname:
            return fname
        else:
            return arg or ''
    
    def parse_limLow(self, elem):
        """Limit lower - for limits with subscripts, also underbrace labels"""
        base_elem = elem.find('m:e', self.ns)
        lim_elem = elem.find('m:lim', self.ns)

        base = self.parse(base_elem) if base_elem is not None else ''
        lim = self.parse(lim_elem) if lim_elem is not None else ''

        # underbrace/overbrace: label goes in _{text}
        if '\\underbrace' in base or '\\overbrace' in base:
            return base + '_{\\text{' + lim + '}}'

        # Convert lim to LaTeX if needed
        if base == 'lim':
            base = '\\lim'
        elif not base.startswith('\\'):
            base = self.convert_function_names(base)

        # If base contains \lim followed by more content (e.g. \lim \frac{...}{...}),
        # the subscript should attach to \lim, not to the end of the whole expression
        if '\\lim ' in base:
            return base.replace('\\lim ', '\\lim_{' + lim + '} ', 1)

        return base + '_{' + lim + '}'

    
    def parse_acc(self, elem):
        """Accents"""
        chr_elem = elem.find('.//m:accPr/m:chr', self.ns)
        base_elem = elem.find('m:e', self.ns)
        
        base = self.parse(base_elem) if base_elem is not None else ''
        
        if chr_elem is not None:
            acc_val = chr_elem.get(f'{{{self.ns["m"]}}}val', '')
            accent_map = {
                '̂': 'hat', '̃': 'tilde', '̄': 'bar',
                '̇': 'dot', '̈': 'ddot', '⃗': 'vec',
            }
            latex_acc = accent_map.get(acc_val, 'hat')
            #return f'\\{latex_acc}{{{base}}}'
            return '\\' + latex_acc + '{' + base + '}'

        
        #return f'\\hat{{{base}}}'
        return LatexCommands.format('hat', base)        

    
    def parse_eqArr(self, elem):
        """Equation array for piecewise functions"""
        parts = []
        for child in elem:
            if child.tag.endswith('e'):
                part = self.parse(child)
                if part and part.strip():
                    parts.append(part.strip())
        
        # Format for cases environment
        formatted_parts = []
        for part in parts:
            # Look for patterns like "a, n odd" or "a, &n even"
            if ',' in part:
                pieces = part.split(',', 1)  # Split on first comma only
                value = pieces[0].strip()
                condition = pieces[1].strip() if len(pieces) > 1 else ''
                
                # Remove leading & if present
                if condition.startswith('&'):
                    condition = condition[1:].strip()
                
                # Format as value & condition
                if condition:
                    if 'odd' in condition or 'even' in condition:
                        formatted_parts.append(f'{value}, & \\text{{{condition}}}')
                    else:
                        formatted_parts.append(f'{value}, & {condition}')
                else:
                    formatted_parts.append(value)
            else:
                formatted_parts.append(part)
        
        return ' \\\\ '.join(formatted_parts)

    def parse_groupChr(self, elem):
        """Group character - underbrace, overbrace, arrows, brackets, etc.

        m:groupChr places a stretchy character above or below content.
        Common characters:
          ⏟ U+23DF  bottom curly bracket (underbrace)
          ⏞ U+23DE  top curly bracket (overbrace)
          { }       curly braces (same as underbrace/overbrace)
          ⎵ U+23B5  bottom square bracket
          ⎴ U+23B4  top square bracket
          → ← ↔    arrows (horizontal)
        """
        chr_elem = elem.find('.//m:groupChrPr/m:chr', self.ns)
        pos_elem = elem.find('.//m:groupChrPr/m:pos', self.ns)
        base_elem = elem.find('m:e', self.ns)

        base = self.parse(base_elem) if base_elem is not None else ''

        group_char = '\u23DF'  # default: bottom curly bracket
        if chr_elem is not None:
            group_char = chr_elem.get(f'{{{self.ns["m"]}}}val', '\u23DF')

        pos = 'bot'
        if pos_elem is not None:
            pos = pos_elem.get(f'{{{self.ns["m"]}}}val', 'bot')

        # Curly braces → underbrace/overbrace
        curly = {'\u23DF', '\u23DE', '{', '}'}
        # Arrows
        arrows = {
            '\u2192': r'\rightarrow',   # →
            '\u2190': r'\leftarrow',    # ←
            '\u2194': r'\leftrightarrow',  # ↔
            '\u21D2': r'\Rightarrow',   # ⇒
            '\u21D0': r'\Leftarrow',    # ⇐
            '\u21D4': r'\Leftrightarrow',  # ⇔
        }

        if group_char in curly:
            if pos == 'top':
                return f'\\overbrace{{{base}}}'
            return f'\\underbrace{{{base}}}'
        elif group_char in arrows:
            arrow = arrows[group_char]
            if pos == 'top':
                return f'\\overset{{{arrow}}}{{{base}}}'
            return f'\\underset{{{arrow}}}{{{base}}}'
        else:
            # Generic: use xoverline/xunderline style with the char
            if pos == 'top':
                return f'\\overset{{{group_char}}}{{{base}}}'
            return f'\\underset{{{group_char}}}{{{base}}}'

    def parse_default(self, elem):
        """Default handler - process children sequentially"""
        results = []
        for child in elem:
            result = self.parse(child)
            if result:
                results.append(result)
        return ''.join(results)
    
    def parse_e(self, elem):
        """Element container"""
        results = []
        for child in elem:
            result = self.parse(child)
            if result:
                results.append(result)
        return ''.join(results)
    # Aliases for simple pass-through elements
    parse_num = parse_default
    parse_den = parse_default
    parse_sub = parse_default
    parse_sup = parse_default
    parse_lim = parse_default
    def parse_limUpp(self, elem):
        """Limit upper - for superscripts, also overbrace labels"""
        base_elem = elem.find('m:e', self.ns)
        lim_elem = elem.find('m:lim', self.ns)

        base = self.parse(base_elem) if base_elem is not None else ''
        lim = self.parse(lim_elem) if lim_elem is not None else ''

        # overbrace/underbrace: label goes in ^{text}
        if '\\overbrace' in base or '\\underbrace' in base:
            return base + '^{\\text{' + lim + '}}'

        return base + '^{' + lim + '}'
    parse_mr = parse_default


    def apply_post_processing(self, latex):
        """Apply all post-processing fixes"""
        # Strip invisible Unicode characters inserted by Word
        latex = latex.replace('\u2061', '')  # FUNCTION APPLICATION
        latex = latex.replace('\u2062', '')  # INVISIBLE TIMES
        latex = latex.replace('\u2063', '')  # INVISIBLE SEPARATOR
        # All the fixes from process_word_document
        latex = re.sub(r'\\binom([a-zA-Z])([a-zA-Z])', r'\\binom{\1}{\2}', latex)
        latex = re.sub(r'(e\^{[^}]+}[a-z]+)(.*?)\1', r'\1\2', latex)
        latex = re.sub(r'([a-zA-Z]+)\\left\(([^)]+)\\right\)\1', r'\1\\left(\2\\right)', latex)
        latex = re.sub(r'\\partial([a-zA-Z])', r'\\partial \1', latex)
        latex = re.sub(r'\\upsilon([a-zA-Z])', r'\\upsilon \1', latex)
        latex = re.sub(r'\\gamma([a-zA-Z])', r'\\gamma \1', latex)
        latex = re.sub(r'\\rightarrow([A-Z][a-z])', r'\\rightarrow \1', latex)
        latex = latex.replace('⋅', r'\cdot')
        latex = re.sub(r'(\\lim[^}]*})\s*\\lim\s', r'\1 ', latex)
        latex = re.sub(r'(\\exists|\\forall)([a-zA-Z])', r'\1 \2', latex)
        latex = re.sub(r'\\left\(\\binom\{([^}]+)\}\{([^}]+)\}\\right\)', r'\\binom{\1}{\2}', latex)
        latex = re.sub(r'\\cdot([A-Za-z])', r'\\cdot \1', latex)
        latex = re.sub(r'(\\approx|\\equiv|\\sim)(\d)', r'\1 \2', latex)
        latex = re.sub(r'\\cdot([A-Za-z])', r'\\cdot \1', latex)
        return latex
