"""
OMML to MathML Converter
=========================

Converts Office Math Markup Language (OMML) elements to MathML HTML strings.
MathML renders natively in all modern browsers without JavaScript.

This is the MathML equivalent of DirectOmmlToLatex (omml_2_latex.py).
Same recursive parsing approach, different output format.
"""

import re
from lxml import etree


# Unicode accent characters mapped to MathML combining characters
ACCENT_MAP = {
    '\u0302': '\u0302',   # combining circumflex (hat)
    '\u0303': '\u0303',   # combining tilde
    '\u0304': '\u0304',   # combining macron (bar)
    '\u0307': '\u0307',   # combining dot above
    '\u0308': '\u0308',   # combining diaeresis (ddot)
    '\u20D7': '\u2192',   # combining right arrow (vec) -> right arrow
}

# Named accent fallbacks
ACCENT_NAME_MAP = {
    'hat': '\u005E',      # ^
    'tilde': '\u007E',    # ~
    'bar': '\u00AF',      # macron
    'dot': '\u02D9',      # dot above
    'ddot': '\u00A8',     # diaeresis
    'vec': '\u2192',      # right arrow
}

# Function names that should be rendered upright in MathML
FUNCTION_NAMES = {
    'sin', 'cos', 'tan', 'sec', 'csc', 'cot',
    'arcsin', 'arccos', 'arctan',
    'sinh', 'cosh', 'tanh',
    'log', 'ln', 'exp', 'lim', 'sup', 'inf',
    'min', 'max', 'det', 'dim', 'gcd', 'lcm',
    'arg', 'deg', 'hom', 'ker', 'Pr',
}


class OmmlToMathMLConverter:
    """Converts OMML XML elements to MathML HTML strings."""

    def __init__(self):
        self.ns = {
            'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        }

    def convert(self, omml_element, is_display=False):
        """Convert an m:oMath or m:oMathPara element to MathML HTML string.

        Args:
            omml_element: lxml element (m:oMath or m:oMathPara)
            is_display: True for block (display) equations, False for inline

        Returns:
            MathML HTML string
        """
        if omml_element is None:
            return ''

        tag = self._tag_name(omml_element)

        # If it's an oMathPara, it's always display mode
        if tag == 'oMathPara':
            is_display = True
            # Find the oMath child
            omath = omml_element.find('m:oMath', namespaces=self.ns)
            if omath is not None:
                omml_element = omath
            else:
                # Parse children directly
                inner = self._parse_children(omml_element)
                return self._wrap_math(inner, is_display)

        inner = self._parse_children(omml_element)
        return self._wrap_math(inner, is_display)

    def _wrap_math(self, content, is_display=False):
        """Wrap content in <math> tags."""
        display_attr = ' display="block"' if is_display else ''
        return f'<math xmlns="http://www.w3.org/1998/Math/MathML"{display_attr}><mrow>{content}</mrow></math>'

    def _tag_name(self, elem):
        """Get the local tag name without namespace."""
        if elem is None:
            return ''
        tag = elem.tag
        if '}' in tag:
            return tag.split('}')[-1]
        return tag

    def _parse(self, elem):
        """Recursively parse an OMML element to MathML."""
        if elem is None:
            return ''

        tag = self._tag_name(elem)
        handler = getattr(self, f'_parse_{tag}', None)
        if handler:
            return handler(elem)
        return self._parse_default(elem)

    def _parse_children(self, elem):
        """Parse all children and concatenate."""
        parts = []
        for child in elem:
            result = self._parse(child)
            if result:
                parts.append(result)
        return ''.join(parts)

    def _parse_default(self, elem):
        """Default: parse children."""
        return self._parse_children(elem)

    # ── Container elements ──

    def _parse_oMath(self, elem):
        """m:oMath - math container."""
        return self._parse_children(elem)

    def _parse_oMathPara(self, elem):
        """m:oMathPara - display math paragraph."""
        return self._parse_children(elem)

    def _parse_e(self, elem):
        """m:e - expression element (generic container)."""
        return self._parse_children(elem)

    def _parse_num(self, elem):
        """m:num - numerator."""
        return self._parse_children(elem)

    def _parse_den(self, elem):
        """m:den - denominator."""
        return self._parse_children(elem)

    def _parse_sub(self, elem):
        """m:sub - subscript content."""
        return self._parse_children(elem)

    def _parse_sup(self, elem):
        """m:sup - superscript content."""
        return self._parse_children(elem)

    def _parse_lim(self, elem):
        """m:lim - limit content."""
        return self._parse_children(elem)

    def _parse_deg(self, elem):
        """m:deg - degree (for radicals)."""
        return self._parse_children(elem)

    def _parse_fName(self, elem):
        """m:fName - function name."""
        return self._parse_children(elem)

    # ── Run element (text) ──

    def _parse_r(self, elem):
        """m:r - run element containing text.

        Classifies each character/token as:
        - <mn> for numbers
        - <mo> for operators
        - <mi> for identifiers (variables, Greek letters)
        """
        # Check for double-struck (blackboard bold) formatting
        rpr = elem.find('m:rPr', self.ns)
        is_double_struck = False
        if rpr is not None:
            scr = rpr.find('m:scr', self.ns)
            if scr is not None and scr.get(f'{{{self.ns["m"]}}}val') == 'double-struck':
                is_double_struck = True

        # Also check w:rPr for double-struck
        w_rpr = elem.find('w:rPr', self.ns)
        if w_rpr is not None and not is_double_struck:
            # Some docs use w:rFonts with double-struck fonts
            pass

        # Get text content
        texts = elem.xpath('.//m:t/text()', namespaces=self.ns)
        text = ''.join(texts)

        if not text:
            return ''

        # Replace minus sign
        text = text.replace('\u2212', '-')

        if is_double_struck:
            return f'<mi mathvariant="double-struck">{self._escape(text)}</mi>'

        return self._classify_and_wrap(text)

    def _classify_and_wrap(self, text):
        """Classify text and wrap in appropriate MathML tags."""
        if not text:
            return ''

        text = text.strip()
        if not text:
            return ''

        parts = []
        i = 0
        while i < len(text):
            # Try to match a number (possibly with decimal point)
            num_match = re.match(r'^(\d+\.?\d*)', text[i:])
            if num_match:
                parts.append(f'<mn>{self._escape(num_match.group(1))}</mn>')
                i += len(num_match.group(1))
                continue

            char = text[i]

            # Operators and punctuation
            if char in '+-=<>()[]{}|/\\!@#$%&*,;:.?' or char in '\u2260\u2264\u2265\u00b1\u00d7\u00f7\u22c5\u2248\u2261\u223c\u2208\u2209\u2282\u2286\u222a\u2229\u2205\u2227\u2228\u00ac\u2200\u2203\u2192\u2190\u2194\u21d2\u27f9\u27f8\u2202\u2207\u221e\u2220\u22a5\u2225\u2026\u2234\u2235\u2213\u22c5\u2211\u220f\u222b':
                parts.append(f'<mo>{self._escape(char)}</mo>')
                i += 1
                continue

            # Check for function names at this position
            func_found = False
            for func_name in sorted(FUNCTION_NAMES, key=len, reverse=True):
                if text[i:].startswith(func_name):
                    # Make sure it's a whole word
                    end_pos = i + len(func_name)
                    if end_pos >= len(text) or not text[end_pos].isalpha():
                        parts.append(f'<mi mathvariant="normal">{func_name}</mi>')
                        i = end_pos
                        func_found = True
                        break
            if func_found:
                continue

            # Single letter = identifier (italic by default)
            if char.isalpha():
                parts.append(f'<mi>{self._escape(char)}</mi>')
                i += 1
                continue

            # Greek and other Unicode math characters -> mi or mo
            if self._is_greek(char):
                parts.append(f'<mi>{self._escape(char)}</mi>')
                i += 1
                continue

            # Default: treat as operator
            parts.append(f'<mo>{self._escape(char)}</mo>')
            i += 1

        return ''.join(parts)

    def _is_greek(self, char):
        """Check if a character is a Greek letter."""
        cp = ord(char)
        return (0x0391 <= cp <= 0x03C9) or (0x1D400 <= cp <= 0x1D7FF)

    # ── Fraction ──

    def _parse_f(self, elem):
        """m:f - fraction -> <mfrac>"""
        num_elem = elem.find('m:num', self.ns)
        den_elem = elem.find('m:den', self.ns)

        num = self._parse(num_elem) if num_elem is not None else ''
        den = self._parse(den_elem) if den_elem is not None else ''

        return f'<mfrac><mrow>{num}</mrow><mrow>{den}</mrow></mfrac>'

    # ── Superscript / Subscript ──

    def _parse_sSup(self, elem):
        """m:sSup - superscript -> <msup>"""
        base_elem = elem.find('m:e', self.ns)
        sup_elem = elem.find('m:sup', self.ns)

        base = self._parse(base_elem) if base_elem is not None else ''
        sup = self._parse(sup_elem) if sup_elem is not None else ''

        return f'<msup><mrow>{base}</mrow><mrow>{sup}</mrow></msup>'

    def _parse_sSub(self, elem):
        """m:sSub - subscript -> <msub>"""
        base_elem = elem.find('m:e', self.ns)
        sub_elem = elem.find('m:sub', self.ns)

        base = self._parse(base_elem) if base_elem is not None else ''
        sub = self._parse(sub_elem) if sub_elem is not None else ''

        return f'<msub><mrow>{base}</mrow><mrow>{sub}</mrow></msub>'

    def _parse_sSubSup(self, elem):
        """m:sSubSup - subscript + superscript -> <msubsup>"""
        base_elem = elem.find('m:e', self.ns)
        sub_elem = elem.find('m:sub', self.ns)
        sup_elem = elem.find('m:sup', self.ns)

        base = self._parse(base_elem) if base_elem is not None else ''
        sub = self._parse(sub_elem) if sub_elem is not None else ''
        sup = self._parse(sup_elem) if sup_elem is not None else ''

        return f'<msubsup><mrow>{base}</mrow><mrow>{sub}</mrow><mrow>{sup}</mrow></msubsup>'

    # ── Radical (square root / nth root) ──

    def _parse_rad(self, elem):
        """m:rad - radical -> <msqrt> or <mroot>"""
        deg_elem = elem.find('m:deg', self.ns)
        expr_elem = elem.find('m:e', self.ns)

        expr = self._parse(expr_elem) if expr_elem is not None else ''

        # Check if degree is hidden (square root)
        deg_hide = elem.find('.//m:degHide', self.ns)
        if deg_hide is not None and deg_hide.get(f'{{{self.ns["m"]}}}val') == '1':
            return f'<msqrt><mrow>{expr}</mrow></msqrt>'

        if deg_elem is not None:
            deg_text = self._parse(deg_elem)
            if deg_text and deg_text.strip():
                return f'<mroot><mrow>{expr}</mrow><mrow>{deg_text}</mrow></mroot>'

        return f'<msqrt><mrow>{expr}</mrow></msqrt>'

    # ── N-ary operators (integral, sum, product) ──

    def _parse_nary(self, elem):
        """m:nary - n-ary operator -> <munderover>/<msubsup> + operator"""
        # Get the operator character
        chr_elem = elem.find('.//m:naryPr/m:chr', self.ns)
        if chr_elem is not None:
            op_char = chr_elem.get(f'{{{self.ns["m"]}}}val', '\u222B')
        else:
            op_char = '\u222B'  # default: integral

        # Check limLoc (limit location): undOvr (under/over) or subSup
        lim_loc_elem = elem.find('.//m:naryPr/m:limLoc', self.ns)
        lim_loc = 'undOvr'  # default
        if lim_loc_elem is not None:
            lim_loc = lim_loc_elem.get(f'{{{self.ns["m"]}}}val', 'undOvr')

        sub_elem = elem.find('m:sub', self.ns)
        sup_elem = elem.find('m:sup', self.ns)
        expr_elem = elem.find('m:e', self.ns)

        sub = self._parse(sub_elem) if sub_elem is not None else ''
        sup = self._parse(sup_elem) if sup_elem is not None else ''
        expr = self._parse(expr_elem) if expr_elem is not None else ''

        op_ml = f'<mo>{self._escape(op_char)}</mo>'

        # Build the operator with limits
        if sub and sup:
            if lim_loc == 'subSup':
                operator = f'<msubsup>{op_ml}<mrow>{sub}</mrow><mrow>{sup}</mrow></msubsup>'
            else:
                operator = f'<munderover>{op_ml}<mrow>{sub}</mrow><mrow>{sup}</mrow></munderover>'
        elif sub:
            if lim_loc == 'subSup':
                operator = f'<msub>{op_ml}<mrow>{sub}</mrow></msub>'
            else:
                operator = f'<munder>{op_ml}<mrow>{sub}</mrow></munder>'
        elif sup:
            if lim_loc == 'subSup':
                operator = f'<msup>{op_ml}<mrow>{sup}</mrow></msup>'
            else:
                operator = f'<mover>{op_ml}<mrow>{sup}</mrow></mover>'
        else:
            operator = op_ml

        if expr:
            return f'{operator}<mrow>{expr}</mrow>'
        return operator

    # ── Delimiters (parentheses, brackets, etc.) ──

    def _parse_d(self, elem):
        """m:d - delimiters -> <mrow> with <mo> or matrix"""
        beg_chr_elem = elem.find('.//m:dPr/m:begChr', self.ns)
        end_chr_elem = elem.find('.//m:dPr/m:endChr', self.ns)

        open_d = '('
        close_d = ')'

        if beg_chr_elem is not None:
            open_d = beg_chr_elem.get(f'{{{self.ns["m"]}}}val', '(')
        if end_chr_elem is not None:
            close_d = end_chr_elem.get(f'{{{self.ns["m"]}}}val', ')')

        # Get direct e children
        e_children = [child for child in elem if self._tag_name(child) == 'e']

        if not e_children:
            return ''

        # Check first e child for matrix or equation array
        first_e = e_children[0]
        for grandchild in first_e:
            gc_tag = self._tag_name(grandchild)
            if gc_tag == 'm':
                # Matrix inside delimiters
                matrix_content = self._parse_matrix_element(grandchild)
                return f'<mrow><mo>{self._escape(open_d)}</mo>{matrix_content}<mo>{self._escape(close_d)}</mo></mrow>'
            elif gc_tag == 'eqArr':
                # Equation array (piecewise)
                content = self._parse_eqArr(grandchild)
                if open_d == '{' and (not close_d or close_d == ''):
                    # Piecewise / cases
                    return f'<mrow><mo>{self._escape(open_d)}</mo>{content}</mrow>'
                return content

        # Regular delimiters - parse all e children
        if len(e_children) == 1:
            inner = self._parse(e_children[0])
        else:
            # Multiple e children separated by implicit comma/separator
            parts = []
            for e_child in e_children:
                parts.append(self._parse(e_child))
            inner = '<mo>,</mo>'.join(parts)

        open_mo = f'<mo>{self._escape(open_d)}</mo>' if open_d else ''
        close_mo = f'<mo>{self._escape(close_d)}</mo>' if close_d else ''

        return f'<mrow>{open_mo}<mrow>{inner}</mrow>{close_mo}</mrow>'

    # ── Matrix ──

    def _parse_m(self, elem):
        """m:m - standalone matrix -> <mtable>"""
        return self._parse_matrix_element(elem)

    def _parse_matrix_element(self, elem):
        """Parse matrix rows and cells into <mtable>."""
        rows = []
        for child in elem:
            if self._tag_name(child) == 'mr':
                cols = []
                for cell in child:
                    if self._tag_name(cell) == 'e':
                        cell_content = self._parse(cell)
                        cols.append(f'<mtd>{cell_content}</mtd>')
                if cols:
                    rows.append(f'<mtr>{"".join(cols)}</mtr>')

        if rows:
            return f'<mtable>{"".join(rows)}</mtable>'
        return ''

    def _parse_mr(self, elem):
        """m:mr - matrix row (handled by _parse_matrix_element)."""
        return self._parse_default(elem)

    # ── Accent ──

    def _parse_acc(self, elem):
        """m:acc - accent -> <mover>"""
        chr_elem = elem.find('.//m:accPr/m:chr', self.ns)
        base_elem = elem.find('m:e', self.ns)

        base = self._parse(base_elem) if base_elem is not None else ''

        # Get accent character
        accent_char = '\u0302'  # default: hat
        if chr_elem is not None:
            acc_val = chr_elem.get(f'{{{self.ns["m"]}}}val', '')
            if acc_val:
                accent_char = ACCENT_MAP.get(acc_val, acc_val)

        return f'<mover><mrow>{base}</mrow><mo>{self._escape(accent_char)}</mo></mover>'

    # ── Function ──

    def _parse_func(self, elem):
        """m:func - function (sin, cos, lim, etc.)"""
        fname_elem = elem.find('m:fName', self.ns)
        arg_elem = elem.find('m:e', self.ns)

        fname = ''
        if fname_elem is not None:
            # Check for limLow inside function name
            limlower = fname_elem.find('.//m:limLow', self.ns)
            if limlower is not None:
                fname = self._parse_limLow(limlower)
            else:
                # Get function name text
                fname_text = self._extract_text(fname_elem)
                if fname_text.strip() in FUNCTION_NAMES:
                    fname = f'<mi mathvariant="normal">{self._escape(fname_text.strip())}</mi>'
                else:
                    fname = self._parse(fname_elem)

        arg = self._parse(arg_elem) if arg_elem is not None else ''

        # Lim doesn't get parentheses
        if fname and arg:
            fname_text = self._extract_text(fname_elem) if fname_elem is not None else ''
            if 'lim' in fname_text.lower():
                return f'{fname}<mo>\u2061</mo><mrow>{arg}</mrow>'
            return f'{fname}<mo>\u2061</mo><mrow>{arg}</mrow>'
        elif fname:
            return fname
        return arg

    # ── Limit lower ──

    def _parse_limLow(self, elem):
        """m:limLow - lower limit -> <munder>"""
        base_elem = elem.find('m:e', self.ns)
        lim_elem = elem.find('m:lim', self.ns)

        base = ''
        if base_elem is not None:
            base_text = self._extract_text(base_elem)
            if base_text.strip() in FUNCTION_NAMES or base_text.strip() == 'lim':
                base = f'<mi mathvariant="normal">{self._escape(base_text.strip())}</mi>'
            else:
                base = self._parse(base_elem)

        lim = self._parse(lim_elem) if lim_elem is not None else ''

        return f'<munder><mrow>{base}</mrow><mrow>{lim}</mrow></munder>'

    # ── Limit upper ──

    def _parse_limUpp(self, elem):
        """m:limUpp - upper limit -> <mover>"""
        base_elem = elem.find('m:e', self.ns)
        lim_elem = elem.find('m:lim', self.ns)

        base = self._parse(base_elem) if base_elem is not None else ''
        lim = self._parse(lim_elem) if lim_elem is not None else ''

        return f'<mover><mrow>{base}</mrow><mrow>{lim}</mrow></mover>'

    # ── Equation array (piecewise / aligned) ──

    def _parse_eqArr(self, elem):
        """m:eqArr - equation array -> <mtable>"""
        rows = []
        for child in elem:
            if self._tag_name(child) == 'e':
                content = self._parse(child)
                if content and content.strip():
                    rows.append(f'<mtr><mtd>{content}</mtd></mtr>')

        if rows:
            return f'<mtable>{"".join(rows)}</mtable>'
        return ''

    # ── Bar (overbar / underbar) ──

    def _parse_bar(self, elem):
        """m:bar - bar -> <mover> or <munder>"""
        bar_pr = elem.find('.//m:barPr/m:pos', self.ns)
        base_elem = elem.find('m:e', self.ns)
        base = self._parse(base_elem) if base_elem is not None else ''

        pos = 'top'
        if bar_pr is not None:
            pos = bar_pr.get(f'{{{self.ns["m"]}}}val', 'top')

        bar_char = '<mo>\u00AF</mo>'  # macron

        if pos == 'bot':
            return f'<munder><mrow>{base}</mrow>{bar_char}</munder>'
        return f'<mover><mrow>{base}</mrow>{bar_char}</mover>'

    # ── Border box ──

    def _parse_borderBox(self, elem):
        """m:borderBox - bordered box -> <menclose>"""
        expr_elem = elem.find('m:e', self.ns)
        expr = self._parse(expr_elem) if expr_elem is not None else ''
        return f'<menclose notation="box"><mrow>{expr}</mrow></menclose>'

    # ── Group character ──

    def _parse_groupChr(self, elem):
        """m:groupChr - group character (underbrace, overbrace, etc.)"""
        chr_elem = elem.find('.//m:groupChrPr/m:chr', self.ns)
        pos_elem = elem.find('.//m:groupChrPr/m:pos', self.ns)
        base_elem = elem.find('m:e', self.ns)

        base = self._parse(base_elem) if base_elem is not None else ''

        group_char = '\u23DF'  # default: bottom curly bracket
        if chr_elem is not None:
            group_char = chr_elem.get(f'{{{self.ns["m"]}}}val', '\u23DF')

        pos = 'bot'
        if pos_elem is not None:
            pos = pos_elem.get(f'{{{self.ns["m"]}}}val', 'bot')

        char_ml = f'<mo>{self._escape(group_char)}</mo>'

        if pos == 'top':
            return f'<mover><mrow>{base}</mrow>{char_ml}</mover>'
        return f'<munder><mrow>{base}</mrow>{char_ml}</munder>'

    # ── Phantom ──

    def _parse_phant(self, elem):
        """m:phant - phantom (invisible spacing)"""
        expr_elem = elem.find('m:e', self.ns)
        expr = self._parse(expr_elem) if expr_elem is not None else ''
        return f'<mphantom><mrow>{expr}</mrow></mphantom>'

    # ── Pre-sub/superscript ──

    def _parse_sPre(self, elem):
        """m:sPre - pre-subscript/superscript -> <mmultiscripts>"""
        base_elem = elem.find('m:e', self.ns)
        sub_elem = elem.find('m:sub', self.ns)
        sup_elem = elem.find('m:sup', self.ns)

        base = self._parse(base_elem) if base_elem is not None else ''
        sub = self._parse(sub_elem) if sub_elem is not None else '<none/>'
        sup = self._parse(sup_elem) if sup_elem is not None else '<none/>'

        return f'<mmultiscripts><mrow>{base}</mrow><mprescripts/><mrow>{sub}</mrow><mrow>{sup}</mrow></mmultiscripts>'

    # ── Skip property elements ──

    def _parse_oMathParaPr(self, elem):
        return ''

    def _parse_rPr(self, elem):
        return ''

    def _parse_ctrlPr(self, elem):
        return ''

    def _parse_fPr(self, elem):
        return ''

    def _parse_sSupPr(self, elem):
        return ''

    def _parse_sSubPr(self, elem):
        return ''

    def _parse_sSubSupPr(self, elem):
        return ''

    def _parse_radPr(self, elem):
        return ''

    def _parse_naryPr(self, elem):
        return ''

    def _parse_dPr(self, elem):
        return ''

    def _parse_mPr(self, elem):
        return ''

    def _parse_accPr(self, elem):
        return ''

    def _parse_funcPr(self, elem):
        return ''

    def _parse_limLowPr(self, elem):
        return ''

    def _parse_limUppPr(self, elem):
        return ''

    def _parse_barPr(self, elem):
        return ''

    def _parse_borderBoxPr(self, elem):
        return ''

    def _parse_groupChrPr(self, elem):
        return ''

    def _parse_phantPr(self, elem):
        return ''

    def _parse_sPrePr(self, elem):
        return ''

    def _parse_eqArrPr(self, elem):
        return ''

    def _parse_mcs(self, elem):
        return ''

    # ── Helpers ──

    def _extract_text(self, elem):
        """Extract all text content from an element."""
        texts = elem.xpath('.//m:t/text()', namespaces=self.ns)
        if not texts:
            texts = elem.xpath('.//w:t/text()', namespaces=self.ns)
        return ''.join(texts)

    def _escape(self, text):
        """Escape HTML special characters."""
        if not text:
            return ''
        return (text
                .replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;')
                .replace('"', '&quot;'))
