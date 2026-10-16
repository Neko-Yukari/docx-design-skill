#!/usr/bin/env python3
"""
Convert LaTeX formulas to OMML for .docx embedding.
Uses latex2mathml for robust LaTeX parsing, then converts MathML to OMML.

Usage:
    from latex2omml import latex_to_omml, insert_formula, insert_formula_display

    # Inline formula (inside a run, for mixed text+math)
    p = doc.add_paragraph()
    p.add_run('The formula: ')
    insert_formula(p, r"E = mc^2")

    # Display formula (block-level, centered equation)
    p = doc.add_paragraph()
    insert_formula_display(p, r"x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}")

CLI: python latex2omml.py "E = mc^2"
"""
import sys
from xml.etree.ElementTree import Element, SubElement, tostring, fromstring
from latex2mathml import converter as _latex_converter

OMML_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
MATHML_NS = 'http://www.w3.org/1998/Math/MathML'
XML_NS = 'http://www.w3.org/XML/1998/namespace'
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def _tag(elem):
    return elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag


def _extract_text(elem):
    parts = []
    if elem.text: parts.append(elem.text)
    for c in elem:
        parts.append(_extract_text(c))
        if c.tail: parts.append(c.tail)
    return ''.join(parts)


_NARY_ABORT_OPS = {
    '=', '\u2260', '<', '>', '\u2264', '\u2265', '\u2248', '\u221D', ',', ';', ':',
}


def _is_mi_text(elem, text):
    return _tag(elem) == 'mi' and elem.text == text


def _has_differential(elem):
    """True if elem's direct children contain a differential (mi 'd' + variable)."""
    kids = list(elem)
    for j in range(len(kids) - 1):
        if _is_mi_text(kids[j], 'd') and _tag(kids[j + 1]) == 'mi':
            return True
    return False


def _attach_nary_operands(parent):
    """Attach following siblings as the n-ary operand (munderover 4th child).

    latex2mathml keeps the integrand/summand as later siblings of the limits
    element, leaving m:nary/m:e empty — Word renders a detached gap. This
    standalone variant also promotes msubsup/msub with an n-ary operator base
    to munderover (docx_tool._fix_mathml does that in a separate pass).
    Attachment runs right-to-left so nested operators bind to their own
    operand first. Conservative boundary policy:
    - ∫/∮: consume siblings up to and including the first top-level
      differential (mi 'd' + variable, or a brace/mfenced group containing
      one); all-or-nothing — if a relation operator or the end of the row
      comes first, consume nothing (old rendering, no meaning change).
      The consumed siblings are wrapped in a single mrow (kids[3] slot).
    - ∑/∏: consume exactly ONE following sibling (never an operator), so
      '= C' and adjacent terms are never swallowed.
    - Bare mo ∫/∮ (no limits): same differential rule, but the operator is
      first promoted to munderover[mo, mrow(), mrow(), operand] so OMML
      emits a real m:nary.
    Note: xml.etree insert/append does NOT reparent — every move here is
    paired with an explicit parent.remove().
    """
    def _wrap_operand(dest_elem, insert_at, nodes):
        wrap = Element(f'{{{MATHML_NS}}}mrow')
        for n in nodes:
            wrap.append(n)
        dest_elem.insert(insert_at, wrap)

    def _promote_msubsup(elem):
        """msubsup/msub with n-ary mo base -> munderover (3 kids), else None."""
        kids = list(elem)
        if not kids or _tag(kids[0]) != 'mo':
            return None
        if kids[0].text not in ('\u222B', '\u222E', '\u2211', '\u220F'):
            return None
        muo = Element(f'{{{MATHML_NS}}}munderover')
        muo.append(kids[0])
        sub = Element(f'{{{MATHML_NS}}}mrow')
        if len(kids) > 1:
            sub.append(kids[1])
        sup = Element(f'{{{MATHML_NS}}}mrow')
        if len(kids) > 2:
            sup.append(kids[2])
        muo.append(sub)
        muo.append(sup)
        return muo

    def _find_integral_end(row, start):
        """Index of last consumed sibling (inclusive), or None."""
        for k in range(start, len(row)):
            cur = row[k]
            if _is_mi_text(cur, 'd') and k + 1 < len(row) and _tag(row[k + 1]) == 'mi':
                return k + 1
            if _tag(cur) in ('mrow', 'mfenced') and _has_differential(cur):
                return k
            if _tag(cur) == 'mo' and cur.text in _NARY_ABORT_OPS:
                return None
        return None

    i = len(parent) - 1
    while i >= 0:
        elem = parent[i]
        promoted = None
        if _tag(elem) in ('msubsup', 'msub'):
            promoted = _promote_msubsup(elem)
        is_muo = _tag(elem) == 'munderover' or promoted is not None
        is_bare_int = (_tag(elem) == 'mo'
                       and elem.text in ('\u222B', '\u222E'))
        if not (is_muo or is_bare_int):
            i -= 1
            continue
        if promoted is not None:
            op = list(promoted)[0].text
        elif is_muo:
            kids = list(elem)
            if len(kids) != 3 or _tag(kids[0]) != 'mo':
                i -= 1
                continue
            op = kids[0].text
            if op not in ('\u222B', '\u222E', '\u2211', '\u220F'):
                i -= 1
                continue
        else:
            op = elem.text
        row = list(parent)
        if i + 1 >= len(row):
            i -= 1
            continue
        if op in ('\u2211', '\u220F'):
            single = row[i + 1]
            if _tag(single) == 'mspace' and i + 2 < len(row):
                single = row[i + 2]
            if _tag(single) not in ('mo', 'mspace'):
                if promoted is not None:
                    promoted.append(single)
                    parent.remove(elem)
                    parent.remove(single)
                    parent.insert(i, promoted)
                else:
                    elem.append(single)
                    parent.remove(single)
        else:
            end = _find_integral_end(row, i + 1)
            if end is not None:
                moved = row[i + 1:end + 1]
                if is_bare_int:
                    muo = Element(f'{{{MATHML_NS}}}munderover')
                    muo.append(elem)
                    muo.append(Element(f'{{{MATHML_NS}}}mrow'))  # empty sub
                    muo.append(Element(f'{{{MATHML_NS}}}mrow'))  # empty sup
                    _wrap_operand(muo, 3, moved)
                    parent.remove(elem)
                    for n in moved:
                        parent.remove(n)
                    parent.insert(i, muo)
                else:
                    dest = promoted if promoted is not None else elem
                    _wrap_operand(dest, 3, moved)
                    for n in moved:
                        parent.remove(n)
                    if promoted is not None:
                        parent.remove(elem)
                        parent.insert(i, promoted)
        i -= 1


def _fix_nary_operands(mathml_root):
    """Apply n-ary operand attachment to every grouping element."""
    def _recurse(elem):
        for child in list(elem):
            _recurse(child)
        if _tag(elem) in ('mrow', 'math', 'mstyle'):
            _attach_nary_operands(elem)
    _recurse(mathml_root)
    return mathml_root


def _add_text_run(parent, text):
    r = SubElement(parent, f'{{{OMML_NS}}}r')
    t = SubElement(r, f'{{{OMML_NS}}}t')
    t.text = text
    if text.startswith(' ') or text.endswith(' '):
        t.set(f'{{{XML_NS}}}space', 'preserve')


def _convert(ml_elem, omml_parent):
    tag = _tag(ml_elem)
    kids = list(ml_elem)

    if tag == 'mrow':
        for c in kids: _convert(c, omml_parent)

    elif tag == 'msup':
        e = SubElement(omml_parent, f'{{{OMML_NS}}}sSup')
        SubElement(e, f'{{{OMML_NS}}}sSupPr')
        ce = SubElement(e, f'{{{OMML_NS}}}e'); cs = SubElement(e, f'{{{OMML_NS}}}sup')
        if kids: _convert(kids[0], ce)
        if len(kids) > 1: _convert(kids[1], cs)

    elif tag == 'msub':
        e = SubElement(omml_parent, f'{{{OMML_NS}}}sSub')
        SubElement(e, f'{{{OMML_NS}}}sSubPr')
        ce = SubElement(e, f'{{{OMML_NS}}}e'); cs = SubElement(e, f'{{{OMML_NS}}}sub')
        if kids: _convert(kids[0], ce)
        if len(kids) > 1: _convert(kids[1], cs)

    elif tag == 'msubsup':
        e = SubElement(omml_parent, f'{{{OMML_NS}}}sSubSup')
        SubElement(e, f'{{{OMML_NS}}}sSubSupPr')
        ce = SubElement(e, f'{{{OMML_NS}}}e')
        csub = SubElement(e, f'{{{OMML_NS}}}sub')
        csup = SubElement(e, f'{{{OMML_NS}}}sup')
        if kids: _convert(kids[0], ce)
        if len(kids) > 1: _convert(kids[1], csub)
        if len(kids) > 2: _convert(kids[2], csup)

    elif tag == 'mfrac':
        e = SubElement(omml_parent, f'{{{OMML_NS}}}f')
        fpr = SubElement(e, f'{{{OMML_NS}}}fPr')
        # \binom arrives as <mfrac linethickness="0"> — must render without a bar
        try:
            no_bar = float(ml_elem.get('linethickness', '')) == 0.0
        except ValueError:
            no_bar = False
        if no_bar:
            SubElement(fpr, f'{{{OMML_NS}}}type').set(f'{{{OMML_NS}}}val', 'noBar')
        num = SubElement(e, f'{{{OMML_NS}}}num')
        den = SubElement(e, f'{{{OMML_NS}}}den')
        if kids: _convert(kids[0], num)
        if len(kids) > 1: _convert(kids[1], den)

    elif tag == 'msqrt':
        e = SubElement(omml_parent, f'{{{OMML_NS}}}rad')
        SubElement(e, f'{{{OMML_NS}}}radPr')
        SubElement(e, f'{{{OMML_NS}}}deg')
        ce = SubElement(e, f'{{{OMML_NS}}}e')
        for c in kids: _convert(c, ce)

    elif tag == 'mroot':
        e = SubElement(omml_parent, f'{{{OMML_NS}}}rad')
        SubElement(e, f'{{{OMML_NS}}}radPr')
        deg = SubElement(e, f'{{{OMML_NS}}}deg')
        ce = SubElement(e, f'{{{OMML_NS}}}e')
        if len(kids) > 1: _convert(kids[1], deg)
        if kids: _convert(kids[0], ce)

    elif tag == 'munderover':
        op = _extract_text(kids[0]) if kids else ''
        nary_ops = {'\u222B': '\u222B', '\u2211': '\u2211', '\u220F': '\u220F', '\u222E': '\u222E'}
        if op in nary_ops:
            e = SubElement(omml_parent, f'{{{OMML_NS}}}nary')
            pr = SubElement(e, f'{{{OMML_NS}}}naryPr')
            SubElement(pr, f'{{{OMML_NS}}}chr').set(f'{{{OMML_NS}}}val', nary_ops[op])
            se = SubElement(e, f'{{{OMML_NS}}}sub'); pe = SubElement(e, f'{{{OMML_NS}}}sup'); ee = SubElement(e, f'{{{OMML_NS}}}e')
            if len(kids) > 1: _convert(kids[1], se)
            if len(kids) > 2: _convert(kids[2], pe)
            if len(kids) > 3: _convert(kids[3], ee)
        else:
            e = SubElement(omml_parent, f'{{{OMML_NS}}}sSubSup')
            SubElement(e, f'{{{OMML_NS}}}sSubSupPr')
            ce = SubElement(e, f'{{{OMML_NS}}}e'); csub = SubElement(e, f'{{{OMML_NS}}}sub'); csup = SubElement(e, f'{{{OMML_NS}}}sup')
            if kids: _convert(kids[0], ce)
            if len(kids) > 1: _convert(kids[1], csub)
            if len(kids) > 2: _convert(kids[2], csup)

    elif tag == 'munder':
        e = SubElement(omml_parent, f'{{{OMML_NS}}}sSub')
        SubElement(e, f'{{{OMML_NS}}}sSubPr')
        ce = SubElement(e, f'{{{OMML_NS}}}e'); cs = SubElement(e, f'{{{OMML_NS}}}sub')
        if kids: _convert(kids[0], ce)
        if len(kids) > 1: _convert(kids[1], cs)

    elif tag == 'mover':
        e = SubElement(omml_parent, f'{{{OMML_NS}}}sSup')
        SubElement(e, f'{{{OMML_NS}}}sSupPr')
        ce = SubElement(e, f'{{{OMML_NS}}}e'); cs = SubElement(e, f'{{{OMML_NS}}}sup')
        if kids: _convert(kids[0], ce)
        if len(kids) > 1: _convert(kids[1], cs)

    elif tag in ('mi', 'mn', 'mo', 'mtext'):
        t = _extract_text(ml_elem)
        if t: _add_text_run(omml_parent, t)

    elif tag == 'mspace':
        _add_text_run(omml_parent, ' ')

    elif tag == 'mfenced':
        op = ml_elem.get('open', '('); cl = ml_elem.get('close', ')')
        _add_text_run(omml_parent, op)
        for c in kids: _convert(c, omml_parent)
        _add_text_run(omml_parent, cl)

    elif tag == 'none':
        pass

    else:
        for c in kids: _convert(c, omml_parent)


def _mathml_to_omml_inner(mathml_root, root_elem):
    """Shared conversion: fills root_elem with oMath children from MathML."""
    for child in list(mathml_root):
        t = _tag(child)
        if t in ('mrow', 'mstyle'):
            om = SubElement(root_elem, f'{{{OMML_NS}}}oMath')
            _convert(child, om)
        elif t == 'semantics':
            for sc in child:
                if _tag(sc) == 'mrow':
                    om = SubElement(root_elem, f'{{{OMML_NS}}}oMath')
                    _convert(sc, om)
                    break
    return root_elem


def mathml_to_omml(mathml_root):
    """Display equation: returns <m:oMathPara> element."""
    return _mathml_to_omml_inner(mathml_root, Element(f'{{{OMML_NS}}}oMathPara'))


def mathml_to_omml_inline(mathml_root):
    """Inline equation: returns a single <m:oMath> element (no oMathPara wrapper).

    Converts MathML content directly into the oMath root. Never nests another
    oMath inside it: Word rejects <m:oMath><m:oMath>...</m:oMath></m:oMath>.
    Inline OMML generates m:oMath directly. Per OOXML spec, m:oMathPara is ONLY for
    block-level display equations and MUST be a direct child of w:p, never inside w:r.
    """
    root = Element(f'{{{OMML_NS}}}oMath')
    for child in list(mathml_root):
        t = _tag(child)
        if t == 'semantics':
            for sc in child:
                if _tag(sc) == 'mrow':
                    _convert(sc, root)
                    break
        else:
            _convert(child, root)
    return root


def latex_to_omml(latex_str, inline=False):
    ml = _latex_converter.convert(latex_str)
    if 'xmlns' not in ml[:200]:
        ml = ml.replace('<math', f'<math xmlns="{MATHML_NS}"', 1)
    root = fromstring(ml)
    _fix_nary_operands(root)
    converter = mathml_to_omml_inline if inline else mathml_to_omml
    return tostring(converter(root), encoding='unicode')


def insert_formula(paragraph, latex_str):
    """Insert an INLINE formula (mixed with text in the same paragraph).
    
    Generates <m:oMath> and appends to the last w:r in the paragraph.
    This is correct for inline math: w:r > m:oMath
    """
    from lxml import etree
    run = paragraph.runs[-1] if paragraph.runs else paragraph.add_run('')
    omml_xml = latex_to_omml(latex_str, inline=True)
    run._element.append(etree.fromstring(omml_xml))
    return run


def insert_formula_display(paragraph, latex_str):
    """Insert a DISPLAY (block-level) formula as a standalone equation.
    
    Generates <m:oMathPara> and inserts it directly into the paragraph (w:p),
    NOT inside any w:r. Per OOXML spec: w:p > m:oMathPara > m:oMath
    This is what Word uses for centered, block-level equations with Equation Tools.
    """
    from lxml import etree
    p_elem = paragraph._p
    omml_xml = latex_to_omml(latex_str, inline=False)
    omp_elem = etree.fromstring(omml_xml)
    p_elem.append(omp_elem)
    return omp_elem


if __name__ == '__main__':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

    if len(sys.argv) < 2:
        print("Usage: python latex2omml.py 'E = mc^2'")
        print("       python latex2omml.py --file formulas.txt")
        sys.exit(1)
    if sys.argv[1] == '--file':
        with open(sys.argv[2], 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#'):
                    print("--- " + line)
                    print(latex_to_omml(line))
                    print()
    else:
        print(latex_to_omml(' '.join(sys.argv[1:])))
