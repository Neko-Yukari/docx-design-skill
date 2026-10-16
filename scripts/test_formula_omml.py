# -*- coding: utf-8 -*-
"""Scoped regression tests for docx skill OMML conversion.

Run:  python test_formula_omml.py
Covers: inline single oMath wrapper, display structure, binomial noBar vs
ordinary fraction bar, common sub/superscripts.  No pytest dependency.
"""
import sys
import zipfile
from pathlib import Path

SCRIPTS = Path(r"C:\Users\81004\.agents\skills\docx\scripts")
sys.path.insert(0, str(SCRIPTS))

from docx_tool import DocxBuilder, latex_to_omml  # noqa: E402
from lxml import etree  # noqa: E402

M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

FAILURES = []


def check(name, cond, detail=""):
    status = "PASS" if cond else "FAIL"
    print(f"[{status}] {name}" + (f"  -- {detail}" if detail and not cond else ""))
    if not cond:
        FAILURES.append(name)


def omml_tree(latex, inline):
    return etree.fromstring(latex_to_omml(latex, inline=inline))


def all_tags(el):
    return [c.tag for c in el.iter() if c.tag.startswith(M)]


# ── 1. Inline: single oMath wrapper, no nesting ──────────────────
for latex in [r"T_{i}", r"E = mc^2", r"\lambda p_{0}^{s}", r"\rho"]:
    root = omml_tree(latex, inline=True)
    check(f"inline root is oMath: {latex!r}", root.tag == M + "oMath", root.tag)
    nested = [t for t in all_tags(root) if root.tag == t and t is not root.tag]
    omaths = [el for el in root.iter() if el.tag == M + "oMath"]
    check(f"inline single oMath (no double-wrap): {latex!r}", len(omaths) == 1,
          f"found {len(omaths)} oMath elements")

# ── 2. Display: valid oMathPara > oMath structure ────────────────
for latex in [r"\frac{a}{b}", r"x = \frac{-b \pm \sqrt{b^2-4ac}}{2a}",
              r"\sum_{m=0}^{\infty} \frac{1}{m!}"]:
    root = omml_tree(latex, inline=False)
    check(f"display root is oMathPara: {latex[:20]!r}", root.tag == M + "oMathPara")
    omaths = [el for el in root.iter() if el.tag == M + "oMath"]
    check(f"display has >=1 oMath child: {latex[:20]!r}",
          len(omaths) >= 1 and all(el.getparent() is root for el in omaths))

# ── 3. Binomial: noBar; ordinary frac: bar (default bar type) ────
binom = omml_tree(r"\binom{n+m}{n}", inline=False)
fracs = [el for el in binom.iter() if el.tag == M + "f"]
check("binom produces m:f", len(fracs) == 1)
if fracs:
    fpr = fracs[0].find(M + "fPr")
    ftype = fpr.find(M + "type") if fpr is not None else None
    check("binom m:f type is noBar",
          ftype is not None and ftype.get(M + "val") == "noBar",
          etree.tostring(fracs[0], encoding="unicode")[:200])
    # paren delimiters around binom (latex2mathml mo &#x28;/&#x29;)
    texts = [t.text or "" for t in binom.iter(M + "t")]
    check("binom wrapped in parentheses", "(" in texts and ")" in texts, str(texts))

frac = omml_tree(r"\frac{a}{b}", inline=False)
ffracs = [el for el in frac.iter() if el.tag == M + "f"]
check("ordinary frac produces m:f", len(ffracs) == 1)
if ffracs:
    fpr = ffracs[0].find(M + "fPr")
    ftype = fpr.find(M + "type") if fpr is not None else None
    check("ordinary frac has NO noBar type (barred)", ftype is None,
          etree.tostring(ffracs[0], encoding="unicode")[:200])

# ── 4. Common sub/superscripts ───────────────────────────────────
sub = omml_tree(r"T_{i}", inline=True)
check("T_{i} uses m:sSub", sub.find(f".//{M}sSub") is not None)
sup = omml_tree(r"x^{2}", inline=True)
check("x^{2} uses m:sSup", sup.find(f".//{M}sSup") is not None)
subsup = omml_tree(r"\lambda p_{0}^{s}", inline=True)
check("p_{0}^{s} uses m:sSubSup", subsup.find(f".//{M}sSubSup") is not None)

# ── 5. Builder integration: document.xml structure after save ────
b = DocxBuilder().preset_chinese(heading_numbering=False)
p1 = b.doc.add_paragraph()
p1.add_run("The delay ")
r = p1.add_run("")
r._element.append(omml_tree(r"T_{i}", inline=True))  # same pattern as insert_formula
p1.add_run(" is exponential.")
b.add_formula(r"\binom{n+m}{n}", display=True)
b.add_formula(r"\frac{a}{b}", display=True)
out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("regression_output.docx")
b.save(str(out))

with zipfile.ZipFile(out) as z:
    doc_xml = etree.fromstring(z.read("word/document.xml"))

body = doc_xml.find(W + "body")
# inline: oMath must not be nested inside another oMath
inline_omaths = doc_xml.findall(f".//{W}r/{M}oMath")
check("saved doc: inline oMath inside w:r", len(inline_omaths) == 1)
for om in doc_xml.iter(M + "oMath"):
    parent = om.getparent()
    nested_omath = any(anc.tag == M + "oMath" for anc in om.iterancestors())
    if om in inline_omaths:
        continue
    check("saved doc: no nested oMath anywhere",
          not any(e is not om and e.tag == M + "oMath" for e in om.iter()))
# display: oMathPara direct child of w:p
paras_with_omp = doc_xml.findall(f".//{W}p/{M}oMathPara")
check("saved doc: 2 display oMathPara as direct w:p children", len(paras_with_omp) == 2)
# binom noBar survives save
saved_types = [t.get(M + "val") for t in doc_xml.iter(M + "type")]
check("saved doc: binom noBar present", "noBar" in saved_types)

# ── 6. N-ary operand attachment (integral/sum m:e nonempty) ──────
def nary_info(latex):
    """Return list of (chr, operand_child_count, operand_texts, sibling_texts_after_root)."""
    root = omml_tree(latex, inline=False)
    out = []
    for n in root.iter(M + "nary"):
        chr_el = n.find(f"{M}naryPr/{M}chr")
        op = chr_el.get(M + "val") if chr_el is not None else ""
        e = n.find(M + "e")
        e_count = len(e) if e is not None else 0
        e_texts = [t.text or "" for t in (e.iter(M + "t") if e is not None else [])]
        out.append((op, e_count, e_texts))
    return out

# definite integral with differential: operand = x^2 dx
info = nary_info(r"\int_0^1 x^{2}\,dx")
check("int_0^1 x^2 dx: one nary", len(info) == 1)
check("int_0^1 x^2 dx: operand nonempty", info[0][1] > 0, str(info))
check("int_0^1 x^2 dx: operand ends with dx",
      "".join(info[0][2]).endswith("dx"), str(info[0][2]))

# Horton-style: F = int (q/A) e^{-t/k} dt — fraction and sSup inside operand
info = nary_info(r"F = \int_0^{t} \frac{q}{A} e^{-t/k}\,dt")
check("Horton: operand contains fraction+exp",
      info[0][1] >= 2 and any(t for t in info[0][2]), str(info))

# sum: exactly one sibling consumed (a_i), '= C' stays outside
info = nary_info(r"\sum_{i=1}^{n} a_i = C")
check("sum a_i = C: operand is exactly a_i (1 child)", info[0][1] == 1, str(info))
root = omml_tree(r"\sum_{i=1}^{n} a_i = C", inline=False)
all_texts = [t.text or "" for t in root.iter(M + "t")]
check("sum a_i = C: '= C' NOT inside nary operand", "=" in all_texts, str(all_texts))

# sum without explicit operand target: operator followed by operator → no attach
info = nary_info(r"\sum_{i=1}^{n} + a_i")
check("sum followed by operator: operand stays empty", info[0][1] == 0, str(info))

# two adjacent integrals: each gets its own operand, no cross-contamination
info = nary_info(r"\int_0^1 x\,dx + \int_1^2 y\,dy")
check("two adjacent integrals: both operands nonempty",
      len(info) == 2 and info[0][1] > 0 and info[1][1] > 0, str(info))
check("two adjacent integrals: first operand ends dx, second dy",
      "".join(info[0][2]).endswith("dx") and "".join(info[1][2]).endswith("dy"),
      str(info))

# nested integral: inner binds first, outer operand contains inner result
info = nary_info(r"\int_0^1 \int_0^x y\,dy\,dx")
check("nested integrals: two nary", len(info) == 2)
check("nested integrals: inner operand nonempty", info[1][1] > 0, str(info))
check("nested integrals: outer operand nonempty (contains inner)",
      info[0][1] > 0, str(info))

# integral without limits followed by dx: still attaches
info = nary_info(r"\int x\,dx")
check("no-limit integral: operand attached", info[0][1] > 0, str(info))

# integral with NO differential: conservative no-attach (relation present anyway)
info = nary_info(r"\int_0^1 f(x) = C")
check("integral without dx before '=': operand stays empty", info[0][1] == 0, str(info))

# no-limit integral, no differential at all → plain operator run, no nary
info = nary_info(r"\int f(x)")
check("no-limit integral without dx: no nary emitted (plain run)", len(info) == 0, str(info))

# sum followed by fraction (single structural unit)
info = nary_info(r"\sum_{i=1}^{n} \frac{1}{i}")
check("sum of fraction: operand is the fraction (1 child)", info[0][1] == 1, str(info))

# product
info = nary_info(r"\prod_{i=1}^{n} a_i")
check("prod: operand attached (1 child)", info[0][1] == 1, str(info))

# standalone latex2omml module parity
sys.path.insert(0, str(SCRIPTS))
import importlib
import latex2omml as l2o
importlib.reload(l2o)
from latex2mathml import converter as _lc
from xml.etree.ElementTree import fromstring as _fs
ml = _lc.convert(r"\int_0^1 x^{2}\,dx")
root = l2o.mathml_to_omml(l2o._fix_nary_operands(_fs(ml)))
narys = list(root.iter(M + "nary"))
check("standalone latex2omml: int operand nonempty",
      len(narys) == 1 and narys[0].find(M + "e") is not None
      and len(narys[0].find(M + "e")) > 0)

print()
if FAILURES:
    print(f"{len(FAILURES)} FAILURE(S): {FAILURES}")
    sys.exit(1)
print("ALL TESTS PASSED")
