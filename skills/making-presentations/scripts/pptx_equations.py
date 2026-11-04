# /// script
# requires-python = ">=3.10"
# dependencies = ["latex2mathml", "lxml"]
# ///
"""Turn $$latex$$ paragraphs in a .pptx into native, editable PowerPoint equations.

    uv run pptx_equations.py IN.pptx OUT.pptx [--xsl PATH-TO-MML2OMML.XSL]

Every text-box paragraph whose whole text is $$...$$ becomes an equation: LaTeX -> MathML
(latex2mathml) -> OMML (Office's own MML2OMML.XSL) -> the structure PowerPoint writes itself:
the shape wrapped in mc:AlternateContent, the paragraph holding a14:m > m:oMathPara > m:oMath,
and every math run in Cambria Math. The original shape, $$latex$$ and all, stays as the
mc:Fallback, so a reader without equation support still shows the source.
"""
import argparse
import copy
import re
import sys
import zipfile
from pathlib import Path

import latex2mathml.converter
from lxml import etree

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
}
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XSL_CANDIDATES = [
    r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files\Microsoft Office\Office16\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\Office16\MML2OMML.XSL",
    r"C:\Program Files\Microsoft Office\root\Office15\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\root\Office15\MML2OMML.XSL",
    r"C:\Program Files\Microsoft Office\Office15\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\Office15\MML2OMML.XSL",
]
EQUATION = re.compile(r"^\s*\$\$(.+)\$\$\s*$", re.S)
SLIDE = re.compile(r"^ppt/slides/slide\d+\.xml$")


def q(tag):
    """'a:p' -> '{namespace}p'."""
    prefix, name = tag.split(":")
    return "{%s}%s" % (NS[prefix], name)


def latex_to_omath(latex, xslt):
    # $$ is display math in LaTeX: "block" puts the limits of a sum under and over it.
    mathml = latex2mathml.converter.convert(latex, display="block")
    # The XSL drops <mspace> (\quad, \,), so the gap would vanish; an em space keeps it.
    mathml = re.sub(r"<mspace[^>]*/>", "<mo>&#x2003;</mo>", mathml)
    omath = xslt(etree.fromstring(mathml)).getroot()
    # Word run properties (w:rPr) mean nothing to PowerPoint; drop every w: element.
    for el in omath.xpath("//*[namespace-uri()=$w]", w=W_NS):
        el.getparent().remove(el)
    return omath


def math_rpr(source_rpr):
    """PowerPoint's run properties for a math run: the text's language, size and fill, in Cambria Math."""
    rpr = etree.Element(q("a:rPr"))
    if source_rpr is not None:
        for key in ("lang", "sz"):
            if source_rpr.get(key):
                rpr.set(key, source_rpr.get(key))
        fill = source_rpr.find(q("a:solidFill"))
        if fill is not None:
            rpr.append(copy.deepcopy(fill))
    etree.SubElement(rpr, q("a:latin"), typeface="Cambria Math")
    return rpr


def to_equation(paragraph, omath):
    """Replace the paragraph's runs with the equation, keeping its paragraph properties."""
    first_run_rpr = paragraph.find("a:r/a:rPr", NS)
    for run in paragraph.xpath("a:r | a:br | a:fld", namespaces=NS):
        paragraph.remove(run)
    for mrun in omath.iter(q("m:r")):
        # a:rPr goes after m:rPr (if any) and before m:t, as PowerPoint writes it.
        mrpr = mrun.find(q("m:rPr"))
        mrun.insert(0 if mrpr is None else 1, math_rpr(first_run_rpr))
    m = etree.Element(q("a14:m"), nsmap={"a14": NS["a14"]})
    para = etree.SubElement(m, q("m:oMathPara"), nsmap={"m": NS["m"]})
    jc = etree.SubElement(etree.SubElement(para, q("m:oMathParaPr")), q("m:jc"))
    jc.set(q("m:val"), "centerGroup")
    para.append(omath)
    ppr = paragraph.find(q("a:pPr"))
    paragraph.insert(0 if ppr is None else 1, m)


def convert_slide(xml, xslt, name="slide"):
    """Returns (new xml bytes, number of equations), or (None, 0) when the slide has none."""
    root = etree.fromstring(xml)
    count = 0
    # Shapes already inside mc:AlternateContent are skipped: they are equations or other
    # versioned content PowerPoint wrote, and nesting would break them.
    for sp in root.xpath("//p:sp[not(ancestor::mc:AlternateContent)]", namespaces=NS):
        fallback = copy.deepcopy(sp)
        found = []
        for paragraph in sp.xpath("p:txBody/a:p", namespaces=NS):
            text = "".join(paragraph.xpath("a:r/a:t/text()", namespaces=NS))
            match = EQUATION.match(text)
            if match:
                found.append((paragraph, match.group(1).strip()))
        if not found:
            continue
        for paragraph, latex in found:
            try:
                omath = latex_to_omath(latex, xslt)
            except Exception as e:
                raise ValueError(f"{name}: cannot convert {latex!r}: {e}") from e
            to_equation(paragraph, omath)
        alt = etree.Element(q("mc:AlternateContent"), nsmap={"mc": NS["mc"], "a14": NS["a14"]})
        sp.addprevious(alt)
        choice = etree.SubElement(alt, q("mc:Choice"), Requires="a14")
        choice.append(sp)
        etree.SubElement(alt, q("mc:Fallback")).append(fallback)
        count += len(found)
    if not count:
        return None, 0
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True), count


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("input")
    ap.add_argument("output")
    ap.add_argument("--xsl", help="path to MML2OMML.XSL (default: Office's own)")
    args = ap.parse_args()
    xsl = args.xsl or next((p for p in XSL_CANDIDATES if Path(p).exists()), None)
    if not xsl or not Path(xsl).exists():
        sys.exit("MML2OMML.XSL not found; install Office or pass --xsl")
    if Path(args.input).resolve() == Path(args.output).resolve():
        sys.exit("write to a new file: OUT must differ from IN")
    xslt = etree.XSLT(etree.parse(xsl))

    try:
        total = write_output(args, xslt)
    except Exception as e:
        Path(args.output).unlink(missing_ok=True)   # never leave a half-written deck
        sys.exit(f"failed, no output written: {e}")
    print(f"{total} equation(s) -> {args.output}")


def write_output(args, xslt):
    total = 0
    with zipfile.ZipFile(args.input) as zin, zipfile.ZipFile(args.output, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if SLIDE.match(item.filename):
                new, n = convert_slide(data, xslt, item.filename)
                if n:
                    data, total = new, total + n
                    print(f"{item.filename}: {n} equation(s)")
            zout.writestr(item, data)
    return total


if __name__ == "__main__":
    main()
