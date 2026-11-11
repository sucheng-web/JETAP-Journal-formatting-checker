"""Low-level OOXML helpers shared by the JETAP checker and fixer."""
from __future__ import annotations

import re
from typing import Any, Callable, Dict, List, Optional, Tuple

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

_CJK_RE = re.compile(r"[⺀-鿿가-힯豈-﫿＀-￯]")
_CS_RE = re.compile(r"[֐-ࣿऀ-෿฀-໿]")

ROMAN = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
         (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]


def to_roman(n: int) -> str:
    out = ""
    for value, sym in ROMAN:
        while n >= value:
            out += sym
            n -= value
    return out


def from_roman(s: str) -> Optional[int]:
    s = s.upper()
    if not s or not re.fullmatch(r"[IVXLCDM]+", s):
        return None
    vals = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100, "D": 500, "M": 1000}
    total = 0
    for i, ch in enumerate(s):
        v = vals[ch]
        if i + 1 < len(s) and vals[s[i + 1]] > v:
            total -= v
        else:
            total += v
    return total if to_roman(total) == s else None


# --------------------------------------------------------------------------
# style / property resolution
# --------------------------------------------------------------------------
def _styles_root(document):
    return document.styles.element


def _style_by_id(document, style_id: str):
    for st in _styles_root(document).findall(qn("w:style")):
        if st.get(qn("w:styleId")) == style_id:
            return st
    return None


def _default_style(document, kind: str = "paragraph"):
    for st in _styles_root(document).findall(qn("w:style")):
        if st.get(qn("w:type")) == kind and st.get(qn("w:default")) in ("1", "true"):
            return st
    return None


def style_chain(document, p_el) -> List[Any]:
    """Paragraph style elements from the paragraph's style up through basedOn."""
    chain = []
    ppr = p_el.find(qn("w:pPr"))
    sid = None
    if ppr is not None:
        ps = ppr.find(qn("w:pStyle"))
        if ps is not None:
            sid = ps.get(qn("w:val"))
    st = _style_by_id(document, sid) if sid else _default_style(document)
    seen = set()
    while st is not None and id(st) not in seen:
        seen.add(id(st))
        chain.append(st)
        based = st.find(qn("w:basedOn"))
        st = _style_by_id(document, based.get(qn("w:val"))) if based is not None else None
    return chain


def paragraph_style_name(document, p_el) -> str:
    chain = style_chain(document, p_el)
    if not chain:
        return ""
    name = chain[0].find(qn("w:name"))
    return name.get(qn("w:val")) if name is not None else ""


def _doc_defaults(document, which: str):
    dd = _styles_root(document).find(qn("w:docDefaults"))
    if dd is None:
        return None
    holder = dd.find(qn("w:rPrDefault" if which == "r" else "w:pPrDefault"))
    if holder is None:
        return None
    return holder.find(qn("w:rPr" if which == "r" else "w:pPr"))


def _ppr_sources(document, p_el):
    ppr = p_el.find(qn("w:pPr"))
    if ppr is not None:
        yield ppr
    for st in style_chain(document, p_el):
        sp = st.find(qn("w:pPr"))
        if sp is not None:
            yield sp
    dd = _doc_defaults(document, "p")
    if dd is not None:
        yield dd


def effective_ppr_attr(document, p_el, tag: str, attr: str = "w:val") -> Optional[str]:
    for src in _ppr_sources(document, p_el):
        el = src.find(qn(tag))
        if el is not None:
            val = el.get(qn(attr))
            if val is not None:
                return val
    return None


def effective_alignment(document, p_el) -> str:
    val = effective_ppr_attr(document, p_el, "w:jc")
    return {"left": "LEFT", "start": "LEFT", "center": "CENTER", "right": "RIGHT",
            "end": "RIGHT", "both": "JUSTIFY", "distribute": "JUSTIFY"}.get(val or "left", "LEFT")


def effective_spacing(document, p_el) -> Dict[str, Any]:
    """before/after in pt, line as (value, rule)."""
    out: Dict[str, Any] = {"before": 0.0, "after": 0.0, "line": 240, "rule": "auto",
                           "lines_based": False}
    found = {"before": False, "after": False, "line": False}
    for src in _ppr_sources(document, p_el):
        sp = src.find(qn("w:spacing"))
        if sp is None:
            continue
        for key, att, lines_att in (("before", "w:before", "w:beforeLines"),
                                    ("after", "w:after", "w:afterLines")):
            if found[key]:
                continue
            if sp.get(qn(lines_att)) is not None:
                out[key] = int(sp.get(qn(lines_att))) / 100.0 * 12.0
                out["lines_based"] = True
                found[key] = True
            elif sp.get(qn(att)) is not None:
                out[key] = int(sp.get(qn(att))) / 20.0
                found[key] = True
        if not found["line"] and sp.get(qn("w:line")) is not None:
            out["line"] = int(sp.get(qn("w:line")))
            out["rule"] = sp.get(qn("w:lineRule")) or "auto"
            found["line"] = True
    return out


def effective_indent(document, p_el) -> Dict[str, Optional[int]]:
    out = {"left": 0, "first": 0, "hanging": 0}
    seen = set()
    for src in _ppr_sources(document, p_el):
        ind = src.find(qn("w:ind"))
        if ind is None:
            continue
        for key, atts in (("left", ("w:left", "w:start")), ("first", ("w:firstLine",)),
                          ("hanging", ("w:hanging",))):
            if key in seen:
                continue
            for att in atts:
                if ind.get(qn(att)) is not None:
                    out[key] = int(ind.get(qn(att)))
                    seen.add(key)
                    if key == "first":
                        seen.add("hanging")
                    if key == "hanging":
                        seen.add("first")
                        out["first"] = 0
                    break
    return out


def has_numbering(p_el) -> bool:
    ppr = p_el.find(qn("w:pPr"))
    if ppr is None:
        return False
    num = ppr.find(qn("w:numPr"))
    if num is None:
        return False
    nid = num.find(qn("w:numId"))
    return nid is not None and nid.get(qn("w:val")) not in (None, "0")


def style_has_numbering(document, p_el) -> bool:
    for st in style_chain(document, p_el):
        sp = st.find(qn("w:pPr"))
        if sp is not None:
            num = sp.find(qn("w:numPr"))
            if num is not None:
                nid = num.find(qn("w:numId"))
                if nid is not None:
                    return nid.get(qn("w:val")) not in (None, "0")
    return False


# ---- runs ---------------------------------------------------------------
def run_elements(p_el) -> List[Any]:
    """All w:r in a paragraph, including those inside hyperlinks / smartTags."""
    return [r for r in p_el.iter(qn("w:r"))]


def run_text(r_el) -> str:
    return "".join(t.text or "" for t in r_el.findall(qn("w:t")))


def run_has_drawing(r_el) -> bool:
    return (r_el.find(qn("w:drawing")) is not None or r_el.find(qn("w:pict")) is not None
            or r_el.find(".//" + qn("w:drawing")) is not None
            or r_el.find(qn("w:object")) is not None)


def _rpr_sources(document, p_el, r_el):
    rpr = r_el.find(qn("w:rPr"))
    if rpr is not None:
        yield rpr
        rs = rpr.find(qn("w:rStyle"))
        if rs is not None:
            st = _style_by_id(document, rs.get(qn("w:val")))
            seen = set()
            while st is not None and id(st) not in seen:
                seen.add(id(st))
                srpr = st.find(qn("w:rPr"))
                if srpr is not None:
                    yield srpr
                based = st.find(qn("w:basedOn"))
                st = _style_by_id(document, based.get(qn("w:val"))) if based is not None else None
    for st in style_chain(document, p_el):
        srpr = st.find(qn("w:rPr"))
        if srpr is not None:
            yield srpr
    dd = _doc_defaults(document, "r")
    if dd is not None:
        yield dd


def _toggle(el) -> bool:
    if el is None:
        return False
    v = el.get(qn("w:val"))
    return v not in ("0", "false", "off")


def effective_run_props(document, p_el, r_el) -> Dict[str, Any]:
    props: Dict[str, Any] = {"font": None, "size": None, "bold": None, "italic": None,
                             "vert": None, "caps": None}
    for src in _rpr_sources(document, p_el, r_el):
        if props["font"] is None:
            rf = src.find(qn("w:rFonts"))
            if rf is not None:
                f = rf.get(qn("w:ascii")) or rf.get(qn("w:hAnsi"))
                if f:
                    props["font"] = f
                elif rf.get(qn("w:asciiTheme")):
                    props["font"] = "(theme)"
        if props["size"] is None:
            sz = src.find(qn("w:sz"))
            if sz is not None and sz.get(qn("w:val")):
                props["size"] = int(sz.get(qn("w:val"))) / 2.0
        if props["bold"] is None:
            b = src.find(qn("w:b"))
            if b is not None:
                props["bold"] = _toggle(b)
        if props["italic"] is None:
            i = src.find(qn("w:i"))
            if i is not None:
                props["italic"] = _toggle(i)
        if props["vert"] is None:
            va = src.find(qn("w:vertAlign"))
            if va is not None:
                props["vert"] = va.get(qn("w:val"))
    for k in ("bold", "italic"):
        if props[k] is None:
            props[k] = False
    return props


def paragraph_text(p_el) -> str:
    """Visible text of a paragraph (w:t inside runs, hyperlinks, fields)."""
    parts = []
    for node in p_el.iter():
        if node.tag == qn("w:t"):
            parts.append(node.text or "")
        elif node.tag == qn("w:tab"):
            if node.getparent() is not None and node.getparent().tag == qn("w:r"):
                parts.append("\t")
    return "".join(parts)


def paragraph_has_drawing(p_el) -> bool:
    return (p_el.find(".//" + qn("w:drawing")) is not None
            or p_el.find(".//" + qn("w:pict")) is not None
            or p_el.find(".//" + qn("w:object")) is not None)


def paragraph_has_math(p_el) -> bool:
    return p_el.find(".//{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath") is not None


def paragraph_has_sectpr(p_el) -> bool:
    ppr = p_el.find(qn("w:pPr"))
    return ppr is not None and ppr.find(qn("w:sectPr")) is not None


def is_plain_empty(p_el) -> bool:
    """True only for a paragraph with nothing in it (no text, section break,
    drawing, field, bookmark with content ...)."""
    if paragraph_text(p_el).strip():
        return False
    if paragraph_has_sectpr(p_el) or paragraph_has_drawing(p_el) or paragraph_has_math(p_el):
        return False
    for child in p_el:
        if child.tag == qn("w:pPr"):
            continue
        if child.tag == qn("w:r"):
            for sub in child:
                if sub.tag == qn("w:rPr"):
                    continue
                if sub.tag == qn("w:t"):
                    continue
                if sub.tag == qn("w:tab"):
                    continue
                return False
            continue
        if child.tag in (qn("w:bookmarkStart"), qn("w:bookmarkEnd"), qn("w:proofErr")):
            continue
        return False
    return True


# --------------------------------------------------------------------------
# editing helpers
# --------------------------------------------------------------------------
def get_or_add(parent, tag: str, before: Tuple[str, ...] = ()):  # noqa: D401
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        parent.append(el)
    return el


# Order of children inside w:pPr (CT_PPr) so inserted elements stay schema-valid.
_PPR_ORDER = ["w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr",
              "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd",
              "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
              "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
              "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
              "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
              "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
              "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange"]

_RPR_ORDER = ["w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps",
              "w:smallCaps", "w:strike", "w:dstrike", "w:outline", "w:shadow",
              "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid", "w:vanish",
              "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern", "w:position",
              "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd",
              "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
              "w:eastAsianLayout", "w:specVanish", "w:oMath"]


def _insert_ordered(parent, el, order: List[str]):
    tag = el.tag
    names = [qn(t) for t in order]
    idx = names.index(tag) if tag in names else len(names)
    for i, child in enumerate(list(parent)):
        if child.tag in names and names.index(child.tag) > idx:
            parent.insert(i, el)
            return el
    parent.append(el)
    return el


def ppr_child(p_el, tag: str, create: bool = True):
    ppr = p_el.find(qn("w:pPr"))
    if ppr is None:
        if not create:
            return None
        ppr = OxmlElement("w:pPr")
        p_el.insert(0, ppr)
    el = ppr.find(qn(tag))
    if el is None and create:
        el = OxmlElement(tag)
        _insert_ordered(ppr, el, _PPR_ORDER)
    return el


def rpr_child(r_el, tag: str, create: bool = True):
    rpr = r_el.find(qn("w:rPr"))
    if rpr is None:
        if not create:
            return None
        rpr = OxmlElement("w:rPr")
        r_el.insert(0, rpr)
    el = rpr.find(qn(tag))
    if el is None and create:
        el = OxmlElement(tag)
        _insert_ordered(rpr, el, _RPR_ORDER)
    return el


def set_jc(p_el, align: str):
    val = {"LEFT": "left", "CENTER": "center", "RIGHT": "right", "JUSTIFY": "both"}[align]
    ppr_child(p_el, "w:jc").set(qn("w:val"), val)


def set_spacing(p_el, before=None, after=None, line=None):
    sp = ppr_child(p_el, "w:spacing")
    if before is not None:
        sp.set(qn("w:before"), str(int(round(before * 20))))
        for a in ("w:beforeLines", "w:beforeAutospacing"):
            if sp.get(qn(a)) is not None:
                del sp.attrib[qn(a)]
    if after is not None:
        sp.set(qn("w:after"), str(int(round(after * 20))))
        for a in ("w:afterLines", "w:afterAutospacing"):
            if sp.get(qn(a)) is not None:
                del sp.attrib[qn(a)]
    if line is not None:
        sp.set(qn("w:line"), str(int(line)))
        sp.set(qn("w:lineRule"), "auto")


def set_indent(p_el, left=None, first=None, hanging=None):
    ind = ppr_child(p_el, "w:ind")
    for a in ("w:start", "w:startChars", "w:leftChars", "w:firstLineChars", "w:hangingChars"):
        if ind.get(qn(a)) is not None:
            del ind.attrib[qn(a)]
    if left is not None:
        ind.set(qn("w:left"), str(int(left)))
    if first is not None:
        ind.set(qn("w:firstLine"), str(int(first)))
        if ind.get(qn("w:hanging")) is not None:
            del ind.attrib[qn("w:hanging")]
    if hanging is not None:
        ind.set(qn("w:hanging"), str(int(hanging)))
        if ind.get(qn("w:firstLine")) is not None:
            del ind.attrib[qn("w:firstLine")]


def set_run_font(r_el, name: str):
    text = run_text(r_el)
    rf = rpr_child(r_el, "w:rFonts")
    for a in ("w:asciiTheme", "w:hAnsiTheme"):
        if rf.get(qn(a)) is not None:
            del rf.attrib[qn(a)]
    rf.set(qn("w:ascii"), name)
    rf.set(qn("w:hAnsi"), name)
    if not _CJK_RE.search(text):
        for a in ("w:eastAsiaTheme",):
            if rf.get(qn(a)) is not None:
                del rf.attrib[qn(a)]
        rf.set(qn("w:eastAsia"), name)
        if rf.get(qn("w:hint")) is not None:
            del rf.attrib[qn("w:hint")]
    if not _CS_RE.search(text):
        if rf.get(qn("w:cstheme")) is not None:
            del rf.attrib[qn("w:cstheme")]
        rf.set(qn("w:cs"), name)


def set_run_size(r_el, pt: float):
    half = str(int(round(pt * 2)))
    rpr_child(r_el, "w:sz").set(qn("w:val"), half)
    rpr_child(r_el, "w:szCs").set(qn("w:val"), half)


def set_run_toggle(r_el, tag: str, value: bool):
    """tag is w:b or w:i (complex-script twin kept in sync)."""
    for t in (tag, tag + "Cs"):
        el = rpr_child(r_el, t)
        if value:
            if el.get(qn("w:val")) is not None:
                del el.attrib[qn("w:val")]
        else:
            el.set(qn("w:val"), "0")


def replace_text_span(p_el, start: int, end: int, new_text: str) -> None:
    """Replace characters [start, end) of the paragraph's run text with
    new_text, keeping the formatting of the run where the span starts."""
    pos = 0
    first_done = False
    for r in run_elements(p_el):
        if run_has_drawing(r):
            continue
        ts = r.findall(qn("w:t"))
        if not ts:
            continue
        for t in ts:
            txt = t.text or ""
            t_start, t_end = pos, pos + len(txt)
            pos = t_end
            if t_end <= start or t_start >= end:
                continue
            a = max(start, t_start) - t_start
            b = min(end, t_end) - t_start
            if not first_done:
                t.text = txt[:a] + new_text + txt[b:]
                first_done = True
            else:
                t.text = txt[:a] + txt[b:]
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    if not first_done and start >= pos:
        # span lies at the very end: append a new run
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = new_text
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r.append(t)
        p_el.append(r)


def regex_replace_in_paragraph(p_el, pattern, func: Callable) -> int:
    """Apply regex replacement across run boundaries; returns replacements."""
    text = "".join(run_text(r) for r in run_elements(p_el) if not run_has_drawing(r))
    matches = list(pattern.finditer(text))
    count = 0
    for m in reversed(matches):
        new = func(m)
        if new is None or new == m.group(0):
            continue
        replace_text_span(p_el, m.start(), m.end(), new)
        count += 1
    return count


def paragraph_is_auto_numbered(document, p_el) -> bool:
    """Direct numPr wins over the style's (numId 0 switches numbering off)."""
    ppr = p_el.find(qn("w:pPr"))
    if ppr is not None:
        num = ppr.find(qn("w:numPr"))
        if num is not None:
            nid = num.find(qn("w:numId"))
            if nid is not None:
                return nid.get(qn("w:val")) not in (None, "0")
    return style_has_numbering(document, p_el)
