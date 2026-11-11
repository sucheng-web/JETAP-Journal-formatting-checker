"""Document-level building blocks for the JETAP fixer: numbering definitions,
masthead table, section layout, running header and the first-page footer."""
from __future__ import annotations

import re
from copy import deepcopy
from io import BytesIO
from typing import Any, Dict, List, Optional

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

from . import spec
from .helpers import run_text

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def E(tag: str, **attrs):
    el = OxmlElement(tag if ":" in tag else "w:" + tag)
    for k, v in attrs.items():
        el.set(qn("w:" + k), str(v))
    return el


# --------------------------------------------------------------------------
# numbering
# --------------------------------------------------------------------------
def _numbering_root(document):
    return document.part.numbering_part.element


def _next_id(root, tag: str, attr: str) -> int:
    ids = [int(e.get(qn(attr))) for e in root.findall(qn(tag)) if e.get(qn(attr)) is not None]
    return (max(ids) + 1) if ids else 1


def _add_abstract_and_num(root, abstract_xml: str) -> int:
    aid = _next_id(root, "w:abstractNum", "w:abstractNumId")
    nid = _next_id(root, "w:num", "w:numId")
    abstract = parse_xml(abstract_xml.replace("{AID}", str(aid)))
    nums = root.findall(qn("w:num"))
    if nums:
        nums[0].addprevious(abstract)
    else:
        root.append(abstract)
    num = parse_xml(f'<w:num {nsdecls("w")} w:numId="{nid}"><w:abstractNumId w:val="{aid}"/></w:num>')
    root.append(num)
    return nid


_FONT_RPR = ('<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
             'w:eastAsia="Times New Roman" w:cs="Times New Roman"/>{extra}</w:rPr>')


def ensure_heading_numbering(document) -> int:
    """Multilevel list: I. / A. / 1) - restarting each level automatically."""
    root = _numbering_root(document)
    xml = (
        f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{{AID}}"><w:multiLevelType w:val="multilevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:suff w:val="tab"/>'
        '<w:lvlText w:val="%1."/><w:lvlJc w:val="right"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="340"/></w:tabs>'
        '<w:ind w:left="0" w:firstLine="216"/></w:pPr>'
        + _FONT_RPR.format(extra='<w:b w:val="0"/><w:i w:val="0"/><w:sz w:val="20"/><w:szCs w:val="20"/>') +
        '</w:lvl>'
        '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="upperLetter"/><w:suff w:val="tab"/>'
        '<w:lvlText w:val="%2."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="284"/></w:tabs>'
        '<w:ind w:left="360" w:hanging="360"/></w:pPr>'
        + _FONT_RPR.format(extra='<w:b w:val="0"/><w:i/><w:sz w:val="20"/><w:szCs w:val="20"/>') +
        '</w:lvl>'
        '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:suff w:val="tab"/>'
        '<w:lvlText w:val="%3)"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
        + _FONT_RPR.format(extra='<w:b w:val="0"/><w:i/><w:sz w:val="20"/><w:szCs w:val="20"/>') +
        '</w:lvl></w:abstractNum>'
    )
    return _add_abstract_and_num(root, xml)


def ensure_reference_numbering(document) -> int:
    root = _numbering_root(document)
    xml = (
        f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{{AID}}"><w:multiLevelType w:val="singleLevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:suff w:val="tab"/>'
        '<w:lvlText w:val="[%1]"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="284"/></w:tabs>'
        '<w:ind w:left="284" w:hanging="284"/></w:pPr>'
        + _FONT_RPR.format(extra='<w:b w:val="0"/><w:i w:val="0"/><w:sz w:val="16"/><w:szCs w:val="16"/>') +
        '</w:lvl></w:abstractNum>'
    )
    return _add_abstract_and_num(root, xml)


def set_numbering(p_el, num_id: int, level: int = 0):
    from .helpers import ppr_child
    num_pr = ppr_child(p_el, "w:numPr")
    for ch in list(num_pr):
        num_pr.remove(ch)
    num_pr.append(E("ilvl", val=level))
    num_pr.append(E("numId", val=num_id))


def clear_numbering(p_el):
    from .helpers import ppr_child
    num_pr = ppr_child(p_el, "w:numPr")
    for ch in list(num_pr):
        num_pr.remove(ch)
    num_pr.append(E("ilvl", val=0))
    num_pr.append(E("numId", val=0))


def set_paragraph_mark(p_el, size: float, bold: bool, italic: bool):
    """Paragraph-mark run properties drive how an automatic number looks."""
    from .helpers import ppr_child, rpr_child
    mark = ppr_child(p_el, "w:rPr")
    for ch in list(mark):
        mark.remove(ch)
    rf = E("rFonts", ascii=spec.FONT, hAnsi=spec.FONT, eastAsia=spec.FONT, cs=spec.FONT)
    mark.append(rf)
    for tag, on in (("b", bold), ("i", italic)):
        el = E(tag)
        if not on:
            el.set(qn("w:val"), "0")
        mark.append(el)
    half = str(int(round(size * 2)))
    mark.append(E("sz", val=half))
    mark.append(E("szCs", val=half))


# --------------------------------------------------------------------------
# elements built from the template spec
# --------------------------------------------------------------------------
def make_run(text: str, size: float, bold=False, italic=False):
    r = E("r")
    rpr = E("rPr")
    rpr.append(E("rFonts", ascii=spec.FONT, hAnsi=spec.FONT, eastAsia=spec.FONT, cs=spec.FONT))
    for tag, on in (("b", bold), ("i", italic)):
        if on:
            rpr.append(E(tag))
    rpr.append(E("sz", val=int(round(size * 2))))
    rpr.append(E("szCs", val=int(round(size * 2))))
    r.append(rpr)
    t = E("t")
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)
    return r


def make_paragraph(text: str, role: str, runs: Optional[List[Any]] = None):
    from .helpers import set_jc, set_spacing, set_indent
    want = spec.ROLES[role]
    p = E("p")
    set_spacing(p, before=want["before"], after=want["after"], line=want["line"])
    if want.get("first") is not None:
        set_indent(p, left=0, first=want["first"])
    set_jc(p, want["align"])
    if runs is None:
        runs = [make_run(text, want["size"], bool(want["bold"]), bool(want["italic"]))]
    for r in runs:
        p.append(r)
    return p


def make_blank(after_pt: float = 6, line: int = 228, size: float = 10):
    from .helpers import set_spacing
    p = E("p")
    set_spacing(p, before=0, after=after_pt, line=line)
    ppr = p.find(qn("w:pPr"))
    mark = E("rPr")
    mark.append(E("rFonts", ascii=spec.FONT, hAnsi=spec.FONT))
    mark.append(E("sz", val=int(size * 2)))
    mark.append(E("szCs", val=int(size * 2)))
    ppr.append(mark)
    return p


def make_masthead_table():
    xml = (
        f'<w:tbl {nsdecls("w")}><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
        '<w:tblBorders><w:top w:val="single" w:sz="8" w:space="0" w:color="auto"/>'
        '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="thinThickLargeGap" w:sz="24" w:space="0" w:color="auto"/>'
        '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:tblBorders>'
        '<w:tblLayout w:type="autofit"/>'
        '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
        '</w:tblPr><w:tblGrid><w:gridCol w:w="9026"/></w:tblGrid><w:tr><w:tc><w:tcPr>'
        '<w:tcW w:w="9026" w:type="dxa"/></w:tcPr>'
        '<w:p><w:pPr><w:spacing w:before="0" w:after="120" w:line="240" w:lineRule="auto"/>'
        '<w:ind w:left="0" w:firstLine="0"/><w:jc w:val="center"/></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>'
        '<w:b/><w:bCs/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr>'
        f'<w:t>{spec.JOURNAL_NAME}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    )
    return parse_xml(xml)


# --------------------------------------------------------------------------
# sections
# --------------------------------------------------------------------------
def all_sectprs(document) -> List[Any]:
    out = []
    for el in document.element.body.iter(qn("w:sectPr")):
        out.append(el)
    return out


def set_cols(sect, columns: int):
    old = sect.find(qn("w:cols"))
    new = E("cols", space=720) if columns == 1 else parse_xml(
        f'<w:cols {nsdecls("w")} w:num="2" w:space="720" w:equalWidth="0">'
        f'<w:col w:w="{spec.PAGE["column_width_tw"]}" w:space="{spec.PAGE["column_gap_tw"]}"/>'
        f'<w:col w:w="{spec.PAGE["column_width_tw"]}" w:space="0"/></w:cols>')
    if old is not None:
        old.addprevious(new)
        sect.remove(old)
    else:
        # schema order: ... pgMar, paperSrc, pgBorders, lnNumType, pgNumType, cols, ...
        anchor = None
        for tag in ("w:pgNumType", "w:lnNumType", "w:pgBorders", "w:paperSrc", "w:pgMar", "w:pgSz"):
            anchor = sect.find(qn(tag))
            if anchor is not None:
                break
        if anchor is not None:
            anchor.addnext(new)
        else:
            sect.append(new)


def set_section_type(sect, kind: Optional[str]):
    old = sect.find(qn("w:type"))
    if old is not None:
        sect.remove(old)
    if kind:
        el = E("type", val=kind)
        refs = [c for c in sect if c.tag in (qn("w:headerReference"), qn("w:footerReference"), qn("w:footnotePr"), qn("w:endnotePr"))]
        if refs:
            refs[-1].addnext(el)
        else:
            sect.insert(0, el)


def apply_page_geometry(sect):
    pg_sz = sect.find(qn("w:pgSz"))
    landscape = pg_sz is not None and (pg_sz.get(qn("w:orient")) == "landscape")
    if pg_sz is None:
        pg_sz = E("pgSz")
        sect.insert(0, pg_sz)
    if not landscape:
        pg_sz.set(qn("w:w"), str(spec.PAGE["width_tw"]))
        pg_sz.set(qn("w:h"), str(spec.PAGE["height_tw"]))
        if pg_sz.get(qn("w:orient")):
            del pg_sz.attrib[qn("w:orient")]
    mar = sect.find(qn("w:pgMar"))
    if mar is None:
        mar = E("pgMar")
        pg_sz.addnext(mar)
    for side in ("top", "right", "bottom", "left"):
        mar.set(qn("w:" + side), str(spec.PAGE["margin_tw"]))
    mar.set(qn("w:header"), str(spec.PAGE["header_tw"]))
    mar.set(qn("w:footer"), str(spec.PAGE["footer_tw"]))
    mar.set(qn("w:gutter"), "0")


def strip_header_footer_refs(sect):
    for tag in ("w:headerReference", "w:footerReference", "w:titlePg"):
        for el in sect.findall(qn(tag)):
            sect.remove(el)


def new_section_break_paragraph(template_sect, columns: int, kind: Optional[str], role: str):
    """Empty paragraph carrying a copy of template_sect's page setup."""
    from .helpers import set_jc, set_spacing
    p = E("p")
    sect = deepcopy(template_sect)
    strip_header_footer_refs(sect)
    for tag in ("w:pgNumType",):
        for el in sect.findall(qn(tag)):
            sect.remove(el)
    apply_page_geometry(sect)
    set_cols(sect, columns)
    set_section_type(sect, kind)
    if role == "front":
        set_spacing(p, before=0, after=0, line=240)
        set_jc(p, "CENTER")
        mark = E("rPr")
        mark.append(E("sz", val=16))
        mark.append(E("szCs", val=16))
        p.find(qn("w:pPr")).append(mark)
    else:
        set_spacing(p, before=0, after=6, line=228)
    p.find(qn("w:pPr")).append(sect)
    return p


# --------------------------------------------------------------------------
# running header
# --------------------------------------------------------------------------
def build_header_paragraph(vol: str, issue: str, year: str):
    p = E("p")
    ppr = E("pPr")
    tabs = E("tabs")
    tabs.append(E("tab", val="center", pos=4513))
    tabs.append(E("tab", val="right", pos=spec.PAGE["text_width_tw"]))
    ppr.append(tabs)
    ppr.append(E("spacing", after=0, line=240, lineRule="auto"))
    ppr.append(E("ind", firstLine=0))
    ppr.append(E("jc", val="left"))
    p.append(ppr)
    p.append(make_run(f"Vol {vol} No {issue} ({year})", 10))
    for _ in range(2):
        r = E("r")
        rpr = E("rPr")
        rpr.append(E("rFonts", ascii=spec.FONT, hAnsi=spec.FONT, cs=spec.FONT))
        rpr.append(E("sz", val=20))
        r.append(rpr)
        r.append(E("tab"))
        p.append(r)
    p.append(make_run("e", 10, italic=True))
    p.append(make_run("-ISSN: 2682-8383", 10))
    return p


def fill_header(header, vol: str, issue: str, year: str):
    header.is_linked_to_previous = False
    el = header._element
    for ch in list(el):
        el.remove(ch)
    el.append(build_header_paragraph(vol, issue, year))


def empty_footer(footer):
    footer.is_linked_to_previous = False
    el = footer._element
    for ch in list(el):
        el.remove(ch)
    p = E("p")
    el.append(p)


# --------------------------------------------------------------------------
# first-page footer (MMU logo + citation / DOI / licence block)
# --------------------------------------------------------------------------
_R_ATTRS = ("embed", "id", "link", "pict", "relid")
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
O_NS = "urn:schemas-microsoft-com:office:office"


def clone_first_page_footer(section, template_doc, meta: Dict[str, str]) -> None:
    tpl = template_doc.sections[0].first_page_footer
    tpl_part = tpl.part
    footer = section.first_page_footer
    footer.is_linked_to_previous = False
    dst_part = footer.part
    new_doi = meta.get("doi_url")

    rid_map: Dict[str, str] = {}
    for rid, rel in tpl_part.rels.items():
        if rel.is_external:
            target = rel.target_ref
            if new_doi and "doi.org/10.33093/jetap" in target:
                target = new_doi
            rid_map[rid] = dst_part.relate_to(target, rel.reltype, is_external=True)
        elif "image" in rel.reltype:
            new_rid, _ = dst_part.get_or_add_image(BytesIO(rel.target_part.blob))
            rid_map[rid] = new_rid

    ftr = footer._element
    for ch in list(ftr):
        ftr.remove(ch)
    for child in tpl._element:
        new_child = deepcopy(child)
        for node in new_child.iter():
            for ns in (R_NS, O_NS):
                for att in _R_ATTRS:
                    key = "{%s}%s" % (ns, att)
                    val = node.get(key)
                    if val in rid_map:
                        node.set(key, rid_map[val])
        ftr.append(new_child)

    # fonts: the template's own text box has no explicit font, so pin it
    for rpr in ftr.iter(qn("w:rPr")):
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = E("rFonts")
            rpr.insert(0, rf)
        for att in ("ascii", "hAnsi", "cs", "eastAsia"):
            rf.set(qn("w:" + att), spec.FONT)
        for att in ("asciiTheme", "hAnsiTheme", "cstheme", "eastAsiaTheme"):
            if rf.get(qn("w:" + att)) is not None:
                del rf.attrib[qn("w:" + att)]

    # the template paragraphs use the template's "Footer" style, which a manuscript may not have
    # (or may name differently): drop the style and pin what it defined, so spacing is identical
    from .helpers import set_spacing, set_indent
    for p in ftr.iter(qn("w:p")):
        ppr = p.find(qn("w:pPr"))
        if ppr is not None:
            ps = ppr.find(qn("w:pStyle"))
            if ps is not None:
                ppr.remove(ps)
        if p.find(".//" + qn("w:p")) is not None:
            continue
        set_spacing(p, before=0, after=0, line=240)
        sp = p.find(qn("w:pPr")).find(qn("w:spacing"))
        sp.set(qn("w:lineRule"), "auto")
        set_indent(p, left=0, first=0)
    for rpr in ftr.iter(qn("w:rPr")):
        rs = rpr.find(qn("w:rStyle"))
        if rs is not None and rs.get(qn("w:val")) in ("Hyperlink", "18"):
            rpr.remove(rs)
            for tag, attrs in (("w:color", {"w:val": "0563C1"}), ("w:u", {"w:val": "single"})):
                if rpr.find(qn(tag)) is None:
                    el = E(tag.split(":")[1], **{k.split(":")[1]: v for k, v in attrs.items()})
                    sz = rpr.find(qn("w:sz"))
                    (sz.addprevious(el) if sz is not None else rpr.append(el))

    # citation line + DOI text (the template's example values are replaced
    # only when the real ones are known)
    year, vol, issue, pages = meta.get("year"), meta.get("volume"), meta.get("issue"), meta.get("pages")
    for p in ftr.iter(qn("w:p")):
        if p.find(".//" + qn("w:p")) is not None:   # outer paragraph holding the text box
            continue
        ts = list(p.iter(qn("w:t")))
        joined = "".join(t.text or "" for t in ts)
        if joined.startswith("Journal of Engineering Technology and Applied Physics (") and year and vol and issue:
            ts[0].text = f"{spec.JOURNAL_NAME} ({year}) {vol}, {issue}, {pages or '1:1-5'}"
            for t in ts[1:]:
                t.text = ""
        elif new_doi and "doi.org/10.33093/jetap" in joined:
            ts[0].text = new_doi
            for t in ts[1:]:
                t.text = ""


# --------------------------------------------------------------------------
# pagination removal
# --------------------------------------------------------------------------
_PAGE_RE = re.compile(r"\b(PAGE|NUMPAGES)\b")


def _has_page_field(el) -> bool:
    for it in el.iter(qn("w:instrText")):
        if _PAGE_RE.search(it.text or ""):
            return True
    for fs in el.iter(qn("w:fldSimple")):
        if _PAGE_RE.search(fs.get(qn("w:instr")) or ""):
            return True
    return False


def remove_page_fields(container) -> int:
    """Remove page-number paragraphs/content controls from a header/footer."""
    removed = 0
    for child in list(container):
        if _has_page_field(child):
            container.remove(child)
            removed += 1
    if not len(container):
        container.append(E("p"))
    return removed


def drop_unused_header_footer_parts(document) -> None:
    """Remove relationships to header/footer parts no section references any more."""
    used = set()
    for sect in all_sectprs(document):
        for tag in ("w:headerReference", "w:footerReference"):
            for ref in sect.findall(qn(tag)):
                rid = ref.get(qn("r:id"))
                if rid:
                    used.add(rid)
    part = document.part
    for rid, rel in list(part.rels.items()):
        if rel.reltype.endswith("/header") or rel.reltype.endswith("/footer"):
            if rid not in used:
                part.rels.pop(rid)
