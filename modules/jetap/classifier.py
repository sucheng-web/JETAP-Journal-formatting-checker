"""Role classifier for JETAP manuscripts.

Position-based first (a title is the first substantive line, the author block
sits between the title and the abstract), then text patterns.  Format alone
(bold, centred ...) is never enough to decide a role - that is exactly the
information the fixer is going to overwrite.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

from docx.oxml.ns import qn

from modules.paragraph_classifier import ClassifiedParagraph, ParagraphType
from . import spec
from .helpers import (
    from_roman, is_plain_empty, paragraph_has_drawing, paragraph_has_math,
    paragraph_has_sectpr, paragraph_style_name, paragraph_text, run_elements,
    effective_run_props, paragraph_is_auto_numbered,
)

_NUM_ROMAN = re.compile(r"^\s*([IVXLC]{1,6})[\.\)]\s+(\S.*)$")
_NUM_MULTI = re.compile(r"^\s*(\d{1,2}(?:\.\d{1,2}){1,3})[\.\)]?\s+(\S.*)$")
_NUM_ARAB = re.compile(r"^\s*(\d{1,2})[\.\)]?\s+(\S.*)$")
_NUM_LETTER = re.compile(r"^\s*([A-Z])[\.\)]\s+(\S.*)$")
_ABSTRACT_RE = re.compile(r"^\s*abstract\b\s*(?:[—–\-:.]\s*)?(.*)$", re.I | re.S)
_KEYWORDS_RE = re.compile(r"^\s*(?:index\s+terms|key\s*words?)\b\s*(?:[—–\-:.]\s*)?(.*)$", re.I | re.S)
_FIG_CAP = re.compile(r"^\s*(?:fig(?:ure)?s?)\.?\s*(\d+)\s*([\.\:\-–—])?", re.I)
_TAB_CAP = re.compile(r"^\s*(?:table|tab\.)\s*\.?\s*(\d+|[IVXLC]{1,6})\b\s*([\.\:\-–—])?", re.I)
_DOI_RE = re.compile(r"^\s*(?:doi\s*:?\s*)?(?:https?://)?(?:dx\.)?doi\.org/|^\s*doi\s*:", re.I)
_REV_RE = re.compile(r"^\s*(?:manuscript\s+)?(?:received|submitted)\s*:|\baccepted\s*:|\bpublished\s*:", re.I)
_CORR_RE = re.compile(r"corresponding\s+author|\bcorrespondence\b|orcid", re.I)
_EQ_NUM = re.compile(r"\(\s*\d+\s*\)\s*$")
_SMALL_WORDS = {"a", "an", "and", "as", "at", "but", "by", "for", "from", "in", "into", "of",
                "on", "or", "the", "to", "via", "with", "vs", "versus", "its", "per"}


@dataclass
class JetapPara:
    index: int
    el: Any
    role: str
    text: str
    level: int = 0                 # heading level 1-3
    number_prefix: str = ""        # typed heading/reference number, e.g. "2.1"
    number_len: int = 0            # characters of the text it spans (incl. spaces)
    component: bool = False        # un-numbered component head
    has_auto_number: bool = False
    note: str = ""
    extra: Dict[str, Any] = field(default_factory=dict)


def _norm(text: str) -> str:
    return re.sub(r"[\s:.—–\-]+$", "", text.strip().lower())


def _word_count(text: str) -> int:
    return len(text.split())


def _headingish(text: str, bold: bool, styled: bool, known: bool) -> bool:
    """Gate that keeps numbered body sentences / list items from being headings."""
    t = text.strip()
    if not t or len(t) > 110 or _word_count(t) > 14:
        return False
    if known or styled:
        return True
    if t.endswith((".", ";", ",", "?", "!")) and not t.isupper():
        return False
    if t.endswith(":") and _word_count(t) > 6:
        return False
    if not t[0].isupper():
        return False
    words = [w for w in re.split(r"\s+", t) if re.search(r"[A-Za-z]", w)]
    if not words:
        return False
    capped = sum(1 for w in words if w[0].isupper() or w.lower() in _SMALL_WORDS)
    return bold or t.isupper() or capped / len(words) >= 0.75


class JetapClassifier:
    def __init__(self, profile: Optional[Dict[str, Any]] = None):
        self.profile = profile or {}

    # ------------------------------------------------------------------
    def classify(self, document) -> List[JetapPara]:
        paras = list(document.paragraphs)
        n = len(paras)
        texts = [paragraph_text(p._p).strip() for p in paras]
        out: List[Optional[JetapPara]] = [None] * n

        abstract_idx = None
        for i, t in enumerate(texts[:80]):
            if t and _ABSTRACT_RE.match(t):
                m = _ABSTRACT_RE.match(t)
                label_only = not m.group(1).strip()
                if label_only or re.match(r"^\s*abstract\s*[—–\-:.]", t, re.I) \
                        or re.match(r"^\s*abstract\s+\S", t, re.I) and len(t) > 60:
                    abstract_idx = i
                    break
        keywords_idx = next((i for i, t in enumerate(texts[:120]) if t and _KEYWORDS_RE.match(t)), None)

        first_heading_idx = None
        for i, t in enumerate(texts):
            if not t or (abstract_idx is not None and i <= abstract_idx):
                continue
            if keywords_idx is not None and i <= keywords_idx:
                continue
            if self._heading_info(document, paras[i], t)[0] is not None:
                first_heading_idx = i
                break

        front_end = next((x for x in (abstract_idx, keywords_idx, first_heading_idx) if x is not None), min(n, 12))

        # ---- front matter ------------------------------------------------
        title_idx = None
        seen_affiliation = False
        for i in range(front_end):
            p = paras[i]
            t = texts[i]
            if paragraph_has_sectpr(p._p) and not t:
                out[i] = JetapPara(i, p._p, "section_break", t)
                continue
            if not t:
                out[i] = JetapPara(i, p._p, "empty" if is_plain_empty(p._p) else "other", t)
                continue
            low = t.lower()
            if title_idx is None:
                if low.startswith("journal of") and len(t) < 90 and not re.search(r"\d{3,}", t):
                    out[i] = JetapPara(i, p._p, "masthead", t)
                    continue
                if re.match(r"^\s*(vol(ume)?\.?\s*\d|e-?issn|issn|©|copyright)", low) \
                        and _word_count(t) < 25:
                    out[i] = JetapPara(i, p._p, "journal_info", t)
                    continue
                title_idx = i
                out[i] = JetapPara(i, p._p, "title", t)
                continue
            if _DOI_RE.match(t):
                out[i] = JetapPara(i, p._p, "doi", t)
            elif _REV_RE.search(t):
                out[i] = JetapPara(i, p._p, "revision", t)
            elif _CORR_RE.search(t) or (t.startswith("*") and "@" in t) or re.search(r"^\W*[\w.\-]+@[\w.\-]+\.\w+\W*$", t):
                out[i] = JetapPara(i, p._p, "corresponding", t)
            elif seen_affiliation or self._looks_like_affiliation(t):
                seen_affiliation = True
                out[i] = JetapPara(i, p._p, "affiliation", t)
            else:
                out[i] = JetapPara(i, p._p, "authors", t)

        # ---- abstract / keywords -------------------------------------------
        i = front_end
        if abstract_idx is not None:
            m = _ABSTRACT_RE.match(texts[abstract_idx])
            label_only = not m.group(1).strip()
            out[abstract_idx] = JetapPara(abstract_idx, paras[abstract_idx]._p,
                                          "abstract_label" if label_only else "abstract",
                                          texts[abstract_idx], note="run-in" if not label_only else "label")
            j = abstract_idx + 1
            while j < n:
                t = texts[j]
                if not t:
                    if is_plain_empty(paras[j]._p) and j + 1 < n and (_KEYWORDS_RE.match(texts[j + 1] or "") or
                                                                       self._heading_info(document, paras[j + 1], texts[j + 1])[0] is not None):
                        break
                    if paragraph_has_sectpr(paras[j]._p):
                        break
                    j += 1
                    continue
                if _KEYWORDS_RE.match(t) or self._heading_info(document, paras[j], t)[0] is not None:
                    break
                out[j] = JetapPara(j, paras[j]._p, "abstract", t)
                j += 1
            i = j
        if keywords_idx is not None:
            m = _KEYWORDS_RE.match(texts[keywords_idx])
            out[keywords_idx] = JetapPara(keywords_idx, paras[keywords_idx]._p, "keywords", texts[keywords_idx])
            if not m.group(1).strip() and keywords_idx + 1 < n and texts[keywords_idx + 1] \
                    and self._heading_info(document, paras[keywords_idx + 1], texts[keywords_idx + 1])[0] is None:
                out[keywords_idx + 1] = JetapPara(keywords_idx + 1, paras[keywords_idx + 1]._p,
                                                  "keywords", texts[keywords_idx + 1], note="content")

        # ---- remainder -----------------------------------------------------
        in_refs = False
        in_decl = False
        pending_caption = None
        for i in range(n):
            if out[i] is not None:
                continue
            p = paras[i]
            t = texts[i]
            el = p._p
            if not t:
                if paragraph_has_sectpr(el):
                    out[i] = JetapPara(i, el, "section_break", t)
                elif paragraph_has_drawing(el):
                    out[i] = JetapPara(i, el, "figure", t)
                elif paragraph_has_math(el):
                    out[i] = JetapPara(i, el, "equation", t)
                else:
                    out[i] = JetapPara(i, el, "empty" if is_plain_empty(el) else "other", t)
                continue

            if pending_caption:
                out[i] = JetapPara(i, el, pending_caption, t, note="title-line")
                pending_caption = None
                continue
            lone = re.match(r"^\s*(table|tab\.|fig(?:ure)?\.?)\s*(\d+|[IVXLC]{1,6})\s*[.:]?\s*$", t, re.I)
            if lone and len(t) < 20:
                kind = "caption_table" if lone.group(1).lower().startswith("tab") else "caption_figure"
                out[i] = JetapPara(i, el, kind, t, note="label-only")
                pending_caption = kind
                continue
            if re.match(r"^\s*(?:notes?|source|sources)\s*[.:]", t, re.I) and len(t) < 300:
                out[i] = JetapPara(i, el, "table_note", t)
                continue

            level, prefix, plen, component, known = self._heading_info(document, p, t)
            if level is not None:
                norm = _norm(self._strip_prefix(t, plen))
                in_refs = norm in ("references", "reference", "bibliography")
                in_decl = component and norm not in ("references", "reference", "bibliography",
                                                     "appendix", "appendices")
                out[i] = JetapPara(i, el, f"heading{level}", t, level=level, number_prefix=prefix,
                                   number_len=plen, component=component,
                                   has_auto_number=paragraph_is_auto_numbered(document, el))
                continue

            cap = self._caption_kind(document, p, t)
            if cap:
                out[i] = JetapPara(i, el, cap, t)
                continue
            if paragraph_has_math(el) or self._looks_like_equation(el, t):
                out[i] = JetapPara(i, el, "equation", t)
                continue
            if paragraph_has_drawing(el) and len(t) < 20:
                out[i] = JetapPara(i, el, "figure", t)
                continue
            if in_refs:
                m = re.match(r"^\s*(\[\s*\d+\s*\]|\(\s*\d+\s*\)|\d+\s*[\.\)])\s*", t)
                out[i] = JetapPara(i, el, "reference", t, number_prefix=m.group(0) if m else "",
                                   number_len=len(m.group(0)) if m else 0,
                                   has_auto_number=paragraph_is_auto_numbered(document, el))
                continue
            if self._is_table_note(document, el, t):
                out[i] = JetapPara(i, el, "table_note", t)
                continue
            out[i] = JetapPara(i, el, "decl_body" if in_decl else "body", t)
        return [o for o in out if o is not None]

    # ------------------------------------------------------------------
    @staticmethod
    def _strip_prefix(text: str, plen: int) -> str:
        return text.strip()[plen:] if plen else text.strip()

    _AFF_STRONG = re.compile(
        r"\b(?:universiti|university|universität|universidad|faculty|fakulti|department|dept\.?|institute|"
        r"institut|school\s+of|college|centre|center|laborator\w*|sdn\.?\s*bhd|hospital|ltd|inc\.|gmbh|academy|"
        r"polytechnic|politeknik|campus|jalan|affiliation)\b", re.I)
    _AFF_WEAK = re.compile(
        r"\b(?:malaysia|singapore|indonesia|india|china|korea|japan|usa|united\s+states|united\s+kingdom|uk|"
        r"australia|thailand|vietnam|pakistan|iraq|iran|saudi\s+arabia|egypt|nigeria|kuala\s+lumpur|selangor|"
        r"penang|johor|melaka|cyberjaya)\b", re.I)

    def _looks_like_affiliation(self, t: str) -> bool:
        if self._AFF_STRONG.search(t):
            return True
        # a leading affiliation marker (superscript / digit) plus a country or city
        return bool(re.match(r"^\s*[¹²³⁰-⁹\d]+\s*[A-Za-z]", t) and self._AFF_WEAK.search(t))

    def _heading_info(self, document, p, text: str):
        """Return (level, prefix, prefix_len, component, known) or (None,)*5."""
        el = p._p
        t = text.strip()
        if re.match(r"^\s*(?:table|tab\.|fig(?:ure)?s?\.?)\s*\.?\s*(?:\d+|[IVXLC]{1,6})\b", t, re.I) or \
                re.match(r"^\s*(?:notes?|source|sources)\s*[.:]", t, re.I) or \
                re.match(r"^\s*(?:received|accepted|published|revised|submitted)\b[^.]{0,3}:", t, re.I):
            return None, "", 0, False, False
        style = paragraph_style_name(document, el).lower()
        styled_level = None
        m = re.match(r"heading\s*(\d)", style)
        if m:
            styled_level = min(int(m.group(1)), 3)
        runs = [r for r in run_elements(el) if "".join(x.text or "" for x in r.findall(qn("w:t"))).strip()]
        bold = bool(runs) and all(effective_run_props(document, el, r)["bold"] for r in runs)

        # numbered patterns ------------------------------------------------
        for rx, kind in ((_NUM_MULTI, "multi"), (_NUM_ROMAN, "roman"), (_NUM_ARAB, "arab"), (_NUM_LETTER, "letter")):
            mm = rx.match(t)
            if not mm:
                continue
            num, rest = mm.group(1), mm.group(2).strip()
            if kind == "roman" and from_roman(num) is None:
                continue
            if kind == "letter" and num == "I":
                continue
            lead = len(t) - len(t.lstrip())
            plen = mm.start(2) - lead if False else (mm.start(2))
            plen = mm.start(2)
            nrest = _norm(rest)
            known = nrest in spec.MAIN_SECTION_TERMS or nrest in spec.COMPONENT_HEADS
            if not _headingish(rest, bold, styled_level is not None, known):
                continue
            if kind == "multi":
                depth = num.count(".") + 1
                level = min(depth, 3)
                if re.fullmatch(r"\d{1,2}\.0+", num):      # "1.0 Introduction" style main heading
                    level = 1
            elif kind == "letter":
                level = 2
            else:
                level = 1
            if level == 1 and styled_level in (2, 3) and kind not in ("roman",):
                level = styled_level
            return level, num, plen, (nrest in spec.COMPONENT_HEADS and level == 1), known

        norm = _norm(t)
        ol = self._outline_level(document, el)
        if ol is not None and ol <= 2 and len(t) < 120 and _word_count(t) <= 16 and not t.endswith((".", ",", ";")):
            comp = norm in spec.COMPONENT_HEADS
            return ol + 1, "", 0, comp, True
        if norm in spec.COMPONENT_HEADS:
            return 1, "", 0, True, True
        if norm in spec.MAIN_SECTION_TERMS:
            return 1, "", 0, False, True
        if styled_level is not None and len(t) < 120 and _word_count(t) <= 16:
            return styled_level, "", 0, False, True
        if (t.isupper() and 1 <= _word_count(t) <= 8 and len(t) >= 4
                and not t.endswith(".") and re.search(r"[A-Z]{3}", t)):
            return 1, "", 0, False, False
        return None, "", 0, False, False

    @staticmethod
    def _outline_level(document, el) -> Optional[int]:
        from .helpers import effective_ppr_attr
        val = effective_ppr_attr(document, el, "w:outlineLvl")
        if val is None:
            return None
        try:
            v = int(val)
        except ValueError:
            return None
        return v if 0 <= v <= 8 else None

    def _caption_kind(self, document, p, t: str) -> Optional[str]:
        el = p._p
        if len(t) > 300:
            return None
        style = paragraph_style_name(document, el).lower()
        fm, tm = _FIG_CAP.match(t), _TAB_CAP.match(t)
        kind = "caption_figure" if fm else "caption_table" if tm else None
        if kind is None:
            if "caption" in style and re.search(r"\b(fig|table)", t, re.I):
                return "caption_table" if re.match(r"\s*tab", t, re.I) else "caption_figure"
            return None
        m = fm or tm
        if m.group(2) or "caption" in style:
            return kind
        # "Table 3 shows ..." is body text unless it sits next to a table/figure
        if _word_count(t) <= 30 and self._adjacent_object(el, kind):
            return kind
        return None

    @staticmethod
    def _is_table_note(document, el, t: str) -> bool:
        prv = el.getprevious()
        while prv is not None and prv.tag == qn("w:p") and is_plain_empty(prv):
            prv = prv.getprevious()
        if prv is None or prv.tag != qn("w:tbl") or len(t) > 200:
            return False
        style = paragraph_style_name(document, el).lower()
        return "footnote" in style or bool(re.match(r"^\s*(?:(?:note|notes|source|sources)\b|\*|[a-z]\))", t, re.I))

    @staticmethod
    def _adjacent_object(el, kind: str) -> bool:
        def sib(e, direction):
            e = e.getprevious() if direction < 0 else e.getnext()
            while e is not None and e.tag == qn("w:p") and is_plain_empty(e):
                e = e.getprevious() if direction < 0 else e.getnext()
            return e

        for d in (-1, 1):
            s = sib(el, d)
            if s is None:
                continue
            if s.tag == qn("w:tbl") and kind == "caption_table":
                return True
            if s.tag == qn("w:p") and paragraph_has_drawing(s) and kind == "caption_figure":
                return True
        return False

    @staticmethod
    def _looks_like_equation(el, t: str) -> bool:
        if "\t" not in paragraph_text(el):
            return False
        if not _EQ_NUM.search(t) or len(t) > 160:
            return False
        letters = len(re.findall(r"[A-Za-z]{3,}", t))
        return letters <= 3

    # ------------------------------------------------------------------
    TYPE_MAP = {
        "masthead": ParagraphType.JOURNAL_HEADER, "journal_info": ParagraphType.JOURNAL_HEADER,
        "title": ParagraphType.PAPER_TITLE,
        "authors": ParagraphType.AUTHOR_INFO, "affiliation": ParagraphType.AUTHOR_INFO,
        "corresponding": ParagraphType.AUTHOR_INFO, "doi": ParagraphType.AUTHOR_INFO,
        "revision": ParagraphType.AUTHOR_INFO,
        "abstract_label": ParagraphType.ABSTRACT_LABEL, "abstract": ParagraphType.ABSTRACT_CONTENT,
        "keywords": ParagraphType.KEYWORDS_CONTENT,
        "heading1": ParagraphType.SECTION_HEADING, "heading2": ParagraphType.SECTION_HEADING,
        "heading3": ParagraphType.SECTION_HEADING,
        "body": ParagraphType.BODY, "decl_body": ParagraphType.BODY, "table_note": ParagraphType.BODY,
        "caption_table": ParagraphType.CAPTION,
        "caption_figure": ParagraphType.CAPTION, "reference": ParagraphType.REFERENCE,
        "equation": ParagraphType.EQUATION, "figure": ParagraphType.FIGURE,
        "empty": ParagraphType.EMPTY, "section_break": ParagraphType.EMPTY,
        "other": ParagraphType.UNKNOWN,
    }

    def to_classified(self, document, jpars: List[JetapPara]) -> List[ClassifiedParagraph]:
        result = []
        for jp in jpars:
            info: Dict[str, Any] = {}
            for r in run_elements(jp.el):
                if "".join(x.text or "" for x in r.findall(qn("w:t"))).strip():
                    props = effective_run_props(document, jp.el, r)
                    info = {"font_name": props["font"], "font_size": props["size"],
                            "bold": props["bold"], "italic": props["italic"]}
                    break
            from .helpers import effective_alignment
            result.append(ClassifiedParagraph(
                index=jp.index, text=jp.text,
                paragraph_type=self.TYPE_MAP.get(jp.role, ParagraphType.UNKNOWN),
                confidence=0.9, font_info=info,
                alignment=effective_alignment(document, jp.el),
                should_fix=jp.role not in ("empty", "section_break", "other", "journal_info"),
                classification_reason=f"jetap:{jp.role}",
            ))
        return result
