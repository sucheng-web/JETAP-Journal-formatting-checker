"""JETAP auto-fixer.

Applies the JETAP template to a manuscript.  Every rule is read from
`spec.ROLES`; nothing is keyed to one sample's wording.  Formatting is
written as direct formatting (the manuscript's own style ids can't be
trusted), preserving author content: bold/italic/underline on body runs,
sub/superscripts, hyperlinks, tabs, drawings, fields and the text itself.
"""
from __future__ import annotations

import re
from copy import deepcopy
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn

from modules.auto_fixer import ChangeRecord
from modules.utils import load_document
from . import spec, structure as st
from . import references as refs_mod
from .checker import _FIG_LABEL, _TAB_LABEL, section_map, sect_columns
from .classifier import JetapClassifier, JetapPara
from .helpers import (
    effective_alignment, effective_indent, effective_run_props, effective_spacing,
    from_roman, has_numbering, is_plain_empty, paragraph_has_drawing, paragraph_has_sectpr,
    paragraph_is_auto_numbered, paragraph_text, ppr_child, regex_replace_in_paragraph,
    replace_text_span, run_elements, run_has_drawing, run_text, set_indent, set_jc,
    set_run_font, set_run_size, set_run_toggle, set_spacing, to_roman,
)

TEMPLATE_PATH = Path(__file__).resolve().parent.parent.parent / "template_profiles" / "assets" / "jetap_template.docx"

_TYPED_HEADING_NUM = re.compile(r"^\s*(?:[IVXLC]{1,6}|\d{1,2}(?:\.\d{1,2}){0,3}|[A-Z])[\.\)]?\s*")
_TYPED_REF_NUM = re.compile(r"^\s*(?:\[\s*\d+\s*\]|\(\s*\d+\s*\)|\d+\s*[\.\)])\s*")
_DOI_META = re.compile(r"10\.33093/jetap\.(\d{4})\.(\d+)\.(\d+)", re.I)
_VOL_META = re.compile(r"vol(?:ume)?\.?\s*(\d+)[\s,;]*(?:no\.?|issue)\s*(\d+)[^\d]{0,12}((?:19|20)\d{2})", re.I)


class JetapFixer:
    def __init__(self, rules: Dict[str, Any], classifications=None, issues_by_category=None,
                 meta: Optional[Dict[str, str]] = None):
        self.rules = rules
        self.classifications = classifications or []
        self.issues_by_category = issues_by_category or {}
        self.meta = dict(meta or {})
        self.changes: List[ChangeRecord] = []
        self.document = None
        self.original_document = None
        self.warnings: List[str] = []
        self._extra_front: List[Tuple[str, Any]] = []
        self.template_doc = None
        self.profile = rules.get("_profile", {})

    # ------------------------------------------------------------------
    def load_manuscript(self, file_path_or_bytes):
        self.document = load_document(file_path_or_bytes)
        self.original_document = deepcopy(self.document)
        return self

    def _template(self):
        if self.template_doc is None:
            src = self.meta.get("template_bytes")
            self.template_doc = Document(BytesIO(src)) if src else Document(str(TEMPLATE_PATH))
        return self.template_doc

    def _rec(self, idx, location, ctype, prop, cur, tgt, preview="", ptype="", evidence="JETAP template"):
        self.changes.append(ChangeRecord(
            paragraph_index=idx, location=location, change_type=ctype,
            before=f"{prop}: {cur}", after=f"{prop}: {tgt}", text_preview=preview[:80],
            property_name=prop, current_value=str(cur), target_value=str(tgt),
            paragraph_type=ptype, evidence=evidence))

    # ------------------------------------------------------------------
    def fix_all(self):
        if self.document is None:
            raise ValueError("No manuscript loaded")
        doc = self.document
        self.jpars = JetapClassifier(self.profile).classify(doc)
        self._detect_meta()

        self._fix_page_setup()
        self._normalise_labels()
        self._merge_captions()
        self._ensure_caption_periods()
        self._fix_in_text_references()
        self._strip_typed_numbers()
        self._convert_references()
        self._strip_reference_dois()
        self._uppercase_headings()
        for jp in self.jpars:
            self._format_paragraph(jp)
        self._apply_heading_numbering()
        self._apply_reference_numbering()
        self._fix_tables()
        self._fix_caption_positions()
        self._insert_front_placeholders()
        self._restructure_front_and_sections()
        self._fix_headers_footers()
        return doc, self.changes

    # ---- meta ----------------------------------------------------------
    def _detect_meta(self):
        m = self.meta
        found_text = []
        for jp in self.jpars:
            if jp.role in ("doi", "journal_info", "masthead"):
                found_text.append(jp.text)
        for sec in self.document.sections:
            for part in (sec.header, sec.first_page_header):
                if part._has_definition:
                    found_text.append(" ".join(t.text or "" for t in part._element.iter(qn("w:t"))))
        blob = " \n".join(found_text)
        d = _DOI_META.search(blob)
        if d:
            m.setdefault("year", d.group(1))
            m.setdefault("volume", d.group(2))
            m.setdefault("issue", d.group(3))
            m.setdefault("doi_url", d.group(0) if d.group(0).startswith("http") else spec.DOI_PREFIX + ".".join(d.groups()))
        v = _VOL_META.search(blob)
        if v:
            m.setdefault("volume", v.group(1))
            m.setdefault("issue", v.group(2))
            m.setdefault("year", v.group(3))
        if m.get("year") and m.get("volume") and m.get("issue") and not m.get("doi_url"):
            m["doi_url"] = f"{spec.DOI_PREFIX}{m['year']}.{m['volume']}.{m['issue']}"
        elif m.get("doi_url") and not m["doi_url"].startswith("http"):
            m["doi_url"] = spec.DOI_PREFIX + m["doi_url"].split("jetap.")[-1]

    # ---- page ----------------------------------------------------------
    def _fix_page_setup(self):
        sec0 = self.document.sections[0]
        before = (sec0.left_margin.inches if sec0.left_margin else 0,
                  sec0.top_margin.inches if sec0.top_margin else 0,
                  sec0.page_width.cm if sec0.page_width else 0)
        for sect in st.all_sectprs(self.document):
            st.apply_page_geometry(sect)
        after_w = sec0.page_width.cm
        if abs(before[0] - 1.0) > 0.05 or abs(before[1] - 1.0) > 0.05:
            self._rec(-1, "Page margins", "margins", "margins", f"{before[0]:.2f} in / {before[1]:.2f} in",
                      "1.00 in all sides")
        if abs(before[2] - 21.0) > 0.2:
            self._rec(-1, "Page size", "layout", "page_size", f"{before[2]:.1f} cm wide", "A4 (21.0 cm)")

    # ---- text conventions ------------------------------------------------
    def _wt_text(self, p_el) -> str:
        return "".join(run_text(r) for r in run_elements(p_el) if not run_has_drawing(r))

    def _normalise_labels(self):
        fig_n = tab_n = 0
        for jp in list(self.jpars):
            el = jp.el
            if jp.role in ("caption_figure", "caption_table") and jp.note == "title-line":
                continue
            if jp.role == "caption_figure":
                fig_n += 1
                text = self._wt_text(el)
                m = _FIG_LABEL.match(text)
                if m:
                    new = f"Fig. {m.group(2)}. "
                    if text[:m.end()] != new:
                        replace_text_span(el, 0, m.end(), new)
                        self._rec(jp.index, f"Figure caption (paragraph {jp.index})", "caption", "caption_label",
                                  text[:m.end()].strip(), new.strip(), jp.text, "caption")
            elif jp.role == "caption_table":
                tab_n += 1
                text = self._wt_text(el)
                m = _TAB_LABEL.match(text)
                if m:
                    num = m.group(2)
                    roman = to_roman(int(num)) if num.isdigit() else num.upper()
                    new = f"Table {roman}. "
                    if text[:m.end()] != new:
                        replace_text_span(el, 0, m.end(), new)
                        self._rec(jp.index, f"Table caption (paragraph {jp.index})", "caption", "caption_label",
                                  text[:m.end()].strip(), new.strip(), jp.text, "caption")
            elif jp.role == "abstract_label":
                self._merge_abstract_label(jp)
            elif jp.role == "abstract" and jp.note == "run-in":
                self._normalise_run_in_label(jp, r"^\s*abstract(\s*)([—–\-:.]?)(\s*)", "Abstract", " — ", italic_label=True)
            elif jp.role == "keywords" and jp.note != "content":
                self._normalise_run_in_label(jp, r"^\s*(key\s*words?|index\s+terms)(\s*)([—–\-:.]?)(\s*)",
                                             "Keywords", "—", italic_label=False, kw=True)

    def _normalise_run_in_label(self, jp, pattern, label, sep, italic_label, kw=False):
        text = self._wt_text(jp.el)
        m = re.match(pattern, text, re.I)
        if not m:
            return
        old = m.group(0)
        new = label + sep
        if kw:
            new = "Keywords—"
        if old == new:
            return
        replace_text_span(jp.el, 0, m.end(), new)
        self._rec(jp.index, f"{label} (paragraph {jp.index})", "structure", "label_separator", repr(old), repr(new),
                  jp.text, jp.role)

    def _merge_abstract_label(self, jp):
        """'Abstract' standing alone -> run-in 'Abstract — ' at the start of the abstract text."""
        nxt = jp.el.getnext()
        while nxt is not None and nxt.tag == qn("w:p") and is_plain_empty(nxt):
            nxt = nxt.getnext()
        if nxt is None or nxt.tag != qn("w:p") or not paragraph_text(nxt).strip():
            return
        label_run = st.make_run("Abstract — ", spec.ROLES["abstract"]["size"], bold=True, italic=False)
        # only the word 'Abstract' is italic in the template
        rpr = label_run.find(qn("w:rPr"))
        label_run.remove(rpr)
        r1 = st.make_run("Abstract", spec.ROLES["abstract"]["size"], bold=True, italic=True)
        r2 = st.make_run(" — ", spec.ROLES["abstract"]["size"], bold=True, italic=False)
        ppr = nxt.find(qn("w:pPr"))
        pos = 0 if ppr is None else list(nxt).index(ppr) + 1
        nxt.insert(pos, r2)
        nxt.insert(pos, r1)
        jp.el.getparent().remove(jp.el)
        self._rec(jp.index, f"Abstract (paragraph {jp.index})", "structure", "abstract_label", "Separate 'Abstract' line",
                  "Run-in 'Abstract — text'", "Abstract", "abstract_label")
        # the next paragraph now carries the label; treat it as the abstract
        for other in self.jpars:
            if other.el is nxt:
                other.role = "abstract"
                other.note = "run-in"
        self.jpars = [j for j in self.jpars if j is not jp]

    def _fix_in_text_references(self):
        fig_rx = re.compile(r"\b(?:Figure|Fig)\.?\s+(\d+)\b")
        tab_rx = re.compile(r"\b(?:Table|TABLE)\s+(\d+)\b")
        for jp in self.jpars:
            if jp.role not in ("body", "decl_body", "abstract"):
                continue
            n_fig = regex_replace_in_paragraph(jp.el, fig_rx, lambda m: f"Fig. {m.group(1)}")
            n_tab = regex_replace_in_paragraph(jp.el, tab_rx, lambda m: f"Table {to_roman(int(m.group(1)))}")
            if n_fig or n_tab:
                self._rec(jp.index, f"Body text (paragraph {jp.index})", "text_convention", "figure_table_reference",
                          "Figure N / Table N", "Fig. N / Table <Roman>", jp.text, jp.role)

    def _strip_typed_numbers(self):
        for jp in self.jpars:
            if jp.role.startswith("heading") and jp.number_prefix:
                text = self._wt_text(jp.el)
                m = _TYPED_HEADING_NUM.match(text)
                if not m:
                    continue
                replace_text_span(jp.el, 0, m.end(), "")
                self._drop_leading_tabs(jp.el)
                self._rec(jp.index, f"Heading {jp.level} (paragraph {jp.index})", "numbering", "heading_number",
                          jp.number_prefix, "automatic numbering", jp.text, jp.role)

    def _uppercase_headings(self):
        for jp in self.jpars:
            if not (jp.role == "heading1" or (jp.role.startswith("heading") and jp.component)):
                # sub-headings (level 2/3) are NOT capitalised; undo ALL-CAPS left by earlier fixes
                if jp.role in ("heading2", "heading3") and jp.text.strip().isupper() and len(jp.text.split()) > 1:
                    first = True
                    small = {"a", "an", "the", "and", "or", "of", "in", "on", "for", "to", "with", "by", "at", "as", "vs"}
                    for r in run_elements(jp.el):
                        for t in r.findall(qn("w:t")):
                            words = re.split(r"(\s+)", t.text or "")
                            out = []
                            for w in words:
                                if w.strip() and len(w) > 1 and w.isupper() and not first and w.lower() in small:
                                    w = w.lower()
                                elif w.strip() and len(w) > 3 and w.isupper():
                                    w = w.capitalize()
                                if w.strip():
                                    first = False
                                out.append(w)
                            t.text = "".join(out)
                    self._rec(jp.index, f"Heading {jp.level} (paragraph {jp.index})", "heading", "capitalization",
                              "ALL CAPITALS", "Title Case", jp.text, jp.role)
                continue
            changed = False
            for r in run_elements(jp.el):
                if run_has_drawing(r):
                    continue
                for t in r.findall(qn("w:t")):
                    up = (t.text or "").upper()
                    if up != t.text:
                        t.text = up
                        changed = True
            if changed:
                self._rec(jp.index, f"Heading {jp.level} (paragraph {jp.index})", "heading", "capitalization",
                          "mixed case", "ALL CAPITALS", jp.text, jp.role)

    # ---- captions: label + title on ONE line, ending with a full stop ------------------------
    def _merge_captions(self):
        jps = self.jpars
        k = 0
        while k < len(jps):
            jp = jps[k]
            if jp.role in ("caption_figure", "caption_table") and jp.note == "label-only":
                title = None
                for j2 in jps[k + 1:k + 4]:
                    if j2.note == "title-line" and j2.el.getparent() is not None:
                        title = j2
                        break
                    if j2.role not in ("empty",):
                        break
                if title is not None:
                    label_text = self._wt_text(jp.el)
                    if label_text and not label_text.endswith(" "):
                        replace_text_span(jp.el, len(label_text), len(label_text), " ")
                    for ch in list(title.el):
                        if ch.tag != qn("w:pPr"):
                            jp.el.append(ch)
                    title.el.getparent().remove(title.el)
                    jps.remove(title)
                    jp.text = (jp.text.strip() + " " + title.text.strip()).strip()
                    jp.note = ""
                    self._rec(jp.index, f"{'Table' if jp.role == 'caption_table' else 'Figure'} caption (paragraph {jp.index})",
                              "caption", "caption_same_line", "label and title on separate lines",
                              "label and title on the same line", jp.text, "caption")
            k += 1

    def _ensure_caption_periods(self):
        for jp in self.jpars:
            if jp.role not in ("caption_figure", "caption_table") or jp.note == "title-line":
                continue
            ts = [t for r in run_elements(jp.el) if not run_has_drawing(r) for t in r.findall(qn("w:t"))]
            while ts and not (ts[-1].text or "").strip():
                ts[-1].text = ""
                ts.pop()
            if not ts:
                continue
            before = "".join(t.text or "" for t in ts)
            last = ts[-1]
            last.text = (last.text or "").rstrip()
            if (last.text or "").strip() == "." and len(ts) > 1:
                last.text = "."
                ts[-2].text = (ts[-2].text or "").rstrip()
            elif last.text.endswith("."):
                last.text = re.sub(r"\s+\.$", ".", last.text)
            now = "".join(t.text or "" for t in ts)
            if not now.endswith("."):
                last.text = last.text + "."
                now += "."
            if now != before:
                last.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                self._rec(jp.index, f"{'Table' if jp.role == 'caption_table' else 'Figure'} caption (paragraph {jp.index})",
                          "caption", "caption_full_stop", before[-30:], now[-30:], jp.text, "caption")

    def _strip_reference_dois(self):
        """The JETAP template's IEEE examples carry no DOI: drop ', doi: ...' left at the end of an entry."""
        rx = re.compile(r",\s*doi:\s*\S+?(?=\.?\s*$)")
        n = 0
        for jp in self.jpars:
            if jp.role == "reference":
                n += regex_replace_in_paragraph(jp.el, rx, lambda m: "")
        if n:
            self._rec(-1, "Reference list", "references", "reference_doi", f"{n} entries with 'doi: ...'",
                      "no DOI (as in the template examples)", "", "reference")

    # ---- references: APA -> IEEE, author-year citations -> [n] ---------------------------------
    def _convert_references(self):
        rjps = [jp for jp in self.jpars if jp.role == "reference" and jp.text.strip()]
        if not rjps:
            return
        entries = []                    # (jp, text) after removing typed numbers
        seen_heads = set()
        for jp in rjps:
            t = _TYPED_REF_NUM.sub("", re.sub(r"\s+", " ", self._wt_text(jp.el)).strip(), count=1)
            entries.append([jp, t])
            seen_heads.add(refs_mod.norm_key(t[:50]))
        glue = re.compile(r"(?<=[0-9a-z/\)])(?=[A-Z][A-Za-z’'\-]+, [A-Z]\.[^()]{0,200}?\((?:19|20)\d\d[a-z]?\)\.)")
        extra = []
        for e in entries:
            m = glue.search(e[1])
            if m and m.start() > 30:
                head, tail = e[1][:m.start()], e[1][m.start():]
                e[1] = head
                if refs_mod.norm_key(tail[:50]) not in seen_heads:
                    extra.append((e[0], tail))
                    seen_heads.add(refs_mod.norm_key(tail[:50]))
        from copy import deepcopy as _dc
        for anchor_jp, tail in extra:
            newp = _dc(anchor_jp.el)
            for ch in list(newp):
                if ch.tag != qn("w:pPr"):
                    newp.remove(ch)
            anchor_jp.el.addnext(newp)
            njp = type(anchor_jp)(-1, newp, "reference", tail)
            self.jpars.insert(self.jpars.index(anchor_jp) + 1, njp)
            entries.append([njp, tail])
        parsed = [refs_mod.parse_apa(t) for _, t in entries]
        ok = [i for i, p in enumerate(parsed) if p]
        if len(ok) < max(1, int(0.5 * len(entries))):
            ieee_rx = re.compile(r"^(?:[A-Z]\.\s?){1,4}\s*[A-Za-z’'\-]")
            if sum(1 for _, t in entries if ieee_rx.match(t)) >= 0.5 * len(entries):
                return                                  # already IEEE style
            self.warnings.append("The reference list could not be recognised as author-year (APA) entries, so it was left "
                                 "unchanged. Convert it to IEEE style manually.")
            return
        refs = [p for p in parsed if p]
        index = refs_mod.build_index(refs)
        pos_map = {i: n for n, i in enumerate(ok)}            # entries index -> refs index
        order = []
        n_cit = n_unres = 0
        for jp in self.jpars:
            if jp.role not in ("body", "decl_body", "abstract", "table_note"):
                continue
            text = self._wt_text(jp.el)
            repl, unres = refs_mod.convert_citations(text, refs, index, order)
            n_unres += unres
            for a, b, new in sorted(repl, reverse=True):
                replace_text_span(jp.el, a, b, new)
            if repl:
                n_cit += len(repl)
                self._rec(jp.index, f"Body text (paragraph {jp.index})", "citations", "citation_style",
                          "author-year", "[n] bracket numbers", jp.text, jp.role)
        # rebuild each reference paragraph in IEEE style
        for n, (jp, t) in enumerate(entries):
            p = parsed[n]
            if not p:
                continue
            for ch in list(jp.el):
                if ch.tag != qn("w:pPr"):
                    jp.el.remove(ch)
            for text, ital in refs_mod.format_ieee(p):
                jp.el.append(st.make_run(text, 8, False, ital))
            self._rec(jp.index, f"Reference (paragraph {jp.index})", "references", "reference_style",
                      t[:80], refs_mod.ieee_text(p)[:80], t, "reference")
            jp.text = refs_mod.ieee_text(p)
        unparsed = [t for (jp, t), p in zip(entries, parsed) if not p]
        # order: cited references by first citation, then the rest
        if order:
            cited = [next(e for e, rn in pos_map.items() if rn == ri) for ri in order]
            rest = [i for i in range(len(entries)) if i not in cited]
            new_order = cited + rest
            els = [entries[i][0].el for i in new_order]
            anchor = entries[0][0].el.getprevious()
            parent = entries[0][0].el.getparent()
            first_pos = min(list(parent).index(e[0].el) for e in entries)
            for el in els:
                parent.remove(el)
            for k, el in enumerate(els):
                parent.insert(first_pos + k, el)
            self._rec(rjps[0].index, "Reference list", "references", "reference_order", "alphabetical",
                      "order of first citation", "", "reference")
        if n_unres:
            self.warnings.append(f"{n_unres} in-text citation(s) could not be matched to the reference list and were left "
                                 "in author-year form - please convert them by hand.")
        if unparsed:
            self.warnings.append(f"{len(unparsed)} reference(s) could not be parsed and were left unchanged.")
        uncited = len(entries) - len(order) if order else 0
        if uncited > 0:
            self.warnings.append(f"{uncited} reference(s) are not cited in the text (listed after the cited ones).")

    @staticmethod
    def _drop_leading_tabs(p_el):
        for r in run_elements(p_el):
            for ch in list(r):
                if ch.tag == qn("w:t"):
                    if (ch.text or "").strip():
                        return
                elif ch.tag == qn("w:tab"):
                    r.remove(ch)

    # ---- generic paragraph formatting -------------------------------------
    SKIP_ROLES = {"empty", "section_break", "other", "journal_info", "masthead", "equation",
                  "table_note", "abstract_label", "caption_unknown"}

    def _format_paragraph(self, jp: JetapPara):
        role = jp.role
        if role in self.SKIP_ROLES:
            return
        key = {"caption_table": "caption", "caption_figure": "caption"}.get(role, role)
        want = spec.ROLES.get(key)
        if want is None:
            return
        doc, el = self.document, jp.el
        label = f"{role.replace('_', ' ').title()} (paragraph {jp.index})"
        prev = jp.text
        runs = [r for r in run_elements(el) if not run_has_drawing(r)]
        text_runs = [r for r in runs if run_text(r).strip()]

        # fonts / sizes
        n_font = n_size = 0
        for r in runs:
            props = effective_run_props(doc, el, r)
            has_text = bool(run_text(r))
            if has_text and (not props["font"] or props["font"] == "(theme)" or props["font"].lower() != spec.FONT.lower()):
                set_run_font(r, spec.FONT)
                n_font += 1
            if want["size"] and has_text and (props["size"] is None or abs(props["size"] - want["size"]) > 0.01):
                if props["vert"] in ("superscript", "subscript") and props["size"] and abs(props["size"] - want["size"]) < 0.51:
                    continue
                set_run_size(r, want["size"])
                n_size += 1
        if n_font:
            self._rec(jp.index, label, role, "font_name", "other fonts", spec.FONT, prev, role)
        if n_size:
            self._rec(jp.index, label, role, "font_size", "other sizes", f"{want['size']:g}pt", prev, role)

        for tag, name in (("w:b", "bold"), ("w:i", "italic")):
            tgt = want.get(name)
            if tgt is None:
                continue
            changed = 0
            for r in text_runs:
                props = effective_run_props(doc, el, r)
                if props[name] != tgt:
                    set_run_toggle(r, tag, tgt)
                    changed += 1
            if changed:
                self._rec(jp.index, label, role, name, "Not " + name if tgt else name.title(),
                          name.title() if tgt else "Not " + name, prev, role)
        if role == "abstract" and jp.note == "run-in":
            self._ensure_abstract_label_italic(el)

        # paragraph properties
        if want.get("align") and effective_alignment(doc, el) != want["align"]:
            cur = effective_alignment(doc, el)
            set_jc(el, want["align"])
            self._rec(jp.index, label, role, "alignment", cur, want["align"], prev, role)

        sp = effective_spacing(doc, el)
        if role == "decl_body":
            self._line_and_indent(jp, want, sp, label, prev)
            return
        comp = role.startswith("heading") and jp.component
        if comp:
            tgt_before, tgt_after = want["before"], want["after"]
            ok_before = -0.1 <= sp["before"] <= 8.6
            ok_after = 3.4 <= sp["after"] <= 6.6
        else:
            tgt_before, tgt_after = want["before"], want["after"]
            ok_before = abs(sp["before"] - tgt_before) <= 0.6
            ok_after = abs(sp["after"] - tgt_after) <= 0.6
        if sp["lines_based"]:
            ok_before = ok_after = False
        if not ok_after:
            set_spacing(el, after=tgt_after)
            self._rec(jp.index, label, role, "space_after", f"{sp['after']:g}pt", f"{tgt_after:g}pt", prev, role)
        if not ok_before:
            set_spacing(el, before=tgt_before)
            self._rec(jp.index, label, role, "space_before", f"{sp['before']:g}pt", f"{tgt_before:g}pt", prev, role)
        self._line_and_indent(jp, want, sp, label, prev)

    def _line_and_indent(self, jp, want, sp, label, prev):
        role, el, doc = jp.role, jp.el, self.document
        if role in spec.LINE_095_ROLES:
            bad_line = sp["rule"] != "auto" or abs(sp["line"] - want["line"]) > 3
        else:
            bad_line = sp["rule"] != "auto" or not (224 <= sp["line"] <= 244)
        if bad_line:
            set_spacing(el, line=want["line"])
            self._rec(jp.index, label, role, "line_spacing", f"{sp['line'] / 240:.2f} ({sp['rule']})",
                      f"multiple {want['line'] / 240:.2f}", prev, role)
        ind = effective_indent(doc, el)
        if role in spec.FIRST_LINE_ROLES and want.get("first") is not None and not has_numbering(el) and ind["left"] < 60:
            first = ind["first"] - ind["hanging"]
            if abs(first - want["first"]) > 20:
                set_indent(el, left=ind["left"], first=want["first"])
                self._rec(jp.index, label, role, "first_line_indent", f"{first / 1440:.2f} in",
                          f"{want['first'] / 1440:.2f} in (special first line)", prev, role)
        elif role == "reference":
            if abs(ind["left"] - want["left"]) > 60 or abs(ind["hanging"] - want["hanging"]) > 60:
                set_indent(el, left=want["left"], hanging=want["hanging"])
                self._rec(jp.index, label, role, "hanging_indent",
                          f"left {ind['left'] / 1440:.2f} in / hanging {ind['hanging'] / 1440:.2f} in",
                          f"left {want['left'] / 1440:.2f} in / hanging {want['hanging'] / 1440:.2f} in", prev, role)

    @staticmethod
    def _ensure_abstract_label_italic(p_el):
        """Template: the word 'Abstract' is italic, the text after it is not."""
        remaining = len("Abstract")
        for r in run_elements(p_el):
            txt = run_text(r)
            if not txt:
                continue
            if remaining <= 0:
                break
            if txt.strip().lower().startswith("abstract") or remaining < len("Abstract"):
                set_run_toggle(r, "w:i", True)
            remaining -= len(txt)

    # ---- numbering ---------------------------------------------------------------
    def _apply_heading_numbering(self):
        heads = [jp for jp in self.jpars if jp.role.startswith("heading")]
        if not heads:
            return
        num_id = None
        for jp in heads:
            el = jp.el
            want = spec.ROLES[jp.role]
            self._set_outline(el, jp.level - 1)
            if jp.component:
                if paragraph_is_auto_numbered(self.document, el) or True:
                    st.clear_numbering(el)
                set_indent(el, left=216, first=216) if jp.level == 1 else None
                st.set_paragraph_mark(el, want["size"], False, False)
                continue
            if num_id is None:
                num_id = st.ensure_heading_numbering(self.document)
            st.set_numbering(el, num_id, jp.level - 1)
            if jp.level == 1:
                set_indent(el, left=0, first=216)
            elif jp.level == 2:
                set_indent(el, left=360, hanging=360)
            else:
                set_indent(el, left=720, hanging=360)
            st.set_paragraph_mark(el, want["size"], False, jp.level > 1)
            ppr = ppr_child(el, "w:tabs")
            for ch in list(ppr):
                ppr.remove(ch)
            ppr.append(st.E("tab", val="left", pos=340 if jp.level == 1 else 284))
            if jp.number_prefix or not jp.has_auto_number:
                self._rec(jp.index, f"Heading {jp.level} (paragraph {jp.index})", "numbering", "heading_numbering",
                          jp.number_prefix or "none", "I. / A. automatic", jp.text, jp.role)

    @staticmethod
    def _set_outline(el, level: int):
        ppr_child(el, "w:outlineLvl").set(qn("w:val"), str(level))

    def _apply_reference_numbering(self):
        refs = [jp for jp in self.jpars if jp.role == "reference" and jp.text.strip()]
        if not refs:
            return
        num_id = st.ensure_reference_numbering(self.document)
        for jp in refs:
            text = self._wt_text(jp.el)
            m = _TYPED_REF_NUM.match(text)
            if m:
                replace_text_span(jp.el, 0, m.end(), "")
                self._drop_leading_tabs(jp.el)
            st.set_numbering(jp.el, num_id, 0)
            set_indent(jp.el, left=284, hanging=284)
            ppr = ppr_child(jp.el, "w:tabs")
            for ch in list(ppr):
                ppr.remove(ch)
            ppr.append(st.E("tab", val="left", pos=284))
            st.set_paragraph_mark(jp.el, 8, False, False)
        self._rec(refs[0].index, "Reference list", "numbering", "reference_numbering",
                  "typed / none", "[1], [2] ... automatic", refs[0].text, "reference")

    # ---- tables / figures --------------------------------------------------------
    def _table_is_content(self, tbl) -> bool:
        rows = tbl.findall(qn("w:tr"))
        cells = rows[0].findall(qn("w:tc")) if rows else []
        txt = "".join(t.text or "" for t in tbl.iter(qn("w:t"))).lower()
        if "journal of engineering technology" in txt and len(rows) == 1:
            return False
        return not (len(rows) == 1 and len(cells) == 1)

    def _fix_tables(self):
        n = 0
        for tbl in self.document.element.body.iter(qn("w:tbl")):
            if tbl.getparent() is not self.document.element.body or not self._table_is_content(tbl):
                continue
            n += 1
            loc = f"Table {n}"
            tpr = tbl.find(qn("w:tblPr"))
            if tpr is None:
                tpr = st.E("tblPr")
                tbl.insert(0, tpr)
            jc = tpr.find(qn("w:jc"))
            if jc is None:
                jc = st.E("jc")
                look = tpr.find(qn("w:tblLook"))
                (look.addprevious(jc) if look is not None else tpr.append(jc))
            if jc.get(qn("w:val")) != "center":
                jc.set(qn("w:val"), "center")
                self._rec(-1, loc, "tables", "table_alignment", "not centred", "centre", "", "table")
            if self._set_all_borders(tbl, tpr):
                self._rec(-1, loc, "tables", "table_borders", "partial / none", "all borders (grid)", "", "table")
            rows = tbl.findall(qn("w:tr"))
            changed = Counter_ = 0
            for ri, row in enumerate(rows):
                header = ri == 0 and len(rows) > 1
                for p in row.iter(qn("w:p")):
                    if not paragraph_text(p).strip() and not paragraph_has_drawing(p):
                        continue
                    for r in run_elements(p):
                        if run_has_drawing(r) or not run_text(r):
                            continue
                        props = effective_run_props(self.document, p, r)
                        if not props["font"] or props["font"] == "(theme)" or props["font"].lower() != spec.FONT.lower():
                            set_run_font(r, spec.FONT); changed += 1
                        if props["size"] is None or (abs(props["size"] - 8) > 0.01 and props["vert"] not in ("superscript", "subscript")):
                            set_run_size(r, 8); changed += 1
                        if header and not props["bold"] and run_text(r).strip():
                            set_run_toggle(r, "w:b", True); changed += 1
                    if not paragraph_has_drawing(p):
                        stray = 0
                        for tab in list(p.iter(qn("w:tab"))):          # tab characters inside cells push text out
                            if tab.getparent().tag == qn("w:r"):
                                tab.getparent().remove(tab)
                                stray += 1
                        if stray:
                            changed += 1
                        set_jc(p, "CENTER")
                        set_spacing(p, before=0, after=0, line=240)
                        if not has_numbering(p):
                            set_indent(p, first=0)
            if changed:
                self._rec(-1, loc, "tables", "table_text", "mixed fonts / sizes / header weight",
                          "Times New Roman 8pt, bold header row, centred", "", "table")
        # figure paragraphs: centred, no first-line indent
        for jp in self.jpars:
            if jp.role == "figure":
                set_jc(jp.el, "CENTER")
                set_indent(jp.el, first=0)

    _BORDER_SIDES = ("top", "left", "bottom", "right", "insideH", "insideV")

    def _set_all_borders(self, tbl, tpr) -> bool:
        """Every table / cell edge, horizontal and vertical, is a single 0.5pt black line."""
        changed = False
        want = {s_: ("single", "4", "000000") for s_ in self._BORDER_SIDES}
        b = tpr.find(qn("w:tblBorders"))
        ok = b is not None and all(
            (b.find(qn("w:" + s_)) is not None and b.find(qn("w:" + s_)).get(qn("w:val")) == "single"
             and b.find(qn("w:" + s_)).get(qn("w:sz")) not in (None, "0")) for s_ in self._BORDER_SIDES)
        if not ok:
            if b is not None:
                tpr.remove(b)
            b = st.parse_xml(f'<w:tblBorders {st.nsdecls("w")}>' + "".join(
                f'<w:{s_} w:val="single" w:sz="4" w:space="0" w:color="000000"/>' for s_ in self._BORDER_SIDES) + "</w:tblBorders>")
            nxt = None
            for tag in ("shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription"):
                nxt = tpr.find(qn("w:" + tag))
                if nxt is not None:
                    break
            (nxt.addprevious(b) if nxt is not None else tpr.append(b))
            changed = True
        # row-level overrides (w:tblPrEx) would still hide edges: drop their border part
        for ex in tbl.iter(qn("w:tblPrEx")):
            eb = ex.find(qn("w:tblBorders"))
            if eb is not None:
                ex.remove(eb)
                changed = True
            if len(ex) == 0:
                ex.getparent().remove(ex)
        order_after = ("tcW", "gridSpan", "hMerge", "vMerge")
        for tc in tbl.iter(qn("w:tc")):
            tcpr = tc.find(qn("w:tcPr"))
            if tcpr is None:
                tcpr = st.E("tcPr")
                tc.insert(0, tcpr)
            new = st.parse_xml(f'<w:tcBorders {st.nsdecls("w")}>' + "".join(
                f'<w:{s_} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                for s_ in ("top", "left", "bottom", "right")) + "</w:tcBorders>")
            cb = tcpr.find(qn("w:tcBorders"))
            if cb is not None:
                good = all(cb.find(qn("w:" + s_)) is not None and cb.find(qn("w:" + s_)).get(qn("w:val")) == "single"
                           for s_ in ("top", "left", "bottom", "right"))
                if good:
                    continue
                cb.addprevious(new)
                tcpr.remove(cb)
            else:
                anchor = None
                for ch in tcpr:
                    if ch.tag.split("}")[1] in order_after:
                        anchor = ch
                (anchor.addnext(new) if anchor is not None else tcpr.insert(0, new))
            changed = True
        return changed

    def _table_has_borders(self, tbl, tpr) -> bool:
        b = tpr.find(qn("w:tblBorders"))
        if b is not None and any(c.get(qn("w:val")) not in (None, "nil", "none") for c in b):
            return True
        for bt in tbl.iter(qn("w:tcBorders")):
            if any(c.get(qn("w:val")) not in (None, "nil", "none") for c in bt):
                return True
        ts = tpr.find(qn("w:tblStyle"))
        if ts is not None:
            name = ts.get(qn("w:val")) or ""
            sty = None
            for s in self.document.styles.element.findall(qn("w:style")):
                if s.get(qn("w:styleId")) == name:
                    sty = s
                    break
            if sty is not None:
                sn = (sty.find(qn("w:name")).get(qn("w:val")) if sty.find(qn("w:name")) is not None else "").lower()
                if "grid" in sn:
                    return True
                for b in sty.iter(qn("w:tblBorders")):
                    if any(c.get(qn("w:val")) not in (None, "nil", "none") for c in b):
                        return True
        return False

    def _fix_caption_positions(self):
        def sib(e, d):
            e = e.getprevious() if d < 0 else e.getnext()
            while e is not None and e.tag == qn("w:p") and is_plain_empty(e):
                e = e.getprevious() if d < 0 else e.getnext()
            return e

        for jp in self.jpars:
            el = jp.el
            if el.getparent() is None or jp.note in ("label-only", "title-line"):
                continue
            if jp.role == "caption_table":
                nxt, prv = sib(el, 1), sib(el, -1)
                if (nxt is None or nxt.tag != qn("w:tbl")) and prv is not None and prv.tag == qn("w:tbl"):
                    before_tbl = sib(prv, -1)
                    if before_tbl is None or not paragraph_text(before_tbl).strip().lower().startswith("table"):
                        prv.addprevious(el)
                        self._rec(jp.index, f"Table caption (paragraph {jp.index})", "caption", "caption_position",
                                  "below table", "above table", jp.text, "caption")
            elif jp.role == "caption_figure":
                nxt, prv = sib(el, 1), sib(el, -1)
                if (prv is None or prv.tag != qn("w:p") or not paragraph_has_drawing(prv)) \
                        and nxt is not None and nxt.tag == qn("w:p") and paragraph_has_drawing(nxt):
                    nxt.addnext(el)
                    self._rec(jp.index, f"Figure caption (paragraph {jp.index})", "caption", "caption_position",
                              "above figure", "below figure", jp.text, "caption")

    # ---- front matter ---------------------------------------------------------------
    FRONT_ROLES = ("title", "authors", "affiliation", "corresponding", "doi", "revision")

    def _front_elements(self) -> List[Tuple[str, Any]]:
        items = [(j.role, j.el) for j in self.jpars if j.role in self.FRONT_ROLES and j.el.getparent() is not None]
        items += [(r, e) for r, e in self._extra_front if e.getparent() is not None]
        body = list(self.document.element.body)
        pos = {id(e): i for i, e in enumerate(body)}
        return sorted(items, key=lambda it: pos.get(id(it[1]), 10 ** 9))

    def _insert_front_placeholders(self):
        front = self._front_elements()
        roles = [r for r, _ in front]
        if "title" not in roles:
            return
        m = self.meta
        # DOI line
        if "doi" not in roles:
            anchor = next((el for r, el in reversed(front) if r in ("corresponding", "affiliation", "authors", "title")), None)
            doi_text = m.get("doi_url") or f"{spec.DOI_PREFIX}20XX.X.X"
            p = st.make_paragraph(doi_text + " ", "doi")
            anchor.addnext(p)
            self._extra_front.append(("doi", p))
            self._rec(-1, "Front matter", "structure", "doi_line", "missing", doi_text, "", "doi",
                      "Placeholder from the template - replace with the assigned DOI")
            front = self._front_elements()
            doi_el = p
        else:
            doi_el = next(el for r, el in front if r == "doi")
        if "revision" not in roles:
            text = "Manuscript Received: XX Mon 20XX. Revised: XX Mon 20XX. Accepted: XX Mon 20XX. Published: XX Mon 20XX"
            p = st.make_paragraph(text, "revision")
            doi_el.addnext(p)
            self._extra_front.append(("revision", p))
            self._rec(-1, "Front matter", "structure", "manuscript_dates_line", "missing", text, "", "revision",
                      "Placeholder from the template - filled in by the editorial office")

    def _restructure_front_and_sections(self):
        doc = self.document
        body = doc.element.body
        final_sect = body.find(qn("w:sectPr"))
        if final_sect is None:
            return
        # 1. remove old masthead / journal-info lines (their content lives in the new masthead + header)
        removed = 0
        for jp in self.jpars:
            if jp.role in ("masthead", "journal_info") and jp.el.getparent() is not None and not paragraph_has_sectpr(jp.el):
                jp.el.getparent().remove(jp.el)
                removed += 1
        for tbl in list(body.findall(qn("w:tbl"))):
            txt = "".join(t.text or "" for t in tbl.iter(qn("w:t"))).lower()
            rows = tbl.findall(qn("w:tr"))
            if len(rows) <= 2 and (("journal of" in txt and "technology" in txt) or re.search(r"vol\w*\.?\s*\d.*issn", txt)):
                tbl.getparent().remove(tbl)
                removed += 1
        if removed:
            self._rec(-1, "Top of page 1", "structure", "journal_banner", "manuscript's own banner / volume lines",
                      "template masthead + running header", "", "journal_header")

        # 1b. blank / section-break-only paragraphs in front of the title (the masthead brings its own)
        title_el = next((j.el for j in self.jpars if j.role == "title"), None)
        if title_el is not None:
            e = body[0] if len(body) else None
            while e is not None and e is not title_el:
                nxt = e.getnext()
                if e.tag == qn("w:p") and not paragraph_text(e).strip() and not paragraph_has_drawing(e):
                    body.remove(e)
                elif e.tag != qn("w:p"):
                    break
                e = nxt

        # 2. blank lines + section break between the author block and the abstract
        front = self._front_elements()
        last_front = front[-1][1] if front else None
        if front and title_el is not None:
            e = title_el.getnext()
            while e is not None and e is not last_front:
                nxt = e.getnext()
                if e.tag == qn("w:p") and is_plain_empty(e):
                    body.remove(e)
                e = nxt
        first_after = None
        if last_front is not None:
            e = last_front.getnext()
            mids = []
            while e is not None and e.tag == qn("w:p") and (is_plain_empty(e) or (paragraph_has_sectpr(e) and not paragraph_text(e).strip())):
                mids.append(e)
                e = e.getnext()
            first_after = e
            front_break = next((m for m in mids if paragraph_has_sectpr(m)), None)
            for m in mids:
                if m is not front_break:
                    body.remove(m)
            if front_break is None:
                front_break = st.new_section_break_paragraph(final_sect, 1, "continuous", "front")
                (first_after.addprevious(front_break) if first_after is not None else last_front.addnext(front_break))
                self._rec(-1, "Between author block and abstract", "structure", "section_break",
                          "none", "continuous break: single-column title block, two-column body", "", "layout")
            else:
                self._style_front_break(front_break, final_sect)
            for _ in range(2):
                front_break.addprevious(st.make_blank())
        else:
            front_break = None

        # 2b. the template spaces text with paragraph spacing, never blank lines
        removed_blank = 0
        started = False
        for e in list(body):
            if e is front_break:
                started = True
                continue
            if e.tag == qn("w:p") and is_plain_empty(e) and (started or front_break is None) and e.getnext() is not None \
                    and e.getnext().tag != qn("w:sectPr"):
                body.remove(e)
                removed_blank += 1
        if front is not None and last_front is not None:
            pass
        if removed_blank:
            self._rec(-1, "Body", "spacing", "blank_lines", f"{removed_blank} blank paragraph(s)",
                      "removed - spacing comes from paragraph spacing", "", "body")

        # 3. masthead table + first-section break at the very top of the body
        masthead = st.make_masthead_table()
        first_break = st.new_section_break_paragraph(final_sect, 1, None, "first")
        body.insert(0, masthead)
        masthead.addnext(first_break)
        self._rec(-1, "Top of page 1", "structure", "masthead", "missing / different",
                  "Journal title masthead (28pt bold, framed by rules)", "", "journal_header")

        # 4. column layout for every section; headers/footers are rebuilt below
        sects = [s for s in st.all_sectprs(doc)]
        order = []
        for s in sects:
            order.append(s)
        break_sect = front_break.find(qn("w:pPr")).find(qn("w:sectPr")) if front_break is not None else None
        seen_break = break_sect is None
        for s in order:
            if s is order[0]:
                st.set_cols(s, 1)
                continue
            if not seen_break:
                st.set_cols(s, 1)
                if s is break_sect:
                    seen_break = True
                    st.set_section_type(s, "continuous")
                continue
            st.set_cols(s, 2)
            st.set_section_type(s, "continuous")
            st.apply_page_geometry(s)
        self._rec(-1, "Body", "layout", "columns", "single column", "two columns (3.0 in, 0.27 in gap)", "", "layout")
        self._fit_objects_to_column()
        self._first_break_sect = first_break.find(qn("w:pPr")).find(qn("w:sectPr"))

    # ---- fit figures / tables to the column --------------------------------------------
    def _fit_objects_to_column(self):
        max_tw = spec.COLUMN_WIDTH_TW
        max_emu = max_tw * 635 - 6000
        n_fig = n_tab = 0
        for p in self.document.element.body.iter(qn("w:p")):
            if p.getparent() is not self.document.element.body:
                continue
            exts = []
            for tag in ("wp:inline", "wp:anchor"):
                for d in p.iter(qn(tag)):
                    e = d.find(qn("wp:extent"))
                    if e is not None:
                        exts.append((d, e))
            if not exts:
                continue
            total = sum(int(e.get("cx")) for _, e in exts)
            if total <= max_emu:
                continue
            k = max_emu / total
            for d, e in exts:
                for ext in d.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}ext"):
                    if ext.get("cx") is not None:
                        ext.set("cx", str(int(int(ext.get("cx")) * k)))
                        ext.set("cy", str(int(int(ext.get("cy")) * k)))
                e.set("cx", str(int(int(e.get("cx")) * k)))
                e.set("cy", str(int(int(e.get("cy")) * k)))
            n_fig += 1
            self._rec(-1, "Figure", "figure", "figure_size", f"{total / 914400:.2f} in wide",
                      f"{total * k / 914400:.2f} in (column width 3.0 in)", paragraph_text(p), "figure")
        for tbl in self.document.element.body.findall(qn("w:tbl")):
            if not self._table_is_content(tbl):
                continue
            grid = tbl.find(qn("w:tblGrid"))
            cols = [int(g.get(qn("w:w"))) for g in grid.findall(qn("w:gridCol"))] if grid is not None else []
            total = sum(cols)
            if not cols or total <= max_tw + 20:
                continue
            k = max_tw / total
            for g in grid.findall(qn("w:gridCol")):
                g.set(qn("w:w"), str(int(int(g.get(qn("w:w"))) * k)))
            for tc in tbl.iter(qn("w:tcW")):
                if tc.get(qn("w:type")) == "dxa" and tc.get(qn("w:w")):
                    tc.set(qn("w:w"), str(int(int(tc.get(qn("w:w"))) * k)))
            tpr = tbl.find(qn("w:tblPr"))
            tw = tpr.find(qn("w:tblW"))
            if tw is None:
                tw = st.E("tblW")
                tpr.insert(1 if tpr.find(qn("w:tblStyle")) is not None else 0, tw)
            tw.set(qn("w:w"), str(max_tw))
            tw.set(qn("w:type"), "dxa")
            ind = tpr.find(qn("w:tblInd"))
            if ind is not None:
                tpr.remove(ind)
            if tpr.find(qn("w:tblLayout")) is None:
                lay = st.E("tblLayout", type="fixed")
                look = tpr.find(qn("w:tblLook"))
                (look.addprevious(lay) if look is not None else tpr.append(lay))
            n_tab += 1
            self._rec(-1, f"Table", "table", "table_width", f"{total / 1440:.2f} in wide",
                      f"{max_tw / 1440:.2f} in (column width)", "", "table")

    def _style_front_break(self, p, final_sect):
        set_spacing(p, before=0, after=0, line=240)
        set_jc(p, "CENTER")

    # ---- headers / footers ------------------------------------------------------------
    def _fix_headers_footers(self):
        doc = self.document
        m = self.meta
        # drop every existing header/footer reference so the template's header/footer is the only one
        first_sect = self._first_break_sect
        for s in st.all_sectprs(doc):
            if s is first_sect:
                continue
            st.strip_header_footer_refs(s)
        section = doc.sections[0]
        section.different_first_page_header_footer = True
        vol, issue, year = m.get("volume", "X"), m.get("issue", "X"), m.get("year", "20XX")
        st.fill_header(section.first_page_header, vol, issue, year)
        st.fill_header(section.header, vol, issue, year)
        st.empty_footer(section.footer)
        st.clone_first_page_footer(section, self._template(), {**m})
        st.drop_unused_header_footer_parts(doc)
        self._rec(-1, "Page header", "page_header", "running_header", "manuscript's own / none",
                  f"Vol {vol} No {issue} ({year})   e-ISSN: 2682-8383", "", "journal_header")
        self._rec(-1, "First-page footer", "page_header", "first_page_footer", "none / different",
                  "MMU logo, citation, DOI, CC BY-NC-ND licence", "", "journal_header")
        if "volume" not in m:
            self.warnings.append("Volume / issue / year were not found in the manuscript: the header shows the "
                                 "template placeholders 'Vol X No X (20XX)'. Fill them in the sidebar and fix again.")

    # ---- outputs ----------------------------------------------------------------------
    def get_fixed_document_bytes(self) -> bytes:
        buf = BytesIO()
        self.document.save(buf)
        buf.seek(0)
        return buf.getvalue()

    def get_changes_summary(self) -> Dict[str, Any]:
        summary = {"total_changes": len(self.changes), "changes_by_type": {}, "changes_list": []}
        for c in self.changes:
            summary["changes_by_type"][c.change_type] = summary["changes_by_type"].get(c.change_type, 0) + 1
            summary["changes_list"].append({"location": c.location, "type": c.change_type, "before": c.before,
                                            "after": c.after, "property_name": c.property_name,
                                            "current_value": c.current_value, "target_value": c.target_value,
                                            "paragraph_type": c.paragraph_type, "evidence": c.evidence,
                                            "preview": c.text_preview})
        return summary

    def get_change_records(self):
        return self.changes

    def get_highlighted_document_bytes(self) -> bytes:
        doc = deepcopy(self.original_document)
        paras = doc.paragraphs
        seen = set()
        for c in self.changes:
            i = c.paragraph_index
            if i < 0 or i >= len(paras) or i in seen:
                continue
            seen.add(i)
            for run in paras[i].runs:
                if run.text.strip():
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        buf = BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf.getvalue()
