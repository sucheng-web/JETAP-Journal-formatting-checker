"""JETAP template compliance checker (same public shape as ManuscriptChecker)."""
from __future__ import annotations

import re
from collections import Counter, defaultdict
from io import BytesIO
from typing import Any, Dict, List, Optional

from docx.oxml.ns import qn

from modules.manuscript_checker import CheckResult, FormatIssue
from modules.utils import calculate_compliance_score, load_document, truncate_text
from . import spec
from .classifier import JetapClassifier, JetapPara
from .helpers import (
    effective_alignment, effective_indent, effective_run_props, effective_spacing,
    from_roman, paragraph_has_drawing, paragraph_text, run_elements, run_has_drawing,
    run_text, to_roman, is_plain_empty,
)

CATEGORIES = ["margins", "layout", "journal_header", "title", "author_info", "body_text",
              "headings", "structure", "tables", "figures", "references", "line_spacing", "other"]

ROLE_CATEGORY = {
    "title": "title", "authors": "author_info", "affiliation": "author_info",
    "corresponding": "author_info", "doi": "author_info", "revision": "author_info",
    "abstract": "body_text", "keywords": "body_text", "heading1": "headings",
    "heading2": "headings", "heading3": "headings", "body": "body_text", "decl_body": "body_text",
    "caption_table": "tables", "caption_figure": "figures", "reference": "references",
}
ROLE_LABEL = {
    "title": "Title", "authors": "Author names", "affiliation": "Affiliation",
    "corresponding": "Corresponding author line", "doi": "DOI line",
    "revision": "Manuscript dates line", "abstract": "Abstract", "keywords": "Keywords",
    "heading1": "Heading 1", "heading2": "Heading 2", "heading3": "Heading 3",
    "body": "Body text", "decl_body": "Declaration text", "caption_table": "Table caption", "caption_figure": "Figure caption",
    "reference": "Reference",
}
SPEC_KEY = {"caption_table": "caption", "caption_figure": "caption"}

_FIG_LABEL = re.compile(r"^\s*(fig(?:ure)?s?)\.?\s*(\d+)\s*([\.\:\-–—])?\s*", re.I)
_TAB_LABEL = re.compile(r"^\s*(table|tab\.)\s*\.?\s*(\d+|[IVXLC]{1,6})\b\s*([\.\:\-–—])?\s*", re.I)
_PAGE_FIELD = re.compile(r"\bPAGE\b|\bNUMPAGES\b")

_PART_ALIASES = {
    "abstract": ["abstract"],
    "keywords": ["keywords", "keyword", "key words", "index terms"],
    "introduction": ["introduction", "background"],
    "conclusion": ["conclusion", "conclusions", "concluding remarks", "summary"],
    "acknowledgement": ["acknowledgement", "acknowledgements", "acknowledgment", "acknowledgments"],
    "funding": ["funding", "funding statement", "funding statements"],
    "author_contribution": ["author contribution", "author contributions", "authors contribution",
                            "authors contributions", "author contribution statement"],
    "conflict": ["conflict of interest", "conflict of interests", "conflicts of interest",
                 "conflict of interest statement"],
    "ethics": ["ethics statement", "ethics statements", "ethical statement", "ethics"],
    "data": ["data availability", "data availability statement", "data access statement"],
    "references": ["references", "reference", "bibliography"],
}


def section_map(document):
    mapping, sects, pending = {}, [], []
    for ch in document.element.body:
        if ch.tag == qn("w:sectPr"):
            for e in pending:
                mapping[id(e)] = ch
            sects.append(ch)
            pending = []
            continue
        pending.append(ch)
        if ch.tag == qn("w:p"):
            ppr = ch.find(qn("w:pPr"))
            sp = ppr.find(qn("w:sectPr")) if ppr is not None else None
            if sp is not None:
                for e in pending:
                    mapping[id(e)] = sp
                sects.append(sp)
                pending = []
    return mapping, sects


def sect_columns(sp) -> Dict[str, int]:
    cols = sp.find(qn("w:cols")) if sp is not None else None
    num, space = 1, 720
    if cols is not None:
        num = int(cols.get(qn("w:num")) or 1)
        space = int(cols.get(qn("w:space")) or 720)
        kids = cols.findall(qn("w:col"))
        if kids and kids[0].get(qn("w:space")) is not None:
            space = int(kids[0].get(qn("w:space")))
    return {"num": num, "space": space}


def part_key(heading_text: str) -> Optional[str]:
    norm = re.sub(r"[\s:.—–\-]+$", "", heading_text.strip().lower())
    norm = norm.replace("&", "and")
    for key, aliases in _PART_ALIASES.items():
        if norm in aliases:
            return key
    return None


def heading_text_only(jp: JetapPara) -> str:
    return jp.text.strip()[jp.number_len:] if jp.number_len else jp.text.strip()


class JetapChecker:
    CATEGORIES = CATEGORIES

    def __init__(self, rules: Dict[str, Any], llm_integration=None):
        self.rules = rules
        self.llm = llm_integration
        self.profile = rules.get("_profile", {})
        self.document = None
        self.classifier = JetapClassifier(self.profile)
        self.issues: Dict[str, List[FormatIssue]] = {c: [] for c in CATEGORIES}
        self.jpars: List[JetapPara] = []
        self.classifications = []

    def load_manuscript(self, file_path_or_bytes):
        self.document = load_document(file_path_or_bytes)
        return self

    # ------------------------------------------------------------------
    def _add(self, category, location, idx, description, current, expected, severity="warning", preview=""):
        self.issues[category].append(FormatIssue(
            category=category, location=location, paragraph_index=idx, description=description,
            current_value=str(current), expected_value=str(expected), severity=severity,
            text_preview=preview))

    def check_all(self) -> CheckResult:
        if not self.document:
            raise ValueError("No manuscript loaded. Call load_manuscript() first.")
        self.issues = {c: [] for c in CATEGORIES}
        self.jpars = self.classifier.classify(self.document)
        self.classifications = self.classifier.to_classified(self.document, self.jpars)
        self._sec_map, self._sects = section_map(self.document)

        self._check_margins_and_page()
        self._check_columns()
        self._check_masthead_header_footer()
        for jp in self.jpars:
            key = ROLE_CATEGORY.get(jp.role)
            if key:
                self._check_paragraph(jp)
        self._check_front_matter_content()
        self._check_headings()
        self._check_body_conventions()
        self._check_tables_and_captions()
        self._check_references()
        structure = self._check_structure()
        self._check_other()

        score = calculate_compliance_score(self.issues)
        total = sum(len(v) for v in self.issues.values())
        return CheckResult(
            is_compliant=total == 0, compliance_score=score, total_issues=total,
            issues_by_category=self.issues, classifications=self.classifications,
            document_structure=structure, statistics=self._statistics())

    # ---- page / layout -------------------------------------------------
    def _check_margins_and_page(self):
        sec = self.document.sections[0]
        for name, val in (("left", sec.left_margin), ("right", sec.right_margin),
                          ("top", sec.top_margin), ("bottom", sec.bottom_margin)):
            inches = val.inches if val is not None else 0
            if abs(inches - 1.0) > 0.05:
                self._add("margins", f"{name.title()} margin", -1,
                          f"{name.title()} margin differs from the template",
                          f"{inches:.2f} in", "1.00 in", "error")
        w = sec.page_width.cm if sec.page_width else 0
        h = sec.page_height.cm if sec.page_height else 0
        if abs(w - 21.0) > 0.2 or abs(h - 29.7) > 0.2:
            self._add("layout", "Page size", -1, "Page size is not A4", f"{w:.1f} x {h:.1f} cm",
                      "A4 (21.0 x 29.7 cm)", "error")

    def _check_columns(self):
        body_roles = {"heading1", "heading2", "heading3", "body", "reference", "caption_table",
                      "caption_figure", "equation", "figure"}
        front_roles = {"title", "authors", "affiliation", "corresponding", "doi", "revision"}
        body_cols, front_cols = Counter(), Counter()
        for jp in self.jpars:
            sp = self._sec_map.get(id(jp.el))
            cols = sect_columns(sp)["num"] if sp is not None else 1
            if jp.role in body_roles and jp.text.strip():
                body_cols[cols] += 1
            elif jp.role in front_roles:
                front_cols[cols] += 1
        if body_cols and body_cols.most_common(1)[0][0] != 2:
            self._add("layout", "Body columns", -1,
                      "Body text is not set in two columns",
                      f"{body_cols.most_common(1)[0][0]} column(s)",
                      "2 columns (3.0 in each, 0.27 in gap)", "error")
        else:
            for sp in self._sects:
                c = sect_columns(sp)
                if c["num"] == 2 and abs(c["space"] - spec.PAGE["column_gap_tw"]) > 80:
                    self._add("layout", "Column spacing", -1, "Column spacing differs from the template",
                              f"{c['space'] / 1440:.2f} in", "0.27 in", "warning")
                    break
        if front_cols and front_cols.most_common(1)[0][0] != 1:
            self._add("layout", "Title block columns", -1,
                      "Title, authors and affiliations must span the full page width (single column)",
                      f"{front_cols.most_common(1)[0][0]} columns", "1 column", "warning")

    # ---- masthead / header / footer -------------------------------------
    def _top_masthead(self):
        for ch in self.document.element.body:
            if ch.tag == qn("w:tbl"):
                txt = "".join(t.text or "" for t in ch.iter(qn("w:t"))).lower()
                return ch if "journal of engineering technology" in txt else None
            if ch.tag == qn("w:p") and paragraph_text(ch).strip():
                return ch if paragraph_text(ch).strip().lower().startswith(
                    "journal of engineering technology") else None
        return None

    def _check_masthead_header_footer(self):
        m = self._top_masthead()
        if m is None:
            self._add("journal_header", "Journal title masthead", -1,
                      "Journal title masthead (JOURNAL OF ENGINEERING TECHNOLOGY AND APPLIED PHYSICS) is missing from the top of page 1",
                      "Missing", "28pt bold, centred, framed by rules", "error")
        else:
            sizes = []
            for r in m.iter(qn("w:r")):
                if run_text(r).strip():
                    sizes.append(effective_run_props(self.document, r.getparent(), r)["size"])
            sizes = [s for s in sizes if s]
            if sizes and abs(Counter(sizes).most_common(1)[0][0] - 28) > 0.5:
                self._add("journal_header", "Journal title masthead", -1,
                          "Journal title masthead font size differs from the template",
                          f"{Counter(sizes).most_common(1)[0][0]:g}pt", "28pt", "warning")

        sec = self.document.sections[0]
        if not sec.different_first_page_header_footer:
            self._add("journal_header", "First page header", -1,
                      "Different first page header/footer is not enabled", "Off", "On", "warning")
        h_text = self._part_text(sec.first_page_header) if sec.different_first_page_header_footer else ""
        if not re.search(r"vol", h_text, re.I) or "issn" not in h_text.lower():
            self._add("journal_header", "First page header", -1,
                      "First-page header lacks the 'Vol X No X (20XX)  e-ISSN: 2682-8383' line",
                      truncate_text(h_text, 60) or "Empty", "Vol X No X (20XX) ... e-ISSN: 2682-8383", "warning")
        f_text = self._part_text(sec.first_page_footer) if sec.different_first_page_header_footer else ""
        if "doi.org" not in f_text.lower() or "mmu" not in f_text.lower():
            self._add("journal_header", "First page footer", -1,
                      "First-page footer lacks the MMU Press citation, DOI and licence block",
                      truncate_text(f_text, 60) or "Empty", "Citation, DOI, CC BY-NC-ND licence, MMU logo", "warning")

        pagination = False
        for s in self.document.sections:
            for part in (s.header, s.first_page_header, s.even_page_header,
                         s.footer, s.first_page_footer, s.even_page_footer):
                if part._has_definition and self._has_page_field(part._element):
                    pagination = True
        if self._has_page_field(self.document.element.body):
            pagination = True
        if pagination:
            self._add("journal_header", "Pagination", -1,
                      "Page numbers are present; the template says not to add any kind of pagination",
                      "Page number field found", "No pagination", "warning")

    @staticmethod
    def _part_text(part) -> str:
        if not part._has_definition:
            return ""
        return " ".join((t.text or "") for t in part._element.iter(qn("w:t")))

    @staticmethod
    def _has_page_field(el) -> bool:
        for it in el.iter(qn("w:instrText")):
            if _PAGE_FIELD.search(it.text or ""):
                return True
        for fs in el.iter(qn("w:fldSimple")):
            if _PAGE_FIELD.search(fs.get(qn("w:instr")) or ""):
                return True
        return False

    # ---- generic paragraph format ---------------------------------------
    def _text_runs(self, jp):
        return [r for r in run_elements(jp.el) if run_text(r).strip() and not run_has_drawing(r)]

    def _check_paragraph(self, jp: JetapPara):
        key = SPEC_KEY.get(jp.role, jp.role)
        want = spec.ROLES[key]
        cat = ROLE_CATEGORY[jp.role]
        label = ROLE_LABEL[jp.role]
        loc = f"{label} (paragraph {jp.index})"
        prev = truncate_text(jp.text, 60)
        doc = self.document
        runs = self._text_runs(jp)
        if not runs:
            return
        props = [(r, effective_run_props(doc, jp.el, r)) for r in runs]

        bad_fonts = Counter(p["font"] for _, p in props
                            if p["font"] and p["font"] != "(theme)" and p["font"].lower() != spec.FONT.lower())
        if bad_fonts:
            self._add(cat, loc, jp.index, f"{label} font is not {spec.FONT}", bad_fonts.most_common(1)[0][0],
                      spec.FONT, "error", prev)
        elif any(p["font"] == "(theme)" for _, p in props):
            self._add(cat, loc, jp.index, f"{label} font is a theme font, not {spec.FONT}", "Theme font",
                      spec.FONT, "warning", prev)

        sizes = [p["size"] for _, p in props if p["size"] and p["vert"] not in ("superscript", "subscript")]
        bad_sizes = Counter(s for s in sizes if abs(s - want["size"]) > 0.5)
        if bad_sizes:
            self._add(cat, loc, jp.index, f"{label} font size is not {want['size']:g}pt",
                      f"{bad_sizes.most_common(1)[0][0]:g}pt", f"{want['size']:g}pt", "error", prev)

        weights = [(len(run_text(r).strip()), p) for r, p in props]
        total_chars = max(1, sum(w for w, _ in weights))
        if want.get("bold") is not None:
            bad_chars = sum(w for w, p in weights if p["bold"] != want["bold"])
            # a short tail (e.g. "(include up to 5 keywords)") may differ
            if bad_chars / total_chars > 0.5:
                self._add(cat, loc, jp.index,
                          f"{label} bold setting differs from the template",
                          "Bold" if not want["bold"] else "Not bold", "Bold" if want["bold"] else "Not bold",
                          "warning", prev)
        if want.get("italic") is not None:
            bad_chars = sum(w for w, p in weights if p["italic"] != want["italic"])
            if bad_chars / total_chars > 0.5 and jp.role not in ("authors",):
                self._add(cat, loc, jp.index, f"{label} italic setting differs from the template",
                          "Italic" if not want["italic"] else "Not italic",
                          "Italic" if want["italic"] else "Not italic", "warning", prev)

        align = effective_alignment(doc, jp.el)
        if want.get("align") and align != want["align"]:
            self._add(cat, loc, jp.index, f"{label} alignment is not {want['align'].lower()}",
                      align, want["align"], "warning", prev)

        sp = effective_spacing(doc, jp.el)
        if jp.role == "decl_body":
            self._check_line_indent(jp, want, sp, loc, prev, cat)
            return
        if jp.role.startswith("heading") and jp.component:
            ok_after, ok_before = 3.4 <= sp["after"] <= 6.6, -0.1 <= sp["before"] <= 8.6
        else:
            ok_after, ok_before = abs(sp["after"] - want["after"]) <= 0.6, abs(sp["before"] - want["before"]) <= 0.6
        if not ok_after or sp["lines_based"]:
            self._add("line_spacing" if False else cat, loc, jp.index,
                      f"{label} spacing after is not {want['after']:g}pt",
                      f"{sp['after']:g}pt" + (" (set in lines)" if sp["lines_based"] else ""),
                      f"{want['after']:g}pt", "warning", prev)
        if not ok_before:
            self._add(cat, loc, jp.index, f"{label} spacing before is not {want['before']:g}pt",
                      f"{sp['before']:g}pt", f"{want['before']:g}pt", "warning", prev)
        # the template itself mixes 0.95 and single spacing, so both are accepted
        if jp.role in spec.LINE_095_ROLES:
            line_ok = (sp["rule"] == "auto" and abs(sp["line"] - want["line"]) <= 3)
        else:
            line_ok = (sp["rule"] == "auto" and 224 <= sp["line"] <= 244)
        if not line_ok:
            cur = f"{sp['line'] / 240:.2f}" + ("" if sp["rule"] == "auto" else f" ({sp['rule']})")
            self._add("line_spacing", loc, jp.index, f"{label} line spacing differs from the template",
                      cur, f"{want['line'] / 240:.2f}", "warning", prev)

        self._check_indent(jp, want, loc, prev, cat)
        if jp.role == "reference":
            ind = effective_indent(doc, jp.el)
            if abs(ind["left"] - want["left"]) > 60 or abs(ind["hanging"] - want["hanging"]) > 60:
                self._add(cat, loc, jp.index, "Reference hanging indent differs from the template",
                          f"left {ind['left'] / 1440:.2f} in, hanging {ind['hanging'] / 1440:.2f} in",
                          f"left {want['left'] / 1440:.2f} in, hanging {want['hanging'] / 1440:.2f} in",
                          "warning", prev)

    def _check_line_indent(self, jp, want, sp, loc, prev, cat):
        label = ROLE_LABEL[jp.role]
        if not (sp["rule"] == "auto" and abs(sp["line"] - want["line"]) <= 3):
            self._add("line_spacing", loc, jp.index, f"{label} line spacing is not multiple 0.95",
                      f"{sp['line'] / 240:.2f}", "multiple 0.95", "warning", prev)
        self._check_indent(jp, want, loc, prev, cat)

    def _check_indent(self, jp, want, loc, prev, cat):
        if jp.role not in spec.FIRST_LINE_ROLES or want.get("first") is None:
            return
        ind = effective_indent(self.document, jp.el)
        first = ind["first"] - ind["hanging"]
        if abs(first - want["first"]) > 20 and not self._is_list_like(jp) and ind["left"] < 60:
            self._add(cat, loc, jp.index,
                      f"{ROLE_LABEL[jp.role]} first-line indent is not 0.2 in (special first line)",
                      f"{first / 1440:.2f} in", f"{want['first'] / 1440:.2f} in", "warning", prev)

    @staticmethod
    def _is_list_like(jp) -> bool:
        from .helpers import has_numbering
        return has_numbering(jp.el)

    # ---- front matter content --------------------------------------------
    def _check_front_matter_content(self):
        roles = {jp.role: jp for jp in reversed(self.jpars)}
        by_role = defaultdict(list)
        for jp in self.jpars:
            by_role[jp.role].append(jp)
        if not by_role["title"]:
            self._add("title", "Title", -1, "Paper title was not found", "Missing", "24pt, centred", "error")
        elif by_role["title"][0].text.strip() and re.search(r"[^\x00-\x7F–—‘’“”]", by_role["title"][0].text) \
                and False:
            pass
        if not by_role["authors"]:
            self._add("author_info", "Author names", -1, "Author names were not found below the title",
                      "Missing", "Author names, 11pt, centred", "error")
        if not by_role["affiliation"]:
            self._add("author_info", "Affiliation", -1, "Affiliation lines were not found",
                      "Missing", "Italic 10pt, centred", "warning")
        corr = by_role["corresponding"]
        if not corr:
            self._add("author_info", "Corresponding author", -1,
                      "Corresponding author line (email and ORCiD) was not found",
                      "Missing", "*Corresponding author's email address: ..., ORCiD: ...", "warning")
        else:
            txt = " ".join(c.text for c in corr)
            if "@" not in txt:
                self._add("author_info", "Corresponding author", corr[0].index,
                          "Corresponding author email address is missing", "No email", "Email address", "warning",
                          truncate_text(txt, 60))
            if "orcid" not in txt.lower():
                self._add("author_info", "Corresponding author", corr[0].index,
                          "Corresponding author ORCiD is missing", "No ORCiD", "ORCiD: xxxx-xxxx-xxxx-xxxx",
                          "warning", truncate_text(txt, 60))
        if not by_role["doi"]:
            self._add("author_info", "DOI line", -1,
                      "DOI line (https://doi.org/10.33093/jetap.20XX.X.X) is missing below the author block",
                      "Missing", "DOI line, 8pt centred", "info")
        if not by_role["revision"]:
            self._add("author_info", "Manuscript dates line", -1,
                      "Manuscript Received/Revised/Accepted/Published line is missing",
                      "Missing", "Manuscript Received: XX Mon 20XX. Revised: ... Published: ...", "info")
        for jp in by_role["abstract_label"]:
            self._add("body_text", f"Abstract (paragraph {jp.index})", jp.index,
                      "Abstract label stands alone; the template runs it in as 'Abstract — text'",
                      "Separate 'Abstract' line", "Abstract — text (italic label, bold text)", "warning", jp.text[:40])
        for jp in by_role["abstract"]:
            if jp.note == "run-in":
                m = re.match(r"^\s*abstract(\s*)([—–\-:.]?)(\s*)", jp.text, re.I)
                if m and not (m.group(2) == "—" and m.group(1) == " " and m.group(3) == " "):
                    self._add("body_text", f"Abstract (paragraph {jp.index})", jp.index,
                              "Abstract label separator should be ' — ' (space, em dash, space)",
                              repr(m.group(0)), "'Abstract — '", "warning", truncate_text(jp.text, 40))
        for jp in by_role["keywords"]:
            if jp.note == "content":
                continue
            m = re.match(r"^\s*(key\s*words?|index\s+terms)(\s*)([—–\-:.]?)(\s*)", jp.text, re.I)
            if m and not (m.group(1) == "Keywords" and m.group(2) == "" and m.group(3) == "—"):
                self._add("body_text", f"Keywords (paragraph {jp.index})", jp.index,
                          "Keywords label should read 'Keywords—' followed directly by the terms",
                          repr(m.group(0)), "'Keywords—'", "warning", truncate_text(jp.text, 40))
            terms = [t for t in re.split(r"[,;]", re.sub(r"^\s*(key\s*words?|index\s+terms)\s*[—–\-:.]?\s*", "", jp.text, flags=re.I))
                     if t.strip() and not t.strip().startswith("(")]
            if len(terms) > spec.MAX_KEYWORDS:
                self._add("structure", f"Keywords (paragraph {jp.index})", jp.index,
                          "More than 5 keywords are listed", f"{len(terms)} keywords", "Up to 5 keywords",
                          "warning", truncate_text(jp.text, 60))

    # ---- headings --------------------------------------------------------
    def _check_headings(self):
        heads = [jp for jp in self.jpars if jp.role.startswith("heading")]
        for jp in heads:
            loc = f"Heading {jp.level} (paragraph {jp.index})"
            title = heading_text_only(jp)
            if jp.number_prefix:
                self._add("headings", loc, jp.index,
                          "Heading number is typed manually; the template numbers headings automatically (I., II. / A., B.)",
                          jp.number_prefix, "Automatic numbering", "warning", jp.text[:60])
            elif jp.level == 1 and not jp.component and not jp.has_auto_number:
                self._add("headings", loc, jp.index,
                          "Main heading is not numbered (Roman numeral numbering expected)",
                          "No number", "I., II., III. ...", "warning", jp.text[:60])
            elif jp.level == 2 and not jp.has_auto_number:
                self._add("headings", loc, jp.index,
                          "Sub-heading is not numbered (letter numbering expected)", "No number",
                          "A., B., C. ...", "warning", jp.text[:60])
            if jp.component and jp.has_auto_number:
                self._add("headings", loc, jp.index,
                          "Component heading (Acknowledgement, References, ...) must not be numbered",
                          "Numbered", "Unnumbered", "warning", jp.text[:60])
            if (jp.role == "heading1" or jp.component) and re.search(r"[a-z]", title):
                self._add("headings", loc, jp.index, "Heading text must be in capital letters",
                          title[:50], title[:50].upper(), "warning", jp.text[:60])
        # sub-heading rules: at least two sub-topics under a heading, no skipped level
        group: List[JetapPara] = []
        def flush():
            subs = [g for g in group if g.level == 2]
            if len(subs) == 1:
                self._add("headings", f"Heading 2 (paragraph {subs[0].index})", subs[0].index,
                          "A single sub-heading is used; the template requires at least two sub-topics or none",
                          "1 sub-heading", "0 or 2+ sub-headings", "info", subs[0].text[:60])
        for jp in heads:
            if jp.level == 1:
                flush()
                group = []
            group.append(jp)
        flush()
        prev_level = 0
        for jp in heads:
            if jp.level == 3 and prev_level < 2:
                self._add("headings", f"Heading 3 (paragraph {jp.index})", jp.index,
                          "Heading 3 is used without a Heading 2 above it", f"level {prev_level}", "Heading 2 first",
                          "info", jp.text[:60])
            prev_level = jp.level

    # ---- body conventions -------------------------------------------------
    def _check_body_conventions(self):
        bodies = [jp for jp in self.jpars if jp.role in ("body", "abstract")]
        numeric = sum(len(re.findall(r"\[\s*\d+(?:\s*[,–\-]\s*\d+)*\s*\]", jp.text)) for jp in bodies)
        author_year = [jp for jp in bodies if re.search(
            r"\(\s*[A-Z][A-Za-z\-À-ſ]+(?:\s+(?:et al\.?|and|&)\s*[A-Z]?[A-Za-z\-]*)?,?\s+(?:19|20)\d{2}[a-z]?\s*\)", jp.text)]
        if author_year and numeric < len(author_year):
            jp = author_year[0]
            self._add("references", "In-text citations", jp.index,
                      "In-text citations use author-year format; the template numbers citations in brackets [1]",
                      "Author-year citations", "[1], [2] ...", "warning", truncate_text(jp.text, 60))
        for jp in bodies:
            if re.search(r"\bFigure\s+\d", jp.text) or re.search(r"\bTable\s+\d", jp.text):
                self._add("figures", f"Body text (paragraph {jp.index})", jp.index,
                          "Figure/table mentions should read 'Fig. 1' and 'Table I' (Roman numeral)",
                          re.search(r"\b(?:Figure|Table)\s+\d", jp.text).group(0), "Fig. N / Table I", "info",
                          truncate_text(jp.text, 60))
                break

    # ---- tables / figures / captions --------------------------------------
    def _top_level_tables(self):
        out = []
        m = self._top_masthead()
        for t in self.document.tables:
            if m is not None and t._tbl is m:
                continue
            rows = t._tbl.findall(qn("w:tr"))
            cells = rows[0].findall(qn("w:tc")) if rows else []
            if len(rows) == 1 and len(cells) == 1:
                continue
            out.append(t)
        return out

    def _check_tables_and_captions(self):
        cap_tables = [jp for jp in self.jpars if jp.role == "caption_table" and jp.note != "title-line"]
        cap_figs = [jp for jp in self.jpars if jp.role == "caption_figure" and jp.note != "title-line"]
        for n, jp in enumerate(cap_tables, 1):
            m = _TAB_LABEL.match(jp.text)
            loc = f"Table caption (paragraph {jp.index})"
            if not m or jp.note == "title-line":
                continue
            num = m.group(2)
            roman_ok = from_roman(num) is not None and not num.isdigit()
            if m.group(1).lower() != "table" or not roman_ok or (m.group(3) or "") != ".":
                want = f"Table {to_roman(int(num)) if num.isdigit() else num.upper()}."
                self._add("tables", loc, jp.index, "Table caption label should be 'Table I.' (Roman numeral, then a full stop)",
                          jp.text[:m.end()].strip(), want, "warning", truncate_text(jp.text, 60))
            elif from_roman(num) != n:
                self._add("tables", loc, jp.index, "Table numbers are not consecutive",
                          f"Table {num}", f"Table {to_roman(n)}", "warning", truncate_text(jp.text, 60))
            nxt = jp.el.getnext()
            while nxt is not None and nxt.tag == qn("w:p") and is_plain_empty(nxt):
                nxt = nxt.getnext()
            if jp.note == "label-only":
                self._add("tables", loc, jp.index, "Table caption: the label 'Table I.' and the title must be on the same line",
                          "Label and title on separate lines", "Table I. Title.", "warning", truncate_text(jp.text, 60))
                continue
            if not jp.text.rstrip().endswith(".") or re.search(r"\s\.$", jp.text.rstrip()):
                self._add("tables", loc, jp.index, "Table caption must end with a full stop (no space before it)",
                          truncate_text(jp.text[-25:], 25), "... title.", "warning", truncate_text(jp.text, 60))
            if nxt is None or nxt.tag != qn("w:tbl"):
                self._add("tables", loc, jp.index, "Table caption should be placed above the table",
                          "Below / detached", "Above the table", "warning", truncate_text(jp.text, 60))
        for n, jp in enumerate(cap_figs, 1):
            m = _FIG_LABEL.match(jp.text)
            loc = f"Figure caption (paragraph {jp.index})"
            if not m or jp.note == "title-line":
                continue
            if m.group(1) != "Fig" or (m.group(3) or "") != ".":
                self._add("figures", loc, jp.index, "Figure caption label should be 'Fig. 1.' (abbreviated, then a full stop)",
                          jp.text[:m.end()].strip(), f"Fig. {m.group(2)}.", "warning", truncate_text(jp.text, 60))
            elif int(m.group(2)) != n:
                self._add("figures", loc, jp.index, "Figure numbers are not consecutive",
                          f"Fig. {m.group(2)}", f"Fig. {n}", "warning", truncate_text(jp.text, 60))
            if jp.note == "label-only":
                self._add("figures", loc, jp.index, "Figure caption: the label 'Fig. 1.' and the title must be on the same line",
                          "Label and title on separate lines", "Fig. 1. Title.", "warning", truncate_text(jp.text, 60))
                continue
            if not jp.text.rstrip().endswith(".") or re.search(r"\s\.$", jp.text.rstrip()):
                self._add("figures", loc, jp.index, "Figure caption must end with a full stop (no space before it)",
                          truncate_text(jp.text[-25:], 25), "... title.", "warning", truncate_text(jp.text, 60))
            prv = jp.el.getprevious()
            while prv is not None and prv.tag == qn("w:p") and is_plain_empty(prv):
                prv = prv.getprevious()
            if prv is not None and prv.tag == qn("w:p") and not paragraph_has_drawing(prv):
                nxt = jp.el.getnext()
                if nxt is not None and nxt.tag == qn("w:p") and paragraph_has_drawing(nxt):
                    self._add("figures", loc, jp.index, "Figure caption should be placed below the figure",
                              "Above the figure", "Below the figure", "warning", truncate_text(jp.text, 60))

        maxw = spec.COLUMN_WIDTH_TW
        for p in self.document.element.body.findall(qn("w:p")):
            tot = sum(int(e.get("cx")) for e in p.iter(qn("wp:extent")))
            if tot / 635 > maxw + 30:
                self._add("figures", "Figure", -1, "Figure is wider than the 3.0 in column",
                          f"{tot / 914400:.2f} in", "3.0 in or less", "warning")
        for n, t in enumerate(self._top_level_tables(), 1):
            loc = f"Table {n}"
            grid = [int(g.get(qn("w:w"))) for g in t._tbl.iter(qn("w:gridCol"))]
            if sum(grid) > maxw + 20:
                self._add("tables", loc, -1, "Table is wider than the 3.0 in column",
                          f"{sum(grid) / 1440:.2f} in", "3.0 in or less", "warning")
            tp = t._tbl.find(qn("w:tblPr"))
            tb = tp.find(qn("w:tblBorders")) if tp is not None else None
            sides = ("top", "left", "bottom", "right", "insideH", "insideV")
            full = tb is not None and all(tb.find(qn("w:" + s_)) is not None and tb.find(qn("w:" + s_)).get(qn("w:val")) == "single"
                                          for s_ in sides)
            for cb in t._tbl.iter(qn("w:tcBorders")):
                if any(c.get(qn("w:val")) in ("nil", "none") for c in cb):
                    full = False
                    break
            for eb in t._tbl.iter(qn("w:tblPrEx")):
                bb = eb.find(qn("w:tblBorders"))
                if bb is not None and any(c.get(qn("w:val")) in ("nil", "none") for c in bb):
                    full = False
                    break
            if not full:
                self._add("tables", loc, -1, "Table must have all borders (every horizontal and vertical line)",
                          "Only some borders", "All borders (grid)", "warning")
            sizes, fonts, aligns = Counter(), Counter(), Counter()
            for p in t._tbl.iter(qn("w:p")):
                for r in run_elements(p):
                    if not run_text(r).strip():
                        continue
                    pr = effective_run_props(self.document, p, r)
                    if pr["size"] and pr["vert"] not in ("superscript", "subscript"):
                        sizes[pr["size"]] += 1
                    if pr["font"] and pr["font"] != "(theme)":
                        fonts[pr["font"]] += 1
                if paragraph_text(p).strip():
                    aligns[effective_alignment(self.document, p)] += 1
            bad = [s for s in sizes if abs(s - 8) > 0.5]
            if bad:
                self._add("tables", loc, -1, "Table text font size is not 8pt",
                          f"{Counter({s: sizes[s] for s in bad}).most_common(1)[0][0]:g}pt", "8pt", "warning")
            badf = [f for f in fonts if f.lower() != spec.FONT.lower()]
            if badf:
                self._add("tables", loc, -1, f"Table text font is not {spec.FONT}", badf[0], spec.FONT, "warning")
            if aligns and aligns.most_common(1)[0][0] != "CENTER":
                self._add("tables", loc, -1, "Table text alignment is not center",
                          aligns.most_common(1)[0][0], "CENTER", "info")
            rows = t._tbl.findall(qn("w:tr"))
            if len(rows) > 1:
                first = rows[0]
                unbold = [r for r in first.iter(qn("w:r")) if run_text(r).strip()
                          and not effective_run_props(self.document, r.getparent(), r)["bold"]]
                if unbold:
                    self._add("tables", loc, -1, "Table column heads (first row) are not bold",
                              "Not bold", "Bold", "warning")

    # ---- references -------------------------------------------------------
    def _check_references(self):
        refs = [jp for jp in self.jpars if jp.role == "reference"]
        if not refs:
            return
        typed = sum(1 for jp in refs if re.match(r"^\s*\[\s*\d+\s*\]", jp.text))
        auto = sum(1 for jp in refs if jp.has_auto_number)
        if typed + auto < len(refs) * 0.6:
            self._add("references", "Reference list", refs[0].index,
                      "References are not numbered with brackets [1], [2] ...",
                      "Unnumbered / other numbering", "[1], [2] ... (automatic numbering)", "warning",
                      truncate_text(refs[0].text, 60))
        elif typed and not auto:
            self._add("references", "Reference list", refs[0].index,
                      "Reference numbers are typed manually instead of automatic numbering",
                      "Typed [n]", "Automatic [n] numbering", "info", truncate_text(refs[0].text, 60))
        ieee_start = re.compile(r"^\s*(?:\[\s*\d+\s*\]\s*)?(?:[A-Z]\.\s?){1,4}\s*[A-Za-z’'\-]")
        et_al = [jp for jp in refs if re.search(r"\bet al\.", re.sub(r"^\s*\[\s*\d+\s*\]\s*", "", jp.text))
                 and not ieee_start.match(jp.text)]
        apa = [jp for jp in refs if not ieee_start.match(jp.text) and (
               re.match(r"^[A-Z][A-Za-z\-']+,\s*[A-Z]\.", jp.text.strip())
               or re.search(r"\(\s*(?:19|20)\d{2}[a-z]?\s*\)\.", jp.text[:140]))]
        if apa:
            self._add("references", "Reference list", apa[0].index,
                      f"{len(apa)} of {len(refs)} references are in author-year (APA) form; the template uses IEEE style "
                      "(initials before surname, year at the end)",
                      truncate_text(apa[0].text, 50),
                      "G. Eason, B. Noble, \u201cTitle,\u201d Journal, vol., pp., year.", "warning", truncate_text(apa[0].text, 60))
        if et_al:
            self._add("references", "Reference list", et_al[0].index,
                      f"{len(et_al)} reference(s) use 'et al.'; give all authors' names unless there are six or more",
                      "et al.", "All author names", "info", truncate_text(et_al[0].text, 60))

    # ---- structure ----------------------------------------------------------
    def _check_structure(self) -> Dict[str, Any]:
        found: Dict[str, Optional[int]] = {k: None for k, _ in spec.REQUIRED_PARTS}
        for jp in self.jpars:
            if jp.role in ("abstract", "abstract_label") and found["abstract"] is None:
                found["abstract"] = jp.index
            elif jp.role == "keywords" and found["keywords"] is None:
                found["keywords"] = jp.index
            elif jp.role.startswith("heading"):
                k = part_key(heading_text_only(jp))
                if k and found.get(k) is None:
                    found[k] = jp.index
        sections: Dict[str, Any] = {}
        for key, label in spec.REQUIRED_PARTS:
            sections[label.lower()] = {"found": found[key] is not None, "index": found[key],
                                       "format_status": "not_checked"}
            if found[key] is None:
                self._add("structure", label, -1, f"Required part '{label}' was not found", "Missing",
                          label, "warning" if key in ("abstract", "keywords", "introduction", "conclusion", "references") else "info")
        order = [k for k, _ in spec.REQUIRED_PARTS]
        positions = [found[k] for k in order if found[k] is not None]
        order_ok = positions == sorted(positions)
        if not order_ok:
            self._add("structure", "Section order", -1, "Sections are not in the template order",
                      "Out of order", " > ".join(label for _, label in spec.REQUIRED_PARTS), "warning")
        app = next((jp for jp in self.jpars if jp.role.startswith("heading") and
                    heading_text_only(jp).lower().startswith("appendi")), None)
        if app and found["acknowledgement"] is not None and app.index > found["acknowledgement"]:
            self._add("structure", "Appendix position", app.index,
                      "Appendix should appear before the Acknowledgement", "After Acknowledgement",
                      "Before Acknowledgement", "warning", app.text[:40])
        return {"sections": sections, "declarations": {}, "order_correct": order_ok,
                "expected_order": [label.lower() for _, label in spec.REQUIRED_PARTS]}

    def _check_other(self):
        run = 0
        for jp in self.jpars:
            if jp.role == "empty":
                run += 1
                if run == 3:
                    self._add("other", f"Blank lines (paragraph {jp.index})", jp.index,
                              "Several consecutive blank paragraphs", "3+ blank paragraphs",
                              "Use paragraph spacing", "info")
            else:
                run = 0

    # ---- stats -----------------------------------------------------------------
    def _statistics(self) -> Dict[str, Any]:
        types = Counter(cp.paragraph_type.value for cp in self.classifications)
        words = sum(len(jp.text.split()) for jp in self.jpars)
        return {
            "total_paragraphs": len(self.document.paragraphs),
            "classified_paragraphs": len(self.classifications),
            "paragraph_types": dict(types), "total_words": words,
            "table_count": len(self._top_level_tables()),
            "figure_count": sum(1 for jp in self.jpars if jp.role == "caption_figure"),
            "reference_count": sum(1 for jp in self.jpars if jp.role == "reference"),
            "pages": len(self.document.sections),
        }
