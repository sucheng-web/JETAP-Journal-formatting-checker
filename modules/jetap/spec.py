"""
JETAP (Journal of Engineering Technology and Applied Physics) format spec.

Every value here was read from the official "MMU Press Journal Format - A4
JETAP rev2026" template: the effective formatting of its styles (Normal,
Title, Name and affiliation, Abstract, Keywords, Heading 1/2, Caption,
Table Head, List of References ...), its section properties (A4, 1" margins,
single-column front matter followed by a two-column body) and its
header/footer parts.

Spacing is stored in the same unit Word uses (twips / 20 = points) so the
fixer can write the exact template values back.
"""
from __future__ import annotations

from copy import deepcopy
from typing import Any, Dict

JOURNAL_NAME = "Journal of Engineering Technology and Applied Physics"
JOURNAL_SHORT = "JETAP"
ISSN_TEXT = "e-ISSN: 2682-8383"
DOI_PREFIX = "https://doi.org/10.33093/jetap."
FONT = "Times New Roman"

# Page / section geometry (twips)
PAGE = {
    "width_tw": 11906, "height_tw": 16838,
    "margin_tw": 1440, "header_tw": 850, "footer_tw": 624,
    "columns": 2, "column_width_tw": 4319, "column_gap_tw": 386,
    "text_width_tw": 9026,
}

# Role keys -> formatting. size in pt, spacing in pt, indents in twips.
#   line: Word "auto" line value in 240ths (228 = 0.95 multiple, 240 = single)
ROLES: Dict[str, Dict[str, Any]] = {
    "masthead": dict(size=28, bold=True, italic=False, align="CENTER", line=240,
                     before=0, after=6, first=0),
    "title": dict(size=24, bold=False, italic=False, align="CENTER", line=228,
                  before=8, after=14, first=289),
    # NOTE: the template's prose says author names 11pt and affiliations
    # italic 10pt, but its own sample paragraphs (style "Name and
    # affiliation") are 9pt italic.  The formatted sample is the ground truth
    # (the template passes its own check); set AUTHOR_BLOCK_SIZES below to
    # (11, 10, 10) to follow the written instruction instead.
    "authors": dict(size=9, bold=False, italic=True, align="CENTER", line=240,
                    before=0, after=0, first=289),
    "affiliation": dict(size=9, bold=False, italic=True, align="CENTER", line=240,
                        before=0, after=0, first=289),
    "corresponding": dict(size=9, bold=False, italic=False, align="CENTER", line=240,
                          before=0, after=0, first=289),
    "doi": dict(size=8, bold=False, italic=False, align="CENTER", line=240,
                before=0, after=0, first=289),
    "revision": dict(size=8, bold=False, italic=True, align="CENTER", line=240,
                     before=0, after=0, first=289),
    "abstract": dict(size=9, bold=True, italic=None, align="JUSTIFY", line=240,
                     before=0, after=6, first=289),
    "keywords": dict(size=9, bold=True, italic=True, align="JUSTIFY", line=240,
                     before=0, after=6, first=288),
    "heading1": dict(size=10, bold=False, italic=False, align="CENTER", line=240,
                     before=8, after=4, first=216),
    "heading2": dict(size=10, bold=False, italic=True, align="JUSTIFY", line=240,
                     before=8, after=4, first=None),
    "heading3": dict(size=10, bold=False, italic=True, align="JUSTIFY", line=240,
                     before=4, after=2, first=289),
    "body": dict(size=10, bold=None, italic=None, align="JUSTIFY", line=228,
                 before=0, after=6, first=289),
    "caption": dict(size=8, bold=False, italic=False, align="CENTER", line=240,
                    before=0, after=0, first=216),
    "reference": dict(size=8, bold=None, italic=None, align="JUSTIFY", line=228,
                      before=0, after=3, first=None, left=284, hanging=284),
    "table_head": dict(size=8, bold=True, italic=None, align="CENTER", line=240,
                       before=0, after=0, first=0),
    "table_cell": dict(size=8, bold=None, italic=None, align="CENTER", line=240,
                       before=0, after=0, first=0),
    "decl_body": dict(size=10, bold=None, italic=None, align="JUSTIFY", line=228,
                      before=None, after=None, first=289),
    "figure": dict(size=10, bold=None, italic=None, align="CENTER", line=240,
                   before=0, after=0, first=0),
}

# Headings that are "component heads": un-numbered, centred.
COMPONENT_HEADS = [
    "acknowledgement", "acknowledgements", "acknowledgment", "acknowledgments",
    "funding", "funding statement", "funding statements",
    "author contribution", "author contributions", "authors contribution",
    "authors contributions", "author contribution statement",
    "conflict of interest", "conflict of interests", "conflicts of interest",
    "conflict of interest statement",
    "ethics statement", "ethics statements", "ethical statement", "ethics",
    "data availability", "data availability statement", "data access statement",
    "appendix", "appendices", "references", "reference", "bibliography",
]

# Known numbered main sections (Heading 1, Roman numeral).
MAIN_SECTION_TERMS = [
    "introduction", "background", "literature review", "related work", "related works",
    "methodology", "methods", "method", "materials and methods", "experimental",
    "experimental setup", "results", "results and discussion", "results and discussions",
    "discussion", "discussions", "conclusion", "conclusions", "analysis", "findings",
    "future work", "summary", "recommendations",
]

# Parts the template contains (checked as "structure").
REQUIRED_PARTS = [
    ("abstract", "Abstract"),
    ("keywords", "Keywords"),
    ("introduction", "Introduction"),
    ("conclusion", "Conclusion"),
    ("acknowledgement", "Acknowledgement"),
    ("funding", "Funding Statement"),
    ("author_contribution", "Author Contributions"),
    ("conflict", "Conflict of Interests"),
    ("ethics", "Ethics Statements"),
    ("data", "Data Availability Statement"),
    ("references", "References"),
]

AFFILIATION_TERMS = [
    "universiti", "university", "faculty", "department", "dept.", "dept ", "institute",
    "institut", "school of", "college", "centre", "center", "laborator", "sdn bhd",
    "sdn. bhd", "hospital", "ltd", "inc.", "gmbh", "academy", "polytechnic", "politeknik",
    "jalan", "campus", "research", "malaysia", "singapore", "indonesia", "india", "china",
    "korea", "japan", "usa", "united states", "united kingdom", "uk", "australia",
    "thailand", "vietnam", "pakistan", "iraq", "iran", "saudi", "egypt", "nigeria",
]

MAX_KEYWORDS = 5

# Roles whose paragraphs all start with a 0.2 in special first-line indent.
FIRST_LINE_ROLES = ("body", "decl_body", "abstract", "keywords", "authors", "affiliation",
                    "corresponding", "doi", "revision")
# Roles forced to "multiple 0.95" line spacing (others keep their own value).
LINE_095_ROLES = ("body", "decl_body", "reference")
# Figures / tables are scaled down to the column width (twips).
COLUMN_WIDTH_TW = 4319

# (authors, affiliation, corresponding) font sizes: template sample = 9pt each,
# template prose = 11 / 10 / 10.
AUTHOR_BLOCK_SIZES = {"authors": 9, "affiliation": 9, "corresponding": 9}
for _k, _v in AUTHOR_BLOCK_SIZES.items():
    ROLES[_k]["size"] = _v

TEMPLATE_RULES_FOR_REPORTS = {
    "margins": {"left": 1.0, "right": 1.0, "top": 1.0, "bottom": 1.0},
    "title": {"font_name": FONT, "font_size": 24, "bold": False, "alignment": "CENTER"},
    "body": {"font_name": FONT, "font_size": 10, "line_spacing": 0.95, "bold": False,
             "space_after": 6, "space_before": 0, "alignment": "JUSTIFY"},
    "heading": {"font_name": FONT, "font_size": 10, "bold": False, "alignment": "CENTER"},
    "subheading": {"font_name": FONT, "font_size": 10, "italic": True, "alignment": "JUSTIFY"},
    "caption": {"font_name": FONT, "font_size": 8, "alignment": "CENTER"},
    "reference": {"font_name": FONT, "font_size": 8, "alignment": "JUSTIFY",
                  "hanging_indent": 0.2, "space_after": 3},
    "layout": {"columns": 2, "page_size": "A4", "orientation": "PORTRAIT"},
}


def build_rules() -> Dict[str, Any]:
    """Rules dict in the shape the rest of the app (report, guidance) expects."""
    rules = deepcopy(TEMPLATE_RULES_FOR_REPORTS)
    rules["_profile"] = {
        "name": "JETAP",
        "source": "default_profile",
        "description": "Journal of Engineering Technology and Applied Physics (MMU Press) template profile",
        "required_sections": [label.lower() for _, label in REQUIRED_PARTS],
    }
    rules["_jetap"] = {"roles": deepcopy(ROLES), "page": deepcopy(PAGE)}
    return rules
