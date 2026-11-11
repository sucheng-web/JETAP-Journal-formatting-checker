"""APA-like reference entries -> IEEE style, and author-year in-text citations -> [n].

Everything here is generic text processing (no hard-coded manuscript wording).  Entries that cannot be
parsed with confidence are returned as ``None`` so the caller can leave them untouched.
"""
from __future__ import annotations

import re
import unicodedata
from typing import Dict, List, Optional, Tuple

Segment = Tuple[str, bool]          # (text, italic)

SMALL_WORDS = {"a", "an", "the", "and", "or", "nor", "but", "of", "in", "on", "at", "to", "for", "by", "as",
               "vs", "vs.", "per", "via", "with", "from", "into", "onto", "over"}
MONTHS = {"january": "Jan.", "february": "Feb.", "march": "Mar.", "april": "Apr.", "may": "May", "june": "Jun.",
          "july": "Jul.", "august": "Aug.", "september": "Sep.", "october": "Oct.", "november": "Nov.",
          "december": "Dec."}

_DATE = r"\((?P<year>(?:19|20)\d{2}[a-z]?|n\.d\.|in press)(?:,\s*(?P<month>[A-Z][a-z]+)(?:\s+(?P<day>\d{1,2}))?)?\)"
_HEAD = re.compile(r"^(?P<authors>.+?)\s*" + _DATE + r"\.?\s*(?P<rest>.*)$", re.S)
_GIVEN = re.compile(r"^(?:[A-Z][A-Za-z\-]*\.?\s*)+$")
_JOURNAL = re.compile(
    r"^(?P<title>.+?[.?!])\s+(?P<journal>[^.?!]{3,200}?),\s*(?P<vol>\d+[A-Za-z]?)"
    r"(?:\s*\((?P<iss>[^)]+)\))?(?:,\s*(?P<pages>[A-Za-z]?\d+(?:\s*[–\-]\s*[A-Za-z]?\d+)?))?\.?$", re.S)
_CHAPTER = re.compile(
    r"^(?P<title>.+?[.?!])\s+In\s+(?P<book>.+?)\s*\((?:(?P<eds>[^()]*?),?\s*\(?Eds?\.\)?\s*,?\s*)?"
    r"pp?\.\s*(?P<pages>[\d]+(?:\s*[–\-]\s*\d+)?)\)\.?\s*(?P<pub>[^.]*?)\.?$", re.S)
_EDITION = re.compile(r"\s*\((\d+(?:st|nd|rd|th)\s+ed\.)\)")


# ------------------------------------------------------------------ helpers
def norm_key(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]", "", s.lower())


def title_case(t: str) -> str:
    """Capitalise every word (never lower-casing what is already capital), keep short function words small."""
    out = []
    for i, w in enumerate(re.split(r"(\s+)", t)):
        if not w.strip() or not w[0].isalpha():
            out.append(w)
            continue
        core = w.lower().strip(".,:;?!()[]\"“”")
        if w[0].islower() and (i == 0 or core not in SMALL_WORDS):
            w = w[0].upper() + w[1:]
        out.append(w)
    return "".join(out)


def clean_url(u: str) -> str:
    u = re.sub(r"\s+", "", u)
    u = re.sub(r"/(?:METRICS|BIBTEX|CITE/\w+|ASSET/.*)$", "", u, flags=re.I)
    return u.rstrip(".,;")


def _initials(given: str) -> str:
    return " ".join(p[0].upper() + "." for p in re.findall(r"[A-Za-z]+", given))


def parse_authors(s: str) -> Optional[Dict]:
    s0 = s.strip()
    s0 = re.sub(r"\bet al\.?", "", s0).strip(" ,")
    toks = [t.strip() for t in re.split(r",\s*", s0.replace("&", ",").replace(" and ", ", ")) if t.strip()]
    toks = [t for t in toks if t]
    if not toks:
        return None
    personal = len(toks) >= 2 and len(toks) % 2 == 0 and all(_GIVEN.match(toks[i]) for i in range(1, len(toks), 2)) \
        and all(not re.match(r"^(?:[A-Z]\.\s*)+$", toks[i]) for i in range(0, len(toks), 2))
    if personal:
        people = [(toks[i].rstrip("."), toks[i + 1]) for i in range(0, len(toks), 2)]
        names = [f"{_initials(g)} {sn}".strip() for sn, g in people]
        n = len(names)
        if n >= 6:
            text = f"{names[0]} et al."
        elif n == 1:
            text = names[0]
        elif n == 2:
            text = f"{names[0]} and {names[1]}"
        else:
            text = ", ".join(names[:-1]) + f" and {names[-1]}"
        return {"text": text, "surnames": [sn for sn, _ in people], "n": n, "org": False}
    org = s0.rstrip(".").strip()
    return {"text": org, "surnames": [org], "n": 1, "org": True}


def parse_apa(text: str) -> Optional[Dict]:
    text = re.sub(r"\s+", " ", text).strip()
    m = _HEAD.match(text)
    if not m or len(m.group("authors")) > 400:
        return None
    au = parse_authors(m.group("authors"))
    if not au:
        return None
    rest = m.group("rest").strip()
    ref = {"authors": au, "year": m.group("year"), "month": m.group("month"), "day": m.group("day"),
           "raw": text, "type": None, "doi": None, "url": None}
    um = re.search(r"(?:https?://|www\.)", rest)
    if um:
        url = clean_url(rest[um.start():])
        rest = rest[:um.start()].strip()
        dm = re.match(r"https?://(?:dx\.)?doi\.org/(.+)$", url, re.I)
        if dm:
            ref["doi"] = dm.group(1)
        else:
            ref["url"] = url if url.lower().startswith("http") else "https://" + url
    rest = _EDITION.sub(lambda mm: " §" + mm.group(1), rest)
    editions = re.findall(r"§(\d+(?:st|nd|rd|th) ed\.)", rest)
    rest = re.sub(r"\s*§\d+(?:st|nd|rd|th) ed\.", "", rest).strip()
    ref["edition"] = editions[0] if editions else None
    rest = rest.rstrip()
    if not rest:
        return None

    jm = _JOURNAL.match(rest)
    cm = _CHAPTER.match(rest)
    if jm and not re.search(r"\bpress\b|publish|university", jm.group("journal"), re.I):
        ref.update(type="article", title=jm.group("title").strip().rstrip("."), journal=jm.group("journal").strip(),
                   vol=jm.group("vol"), iss=jm.group("iss"), pages=(jm.group("pages") or "").replace(" ", ""))
        if jm.group("title").strip()[-1] in "?!":
            ref["title"] = jm.group("title").strip()
        return ref
    if cm:
        ref.update(type="chapter", title=cm.group("title").strip().rstrip("."), book=cm.group("book").strip().rstrip("."),
                   pages=cm.group("pages").replace(" ", ""), pub=(cm.group("pub") or "").strip())
        if cm.group("title").strip()[-1] in "?!":
            ref["title"] = cm.group("title").strip()
        return ref
    sm = re.match(r"^(?P<title>.+?[.?!])\s+In\s+(?P<tail>.+)$", rest, re.S)
    if sm:
        parts = [x.strip() for x in re.split(r"\.\s+(?=[A-Z])", sm.group("tail")) if x.strip()]
        book = " ".join(parts[:-1]) if len(parts) >= 2 else (parts[0] if parts else "")
        pub = parts[-1] if len(parts) >= 2 else ""
        ref.update(type="chapter", title=sm.group("title").strip().rstrip("."), book=book.rstrip("."),
                   pages="", pub=pub.rstrip("."))
        return ref
    sentences = re.split(r"(?<=[.?!])\s+(?=[A-Z\[“\"])", rest)
    sentences = [x.strip() for x in sentences if x.strip()]
    if len(sentences) >= 2:
        ref.update(type="book", title=" ".join(sentences[:-1]).rstrip("."), pub=sentences[-1].rstrip("."))
    else:
        ref.update(type="book", title=sentences[0].rstrip("."), pub="")
    return ref


# ------------------------------------------------------------------ IEEE formatting
def _date_text(ref: Dict) -> str:
    y = ref["year"]
    if y == "in press":
        return "in press"
    if y == "n.d.":
        return "n.d."
    mon = MONTHS.get((ref.get("month") or "").lower())
    y = re.sub(r"[a-z]$", "", y)
    if mon and ref.get("day"):
        return f"{mon} {ref['day']}, {y}"
    if mon:
        return f"{mon} {y}"
    return y


def _pages(p: str) -> str:
    if not p:
        return ""
    if re.search(r"[–\-]", p):
        return "pp. " + re.sub(r"\s*-\s*", "–", p)
    if len(re.sub(r"\D", "", p)) >= 6:
        return "Art. no. " + p
    return "p. " + p


def _tail(ref: Dict, parts: List[str]) -> str:
    bits = [b for b in parts if b]
    bits.append(_date_text(ref))
    s = ", ".join(bits)
    return s + "."


def format_ieee(ref: Dict) -> List[Segment]:
    a = ref["authors"]["text"]
    typ = ref["type"]
    segs: List[Segment] = []
    online = f" [Online]. Available: {ref['url']}" if ref.get("url") else ""
    if typ == "article":
        title = title_case(ref["title"])
        quote = f"“{title}” " if title[-1] in "?!" else f"“{title},” "
        segs = [(f"{a}, {quote}", False), (ref["journal"], True)]
        parts = [f"vol. {ref['vol']}"]
        if ref.get("iss"):
            parts.append(f"no. {ref['iss']}")
        parts.append(_pages(ref.get("pages", "")))
        segs.append((", " + _tail(ref, [p for p in [parts[0]] + parts[1:]]) + online, False))
        return segs
    if typ == "chapter":
        title = title_case(ref["title"])
        quote = f"“{title}” " if title[-1] in "?!" else f"“{title},” "
        segs = [(f"{a}, {quote}in ", False), (title_case(ref["book"]), True)]
        tail_parts = [ref.get("pub") or ""]
        pg = _pages(ref.get("pages", ""))
        s = ". " if ref.get("pub") else ", "
        bits = [b for b in [ref.get("pub")] if b]
        bits.append(_date_text(ref))
        if pg:
            bits.append(pg)
        tail = ", ".join(bits)
        segs.append((s + tail + "." + online if ref.get("pub") else ", " + tail + "." + online, False))
        return segs
    # book / report / web / other
    title = title_case(ref["title"])
    pub = ref.get("pub") or ""
    ed = f", {ref['edition']}" if ref.get("edition") else ""
    is_standalone = bool(pub) and not ref.get("url") and not re.search(r"\b(report|survey|study|statistics|data)\b", title, re.I)
    if is_standalone or (pub and re.search(r"press|publish|books|wiley|springer|sage|routledge|elsevier|oxford|cambridge",
                                            pub, re.I)):
        segs = [(f"{a}, ", False), (title, True),
                (ed + (" " if ed.endswith(".") else ". ") + _tail(ref, [pub]) + online, False)]
    else:
        quote = f"“{title}” " if title[-1] in "?!" else f"“{title},” "
        segs = [(f"{a}, {quote}", False), (_tail(ref, [pub]) + online, False)]
    return segs


def ieee_text(ref: Dict) -> str:
    return "".join(t for t, _ in format_ieee(ref))


# ------------------------------------------------------------------ in-text citations
def ref_keys(ref: Dict) -> Tuple[str, str]:
    return norm_key(ref["authors"]["surnames"][0]), re.sub(r"\s", "", ref["year"].lower())


_YEAR = r"(?:19|20)\d{2}[a-z]?"
_PAREN = re.compile(r"\(([^()]{3,400})\)")
_PIECE = re.compile(r"^(?:see|e\.g\.,?|cf\.|also|compare)?\s*(?P<names>.+?),?\s+(?P<years>" + _YEAR +
                    r"(?:\s*,\s*" + _YEAR + r")*)(?:\s*,\s*(?:pp?\.\s*)?[\d–\-]+)?\s*$", re.I)


def _lookup(index: Dict[Tuple[str, str], List[int]], names: str, year: str, refs: List[Dict]) -> Optional[int]:
    year = year.lower()
    candidates: List[int] = []
    first_words = re.split(r"\s*(?:,|&|\band\b|\bet al\b)\s*", names.strip())[0]
    k_try = [norm_key(first_words), norm_key(names)]
    for (kn, ky), idxs in index.items():
        if ky != year:
            continue
        for kt in k_try:
            if kt and (kt == kn or (len(kn) > 3 and kt.startswith(kn)) or (len(kt) > 3 and kn.startswith(kt))):
                candidates += idxs
                break
    candidates = sorted(set(candidates))
    if len(candidates) == 1:
        return candidates[0]
    if len(candidates) > 1:
        # disambiguate with the second surname
        low = norm_key(names)
        best = [i for i in candidates if all(norm_key(sn) in low for sn in refs[i]["authors"]["surnames"][:2])]
        if len(best) == 1:
            return best[0]
    return None


def convert_citations(text: str, refs: List[Dict], index, order: List[int]) -> Tuple[List[Tuple[int, int, str]], int]:
    """Return (replacements, unresolved) for one paragraph's text; `order` (first-citation sequence) is updated."""
    repl: List[Tuple[int, int, str]] = []
    unresolved = 0
    taken: List[Tuple[int, int]] = []

    def number(i: int) -> int:
        if i not in order:
            order.append(i)
        return order.index(i) + 1

    def fmt(nums: List[int]) -> str:
        nums = sorted(set(nums))
        out, i = [], 0
        while i < len(nums):
            j = i
            while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
                j += 1
            if j - i >= 2:
                out.append(f"[{nums[i]}]–[{nums[j]}]")
            else:
                out += [f"[{n}]" for n in nums[i:j + 1]]
            i = j + 1
        return ", ".join(out)

    found: List[Tuple[int, int, List[int], str]] = []     # start, end, ref indices, kind
    for m in _PAREN.finditer(text):
        inner = m.group(1)
        if not re.search(_YEAR, inner):
            continue
        idxs, ok = [], True
        for piece in re.split(r";\s*", inner):
            pm = _PIECE.match(piece.strip())
            if not pm:
                ok = False
                break
            for y in re.split(r"\s*,\s*", pm.group("years")):
                i = _lookup(index, pm.group("names"), y, refs)
                if i is None:
                    ok = False
                    break
                idxs.append(i)
            if not ok:
                break
        if ok and idxs:
            found.append((m.start(), m.end(), idxs, "paren"))
            taken.append((m.start(), m.end()))
        elif re.search(r"[A-Za-z]{3,}.*" + _YEAR, inner):
            unresolved += 1
    # narrative: Name (2019) / Name et al. (2019) / Name and Name (2019)
    for m in re.finditer(r"(?P<names>[A-Z][\w’'\-]+(?:(?:,\s*|\s+)[A-Za-z][\w’'\-]+){0,6}?(?:\s+et al\.?)?)(?:[’']s)?\s*\((?P<years>" +
                         _YEAR + r"(?:\s*,\s*" + _YEAR + r")*)\)", text):
        if any(a <= m.start() < b for a, b in taken):
            continue
        names = m.group("names")
        words = names.replace(",", ", ").split()
        hit = None
        for k in range(len(words)):                     # try the longest tail that matches a reference
            cand = " ".join(words[k:])
            ids = [_lookup(index, cand, y, refs) for y in re.split(r"\s*,\s*", m.group("years"))]
            if all(i is not None for i in ids):
                hit = (m.start("names") + len(" ".join(words[:k])) + (1 if k else 0), ids)
                break
        if hit:
            start, ids = hit
            yr_start = m.start("years") - 1
            found.append((yr_start, m.end(), ids, "narr"))
        else:
            unresolved += 1
    found.sort(key=lambda f: f[0])
    for a, b, ids, kind in found:
        nums = [number(i) for i in ids]
        repl.append((a, b, fmt(nums) if kind == "paren" else " " + fmt(nums) if False else fmt(nums)))
    return repl, unresolved


def build_index(refs: List[Dict]) -> Dict[Tuple[str, str], List[int]]:
    idx: Dict[Tuple[str, str], List[int]] = {}
    for i, r in enumerate(refs):
        idx.setdefault(ref_keys(r), []).append(i)
        # allow "2019" to match "2019a"
        k = ref_keys(r)
    return idx
