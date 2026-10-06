"""Render one version of a corpus pair as a styled Word document.

The text of every paragraph is fixed by the corpus; this module only decides
how it looks, from the pair's ``layout`` and each block's style and table
position. Nothing here may change the text a reader extracts, and
``run_eval`` checks that for every document it renders.

What a v3 document carries, modelled on the styling of real negotiated
agreements (measured over a private set of 33: most have tables, half have
multi-level legal numbering, most are justified with explicit spacing, many
carry comments, a third have headers and footers with page fields):

- paragraph styles: Title (all caps), Subtitle, Heading 1 to 3, custom
  Clause / Sub-clause / Sub-sub-clause styles, List Bullet, table styles;
- direct multi-level numbering on each paragraph (``w:numPr``), either legal
  (1. / 1.1 / (a) / (i), headings numbered) or simple (1. / (a) / (i));
- tables with a repeating header row;
- run formatting: defined terms bold, Acts and Latin phrases italic,
  placeholders such as ``[NAME]`` highlighted;
- fields and structure: clause cross-references as REF fields to bookmarked
  headings, the agreement date in a content control, email addresses as
  hyperlinks, a page break before the signature section;
- a header, a footer with PAGE and NUMPAGES fields, and Word comments.

A corpus without ``layout`` (v1, v2) renders as before: one plain paragraph
per block in its named style, defined terms in bold.
"""

from __future__ import annotations

import hashlib
import io
import re
from typing import Any

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from generate import STYLES

_MONTHS = (
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
)
_DEFINED_TERM_RE = re.compile(r"(“[^”]{1,60}”|\"[A-Z][^\"]{0,60}\")")

# Inline rules, in priority order. Each match becomes one span; overlaps go to
# the earlier rule.
_RULES: list[tuple[str, re.Pattern[str]]] = [
    ("email", re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+")),
    (
        "date",
        re.compile(r"\b\d{1,2} (?:" + "|".join(_MONTHS) + r") \d{4}\b"),
    ),
    ("ref", re.compile(r"\bclause \d+(?:\.\d+)+\b")),
    ("placeholder", re.compile(r"\[(?:●|[A-Z ]{2,20})\]")),
    ("defined", _DEFINED_TERM_RE),
    (
        "act",
        re.compile(r"\b[A-Z][\w,']*(?:\s(?:\(?[A-Z][\w,()']*|of|and|the|for))*\s(?:Act|Order) \d{4}\b"),
    ),
    ("latin", re.compile(r"\b(?:mutatis mutandis|ejusdem generis|inter alia|pro rata|bona fide)\b")),
]

COMMENT_TEXTS = (
    "Please confirm this reflects the agreed position.",
    "Check against the term sheet.",
    "Counterparty to confirm.",
    "Is this period long enough?",
    "Consider a cap here.",
)


def _spans(text: str) -> list[tuple[str, str]]:
    """``text`` cut into (kind, piece) spans, kind "" for plain text."""
    taken: list[tuple[int, int, str]] = []
    for kind, rx in _RULES:
        for m in rx.finditer(text):
            if any(m.start() < e and s < m.end() for s, e, _ in taken):
                continue
            taken.append((m.start(), m.end(), kind))
    taken.sort()
    out: list[tuple[str, str]] = []
    pos = 0
    for s, e, kind in taken:
        if s > pos:
            out.append(("", text[pos:s]))
        out.append((kind, text[s:e]))
        pos = e
    if pos < len(text):
        out.append(("", text[pos:]))
    return out


def _stable(text: str) -> int:
    return int(hashlib.sha1(text.encode()).hexdigest()[:8], 16)


# Document-level setup ------------------------------------------------------


def _ensure_style(doc: Any, name: str, base: str, *, bold: bool = False, keep_next: bool = False) -> Any:
    if name in [s.name for s in doc.styles]:
        return doc.styles[name]
    style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    style.base_style = doc.styles[base]
    style.font.bold = bold or None
    if keep_next:
        style.paragraph_format.keep_with_next = True
    return style


def _setup_styles(doc: Any, layout: dict[str, str]) -> None:
    font = layout.get("font") or "Times New Roman"
    size = Pt(float(layout.get("size") or 11))
    normal = doc.styles["Normal"]
    normal.font.name = font
    normal.font.size = size
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), font)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    title = doc.styles["Title"]
    title.font.all_caps = True
    title.font.bold = True
    title.font.name = font
    title.font.size = Pt(16)
    title.font.color.rgb = RGBColor(0, 0, 0)
    title.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle = doc.styles["Subtitle"]
    subtitle.font.name = font
    subtitle.font.italic = True
    subtitle.font.color.rgb = RGBColor(0x40, 0x40, 0x40)
    subtitle.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # The template's title and heading styles take their font from the theme,
    # which overrides a plain font name; drop the theme reference.
    for name in ("Title", "Subtitle", "Heading 1", "Heading 2", "Heading 3"):
        rpr = doc.styles[name].element.get_or_add_rPr()
        fonts = rpr.get_or_add_rFonts()
        for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
            fonts.attrib.pop(qn(attr), None)
        fonts.set(qn("w:ascii"), font)
        fonts.set(qn("w:hAnsi"), font)
    for n in (1, 2, 3):
        h = doc.styles[f"Heading {n}"]
        h.font.name = font
        h.font.size = size if n > 1 else Pt(size.pt + 1)
        h.font.bold = True
        h.font.color.rgb = RGBColor(0, 0, 0)
        h.paragraph_format.keep_with_next = True
        h.paragraph_format.space_before = Pt(12 if n == 1 else 6)
    _ensure_style(doc, "Clause", "Normal")
    _ensure_style(doc, "Sub-clause", "Normal")
    _ensure_style(doc, "Sub-sub-clause", "Normal")
    cell = _ensure_style(doc, "Table Text", "Normal")
    cell.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    cell.paragraph_format.space_after = Pt(0)
    _ensure_style(doc, "Table Heading", "Table Text", bold=True, keep_next=True)


# The level each style takes in each numbering scheme, and the levels' format.
_LEVELS = {
    "legal": {"h1": 0, "clause": 1, "sub": 2, "subsub": 3},
    "simple": {"clause": 0, "sub": 1, "subsub": 2},
}
_FORMATS = {
    "legal": [("decimal", "%1."), ("decimal", "%1.%2"), ("lowerLetter", "(%3)"), ("lowerRoman", "(%4)")],
    "simple": [("decimal", "%1."), ("lowerLetter", "(%2)"), ("lowerRoman", "(%3)")],
}


def _setup_numbering(doc: Any, scheme: str) -> int | None:
    """A multi-level list definition for ``scheme``; returns its numId."""
    if scheme not in _FORMATS:
        return None
    numbering = doc.part.numbering_part.element
    abstract_id = 90
    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    mlt = OxmlElement("w:multiLevelType")
    mlt.set(qn("w:val"), "multilevel")
    abstract.append(mlt)
    for ilvl, (fmt, text) in enumerate(_FORMATS[scheme]):
        lvl = OxmlElement("w:lvl")
        lvl.set(qn("w:ilvl"), str(ilvl))
        for tag, val in (("w:start", "1"), ("w:numFmt", fmt), ("w:lvlText", text), ("w:lvlJc", "left")):
            el = OxmlElement(tag)
            el.set(qn("w:val"), val)
            lvl.append(el)
        ppr = OxmlElement("w:pPr")
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), str(720 * (ilvl + 1)))
        ind.set(qn("w:hanging"), "720")
        ppr.append(ind)
        lvl.append(ppr)
        abstract.append(lvl)
    # abstractNum elements must precede num elements.
    first_num = numbering.find(qn("w:num"))
    if first_num is not None:
        first_num.addprevious(abstract)
    else:
        numbering.append(abstract)
    num_id = 90
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    ref = OxmlElement("w:abstractNumId")
    ref.set(qn("w:val"), str(abstract_id))
    num.append(ref)
    numbering.append(num)
    return num_id


def _number(paragraph: Any, num_id: int, ilvl: int) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    numpr = OxmlElement("w:numPr")
    lvl = OxmlElement("w:ilvl")
    lvl.set(qn("w:val"), str(ilvl))
    nid = OxmlElement("w:numId")
    nid.set(qn("w:val"), str(num_id))
    numpr.append(lvl)
    numpr.append(nid)
    ppr.insert_element_before(numpr, *_PPR_AFTER_NUMPR)


_PPR_AFTER_NUMPR = (
    "w:suppressLineNumbers",
    "w:pBdr",
    "w:shd",
    "w:tabs",
    "w:suppressAutoHyphens",
    "w:kinsoku",
    "w:wordWrap",
    "w:overflowPunct",
    "w:topLinePunct",
    "w:autoSpaceDE",
    "w:autoSpaceDN",
    "w:bidi",
    "w:adjustRightInd",
    "w:snapToGrid",
    "w:spacing",
    "w:ind",
    "w:contextualSpacing",
    "w:mirrorIndents",
    "w:suppressOverlap",
    "w:jc",
    "w:textDirection",
    "w:textAlignment",
    "w:textboxTightWrap",
    "w:outlineLvl",
    "w:divId",
    "w:cnfStyle",
    "w:rPr",
    "w:sectPr",
    "w:pPrChange",
)


def _field_runs(paragraph: Any, instr: str, result: str) -> None:
    """A complex field: begin, instruction, separate, result text, end."""
    p = paragraph._p

    def run_with(child: Any) -> None:
        r = OxmlElement("w:r")
        r.append(child)
        p.append(r)

    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    run_with(begin)
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = f" {instr} "
    run_with(it)
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    run_with(sep)
    paragraph.add_run(result)
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run_with(end)


def _hyperlink(paragraph: Any, address: str) -> None:
    r_id = paragraph.part.relate_to(f"mailto:{address}", RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), r_id)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rpr.append(color)
    rpr.append(u)
    run.append(rpr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = address
    run.append(t)
    link.append(run)
    paragraph._p.append(link)


def _content_control(paragraph: Any, text: str, tag: str) -> None:
    sdt = OxmlElement("w:sdt")
    pr = OxmlElement("w:sdtPr")
    alias = OxmlElement("w:alias")
    alias.set(qn("w:val"), tag)
    pr.append(alias)
    t = OxmlElement("w:tag")
    t.set(qn("w:val"), tag)
    pr.append(t)
    sdt.append(pr)
    content = OxmlElement("w:sdtContent")
    run = OxmlElement("w:r")
    wt = OxmlElement("w:t")
    wt.set(qn("xml:space"), "preserve")
    wt.text = text
    run.append(wt)
    content.append(run)
    sdt.append(content)
    paragraph._p.append(sdt)


def _bookmark(paragraph: Any, name: str, bid: int) -> None:
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bid))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bid))
    p = paragraph._p
    ppr = p.find(qn("w:pPr"))
    if ppr is not None:
        ppr.addnext(start)
    else:
        p.insert(0, start)
    p.append(end)


def _fill(paragraph: Any, text: str, bookmarks: list[str], *, rich: bool) -> None:
    """Write ``text`` into ``paragraph``, formatted by the inline rules."""
    if not rich:
        for k, piece in enumerate(_DEFINED_TERM_RE.split(text)):
            if piece:
                paragraph.add_run(piece).bold = k % 2 == 1
        return
    for kind, piece in _spans(text):
        if kind == "":
            paragraph.add_run(piece)
        elif kind == "defined":
            paragraph.add_run(piece).bold = True
        elif kind in ("act", "latin"):
            paragraph.add_run(piece).italic = True
        elif kind == "placeholder":
            paragraph.add_run(piece).font.highlight_color = WD_COLOR_INDEX.YELLOW
        elif kind == "email":
            _hyperlink(paragraph, piece)
        elif kind == "date":
            _content_control(paragraph, piece, "Date")
        elif kind == "ref":
            target = bookmarks[_stable(piece) % len(bookmarks)] if bookmarks else "_Ref000"
            _field_runs(paragraph, f"REF {target} \\r \\h", piece)


_STAMP = "2026-01-01T00:00:00Z"


def _revision(tag: str, ids: list[int]) -> Any:
    ids[0] += 1
    el = OxmlElement(tag)
    el.set(qn("w:id"), str(ids[0]))
    el.set(qn("w:author"), "ground truth")
    el.set(qn("w:date"), _STAMP)
    return el


def _fill_tracked(
    paragraph: Any, segs: list[tuple[str, str]], mark: str | None, bookmarks: list[str], rich: bool, ids: list[int]
) -> None:
    """A paragraph of the reference redline: unchanged text formatted as in
    the versions, inserted and deleted text as w:ins / w:del, and the
    paragraph mark itself tracked when the whole paragraph is new or gone."""
    if mark:
        rpr = OxmlElement("w:rPr")
        rpr.append(_revision("w:ins" if mark == "ins" else "w:del", ids))
        paragraph._p.get_or_add_pPr().append(rpr)
    for kind, text in segs:
        if kind == "equal":
            _fill(paragraph, text, bookmarks, rich=rich)
            continue
        wrapper = _revision("w:ins" if kind == "insert" else "w:del", ids)
        run = OxmlElement("w:r")
        t = OxmlElement("w:t" if kind == "insert" else "w:delText")
        t.set(qn("xml:space"), "preserve")
        t.text = text
        run.append(t)
        wrapper.append(run)
        paragraph._p.append(wrapper)


def _header_footer(doc: Any, layout: dict[str, str]) -> None:
    section = doc.sections[0]
    if layout.get("header"):
        hp = section.header.paragraphs[0]
        hp.text = layout["header"]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        for run in hp.runs:
            run.font.size = Pt(8)
    if layout.get("footer") == "page":
        fp = section.footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fp.add_run("Page ")
        _field_runs(fp, "PAGE", "1")
        fp.add_run(" of ")
        _field_runs(fp, "NUMPAGES", "1")


def _fixed_properties(doc: Any) -> None:
    from datetime import datetime

    cp = doc.core_properties
    cp.author = "kyralaw-redline eval"
    cp.last_modified_by = "kyralaw-redline eval"
    cp.created = cp.modified = datetime(2026, 1, 1)
    cp.revision = 1


# The document --------------------------------------------------------------


def render(
    blocks: list[str],
    styles: list[str] | None,
    cells: list[Any] | None,
    layout: dict[str, str] | None,
    tracked: list[tuple[list[tuple[str, str]], str | None]] | None = None,
) -> bytes:
    """A .docx of one version: ``blocks`` are its paragraphs in reading order
    (table cells row by row), ``styles`` and ``cells`` the parallel style tags
    and table positions.

    With ``tracked``, the document is a redline instead: ``blocks`` are the
    paragraphs of both versions together, and each entry gives that
    paragraph's (kind, text) segments and its mark ("ins", "del" or None). A
    table row whose every cell is inserted (deleted) is a tracked row."""
    doc = Document()
    _fixed_properties(doc)
    styles = styles or ["normal"] * len(blocks)
    cells = cells or [None] * len(blocks)
    position: list[Any] = cells
    rich = bool(layout)
    num_id = None
    scheme = ""
    if layout:
        _setup_styles(doc, layout)
        scheme = layout.get("numbering", "")
        num_id = _setup_numbering(doc, scheme)
        _header_footer(doc, layout)

    # Bookmark every top-level heading; cross-references point at them.
    bookmarks = [f"_Ref{n:03d}" for n, s in enumerate(styles) if s == "h1"] if rich else []
    heading_count = 0
    commentable: list[Any] = []

    ids = [0]

    def fill(para: Any, k: int) -> None:
        if tracked is None:
            _fill(para, blocks[k], bookmarks, rich=rich)
        else:
            segs, mark = tracked[k]
            _fill_tracked(para, segs, mark if position[k] is None else None, bookmarks, rich, ids)

    n = 0
    while n < len(blocks):
        cell = cells[n]
        if cell is not None:
            # A table: every consecutive block of the same table.
            t = cell[0]
            end = n
            while end < len(blocks) and cells[end] is not None and cells[end][0] == t:
                end += 1
            rows: list[list[int]] = []
            last_row = None
            for k in range(n, end):
                if cells[k][1] != last_row:
                    rows.append([])
                    last_row = cells[k][1]
                rows[-1].append(k)
            ncols = max(len(r) for r in rows)
            table = doc.add_table(rows=len(rows), cols=ncols)
            table.style = doc.styles["Table Grid"]
            for ri, row in enumerate(rows):
                if all(styles[k] == "cellhead" for k in row) and rich:
                    trpr = table.rows[ri]._tr.get_or_add_trPr()
                    trpr.append(OxmlElement("w:tblHeader"))
                if tracked is not None:
                    marks = {tracked[k][1] for k in row}
                    if marks in ({"ins"}, {"del"}):
                        trpr = table.rows[ri]._tr.get_or_add_trPr()
                        trpr.append(_revision("w:ins" if marks == {"ins"} else "w:del", ids))
                for ci, k in enumerate(row):
                    para = table.cell(ri, ci).paragraphs[0]
                    para.style = doc.styles[STYLES[styles[k]]] if rich else doc.styles["Normal"]
                    fill(para, k)
            n = end
            continue

        style = styles[n]
        para = doc.add_paragraph(style=doc.styles[STYLES[style]])
        if rich and num_id is not None and style in _LEVELS.get(scheme, {}):
            _number(para, num_id, _LEVELS[scheme][style])
        if rich and style == "h1" and blocks[n].startswith(("Signature", "Schedule")):
            para.paragraph_format.page_break_before = True
        fill(para, n)
        if rich and style == "h1":
            _bookmark(para, bookmarks[heading_count], 1000 + heading_count)
            heading_count += 1
        if style in ("clause", "normal") and len(blocks[n].split()) >= 10 and para.runs:
            commentable.append(para)
        n += 1

    # Comments, on paragraphs chosen by their opening words, so the same
    # paragraph carries the same comment in both versions of a pair.
    wanted = int((layout or {}).get("comments") or 0)
    chosen = sorted(commentable, key=lambda p: _stable(p.text[:40]))[:wanted]
    for para in chosen:
        text = COMMENT_TEXTS[_stable(para.text[:40]) % len(COMMENT_TEXTS)]
        doc.add_comment(para.runs, text=text, author="Reviewer", initials="RV")

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
