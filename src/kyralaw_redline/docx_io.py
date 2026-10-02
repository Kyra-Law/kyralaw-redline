"""Read .docx text in, write a redline .docx with native tracked changes out.

:mod:`kyralaw_redline.diff` computes *what* changed. This module renders it
as OOXML ``w:ins``/``w:del`` markup, so the result opens in Word with real
Track Changes, and provides the minimal reader that turns a .docx into the
text the diff runs on.

**This module computes no diff of its own.** :func:`build_redline_docx` calls
:func:`~kyralaw_redline.diff.compute_redline_segments` and does nothing but
regroup the segments into paragraphs and render them. That is load-bearing:
a second, block-level diff in the renderer once made the shipped document a
measurably different algorithm from the one being evaluated, and nothing
surfaced the gap because the evaluation scored the segments, not the file.
The evaluation in this repository scores the rendered .docx for that reason.

**Known limitation**: the output is rebuilt from plain text, one paragraph
per block, in a blank document. Run formatting, styles, numbering, tables,
headers, footers, footnotes, images and hyperlinks are not carried over. This
produces a content-correct redline, not a formatting-faithful one.
"""

from __future__ import annotations

import io
import itertools
import zipfile
from datetime import datetime, timezone
from typing import Literal
from xml.etree import ElementTree

from docx import Document as _new_document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.xmlchemy import BaseOxmlElement
from docx.text.paragraph import Paragraph

from kyralaw_redline.diff import DiffSegment, compute_redline_segments

DEFAULT_AUTHOR = "kyralaw-redline"

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

View = Literal["accepted", "original"]

# Field instructions (the code behind a field, e.g. 'PAGE \\* MERGEFORMAT',
# not its result) are never body text.
_FIELD_CODES = {f"{_W}instrText", f"{_W}delInstrText"}

# Revision wrappers whose content is absent from each view: pending
# insertions and move destinations do not exist in the original, pending
# deletions and move sources do not exist once accepted.
_SKIPPED: dict[str, set[str]] = {
    "accepted": {f"{_W}del", f"{_W}moveFrom"} | _FIELD_CODES,
    "original": {f"{_W}ins", f"{_W}moveTo"} | _FIELD_CODES,
}


def _paragraph_text(p: ElementTree.Element, view: View) -> str:
    skipped = _SKIPPED[view]
    parts: list[str] = []

    def walk(el: ElementTree.Element) -> None:
        for child in el:
            tag = child.tag
            if tag in skipped:
                continue
            if tag == f"{_W}p":
                # A nested paragraph (inside a text box) is its own block, and
                # is reached by the outer iteration in extract_text.
                continue
            if tag == f"{_W}t" or (tag == f"{_W}delText" and view == "original"):
                parts.append(child.text or "")
            elif tag == f"{_W}tab":
                parts.append("\t")
            elif tag in (f"{_W}br", f"{_W}cr"):
                parts.append("\n")
            elif tag == f"{_W}noBreakHyphen":
                parts.append("\u2011")  # non-breaking hyphen
            else:
                walk(child)

    walk(p)
    return "".join(parts)


# A paragraph whose own mark is revised exists in one view only: once its
# mark is gone, its text runs on into the next paragraph, as in Word.
_ABSENT_MARK: dict[str, str] = {"accepted": f"{_W}del", "original": f"{_W}ins"}


def _mark_revision(p: ElementTree.Element) -> str | None:
    """``w:ins`` or ``w:del`` if the paragraph mark itself is a tracked
    insertion or deletion (``w:pPr/w:rPr/w:ins|w:del``), else None."""
    rpr = p.find(f"{_W}pPr/{_W}rPr")
    if rpr is None:
        return None
    for tag in (f"{_W}del", f"{_W}ins"):
        if rpr.find(tag) is not None:
            return tag
    return None


def _body(docx_bytes: bytes) -> ElementTree.Element:
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        root = ElementTree.fromstring(z.read("word/document.xml"))
    body = root.find(f"{_W}body")
    if body is None:
        raise ValueError("word/document.xml has no w:body element")
    return body


def extract_paragraphs(docx_bytes: bytes, *, view: View = "accepted") -> list[str]:
    """Every paragraph in the main document part, empty ones included, as it
    reads with all pending changes accepted (default) or all rejected
    (``view="original"``).

    Paragraph-mark revisions are applied: a paragraph whose mark is deleted
    (when accepting) or inserted (when rejecting) runs on into the next
    paragraph instead of standing alone. Paragraphs inside table cells are
    included, in document order. See :func:`extract_text` for what else is
    and is not read.
    """
    absent = _ABSENT_MARK[view]
    paragraphs: list[str] = []
    carry = ""
    for p in _body(docx_bytes).iter(f"{_W}p"):
        text = carry + _paragraph_text(p, view)
        if _mark_revision(p) == absent:
            carry = text
            continue
        carry = ""
        paragraphs.append(text)
    if carry:
        paragraphs.append(carry)
    return paragraphs


def extract_text(docx_bytes: bytes, *, view: View = "accepted") -> str:
    """The body text of a .docx, one block per paragraph, blocks joined by a
    blank line.

    Pending tracked changes are resolved as if all accepted (the default:
    inserted text kept, deleted text dropped) or, with ``view="original"``,
    as if all rejected. Diffing a document's own pending markup against
    another document's text is wrong, because the still-present old wording
    inside a deletion pollutes the comparison.

    Every non-empty ``w:p`` in the main document part becomes one block, in
    document order, including paragraphs inside table cells; empty
    paragraphs are dropped (use :func:`extract_paragraphs` to keep them). A
    manual line break inside a paragraph becomes ``"\\n"`` and stays inside
    its block. Headers, footers, footnotes and comments are not read.
    """
    return "\n\n".join(t.strip("\n") for t in extract_paragraphs(docx_bytes, view=view) if t.strip())


def _segments_to_paragraphs(segments: list[DiffSegment]) -> list[list[DiffSegment]]:
    """Regroup a flat, line-level segment list into one list per paragraph.

    :func:`compute_redline_segments` works line by line, so neither boundary
    lines up with a segment: one segment can span several paragraphs, and one
    paragraph can be assembled from several segments of different kinds. The
    paragraph breaks are the blank lines between blocks; they are structure
    rather than content, so they end a paragraph and are never rendered,
    whatever kind of segment they arrived in.

    Line endings inside a paragraph are kept (rendered as ``w:br``); the
    paragraph's own final line ending is dropped, on the old and new side
    independently.
    """
    paragraphs: list[list[DiffSegment]] = []
    current: list[DiffSegment] = []

    def flush() -> None:
        while current and not current[-1].text.strip():
            current.pop()
        if not current:
            return
        # The paragraph's final line ending is structure, on each side
        # separately: a replaced paragraph arrives as delete "old\n" then
        # insert "new\n", and stripping only the last segment would leave the
        # deletion ending in a line break that the original never had.
        for side in (("equal", "delete"), ("equal", "insert")):
            for k in range(len(current) - 1, -1, -1):
                if current[k].kind in side:
                    if current[k].text.endswith("\n"):
                        current[k] = DiffSegment(current[k].kind, current[k].text[:-1])
                    break
        paragraphs.append([seg for seg in current if seg.text])
        current.clear()

    for segment in segments:
        for line in segment.text.splitlines(keepends=True):
            if not line.strip():
                flush()
            elif current and current[-1].kind == segment.kind:
                current[-1] = DiffSegment(segment.kind, current[-1].text + line)
            else:
                current.append(DiffSegment(segment.kind, line))
    flush()
    return paragraphs


def _run(text: str, *, deleted: bool) -> BaseOxmlElement:
    """A ``w:r`` holding ``text``, with each newline rendered as ``w:br``."""
    r = OxmlElement("w:r")
    for k, piece in enumerate(text.split("\n")):
        if k:
            r.append(OxmlElement("w:br"))
        if piece:
            t = OxmlElement("w:delText" if deleted else "w:t")
            t.set(qn("xml:space"), "preserve")
            t.text = piece
            r.append(t)
    return r


def _revision(tag: str, author: str, date: str, rev_id: int) -> BaseOxmlElement:
    el = OxmlElement(f"w:{tag}")
    el.set(qn("w:id"), str(rev_id))
    el.set(qn("w:author"), author)
    el.set(qn("w:date"), date)
    return el


def _tracked(tag: str, text: str, author: str, date: str, rev_id: int) -> BaseOxmlElement:
    wrapper = _revision(tag, author, date, rev_id)
    wrapper.append(_run(text, deleted=tag == "del"))
    return wrapper


def _fill_paragraph(
    paragraph: Paragraph, segments: list[DiffSegment], author: str, date: str, ids: itertools.count[int]
) -> None:
    p = paragraph._p
    # A paragraph that is wholly deleted or wholly inserted has its paragraph
    # mark tracked too. Without that, accepting a deleted paragraph in Word
    # leaves an empty paragraph where it was, and rejecting an inserted one
    # does the same.
    kinds = {seg.kind for seg in segments if seg.text}
    if kinds in ({"delete"}, {"insert"}):
        rpr = OxmlElement("w:rPr")
        rpr.append(_revision("del" if kinds == {"delete"} else "ins", author, date, next(ids)))
        p.get_or_add_pPr().append(rpr)
    for seg in segments:
        if not seg.text:
            continue
        if seg.kind == "equal":
            p.append(_run(seg.text, deleted=False))
        elif seg.kind == "insert":
            p.append(_tracked("ins", seg.text, author, date, next(ids)))
        else:
            p.append(_tracked("del", seg.text, author, date, next(ids)))


def render_segments_docx(
    segments: list[DiffSegment], *, author: str = DEFAULT_AUTHOR, date: datetime | None = None
) -> bytes:
    """Render an already computed segment list as a .docx with native Word
    tracked changes.

    Blank lines in the segment text separate paragraphs. ``date`` stamps every
    revision and defaults to now (UTC); pass a fixed value for reproducible
    output. Exposed so that another diff can be rendered, and scored, through
    exactly the same writer.
    """
    doc = _new_document()
    ids = itertools.count(1)
    stamp = (date or datetime.now(timezone.utc)).strftime("%Y-%m-%dT%H:%M:%SZ")
    for paragraph_segments in _segments_to_paragraphs(segments):
        _fill_paragraph(doc.add_paragraph(), paragraph_segments, author, stamp, ids)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def build_redline_docx(
    old_text: str, new_text: str, *, author: str = DEFAULT_AUTHOR, date: datetime | None = None
) -> bytes:
    """Diff ``old_text`` against ``new_text`` and render the result as a .docx
    with native Word tracked changes.

    Blocks in both texts are separated by blank lines, as :func:`extract_text`
    produces them. See :func:`render_segments_docx` for ``date``.
    """
    return render_segments_docx(compute_redline_segments(old_text, new_text), author=author, date=date)


def redline_docx(
    old_docx: bytes, new_docx: bytes, *, author: str = DEFAULT_AUTHOR, date: datetime | None = None
) -> bytes:
    """Compare two .docx files and return a redline .docx with tracked changes.

    Equivalent to ``build_redline_docx(extract_text(old), extract_text(new))``;
    see both for what is and is not carried over.
    """
    return build_redline_docx(extract_text(old_docx), extract_text(new_docx), author=author, date=date)


def read_tracked_changes(docx_bytes: bytes) -> tuple[list[str], list[str]]:
    """The text of every ``w:ins`` and every ``w:del`` in a .docx body.

    Returns ``(inserted, deleted)``, one string per revision element in
    document order, omitting revisions that carry no text (a revised
    paragraph mark, for instance). Used by the evaluation to score the rendered file rather
    than the segments, and useful for reading a Word Compare output the same
    way.
    """
    body = _body(docx_bytes)

    def text_of(el: ElementTree.Element, text_tag: str) -> str:
        out: list[str] = []
        for node in el.iter():
            if node.tag == f"{_W}{text_tag}":
                out.append(node.text or "")
            elif node.tag in (f"{_W}br", f"{_W}cr"):
                out.append("\n")
            elif node.tag == f"{_W}tab":
                out.append("\t")
        return "".join(out)

    inserted = [t for t in (text_of(el, "t") for el in body.iter(f"{_W}ins")) if t]
    deleted = [t for t in (text_of(el, "delText") for el in body.iter(f"{_W}del")) if t]
    return inserted, deleted
