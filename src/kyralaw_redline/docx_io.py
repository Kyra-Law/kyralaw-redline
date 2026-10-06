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

from kyralaw_redline.diff import DiffSegment, SegmentKind, compute_redline_segments

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


Mark = Literal["equal", "delete", "insert"]


def _break_ends(chars: list[tuple[SegmentKind, str]], visible: set[str]) -> tuple[set[int], set[int]]:
    """For one view, the stream positions of its paragraph-separator characters
    and of the last character of each separator.

    A separator is a run of two or more newlines in that view's text, which is
    how blocks are joined (see :func:`extract_text`). A single newline is a
    line break inside a paragraph, so it is content.
    """
    view = [i for i, (kind, _) in enumerate(chars) if kind in visible]
    sep: set[int] = set()
    ends: set[int] = set()
    run: list[int] = []
    for i in [*view, -1]:
        if i >= 0 and chars[i][1] == "\n":
            run.append(i)
            continue
        if len(run) >= 2:
            sep.update(run)
            ends.add(run[-1])
        run = []
    return sep, ends


def _segments_to_paragraphs(segments: list[DiffSegment]) -> list[tuple[list[DiffSegment], Mark]]:
    """Regroup a flat segment list into Word paragraphs, each with the kind of
    its paragraph mark.

    The segments describe two texts at once: the old one (equal + delete) and
    the new one (equal + insert). Each has its own paragraph breaks, and they
    need not coincide: deleting a paragraph removes a break from the old text
    only, splitting one adds a break to the new text only. So the breaks are
    found in each view separately, and a Word paragraph ends wherever either
    view's break ends. Its mark is untracked when both views break there,
    deleted when only the old one does, inserted when only the new one does;
    :func:`extract_paragraphs` (and Word) then join a paragraph whose mark is
    absent from a view onto the next, which gives back each view exactly.

    Each character is rendered in the views where it is content: plain where it
    is content in both, as a deletion or insertion where it is content in one
    only. Separator characters are never rendered; a newline that separates
    paragraphs in one view but is an in-paragraph line break in the other is
    rendered, as a line break, in the other view only.

    Correct by construction for any segment list, rather than by rules about
    which lines look blank. Two such rules failed: treating every
    whitespace-only line as a break split a paragraph at a lone space between
    two changes, and treating a blank line as a break in both views whatever
    segment it arrived in split or merged paragraphs whenever a break existed on
    one side only.
    """
    chars = [(seg.kind, ch) for seg in segments for ch in seg.text]
    old_sep, old_ends = _break_ends(chars, {"equal", "delete"})
    new_sep, new_ends = _break_ends(chars, {"equal", "insert"})

    # Newlines at the very start or end of either text are not content.
    def edge_newlines(visible: set[str]) -> set[int]:
        view = [i for i, (kind, _) in enumerate(chars) if kind in visible]
        out: set[int] = set()
        for seq in (view, list(reversed(view))):
            for i in seq:
                if chars[i][1] != "\n":
                    break
                out.add(i)
        return out

    old_sep |= edge_newlines({"equal", "delete"})
    new_sep |= edge_newlines({"equal", "insert"})
    last = len(chars) - 1

    paragraphs: list[tuple[list[DiffSegment], Mark]] = []
    current: list[DiffSegment] = []
    for i, (kind, ch) in enumerate(chars):
        in_old = kind in ("equal", "delete") and i not in old_sep
        in_new = kind in ("equal", "insert") and i not in new_sep
        if in_old or in_new:
            out_kind: SegmentKind = "equal" if in_old and in_new else ("delete" if in_old else "insert")
            if current and current[-1].kind == out_kind:
                current[-1] = DiffSegment(out_kind, current[-1].text + ch)
            else:
                current.append(DiffSegment(out_kind, ch))
        ends_old = i in old_ends and i != last
        ends_new = i in new_ends and i != last
        if ends_old or ends_new:
            mark: Mark = "equal" if ends_old and ends_new else ("delete" if ends_old else "insert")
            paragraphs.append((current, mark))
            current = []
    if current or not paragraphs:
        paragraphs.append((current, "equal"))
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
    paragraph: Paragraph,
    segments: list[DiffSegment],
    mark: Mark,
    author: str,
    date: str,
    ids: itertools.count[int],
) -> None:
    p = paragraph._p
    # A paragraph break that exists in one version only is a tracked paragraph
    # mark. A wholly deleted paragraph gets one too, which is what stops
    # accepting it in Word from leaving an empty paragraph behind.
    if mark != "equal":
        rpr = OxmlElement("w:rPr")
        rpr.append(_revision("del" if mark == "delete" else "ins", author, date, next(ids)))
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
    for paragraph_segments, mark in _segments_to_paragraphs(segments):
        _fill_paragraph(doc.add_paragraph(), paragraph_segments, mark, author, stamp, ids)
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
