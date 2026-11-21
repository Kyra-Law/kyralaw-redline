# SPDX-License-Identifier: MIT
# Copyright (c) 2026 Kyra Law Limited
"""The in-place renderer: redline two .docx files inside the revised document's own Word XML.

kyralaw-redline renders one diff two ways:

* **rebuild** (:func:`kyralaw_redline.docx_io.build_redline_docx`) writes the diff
  of the two documents' text into a new, blank document. The redline is
  content-correct, but tables become paragraphs and styles, numbering, headers,
  footers and comments are lost.
* **in-place** (this module) takes the route Word's own Compare takes: the revised
  document *is* the output. Its XML is kept as it is (styles, numbering, tables,
  headers, footers, comments, sections) and the differences from the base
  document are marked inside it as real tracked changes:

  * text only the revised document has is wrapped in ``w:ins``;
  * text only the base document had is written back as ``w:del`` runs, carrying
    the base document's own run formatting;
  * paragraphs, table rows and tables only the base had are copied in from it as
    deletions, next to the neighbours they had there.

**What is compared is the same.** The paragraph alignment and the word diff are
:func:`kyralaw_redline.diff.align_paragraphs`, the same passes and heuristics the
rebuild renderer uses. Only where the result is written differs.

**Both inputs are first resolved to their all-accepted state**, as Word's
Compare does: a change still pending in either document is part of that
document, not part of the comparison, so the output carries only the
comparison's own revisions. The exception is ``keep_pending``, for a revised
document edited directly from the base, which keeps the pending revisions as
they were and checks the result per author (see :func:`redline_in_place`).

The guarantee
-------------
Before anything is returned, the result is read back two ways: accepting every
revision must give the revised document's paragraphs, and rejecting every
revision must give the base document's. A result failing either check is not
returned. :class:`InPlaceRedlineError` is raised instead, so a caller can fall
back to the rebuild renderer rather than hand a reader a redline that misstates
what changed (``redline_docx(..., mode="auto")`` does exactly that).

Known limits
------------
* Formatting-only changes are not marked (Word marks those as ``rPrChange``).
* Headers, footers, footnotes and text boxes are kept from the revised document
  but not compared, the same scope as the rebuild renderer.
* Inside *deleted* text, images, hyperlink targets, fields, bookmarks and
  comment anchors are not carried over; the deleted words and their run
  formatting are.
* A paragraph that moved is shown as a deletion plus an insertion.
"""

from __future__ import annotations

import copy
import functools
import io
import itertools
import posixpath
import zipfile
from collections import Counter
from collections.abc import Callable, Iterable, Iterator
from dataclasses import dataclass, field
from datetime import datetime, timezone

from lxml import etree

from kyralaw_redline.diff import DiffSegment, ParagraphAlignment, align_paragraphs
from kyralaw_redline.docx_io import DEFAULT_AUTHOR

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _w(tag: str) -> str:
    return f"{{{_W_NS}}}{tag}"


_BODY, _P, _R, _T = _w("body"), _w("p"), _w("r"), _w("t")
_TBL, _TR, _TC = _w("tbl"), _w("tr"), _w("tc")
_PPR, _RPR, _TRPR, _TCPR = _w("pPr"), _w("rPr"), _w("trPr"), _w("tcPr")
_TBLPR, _TBLGRID, _TBLPREX = _w("tblPr"), _w("tblGrid"), _w("tblPrEx")
_SECTPR, _SDT, _SDTCONTENT, _CUSTOMXML = _w("sectPr"), _w("sdt"), _w("sdtContent"), _w("customXml")
_INS, _DEL, _MOVEFROM, _MOVETO = _w("ins"), _w("del"), _w("moveFrom"), _w("moveTo")
_DELTEXT = _w("delText")
_ID = _w("id")

# Property-change records and range markers. Accepting a revision removes them;
# none carries text.
_PROPERTY_CHANGES = tuple(
    _w(t)
    for t in (
        "rPrChange",
        "pPrChange",
        "sectPrChange",
        "tblPrChange",
        "tblPrExChange",
        "tblGridChange",
        "trPrChange",
        "tcPrChange",
        "numberingChange",
        "cellIns",
        "cellDel",
        "cellMerge",
        "moveFromRangeStart",
        "moveFromRangeEnd",
        "moveToRangeStart",
        "moveToRangeEnd",
        "customXmlInsRangeStart",
        "customXmlInsRangeEnd",
        "customXmlDelRangeStart",
        "customXmlDelRangeEnd",
        "customXmlMoveFromRangeStart",
        "customXmlMoveFromRangeEnd",
        "customXmlMoveToRangeStart",
        "customXmlMoveToRangeEnd",
    )
)

# Elements that hold runs inside a paragraph. Runs are looked for through these
# and never inside a run, so a text box's paragraphs (which live inside a run's
# drawing) are not read as part of the paragraph that anchors it.
_RUN_CONTAINERS = frozenset(
    _w(t) for t in ("hyperlink", "fldSimple", "smartTag", "customXml", "sdt", "sdtContent", "dir", "bdo")
) | {_INS, _DEL, _MOVETO, _MOVEFROM}

# Run children that are one character of visible text, and what that character is.
_ONE_CHAR = {_w("tab"): "\t", _w("ptab"): "\t", _w("br"): "\n", _w("cr"): "\n", _w("noBreakHyphen"): "-"}

_BLOCKS = (_P, _TBL, _SDT, _CUSTOMXML)

# Comment, bookmark and permission ranges: a start and an end sharing a w:id.
_RANGE_PAIRS = tuple((_w(f"{name}Start"), _w(f"{name}End")) for name in ("commentRange", "bookmark", "perm"))
_ANCHORS = frozenset(tag for pair in _RANGE_PAIRS for tag in pair)

# The schema's order for a paragraph's numbering properties.
_NUMPR_ORDER = {_w(tag): rank for rank, tag in enumerate(("ilvl", "numId", "numberingChange", "ins"))}
_COMMENT_REFERENCE = _w("commentReference")

# Revision markers, each with a w:id of its own. Range starts and ends are left
# out: they pair with each other by id, so neither can be renumbered alone.
_REVISION_MARKERS = frozenset({_INS, _DEL, _MOVEFROM, _MOVETO}) | frozenset(
    tag for tag in _PROPERTY_CHANGES if not tag.endswith(("RangeStart", "RangeEnd"))
)
# The schema's order for the revision markers of a paragraph mark.
_MARK_ORDER = (_INS, _DEL, _MOVEFROM, _MOVETO)


class InPlaceRedlineError(ValueError):
    """The two documents could not be redlined in place.

    Raised for an input that is not a readable .docx, and for a result that
    fails its own read-back check. Either way nothing is returned, and the
    caller is expected to fall back to the rebuild renderer.
    """


# --------------------------------------------------------------------------
# Package I/O
# --------------------------------------------------------------------------


def _parser() -> etree.XMLParser:
    # A document can come from anyone: never resolve entities or touch the network.
    return etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=True)


def _main_part_name(package: zipfile.ZipFile) -> str:
    try:
        rels = etree.fromstring(package.read("_rels/.rels"), _parser())
    except (KeyError, etree.XMLSyntaxError) as exc:
        raise InPlaceRedlineError("not a .docx: no readable package relationships") from exc
    for rel in rels:
        if (rel.get("Type") or "").endswith("/officeDocument"):
            return posixpath.normpath((rel.get("Target") or "").lstrip("/"))
    raise InPlaceRedlineError("not a .docx: the package names no main document")


def _load(docx: bytes) -> tuple[zipfile.ZipFile, str, etree._Element]:
    try:
        package = zipfile.ZipFile(io.BytesIO(docx))
    except zipfile.BadZipFile as exc:
        raise InPlaceRedlineError("not a .docx: not a zip package") from exc
    name = _main_part_name(package)
    try:
        root = etree.fromstring(package.read(name), _parser())
    except (KeyError, etree.XMLSyntaxError) as exc:
        raise InPlaceRedlineError(f"not a .docx: unreadable main document part {name!r}") from exc
    _body(root)
    return package, name, root


def _body(root: etree._Element) -> etree._Element:
    body = root.find(_BODY)
    if body is None:
        raise InPlaceRedlineError("not a .docx: the main document has no body")
    return body


def _repackage(package: zipfile.ZipFile, main_name: str, root: etree._Element) -> bytes:
    xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target:
        for info in package.infolist():
            target.writestr(info, xml if info.filename == main_name else package.read(info.filename))
    return out.getvalue()


# --------------------------------------------------------------------------
# Tree helpers
# --------------------------------------------------------------------------


def _el(tag: str, text: str | None = None) -> etree._Element:
    element = etree.Element(_w(tag))
    if text is not None:
        element.text = text
        element.set(_XML_SPACE, "preserve")
    return element


def _remove(element: etree._Element) -> None:
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def _unwrap(element: etree._Element) -> None:
    parent = element.getparent()
    if parent is None:
        return
    at = parent.index(element)
    for offset, child in enumerate(list(element)):
        parent.insert(at + offset, child)
    parent.remove(element)


def _child_of(element: etree._Element, container: etree._Element) -> etree._Element | None:
    """The ancestor-or-self of ``element`` that is a direct child of ``container``."""
    node: etree._Element | None = element
    while node is not None:
        parent = node.getparent()
        if parent is container:
            return node
        node = parent
    return None


def _rows(table: etree._Element) -> Iterator[etree._Element]:
    for child in table:
        if child.tag == _TR:
            yield child
        elif child.tag == _SDT and (content := child.find(_SDTCONTENT)) is not None:
            yield from _rows(content)
        elif child.tag == _CUSTOMXML:
            yield from _rows(child)


def _cells(row: etree._Element) -> Iterator[etree._Element]:
    for child in row:
        if child.tag == _TC:
            yield child
        elif child.tag == _SDT and (content := child.find(_SDTCONTENT)) is not None:
            yield from _cells(content)
        elif child.tag == _CUSTOMXML:
            yield from _cells(child)


# --------------------------------------------------------------------------
# Accepting the revisions an input already carries
# --------------------------------------------------------------------------


def _accept_all(root: etree._Element) -> None:
    """Resolve every pending revision in place, as accepting them all in Word would."""
    for row in list(root.iter(_TR)):
        row_props = row.find(_TRPR)
        if row_props is None:
            continue
        if row_props.find(_DEL) is not None:
            _remove_keeping_anchors(row)
            continue
        for marker in row_props.findall(_INS):
            row_props.remove(marker)

    for element in list(root.iter(*_PROPERTY_CHANGES)):
        _remove(element)

    merged_away: list[etree._Element] = []
    for paragraph in root.iter(_P):
        mark = paragraph.find(f"{_PPR}/{_RPR}")
        if mark is None:
            continue
        for marker in list(mark):
            if marker.tag in (_INS, _MOVETO):
                mark.remove(marker)
            elif marker.tag in (_DEL, _MOVEFROM):
                mark.remove(marker)
                merged_away.append(paragraph)

    for element in list(root.iter(_DEL, _MOVEFROM)):
        _remove_keeping_anchors(element)
    for element in list(root.iter(_INS, _MOVETO)):
        _unwrap(element)
    # A deleted paragraph mark joins the paragraph to the next one. Reverse
    # order, so a run of them collapses into the paragraph that ends it.
    for paragraph in reversed(merged_away):
        _merge_into_next(paragraph)
    _balance_ranges(root)


def _remove_keeping_anchors(element: etree._Element) -> None:
    """Remove revised-away content, but keep the comment and bookmark anchors in it.

    A comment on text a pending deletion removes is still the reviewer's comment.
    Discarding the deletion with its anchors inside orphaned the other end of
    the comment's range, which is a damaged document to Word. A comment
    reference keeps its run; at row level, where a run cannot stand, only the
    range markers are kept.
    """
    parent = element.getparent()
    if parent is None:
        return
    runs_allowed = parent.tag == _P or parent.tag in _RUN_CONTAINERS
    keep: list[etree._Element] = []
    for node in list(element.iter(*_ANCHORS, _COMMENT_REFERENCE)):
        if node.tag != _COMMENT_REFERENCE:
            keep.append(node)
        elif runs_allowed:
            run = etree.Element(_R)
            source = node.getparent()
            props = source.find(_RPR) if source is not None else None
            if props is not None:
                run.append(copy.deepcopy(props))
            run.append(node)
            keep.append(run)
    at = parent.index(element)
    for offset, node in enumerate(keep):
        parent.insert(at + offset, node)
    parent.remove(element)


def _balance_ranges(root: etree._Element) -> None:
    """Drop range markers whose other end is gone, which Word rejects."""
    for start_tag, end_tag in _RANGE_PAIRS:
        starts = {element.get(_ID): element for element in root.iter(start_tag)}
        ends = {element.get(_ID): element for element in root.iter(end_tag)}
        for key in starts.keys() - ends.keys():
            _remove(starts[key])
        for key in ends.keys() - starts.keys():
            _remove(ends[key])


def _merge_into_next(paragraph: etree._Element) -> None:
    following = paragraph.getnext()
    while following is not None and following.tag not in _BLOCKS:
        following = following.getnext()
    if following is None or following.tag != _P:
        return  # the last paragraph of its container keeps its mark
    at = 1 if following.find(_PPR) is not None else 0
    for offset, child in enumerate([c for c in paragraph if c.tag != _PPR]):
        following.insert(at + offset, child)
    _remove(paragraph)


# --------------------------------------------------------------------------
# Reading paragraphs
# --------------------------------------------------------------------------


def _atom(child: etree._Element) -> str | None:
    """The visible text one run child contributes, or ``None`` for a child with none."""
    tag = child.tag
    if tag == _T:
        return child.text or ""
    if tag in _ONE_CHAR:
        return _ONE_CHAR[tag]
    if tag == _w("sym"):
        try:
            return chr(int(child.get(_w("char"), ""), 16))
        except ValueError:
            return ""
    return None


def _iter_runs(container: etree._Element, skip: frozenset[str] = frozenset()) -> Iterator[etree._Element]:
    for child in container:
        if child.tag == _R:
            yield child
        elif child.tag in _RUN_CONTAINERS and child.tag not in skip:
            yield from _iter_runs(child, skip)


@dataclass
class _Span:
    run: etree._Element
    start: int
    end: int


@dataclass
class _Para:
    element: etree._Element
    # One (table, row, cell) per level of table nesting, outermost first.
    path: tuple[tuple[etree._Element, etree._Element, etree._Element], ...]
    text: str = ""
    spans: list[_Span] = field(default_factory=list)
    # Run containers whose runs are not read: the pending deletions, when pending
    # revisions are kept rather than accepted (see `keep_pending`).
    skip: frozenset[str] = frozenset()

    def reindex(self) -> None:
        parts: list[str] = []
        self.spans = []
        position = 0
        for run in _iter_runs(self.element, self.skip):
            text = "".join(a for child in run if (a := _atom(child)) is not None)
            self.spans.append(_Span(run, position, position + len(text)))
            parts.append(text)
            position += len(text)
        self.text = "".join(parts)


def _collect(
    container: etree._Element,
    path: tuple[tuple[etree._Element, etree._Element, etree._Element], ...],
    out: list[_Para],
    skip: frozenset[str] = frozenset(),
) -> None:
    """Every paragraph of a container in reading order, table cells included."""
    for child in container:
        if child.tag == _P:
            paragraph = _Para(child, path, skip=skip)
            paragraph.reindex()
            out.append(paragraph)
        elif child.tag == _TBL:
            for row in _rows(child):
                for cell in _cells(row):
                    _collect(cell, (*path, (child, row, cell)), out, skip)
        elif child.tag == _SDT:
            content = child.find(_SDTCONTENT)
            if content is not None:
                _collect(content, path, out, skip)
        elif child.tag == _CUSTOMXML:
            _collect(child, path, out, skip)


# --------------------------------------------------------------------------
# Writing revisions
# --------------------------------------------------------------------------


class _Revisions:
    """Mints revision markers with ids no other annotation in the document uses."""

    def __init__(self, roots: tuple[etree._Element, ...], author: str, date: datetime) -> None:
        # Every root whose annotations can end up in the output: the revised
        # document, and the base too when its pending revisions are copied over.
        highest = 0
        for root in roots:
            for element in root.iter(etree.Element):
                value = element.get(_ID)
                if value is not None and value.lstrip("-").isdigit():
                    highest = max(highest, int(value))
        self._ids = itertools.count(highest + 1)
        self._author = author
        self._date = date.strftime("%Y-%m-%dT%H:%M:%SZ")
        #: The ids of the revisions this comparison wrote, so a pending revision the
        #: document already had can be told from one of ours by more than its author.
        self.ours: set[str] = set()

    def next_id(self) -> str:
        return str(next(self._ids))

    def mark(self, kind: str) -> etree._Element:
        element = _el(kind)
        element.set(_ID, self.next_id())
        element.set(_w("author"), self._author)
        element.set(_w("date"), self._date)
        self.ours.add(element.get(_ID, ""))
        return element

    def clone(self, pending: etree._Element) -> etree._Element:
        """A copy of someone else's revision marker, with a fresh id and no content.

        For splitting a pending insertion around an edit made inside it: the two
        halves are still that author's insertion, and an id may not repeat.
        """
        element = etree.Element(pending.tag)
        for key, value in pending.attrib.items():
            element.set(key, value)
        element.set(_ID, self.next_id())
        return element


def _mark_paragraph_mark(paragraph: etree._Element, kind: str, revisions: _Revisions) -> None:
    """Record the paragraph mark itself as inserted or deleted.

    The marker goes first in the mark's run properties, which come after every
    other paragraph property except a section break, behind any marker that
    precedes it in the schema. A mark someone already has pending as deleted is
    not deleted again: under ``keep_pending`` their deletion stays, and covers it.
    """
    props = paragraph.find(_PPR)
    if props is None:
        props = _el("pPr")
        paragraph.insert(0, props)
    mark = props.find(_RPR)
    if mark is None:
        mark = _el("rPr")
        section = props.find(_SECTPR)
        if section is not None:
            section.addprevious(mark)
        else:
            props.append(mark)
    tag = _w(kind)
    if mark.find(tag) is not None:
        return
    earlier = _MARK_ORDER[: _MARK_ORDER.index(tag)]
    index = 0
    while index < len(mark) and mark[index].tag in earlier:
        index += 1
    mark.insert(index, revisions.mark(kind))


def _mark_row(row: etree._Element, kind: str, revisions: _Revisions) -> None:
    props = row.find(_TRPR)
    if props is None:
        props = _el("trPr")
        row.insert(1 if row.find(_TBLPREX) is not None else 0, props)
    props.append(revisions.mark(kind))


def _wrap(runs: list[etree._Element], make: Callable[[], etree._Element]) -> None:
    """Wrap runs in what ``make`` returns, one wrapper per stretch of adjacent siblings."""
    wrapper: etree._Element | None = None
    for run in runs:
        if wrapper is None or run.getprevious() is not wrapper:
            wrapper = make()
            run.addprevious(wrapper)
        wrapper.append(run)


def _wrap_inserted(runs: list[etree._Element], revisions: _Revisions) -> None:
    """Wrap runs in ``w:ins``, one wrapper per stretch of adjacent siblings."""
    _wrap(runs, functools.partial(revisions.mark, "ins"))


def _split_run(run: etree._Element, offset: int) -> None:
    """Keep the first ``offset`` characters in ``run``; move the rest to a new run after it."""
    tail_run = etree.Element(_R)
    props = run.find(_RPR)
    if props is not None:
        tail_run.append(copy.deepcopy(props))
    position = 0
    moving = False
    for child in list(run):
        if child.tag == _RPR:
            continue
        if moving:
            tail_run.append(child)
            continue
        text = _atom(child)
        length = len(text) if text else 0
        if position + length <= offset:
            position += length
            continue
        cut = offset - position
        if child.tag == _T and text and cut > 0:
            tail_run.append(_el("t", text[cut:]))
            child.text = text[:cut]
            child.set(_XML_SPACE, "preserve")
        else:
            tail_run.append(child)
        moving = True
    run.addnext(tail_run)


def _split_at(paragraph: _Para, cuts: set[int]) -> None:
    """Split runs so every cut falls on a run boundary."""
    for cut in sorted(cuts):
        for span in paragraph.spans:
            if span.start < cut < span.end:
                _split_run(span.run, cut - span.start)
                paragraph.reindex()
                break


def _deleted_runs(source: _Para, start: int, end: int) -> list[etree._Element]:
    """Runs carrying ``source.text[start:end]`` as deleted text, in the source's formatting."""
    return [run for _span, run in _deleted_span_runs(source, start, end)]


def _deleted_span_runs(source: _Para, start: int, end: int) -> list[tuple[_Span, etree._Element]]:
    """:func:`_deleted_runs`, each run paired with the source span it was copied from."""
    runs: list[tuple[_Span, etree._Element]] = []
    for span in source.spans:
        low, high = max(span.start, start), min(span.end, end)
        if low >= high:
            continue
        run = etree.Element(_R)
        props = span.run.find(_RPR)
        if props is not None:
            run.append(copy.deepcopy(props))
        position = span.start
        for child in span.run:
            text = _atom(child)
            if not text:
                continue
            child_start, child_end = position, position + len(text)
            position = child_end
            a, b = max(child_start, low), min(child_end, high)
            if a >= b:
                continue
            if child.tag == _T:
                run.append(_el("delText", text[a - child_start : b - child_start]))
            else:
                run.append(copy.deepcopy(child))
        runs.append((span, run))
    return runs


def _outer(run: etree._Element) -> etree._Element:
    """The run, or the ``w:ins`` this module wrapped it in when it is that wrapper's first run."""
    parent = run.getparent()
    if parent is not None and parent.tag == _INS and parent.index(run) == 0:
        return parent
    return run


def _place_inline(paragraph: _Para, spans: list[_Span], position: int, element: etree._Element) -> None:
    """Put a deletion at a character position of a paragraph, between runs."""
    for span in spans:
        if span.start == position and span.end > span.start:
            _outer(span.run).addprevious(element)
            return
    for span in reversed(spans):
        if span.end == position and span.end > span.start:
            _outer(span.run).addnext(element)
            return
    if spans and position == 0:
        _outer(spans[0].run).addprevious(element)
    elif spans:
        _outer(spans[-1].run).addnext(element)
    else:
        paragraph.element.append(element)


# --------------------------------------------------------------------------
# Keeping pending revisions (keep_pending)
# --------------------------------------------------------------------------

_PENDING_INSERTS = (_INS, _MOVETO)
_PENDING_DELETES = frozenset({_DEL, _MOVEFROM})


def _pending_insert(node: etree._Element, ours: set[str]) -> etree._Element | None:
    """The outermost insertion around ``node`` that this comparison did not write, if any."""
    found: etree._Element | None = None
    parent = node.getparent()
    while parent is not None and parent.tag != _P:
        if parent.tag in _PENDING_INSERTS and parent.get(_ID) not in ours:
            found = parent
        parent = parent.getparent()
    return found


def _lift_out_of_pending(run: etree._Element, revisions: _Revisions) -> None:
    """Move a run out of someone else's pending insertion, splitting it around the run.

    An insertion cannot hold another, so our insertion of text inside theirs is
    written the way Word writes it: their insertion up to the edit, ours, then the
    rest of theirs under a fresh id. Only a run directly inside the insertion can be
    moved; one inside a hyperlink or field within it would have to take that with it.
    """
    parent = run.getparent()
    host = _pending_insert(run, revisions.ours)
    if host is None:
        return
    if parent is not host:
        raise InPlaceRedlineError("an edit inside a pending insertion's hyperlink or field cannot be marked")
    index = host.index(run)
    after = list(host)[index + 1 :]
    if index == 0:
        host.addprevious(run)
    else:
        host.addnext(run)
        if after:
            rest = revisions.clone(host)
            rest.extend(after)
            run.addnext(rest)
    if len(host) == 0:
        _remove(host)


def _anchor(spans: list[_Span], position: int) -> tuple[etree._Element | None, str]:
    """The run a change at ``position`` is placed beside, and on which side."""
    for span in spans:
        if span.start == position and span.end > span.start:
            return span.run, "before"
    for span in reversed(spans):
        if span.end == position and span.end > span.start:
            return span.run, "after"
    if spans and position == 0:
        return spans[0].run, "before"
    if spans:
        return spans[-1].run, "after"
    return None, "after"


def _same_author(a: etree._Element, b: etree._Element) -> bool:
    return a.get(_w("author")) == b.get(_w("author"))


@dataclass
class _DeletedGroup:
    """Deleted runs copied from one stretch of a base paragraph."""

    #: The pending insertion the stretch sat in, or None for base text.
    enclosing: etree._Element | None
    runs: list[etree._Element]
    #: Where the stretch starts in the base paragraph's document order.
    order: int


def _wrapped(group: _DeletedGroup, revisions: _Revisions) -> etree._Element:
    """A group as our deletion, inside an insertion by its author if it was theirs."""
    deletion = revisions.mark("del")
    deletion.extend(group.runs)
    if group.enclosing is None:
        return deletion
    theirs = revisions.clone(group.enclosing)
    theirs.append(deletion)
    return theirs


def _place_deletion_keeping(
    paragraph: _Para,
    spans: list[_Span],
    position: int,
    groups: list[_DeletedGroup],
    anchors: list[tuple[int, etree._Element]],
    revisions: _Revisions,
) -> None:
    """Put deleted text back where it was, inside or outside a pending insertion as it was.

    Text that was someone's pending insertion is deleted inside an insertion by that
    author, which is how Word records "inserted by them, deleted by us": rejecting
    their insertion then removes our deletion with it, rather than restoring text the
    base document never had.

    ``anchors`` are the base's pending deletions within or at the edge of the deleted
    text, each as its base-order position and its element in the revised paragraph.
    Like the deleted text they show nothing, so the visible text cannot say which
    goes first; the base document can, so the groups are placed around them in its
    order. Without that, rejecting everything reads their deleted words in the wrong
    place. With no anchor the groups go beside the run at ``position``.
    """
    if anchors:
        items: list[tuple[int, _DeletedGroup | etree._Element]] = [(group.order, group) for group in groups]
        items.extend(anchors)
        items.sort(key=lambda item: item[0])
        first = next(element for _, element in items if isinstance(element, etree._Element))
        previous: etree._Element | None = None
        for _, item in items:
            if isinstance(item, etree._Element):
                previous = item
                continue
            element = _wrapped(item, revisions)
            if previous is None:
                first.addprevious(element)
            else:
                previous.addnext(element)
                previous = element
        return

    run, side = _anchor(spans, position)
    # Placed in reverse when going after the anchor, so the groups keep their order.
    for group in reversed(groups) if side == "after" else groups:
        if run is None:
            paragraph.element.append(_wrapped(group, revisions))
            continue
        host = _pending_insert(run, revisions.ours)
        if group.enclosing is not None and host is not None and _same_author(host, group.enclosing):
            node = _child_of(run, host)
            element = revisions.mark("del")
            element.extend(group.runs)
        else:
            node = host if host is not None else _outer(run)
            element = _wrapped(group, revisions)
        if node is None:
            raise InPlaceRedlineError("could not find where a deletion inside a pending insertion belongs")
        if side == "before":
            node.addprevious(element)
        else:
            node.addnext(element)


def _outermost(run: etree._Element, paragraph: etree._Element) -> etree._Element:
    """The paragraph's child that holds ``run``."""
    node = run
    while (parent := node.getparent()) is not None and parent is not paragraph:
        node = parent
    return node


def _pending_deletions(paragraph: etree._Element) -> list[etree._Element]:
    """A paragraph's pending deletions, outside any other pending revision."""
    return [
        element
        for element in paragraph.iter(*_PENDING_DELETES)
        if not any(
            a.tag in _PENDING_DELETES or a.tag in _PENDING_INSERTS
            for a in element.iterancestors()
            if paragraph in a.iterancestors()
        )
    ]


def _drop_copies(deletions: Iterable[etree._Element]) -> None:
    """Drop pending deletions of the revised document that answer for none of its paired paragraph's.

    Every pending revision in the revised document is one of the base's, since what
    ``keep_pending`` is for writes none of its own, and each of the base's deletions is
    placed from the base side: kept or put back in the paragraph paired with its own
    (:func:`_restore_pending`), or copied into a deleted one
    (:func:`_deleted_content_keeping`). So one the revised document holds anywhere else
    is a second copy, and rejecting everything brought its text back twice (seen when
    an editor emptied a paragraph of everything but a pending deletion: it paired with
    nothing, and the deletion stood both in the deleted copy and in the inserted
    paragraph).
    """
    for deletion in list(deletions):
        _remove_keeping_anchors(deletion)


def _deletion_offsets(paragraph: _Para) -> dict[etree._Element, int]:
    """Each of a paragraph's pending deletions, and the visible character it sits before."""
    order = {element: index for index, element in enumerate(paragraph.element.iter())}
    return {
        deletion: next((s.start for s in paragraph.spans if order[s.run] > order[deletion]), len(paragraph.text))
        for deletion in _pending_deletions(paragraph.element)
    }


def _restore_pending(
    new: _Para, old: _Para, segments: list[DiffSegment], revisions: _Revisions
) -> dict[etree._Element, etree._Element]:
    """Put back the base paragraph's pending revisions that the revised one lost.

    Under ``keep_pending`` the base's pending revisions stand, whatever the revised
    paragraph says about them, because the tool that produced it does not accept or
    reject them. But such a tool may write a segment it edits as plain text, so an
    edit inside someone's pending insertion drops the insertion around the rest of
    their text, and an edit to text they had pending as deleted brings it back. Text that was in a pending
    insertion and is unchanged is therefore wrapped in that insertion again, and a
    pending deletion missing from the revised paragraph is copied back to its place.

    Returns each of the base's pending deletions with the element that now stands
    for it in the revised paragraph, kept or restored, for placing our deletions
    around them.
    """
    if not segments:
        if old.text != new.text:
            return {}
        segments = [DiffSegment("equal", old.text)]

    equal: list[tuple[int, int, int]] = []
    old_position = new_position = 0
    for segment in segments:
        length = len(segment.text)
        if segment.kind == "equal":
            equal.append((old_position, old_position + length, new_position))
        if segment.kind != "insert":
            old_position += length
        if segment.kind != "delete":
            new_position += length

    ranges: list[tuple[int, int, etree._Element]] = []
    for span in old.spans:
        enclosing = _pending_insert(span.run, set())
        if enclosing is None:
            continue
        for start, end, target in equal:
            low, high = max(span.start, start), min(span.end, end)
            if low >= high:
                continue
            low, high = target + low - start, target + high - start
            if ranges and ranges[-1][2] is enclosing and ranges[-1][1] == low:
                ranges[-1] = (ranges[-1][0], high, enclosing)
            else:
                ranges.append((low, high, enclosing))

    def deleted_text(element: etree._Element) -> str:
        return "".join(node.text or "" for node in element.iter(_T, _DELTEXT))

    # The revised paragraph's own pending deletions answer for the base's in order, by
    # their text; a base deletion with none left to answer for it was lost.
    kept: dict[str, list[etree._Element]] = {}
    for element in _pending_deletions(new.element):
        kept.setdefault(deleted_text(element), []).append(element)
    offsets = _deletion_offsets(old)
    standing: dict[etree._Element, etree._Element] = {}
    lost: list[etree._Element] = []
    for deletion in offsets:
        candidates = kept.get(deleted_text(deletion))
        if candidates:
            standing[deletion] = candidates.pop(0)
        else:
            lost.append(deletion)
    _drop_copies(element for left in kept.values() for element in left)

    # Split at every place something goes back, not only at the insertions' edges:
    # placed by a position that falls inside a run, a deletion lands at the end of
    # the paragraph instead.
    positions = {deletion: _map_offset(segments, offsets[deletion]) for deletion in lost}
    cuts = {low for low, _, _ in ranges} | {high for _, high, _ in ranges} | set(positions.values())
    if cuts:
        _split_at(new, cuts)
    for low, high, enclosing in ranges:
        unmarked = [
            s.run
            for s in new.spans
            if low <= s.start < s.end <= high and _pending_insert(s.run, revisions.ours) is None
        ]
        _wrap(unmarked, functools.partial(revisions.clone, enclosing))

    order = {element: index for index, element in enumerate(old.element.iter())}
    for deletion in lost:
        restored = revisions.clone(deletion)
        restored.extend(copy.deepcopy(child) for child in deletion)
        # Beside a deletion already standing at the same place, in the base's order.
        beside = [d for d in standing if offsets[d] == offsets[deletion]]
        later = [d for d in beside if order[d] > order[deletion]]
        earlier = [d for d in beside if order[d] < order[deletion]]
        if later:
            standing[min(later, key=order.__getitem__)].addprevious(restored)
        elif earlier:
            standing[max(earlier, key=order.__getitem__)].addnext(restored)
        else:
            run, side = _anchor(new.spans, positions[deletion])
            if run is None:
                new.element.append(restored)
            elif side == "before":
                _outermost(run, new.element).addprevious(restored)
            else:
                _outermost(run, new.element).addnext(restored)
        standing[deletion] = restored
    return standing


def _map_offset(segments: list[DiffSegment], position: int) -> int:
    """Where a character position of the base paragraph falls in the revised one.

    A position inside deleted text maps to where the deletion is shown, and one at
    the end of an equal stretch stays before an insertion that follows it.
    """
    old = new = 0
    for segment in segments:
        length = len(segment.text)
        if segment.kind == "insert":
            new += length
            continue
        if position <= old + length:
            return new + (position - old if segment.kind == "equal" else 0)
        old += length
        if segment.kind == "equal":
            new += length
    return new


def _deleted_groups(source: _Para, start: int, end: int, order: dict[etree._Element, int]) -> list[_DeletedGroup]:
    """:func:`_deleted_runs`, grouped by the pending insertion each run sat in, if any."""
    groups: list[_DeletedGroup] = []
    for span, run in _deleted_span_runs(source, start, end):
        enclosing = _pending_insert(span.run, set())
        if groups and groups[-1].enclosing is enclosing:
            groups[-1].runs.append(run)
        else:
            groups.append(_DeletedGroup(enclosing, [run], order[span.run]))
    return groups


def _deleted_copy(run: etree._Element) -> etree._Element | None:
    """A whole run as deleted text, in its own formatting; None when it carries no text."""
    out = etree.Element(_R)
    props = run.find(_RPR)
    if props is not None:
        out.append(copy.deepcopy(props))
    carries_text = False
    for child in run:
        text = _atom(child)
        if text:
            out.append(_el("delText", text) if child.tag == _T else copy.deepcopy(child))
            carries_text = True
    return out if carries_text else None


def _deleted_content_keeping(element: etree._Element, revisions: _Revisions) -> list[etree._Element]:
    """A base paragraph's content as deleted text, with its pending revisions kept.

    For a paragraph only the base has, under ``keep_pending``. Base text becomes
    our deletion, text someone had pending as an insertion becomes our deletion
    inside an insertion of theirs (as in :func:`_place_deletion_keeping`), and
    their pending deletions are copied as they were, so rejecting everything still
    gives the base document's own base.
    """
    out: list[etree._Element] = []
    current: tuple[etree._Element | None, etree._Element] | None = None

    def place(element: etree._Element, enclosing: etree._Element | None) -> None:
        if enclosing is None:
            out.append(element)
        else:
            theirs = revisions.clone(enclosing)
            theirs.append(element)
            out.append(theirs)

    def walk(container: etree._Element, enclosing: etree._Element | None) -> None:
        nonlocal current
        for child in container:
            if child.tag == _R:
                run = _deleted_copy(child)
                if run is None:
                    continue
                if current is None or current[0] is not enclosing:
                    current = (enclosing, revisions.mark("del"))
                    place(current[1], enclosing)
                current[1].append(run)
            elif child.tag in _PENDING_DELETES:
                current = None
                kept = revisions.clone(child)
                kept.extend(copy.deepcopy(grandchild) for grandchild in child)
                place(kept, enclosing)
            elif child.tag in _PENDING_INSERTS:
                walk(child, child if enclosing is None else enclosing)
            elif child.tag in _RUN_CONTAINERS:
                walk(child, enclosing)

    walk(element, None)
    return out


def _renumber_repeated(root: etree._Element, revisions: _Revisions) -> None:
    """Give a fresh id to any revision marker whose id an earlier one already has.

    ``keep_pending`` copies the base document's pending revisions into the revised
    one, and the two documents numbered theirs independently. Ours are never
    renumbered: they were minted above every id in both.
    """
    seen: set[str] = set()
    for element in root.iter(*_REVISION_MARKERS):
        value = element.get(_ID)
        if value is None:
            continue
        if value in seen:
            value = revisions.next_id()
            element.set(_ID, value)
        seen.add(value)


def _policy_view(root: etree._Element, ours: set[str], ours_accepted: bool, others_accepted: bool) -> list[str]:
    """Paragraph texts with this comparison's revisions and everyone else's resolved separately.

    The read-back for ``keep_pending``, where "reject everything" is no longer the
    base document: the base had pending revisions of its own. A paragraph whose mark
    is resolved away joins the next one, as accepting a deleted mark or rejecting an
    inserted one does in Word.
    """

    def gone(marker: etree._Element) -> bool:
        accepted = ours_accepted if marker.get(_ID) in ours else others_accepted
        if marker.tag in _PENDING_INSERTS:
            return not accepted
        if marker.tag in _PENDING_DELETES:
            return accepted
        return False

    def text_of(container: etree._Element) -> str:
        parts: list[str] = []
        for child in container:
            if child.tag in _PENDING_INSERTS or child.tag in _PENDING_DELETES:
                if not gone(child):
                    parts.append(text_of(child))
            elif child.tag == _R:
                for grandchild in child:
                    if grandchild.tag == _DELTEXT:
                        parts.append(grandchild.text or "")
                    elif (text := _atom(grandchild)) is not None:
                        parts.append(text)
            elif child.tag in _RUN_CONTAINERS:
                parts.append(text_of(child))
        return "".join(parts)

    def paragraphs(container: etree._Element) -> Iterator[etree._Element]:
        for child in container:
            if child.tag == _P:
                yield child
            elif child.tag == _TBL:
                for row in _rows(child):
                    props = row.find(_TRPR)
                    if props is not None and any(gone(marker) for marker in props):
                        continue
                    for cell in _cells(row):
                        yield from paragraphs(cell)
            elif child.tag == _SDT and (content := child.find(_SDTCONTENT)) is not None:
                yield from paragraphs(content)
            elif child.tag == _CUSTOMXML:
                yield from paragraphs(child)

    texts: list[str] = []
    carried = ""
    for paragraph in paragraphs(_body(root)):
        text = carried + text_of(paragraph)
        mark = paragraph.find(f"{_PPR}/{_RPR}")
        if mark is not None and any(gone(marker) for marker in mark):
            carried = text
            continue
        carried = ""
        if text.strip():
            texts.append(text)
    if carried.strip():
        texts.append(carried)
    return texts


def _verify_keeping(root: etree._Element, ours: set[str], old_root: etree._Element, revised: etree._Element) -> None:
    """The read-back for ``keep_pending``: three views, each against the document it must equal."""
    checks = (
        ("accepting every change does not give the revised document", (True, True), revised, True),
        ("rejecting only this comparison's changes does not give the base document", (False, True), old_root, True),
        ("rejecting every change does not give the base document's own base", (False, False), old_root, False),
    )
    for message, (ours_accepted, others_accepted), expected_root, expected_accepted in checks:
        got = _policy_view(root, ours, ours_accepted, others_accepted)
        want = _policy_view(expected_root, set(), expected_accepted, expected_accepted)
        if got != want:
            raise InPlaceRedlineError(message)


def _apply_edit(
    new: _Para,
    old: _Para,
    segments: list[DiffSegment],
    revisions: _Revisions,
    *,
    keep_pending: bool = False,
    standing: dict[etree._Element, etree._Element] | None = None,
) -> None:
    """Write one paired paragraph's word diff into the revised paragraph.

    ``standing`` is what :func:`_restore_pending` returned for the pair under
    ``keep_pending``: the base's pending deletions and their elements here.
    """
    inserts: list[tuple[int, int]] = []
    deletes: list[tuple[int, int, int]] = []
    new_position = old_position = 0
    for segment in segments:
        length = len(segment.text)
        if segment.kind == "equal":
            new_position += length
            old_position += length
        elif segment.kind == "insert":
            inserts.append((new_position, new_position + length))
            new_position += length
        else:
            deletes.append((new_position, old_position, old_position + length))
            old_position += length
    if new_position != len(new.text) or old_position != len(old.text):
        raise InPlaceRedlineError("word diff does not cover its paragraph")

    _split_at(new, {a for a, _ in inserts} | {b for _, b in inserts} | {p for p, _, _ in deletes})
    spans = list(new.spans)

    def inserted(span: _Span) -> bool:
        if span.end > span.start:
            return any(a <= span.start and span.end <= b for a, b in inserts)
        return any(a < span.start < b for a, b in inserts)

    inserted_runs = [s.run for s in spans if inserted(s)]
    if keep_pending:
        for run in inserted_runs:
            _lift_out_of_pending(run, revisions)
    _wrap_inserted(inserted_runs, revisions)
    if keep_pending:
        order = {element: index for index, element in enumerate(old.element.iter())}
        offsets = _deletion_offsets(old)
    kept = standing or {}
    for position, start, end in deletes:
        if keep_pending:
            groups = _deleted_groups(old, start, end, order)
            anchors = [
                (order[deletion], kept[deletion])
                for deletion, offset in offsets.items()
                if start <= offset <= end and deletion in kept and new.element in kept[deletion].iterancestors()
            ]
            if groups:
                _place_deletion_keeping(new, spans, position, groups, anchors, revisions)
            continue
        runs = _deleted_runs(old, start, end)
        if not runs:
            continue
        wrapper = revisions.mark("del")
        for run in runs:
            wrapper.append(run)
        _place_inline(new, spans, position, wrapper)


# --------------------------------------------------------------------------
# Content only the base document has
# --------------------------------------------------------------------------


@dataclass
class _Context:
    revisions: _Revisions
    old_by_element: dict[etree._Element, _Para]
    deleted: set[etree._Element]
    # Base-document element -> where it is in the output: its counterpart in
    # the revised document, or the deleted copy written for it.
    counterpart: dict[etree._Element, etree._Element]
    placed: set[etree._Element] = field(default_factory=set)
    keep_pending: bool = False


def _deleted_paragraph(source: _Para | None, props: etree._Element | None, ctx: _Context) -> etree._Element:
    """A paragraph whose text and mark are both deleted.

    ``source`` is ``None`` for an empty one: the cell of a deleted row whose
    text turns up elsewhere in the revised document, or a separator.
    """
    paragraph = etree.Element(_P)
    paragraph_props = copy.deepcopy(props) if props is not None else _el("pPr")
    for section in paragraph_props.findall(_SECTPR):
        paragraph_props.remove(section)
    # Some generators write numId before ilvl, which the schema forbids. Word
    # forgives it in the base document, but the copy should not inherit it.
    for numbering in paragraph_props.findall(_w("numPr")):
        numbering[:] = sorted(numbering, key=lambda c: _NUMPR_ORDER.get(c.tag, len(_NUMPR_ORDER)))
    paragraph.append(paragraph_props)
    if source is not None and ctx.keep_pending:
        paragraph.extend(_deleted_content_keeping(source.element, ctx.revisions))
    elif source is not None and source.text:
        wrapper = ctx.revisions.mark("del")
        for run in _deleted_runs(source, 0, len(source.text)):
            wrapper.append(run)
        paragraph.append(wrapper)
    _mark_paragraph_mark(paragraph, "del", ctx.revisions)
    return paragraph


def _deleted_blocks(container: etree._Element, out: etree._Element, ctx: _Context) -> None:
    for child in container:
        if child.tag == _P:
            source = ctx.old_by_element.get(child) if child in ctx.deleted else None
            out.append(_deleted_paragraph(source, child.find(_PPR), ctx))
            ctx.placed.add(child)
        elif child.tag == _TBL:
            out.append(_deleted_table(child, ctx))
        elif child.tag == _SDT and (content := child.find(_SDTCONTENT)) is not None:
            _deleted_blocks(content, out, ctx)
        elif child.tag == _CUSTOMXML:
            _deleted_blocks(child, out, ctx)


def _deleted_cell(cell: etree._Element, ctx: _Context) -> etree._Element:
    out = etree.Element(_TC)
    props = cell.find(_TCPR)
    if props is not None:
        out.append(copy.deepcopy(props))
    _deleted_blocks(cell, out, ctx)
    if len(out) == 0 or out[-1].tag != _P:
        out.append(_deleted_paragraph(None, None, ctx))  # a cell must end with a paragraph
    return out


def _deleted_row(row: etree._Element, ctx: _Context) -> etree._Element:
    out = etree.Element(_TR)
    exceptions = row.find(_TBLPREX)
    if exceptions is not None:
        out.append(copy.deepcopy(exceptions))
    props = row.find(_TRPR)
    out.append(copy.deepcopy(props) if props is not None else _el("trPr"))
    for cell in _cells(row):
        out.append(_deleted_cell(cell, ctx))
    _mark_row(out, "del", ctx.revisions)
    return out


def _deleted_table(table: etree._Element, ctx: _Context) -> etree._Element:
    out = etree.Element(_TBL)
    for child in table:
        if child.tag in (_TBLPR, _TBLGRID):
            out.append(copy.deepcopy(child))
    for row in _rows(table):
        out.append(_deleted_row(row, ctx))
    return out


def _counterpart_in(old_block: etree._Element, new_container: etree._Element, ctx: _Context) -> etree._Element | None:
    """Where a base-document block (or the last of its contents) sits in ``new_container``."""
    target = ctx.counterpart.get(old_block)
    if target is not None:
        return _child_of(target, new_container)
    if old_block.tag in (_SDT, _CUSTOMXML):
        content = old_block.find(_SDTCONTENT) if old_block.tag == _SDT else old_block
        for child in reversed(list(content) if content is not None else []):
            found = _counterpart_in(child, new_container, ctx)
            if found is not None:
                return found
    return None


def _insert_block(
    copy_: etree._Element,
    old_element: etree._Element,
    old_container: etree._Element,
    new_container: etree._Element,
    ctx: _Context,
) -> None:
    """Insert a deleted block after the output position of its nearest earlier sibling."""
    top = _child_of(old_element, old_container)
    previous = top.getprevious() if top is not None else None
    while previous is not None:
        target = _counterpart_in(previous, new_container, ctx)
        if target is not None:
            target.addnext(copy_)
            break
        previous = previous.getprevious()
    else:
        first = 1 if len(new_container) and new_container[0].tag == _TCPR else 0
        new_container.insert(first, copy_)
    if copy_.tag == _TBL:
        # Word joins adjacent tables, and a cell must end with a paragraph.
        preceding = copy_.getprevious()
        if preceding is not None and preceding.tag == _TBL:
            copy_.addprevious(_deleted_paragraph(None, None, ctx))
        following = copy_.getnext()
        if (following is not None and following.tag == _TBL) or (following is None and new_container.tag == _TC):
            copy_.addnext(_deleted_paragraph(None, None, ctx))


def _insert_among(
    copy_: etree._Element,
    old_item: etree._Element,
    old_items: list[etree._Element],
    new_parent: etree._Element,
    new_items: list[etree._Element],
    ctx: _Context,
) -> None:
    """Insert a deleted row or cell next to the output position of its nearest neighbour."""
    at = old_items.index(old_item)
    for earlier in reversed(old_items[:at]):
        target = ctx.counterpart.get(earlier)
        if target is not None and _child_of(target, new_parent) is not None:
            target.addnext(copy_)
            return
    for later in old_items[at + 1 :]:
        target = ctx.counterpart.get(later)
        if target is not None and _child_of(target, new_parent) is not None:
            target.addprevious(copy_)
            return
    if new_items:
        new_items[-1].addnext(copy_)
    else:
        new_parent.append(copy_)


def _mark_cell(cell: etree._Element, kind: str, revisions: _Revisions) -> None:
    """Record a whole cell as inserted (``cellIns``) or deleted (``cellDel``), as Word's Compare does."""
    props = cell.find(_TCPR)
    if props is None:
        props = _el("tcPr")
        cell.insert(0, props)
    props.append(revisions.mark(kind))


def _grid_width(element: etree._Element | None) -> int:
    value = element.get(_w("val")) if element is not None else None
    return int(value) if value and value.isdigit() else (0 if element is None else 1)


def _widen_grid(table: etree._Element, row: etree._Element, source: etree._Element) -> None:
    """Give the table's grid a column for a deleted cell that made a row wider than it."""
    grid = table.find(_TBLGRID)
    if grid is None:
        return
    span = sum(max(1, _grid_width(cell.find(f"{_TCPR}/{_w('gridSpan')}"))) for cell in _cells(row))
    props = row.find(_TRPR)
    if props is not None:
        span += _grid_width(props.find(_w("gridBefore"))) + _grid_width(props.find(_w("gridAfter")))
    width = source.find(f"{_TCPR}/{_w('tcW')}")
    twips = width.get(_w("w")) if width is not None and width.get(_w("type"), "dxa") == "dxa" else None
    while len(grid.findall(_w("gridCol"))) < span:
        column = _el("gridCol")
        column.set(_w("w"), twips if twips and twips.isdigit() else "1000")
        grid.append(column)


def _place_deleted(paragraph: _Para, ctx: _Context, old_body: etree._Element, new_body: etree._Element) -> None:
    """Put a paragraph only the base document has into the output.

    Walks its table nesting from the outside in. The first level with no
    counterpart in the revised document is copied whole, as a deleted table or
    row; if every level has one, the paragraph goes into the corresponding cell.
    """
    old_container, new_container = old_body, new_body
    for table, row, cell in paragraph.path:
        new_table = ctx.counterpart.get(table)
        if new_table is None:
            copy_ = _deleted_table(table, ctx)
            _insert_block(copy_, table, old_container, new_container, ctx)
            ctx.counterpart[table] = copy_
            return
        new_row = ctx.counterpart.get(row)
        if new_row is None or _child_of(new_row, new_table) is None:
            copy_ = _deleted_row(row, ctx)
            _insert_among(copy_, row, list(_rows(table)), new_table, list(_rows(new_table)), ctx)
            ctx.counterpart[row] = copy_
            return
        new_cell = ctx.counterpart.get(cell)
        if new_cell is None or _child_of(new_cell, new_row) is None:
            # A column the revised table no longer has. Word's Compare keeps such
            # a cell in place with its content deleted, and so does this.
            copy_ = _deleted_cell(cell, ctx)
            _mark_cell(copy_, "cellDel", ctx.revisions)
            _insert_among(copy_, cell, list(_cells(row)), new_row, list(_cells(new_row)), ctx)
            _widen_grid(new_table, new_row, cell)
            ctx.counterpart[cell] = copy_
            return
        old_container, new_container = cell, new_cell
    copy_ = _deleted_paragraph(paragraph, paragraph.element.find(_PPR), ctx)
    ctx.placed.add(paragraph.element)
    _insert_block(copy_, paragraph.element, old_container, new_container, ctx)
    ctx.counterpart[paragraph.element] = copy_


def _correspondence(
    steps: list[ParagraphAlignment], old: list[_Para], new: list[_Para], *, weighted: bool = True
) -> dict[etree._Element, etree._Element]:
    """Which revised-document element each base-document element became.

    Paragraphs come straight from the alignment. Tables, rows and cells are
    voted for by the paired paragraphs inside them, one counterpart each, most
    votes first. A pair votes with the length of its text, plus one: a clause
    found in both is far better evidence of which row is which than a "2.50%"
    or an empty cell, which recur from row to row, and counting every pair
    alike once matched a row to its neighbour on the strength of its blank
    cells. Within a pair of rows, a cell nothing voted for then takes the next
    unclaimed cell in reading order: a cell whose text was rewritten entirely
    is still the same cell, and only a cell with no partner left is a removed
    column.
    """
    result: dict[etree._Element, etree._Element] = {}
    votes: Counter[tuple[etree._Element, etree._Element]] = Counter()
    for step in steps:
        if step.old_index is None or step.new_index is None:
            continue
        old_para, new_para = old[step.old_index], new[step.new_index]
        result[old_para.element] = new_para.element
        for old_level, new_level in zip(old_para.path, new_para.path, strict=False):
            for pair in zip(old_level, new_level, strict=True):
                votes[pair] += 1 + len(old_para.text.strip()) if weighted else 1
    taken: set[etree._Element] = set()
    for (old_element, new_element), _count in votes.most_common():
        if old_element in result or new_element in taken or old_element.tag != new_element.tag:
            continue
        result[old_element] = new_element
        taken.add(new_element)
    for old_row, new_row in list(result.items()):
        if old_row.tag != _TR:
            continue
        new_cells = list(_cells(new_row))
        cursor = 0
        for old_cell in _cells(old_row):
            mapped = result.get(old_cell)
            if mapped is not None:
                if mapped in new_cells:
                    cursor = new_cells.index(mapped) + 1
                continue
            if cursor < len(new_cells) and new_cells[cursor] not in taken:
                result[old_cell] = new_cells[cursor]
                taken.add(new_cells[cursor])
                cursor += 1
    return result


def _same_containers(old_para: _Para, new_para: _Para, counterpart: dict[etree._Element, etree._Element]) -> bool:
    """Whether a paired paragraph sits in the counterparts of its own table, row and cell."""
    if len(old_para.path) != len(new_para.path):
        return False
    return all(
        counterpart.get(old_element) is new_element
        for old_level, new_level in zip(old_para.path, new_para.path, strict=True)
        for old_element, new_element in zip(old_level, new_level, strict=True)
    )


def _respect_tables(
    steps: list[ParagraphAlignment], old: list[_Para], new: list[_Para]
) -> tuple[list[ParagraphAlignment], dict[etree._Element, etree._Element]]:
    """The alignment with every pair kept inside corresponding table rows and cells.

    The paragraph alignment reads the documents as one list of paragraphs and does
    not know about rows, so it can pair a cell with identical text in another row:
    one row's "2.50%" with the next row's. Once the rows themselves are matched
    (:func:`_correspondence`), such a pair sits in rows that do not correspond, and
    written as an edit it put text in the wrong row: rejecting every change then no
    longer gave the base document, and the whole redline was refused.

    So, until nothing changes:

    * a pair that does not sit in the counterparts of its own table, row and cell is
      split into a deletion and an insertion, and the rows are matched again;
    * paragraphs left without a partner in a pair of corresponding cells are aligned
      again, by :func:`align_paragraphs` itself on just those cells, so a value that
      stayed in its cell is not shown as deleted and reinserted, and one that changed
      gets its word diff. A pair split once is never formed again, which is what
      makes this end.

    Every word diff still comes from :mod:`kyralaw_redline.diff`.
    """
    refused: set[tuple[int, int]] = set()
    while True:
        counterpart = _correspondence(steps, old, new)
        kept: list[ParagraphAlignment] = []
        for step in steps:
            if (
                step.old_index is not None
                and step.new_index is not None
                and not _same_containers(old[step.old_index], new[step.new_index], counterpart)
            ):
                refused.add((step.old_index, step.new_index))
                kept.append(ParagraphAlignment(step.old_index, None, None))
                kept.append(ParagraphAlignment(None, step.new_index, None))
            else:
                kept.append(step)
        if len(kept) != len(steps):
            steps = kept
            continue
        repaired = _pair_within_cells(steps, old, new, counterpart, refused)
        if repaired is None:
            return steps, counterpart
        steps = repaired


def _pair_within_cells(
    steps: list[ParagraphAlignment],
    old: list[_Para],
    new: list[_Para],
    counterpart: dict[etree._Element, etree._Element],
    refused: set[tuple[int, int]],
) -> list[ParagraphAlignment] | None:
    """Pair the unpartnered paragraphs of corresponding cells; None when there is nothing to pair."""
    old_by_cell: dict[etree._Element, list[int]] = {}
    new_by_cell: dict[etree._Element, list[int]] = {}
    for step in steps:
        if step.old_index is not None and step.new_index is None and old[step.old_index].path:
            target = counterpart.get(old[step.old_index].path[-1][2])
            if target is not None:
                old_by_cell.setdefault(target, []).append(step.old_index)
        elif step.new_index is not None and step.old_index is None and new[step.new_index].path:
            new_by_cell.setdefault(new[step.new_index].path[-1][2], []).append(step.new_index)
    pairs: dict[int, ParagraphAlignment] = {}
    for cell, old_ids in old_by_cell.items():
        new_ids = new_by_cell.get(cell)
        if not new_ids:
            continue
        for sub in align_paragraphs([old[i].text for i in old_ids], [new[j].text for j in new_ids]):
            if sub.old_index is None or sub.new_index is None:
                continue
            i, j = old_ids[sub.old_index], new_ids[sub.new_index]
            if (i, j) not in refused:
                pairs[i] = ParagraphAlignment(i, j, sub.segments)
    if not pairs:
        return None
    partnered = {step.new_index for step in pairs.values()}
    out: list[ParagraphAlignment] = []
    for step in steps:
        if step.new_index is None and step.old_index in pairs:
            out.append(pairs[step.old_index])
        elif step.old_index is None and step.new_index in partnered:
            continue
        else:
            out.append(step)
    return out


# --------------------------------------------------------------------------
# Read-back check
# --------------------------------------------------------------------------


def _view_paragraphs(container: etree._Element, dropped_row: str) -> Iterator[etree._Element]:
    for child in container:
        if child.tag == _P:
            yield child
        elif child.tag == _TBL:
            for row in _rows(child):
                props = row.find(_TRPR)
                if props is not None and props.find(dropped_row) is not None:
                    continue
                for cell in _cells(row):
                    yield from _view_paragraphs(cell, dropped_row)
        elif child.tag == _SDT and (content := child.find(_SDTCONTENT)) is not None:
            yield from _view_paragraphs(content, dropped_row)
        elif child.tag == _CUSTOMXML:
            yield from _view_paragraphs(child, dropped_row)


def _view_text(container: etree._Element, dropped: tuple[str, str]) -> str:
    parts: list[str] = []
    for child in container:
        if child.tag in dropped:
            continue
        if child.tag == _R:
            for grandchild in child:
                if grandchild.tag == _DELTEXT:
                    parts.append(grandchild.text or "")
                elif (text := _atom(grandchild)) is not None:
                    parts.append(text)
        elif child.tag in _RUN_CONTAINERS:
            parts.append(_view_text(child, dropped))
    return "".join(parts)


def _view(root: etree._Element, mode: str) -> list[str]:
    """Paragraph texts with every revision accepted or rejected, read without changing the tree."""
    dropped = (_DEL, _MOVEFROM) if mode == "accept" else (_INS, _MOVETO)
    texts: list[str] = []
    for paragraph in _view_paragraphs(_body(root), dropped[0]):
        text = _view_text(paragraph, dropped)
        mark = paragraph.find(f"{_PPR}/{_RPR}")
        if mark is not None and mark.find(dropped[0]) is not None:
            if text.strip():
                raise InPlaceRedlineError(f"a paragraph whose mark is removed on {mode} still has text")
            continue
        if text.strip():
            texts.append(text)
    return texts


def _verify(root: etree._Element, old_texts: list[str], new_texts: list[str]) -> None:
    if _view(root, "accept") != [t for t in new_texts if t.strip()]:
        raise InPlaceRedlineError("accepting every change does not give the revised document")
    if _view(root, "reject") != [t for t in old_texts if t.strip()]:
        raise InPlaceRedlineError("rejecting every change does not give the base document")


# --------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------


#: How a renderer aligns two documents' paragraphs (:func:`align_paragraphs`).
Aligner = Callable[[list[str], list[str]], list[ParagraphAlignment]]


def redline_in_place(
    old_docx: bytes,
    new_docx: bytes,
    *,
    author: str = DEFAULT_AUTHOR,
    date: datetime | None = None,
    keep_pending: bool = False,
) -> bytes:
    """Redline ``old_docx`` against ``new_docx``, written into ``new_docx`` itself.

    Returns the revised document with the comparison marked as tracked changes.
    Raises :class:`InPlaceRedlineError` for an input that is not a readable
    .docx, or a result that fails the accept/reject read-back.

    ``keep_pending`` is for a revised document that was edited directly from the
    base, pending revisions and all, rather than for comparing two versions: what an
    editing tool produces when it keeps a document's tracked changes through a save
    but writes its own edits untracked. Neither side is accepted first. Both
    are read as a reader sees them with the pending revisions in place, insertions
    shown and deletions hidden, so the only differences are the new edits, and only
    those are marked. The pending revisions stay exactly as they were, and an edit
    inside one is written around it (see :func:`_lift_out_of_pending` and
    :func:`_place_deletion_keeping`). The read-back is per author: accepting
    everything must give the revised document, rejecting only this comparison's
    changes the base document, and rejecting everything the base document's own
    base.
    """
    return _render(old_docx, new_docx, author=author, date=date, keep_pending=keep_pending)


def _render(
    old_docx: bytes,
    new_docx: bytes,
    *,
    author: str = DEFAULT_AUTHOR,
    date: datetime | None = None,
    keep_pending: bool = False,
    align: Aligner = align_paragraphs,
    respect_tables: bool = True,
) -> bytes:
    """:func:`redline_in_place`, with what it aligns by and how it treats tables open.

    The defaults are today's renderer. ``eval/native.py`` freezes the first version
    as a baseline: its own alignment, ``respect_tables=False`` (no
    :func:`_respect_tables`, rows matched by unweighted votes).
    """
    _, _, old_root = _load(old_docx)
    package, main_name, new_root = _load(new_docx)
    skip: frozenset[str] = frozenset()
    revised = None
    if keep_pending:
        skip = _PENDING_DELETES
        revised = copy.deepcopy(new_root)
    else:
        _accept_all(old_root)
        _accept_all(new_root)
    old_body, new_body = _body(old_root), _body(new_root)

    old: list[_Para] = []
    new: list[_Para] = []
    _collect(old_body, (), old, skip)
    _collect(new_body, (), new, skip)
    old_texts = [p.text for p in old]
    new_texts = [p.text for p in new]
    if respect_tables:
        steps, counterpart = _respect_tables(align(old_texts, new_texts), old, new)
    else:
        steps = align(old_texts, new_texts)
        counterpart = _correspondence(steps, old, new, weighted=False)

    revisions = _Revisions(
        (new_root, old_root) if keep_pending else (new_root,), author, date or datetime.now(timezone.utc)
    )
    ctx = _Context(
        revisions=revisions,
        old_by_element={p.element: p for p in old},
        deleted={old[s.old_index].element for s in steps if s.old_index is not None and s.new_index is None},
        counterpart=counterpart,
        keep_pending=keep_pending,
    )

    kept: set[etree._Element] = set()
    for step in steps:
        if step.new_index is None:
            continue
        paragraph = new[step.new_index]
        if step.old_index is None:
            runs = [s.run for s in paragraph.spans]
            if keep_pending:
                _drop_copies(_pending_deletions(paragraph.element))
                for run in runs:
                    _lift_out_of_pending(run, revisions)
            _wrap_inserted(runs, revisions)
            _mark_paragraph_mark(paragraph.element, "ins", revisions)
            continue
        kept.update(element for _, row, cell in paragraph.path for element in (row, cell))
        standing = None
        if keep_pending:
            standing = _restore_pending(paragraph, old[step.old_index], step.segments or [], revisions)
        if step.segments:
            _apply_edit(
                paragraph, old[step.old_index], step.segments, revisions, keep_pending=keep_pending, standing=standing
            )

    # A row or cell of the revised document with nothing from the base in it is
    # an inserted row or cell, not just one holding inserted paragraphs.
    from_base = set(ctx.counterpart.values())
    marked: set[etree._Element] = set()
    for paragraph in new:
        for _, row, cell in paragraph.path:
            if row in marked:
                continue
            if row not in kept and row not in from_base:
                _mark_row(row, "ins", revisions)
                marked.add(row)
            elif cell not in kept and cell not in from_base and cell not in marked:
                _mark_cell(cell, "cellIns", revisions)
                marked.add(cell)

    for step in steps:
        if step.new_index is None and step.old_index is not None:
            paragraph = old[step.old_index]
            if paragraph.element not in ctx.placed:
                _place_deleted(paragraph, ctx, old_body, new_body)

    if revised is not None:
        _renumber_repeated(new_root, revisions)
        _verify_keeping(new_root, revisions.ours, old_root, revised)
    else:
        _verify(new_root, old_texts, new_texts)
    # No etree.cleanup_namespaces() here: it drops declarations whose prefix is
    # used only by name inside mc:Ignorable, and Word rejects a document whose
    # Ignorable list names an undeclared prefix. lxml already reuses the
    # document's own w: declaration for every element added above.
    return _repackage(package, main_name, new_root)
