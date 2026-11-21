# SPDX-License-Identifier: MIT
# Copyright (c) 2026 Kyra Law Limited
"""Read a Word document, or a Word Compare output, the way Word shows it.

``bench/build.py`` checks every reference with this: Word's comparison, with its
changes accepted, must read exactly as ``new.docx``, and with them rejected exactly
as ``old.docx``. It also records the text Word marked inserted and deleted.

* Text is read from runs, in paragraph order, table cells included. ``w:ins`` and
  ``w:moveTo`` are inserted text, ``w:del`` and ``w:moveFrom`` deleted text.
* A paragraph whose mark was inserted runs on into the next one in the old view,
  and one whose mark was deleted in the new view, as in Word. A table row inserted
  (deleted) does not exist in the old (new) view, nor does a cell inserted into or
  deleted from a row that otherwise stays.
* Whitespace at a paragraph's edges is kept, because where Word splits or joins two
  paragraphs, a trailing space is real text ("Prepared by: " joined to the next line).
  Callers trim each finished paragraph themselves.
* Not read as text: field codes, soft hyphens, symbols, note and comment references,
  page and column breaks, and the paragraphs of text boxes.
"""

from __future__ import annotations

import io
import re
import zipfile
from collections.abc import Iterator
from dataclasses import dataclass, field
from typing import Literal
from xml.etree import ElementTree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
View = Literal["old", "new"]

_INSERTING = frozenset({f"{W}ins", f"{W}moveTo"})
_DELETING = frozenset({f"{W}del", f"{W}moveFrom"})
_SKIP = frozenset(
    f"{W}{tag}"
    for tag in (
        "pPr",
        "rPr",
        "instrText",
        "delInstrText",
        "softHyphen",
        "sym",
        "footnoteReference",
        "endnoteReference",
        "commentReference",
    )
)
_CONTAINERS = frozenset({f"{W}sdt", f"{W}sdtContent", f"{W}customXml"})


@dataclass
class Paragraph:
    #: The text as (kind, text): "e" unchanged, "i" inserted, "d" deleted.
    segs: list[tuple[str, str]]
    #: "ins" or "del" when the paragraph mark itself is a revision.
    mark: str | None


@dataclass
class Row:
    change: str | None
    cells: list[list[Paragraph]]
    cell_changes: list[str | None] = field(default_factory=list)


@dataclass
class Table:
    rows: list[Row]


Block = Paragraph | Table


def _mark(ppr: ElementTree.Element | None) -> str | None:
    rpr = ppr.find(f"{W}rPr") if ppr is not None else None
    if rpr is None:
        return None
    if rpr.find(f"{W}del") is not None or rpr.find(f"{W}moveFrom") is not None:
        return "del"
    if rpr.find(f"{W}ins") is not None or rpr.find(f"{W}moveTo") is not None:
        return "ins"
    return None


def _paragraph(p: ElementTree.Element) -> Paragraph:
    segs: list[list[str]] = []

    def add(kind: str, text: str) -> None:
        if not text:
            return
        if segs and segs[-1][0] == kind:
            segs[-1][1] += text
        else:
            segs.append([kind, text])

    def walk(el: ElementTree.Element, kind: str) -> None:
        for c in el:
            if c.tag in _SKIP or c.tag == f"{W}p":
                continue
            if c.tag in _INSERTING:
                walk(c, "i")
            elif c.tag in _DELETING:
                walk(c, "d")
            elif c.tag in (f"{W}t", f"{W}delText"):
                add(kind, c.text or "")
            elif c.tag == f"{W}tab":
                add(kind, "\t")
            elif c.tag in (f"{W}br", f"{W}cr"):
                if c.get(f"{W}type") not in ("page", "column"):
                    add(kind, "\n")
            elif c.tag == f"{W}noBreakHyphen":
                add(kind, "-")
            else:
                walk(c, kind)

    walk(p, "e")
    return Paragraph([(k, t) for k, t in segs if t], _mark(p.find(f"{W}pPr")))


def _change(props: ElementTree.Element | None, deleted: str, inserted: str) -> str | None:
    if props is None:
        return None
    if props.find(f"{W}{deleted}") is not None:
        return "del"
    if props.find(f"{W}{inserted}") is not None:
        return "ins"
    return None


def _children(el: ElementTree.Element, tag: str) -> Iterator[ElementTree.Element]:
    for c in el:
        if c.tag == f"{W}{tag}":
            yield c
        elif c.tag in _CONTAINERS:
            yield from _children(c, tag)


def _cell(tc: ElementTree.Element) -> list[Paragraph]:
    out: list[Paragraph] = []

    def walk(el: ElementTree.Element) -> None:
        for c in el:
            if c.tag == f"{W}p":
                out.append(_paragraph(c))
            elif c.tag == f"{W}tbl":
                for nested in c.iter(f"{W}tc"):
                    walk(nested)
            elif c.tag in _CONTAINERS:
                walk(c)

    walk(tc)
    return out


def _table(tbl: ElementTree.Element) -> Table:
    rows = []
    for tr in _children(tbl, "tr"):
        tcs = list(_children(tr, "tc"))
        rows.append(
            Row(
                change=_change(tr.find(f"{W}trPr"), "del", "ins"),
                cells=[_cell(tc) for tc in tcs],
                cell_changes=[_change(tc.find(f"{W}tcPr"), "cellDel", "cellIns") for tc in tcs],
            )
        )
    return Table(rows)


def read_blocks(docx: bytes) -> list[Block]:
    """The body of a .docx as paragraphs and tables, in order."""
    body = ElementTree.fromstring(zipfile.ZipFile(io.BytesIO(docx)).read("word/document.xml")).find(f"{W}body")
    if body is None:
        raise ValueError("the document has no body")
    blocks: list[Block] = []

    def walk(el: ElementTree.Element) -> None:
        for c in el:
            if c.tag == f"{W}p":
                blocks.append(_paragraph(c))
            elif c.tag == f"{W}tbl":
                blocks.append(_table(c))
            elif c.tag in _CONTAINERS:
                walk(c)

    walk(body)
    return blocks


def paragraphs(blocks: list[Block]) -> Iterator[tuple[Paragraph, str | None]]:
    """Every paragraph in order, with the change of the table row or cell holding it."""
    for block in blocks:
        if isinstance(block, Paragraph):
            yield block, None
            continue
        for row in block.rows:
            for c, cell in enumerate(row.cells):
                change = row.change or (row.cell_changes[c] if c < len(row.cell_changes) else None)
                for p in cell:
                    yield p, change


def view_paragraphs(blocks: list[Block], view: View) -> list[str]:
    """The document as it reads on one side of the comparison, one string per paragraph."""
    absent = "ins" if view == "old" else "del"
    keep = "d" if view == "old" else "i"
    out: list[str] = []
    carry = ""
    for p, change in paragraphs(blocks):
        if change == absent:
            continue
        text = carry + "".join(t for k, t in p.segs if k in ("e", keep))
        if p.mark == absent:
            carry = text
            continue
        carry = ""
        out.append(text)
    if carry:
        out.append(carry)
    return out


def words(text: str) -> list[str]:
    return re.findall(r"\w+", text.lower())
