"""bench/views.py: the read-back that decides whether a Word reference is exact."""

from __future__ import annotations

import io
import sys
from pathlib import Path
from typing import Any

import docx
from docx.oxml import parse_xml

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "bench"))

import views

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _save(document: Any) -> bytes:
    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue()


def _revision(paragraph: Any, kind: str, text: str, revision_id: int) -> None:
    tag = "t" if kind in ("ins", "moveTo") else "delText"
    paragraph._p.append(
        parse_xml(
            f'<w:{kind} xmlns:w="{W}" w:id="{revision_id}" w:author="Word">'
            f'<w:r><w:{tag} xml:space="preserve">{text}</w:{tag}></w:r></w:{kind}>'
        )
    )


def _views(data: bytes) -> tuple[list[str], list[str]]:
    blocks = views.read_blocks(data)
    return views.view_paragraphs(blocks, "old"), views.view_paragraphs(blocks, "new")


def test_moves_count_as_deleted_and_inserted_text() -> None:
    document: Any = docx.Document()
    paragraph = document.add_paragraph("Fee ")
    _revision(paragraph, "moveFrom", "payable monthly", 1)
    _revision(paragraph, "moveTo", "payable quarterly", 2)

    old, new = _views(_save(document))

    assert (old, new) == (["Fee payable monthly"], ["Fee payable quarterly"])


def test_a_paragraph_whose_mark_was_inserted_runs_on_in_the_old_view_and_keeps_its_space() -> None:
    document: Any = docx.Document()
    first = document.add_paragraph("Prepared by: ")
    first._p.get_or_add_pPr().append(parse_xml(f'<w:rPr xmlns:w="{W}"><w:ins w:id="3" w:author="Word"/></w:rPr>'))
    document.add_paragraph("Example Partners LLP")

    old, new = _views(_save(document))

    assert old == ["Prepared by: Example Partners LLP"]
    assert new == ["Prepared by: ", "Example Partners LLP"]


def test_an_inserted_row_does_not_exist_in_the_old_view() -> None:
    document: Any = docx.Document()
    table = document.add_table(rows=2, cols=1)
    table.cell(0, 0).text = "Kept"
    table.cell(1, 0).text = "Added"
    row = table.rows[1]._tr
    row.insert(0, parse_xml(f'<w:trPr xmlns:w="{W}"><w:ins w:id="4" w:author="Word"/></w:trPr>'))

    old, new = _views(_save(document))

    assert "Added" not in old and "Added" in new


def test_page_breaks_and_field_codes_are_not_text() -> None:
    document: Any = docx.Document()
    paragraph = document.add_paragraph("Page ")
    paragraph._p.append(parse_xml(f'<w:r xmlns:w="{W}"><w:instrText>PAGE</w:instrText><w:br w:type="page"/></w:r>'))

    assert _views(_save(document)) == (["Page "], ["Page "])
