"""Tests for the in-place renderer (kyralaw_redline/inplace.py).

The renderer writes the comparison into the revised .docx itself, so these build
real documents with python-docx and read the result's XML back. Accepting and
rejecting the revisions is done here by a reader of its own, independent of the
module's internal check, so a bug shared by the two cannot make both agree.
"""

from __future__ import annotations

import io
import zipfile
from datetime import datetime, timezone
from typing import Any

import docx
import pytest
from docx.oxml import parse_xml
from lxml import etree

from kyralaw_redline import inplace
from kyralaw_redline.diff import align_paragraphs
from kyralaw_redline.docx_io import (
    InPlaceFallbackWarning,
    build_redline_docx,
    extract_text,
    read_tracked_changes,
    redline_docx,
)
from kyralaw_redline.inplace import InPlaceRedlineError, redline_in_place

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC_IGNORABLE = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable"
DATE = datetime(2026, 10, 2, tzinfo=timezone.utc)


def _q(tag: str) -> str:
    return f"{{{W}}}{tag}"


def _save(document: Any) -> bytes:
    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue()


def _document(*paragraphs: str | tuple[str, str]) -> Any:
    document = docx.Document()
    for item in paragraphs:
        text, style = (item, None) if isinstance(item, str) else item
        document.add_paragraph(text, style=style)
    return document


def _table(document: Any, rows: list[list[str]]) -> None:
    table = document.add_table(rows=len(rows), cols=len(rows[0]))
    for r, values in enumerate(rows):
        for c, value in enumerate(values):
            table.cell(r, c).text = value


def _main(data: bytes) -> etree._Element:
    return etree.fromstring(zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml"))


def _redline(old: Any, new: Any) -> etree._Element:
    return _main(redline_in_place(_save(old), _save(new), date=DATE))


def _texts(root: etree._Element, view: str) -> list[str]:
    """Non-empty paragraph texts with every revision accepted or rejected."""
    gone = _q("del") if view == "accept" else _q("ins")
    texts: list[str] = []
    for paragraph in root.iter(_q("p")):
        cell = paragraph.getparent()
        if cell is not None and cell.tag == _q("tc"):
            if cell.find(f"{_q('tcPr')}/{_q('cellDel' if view == 'accept' else 'cellIns')}") is not None:
                continue
            row = cell.getparent()
            if row is not None and row.find(f"{_q('trPr')}/{gone}") is not None:
                continue
        parts: list[str] = []
        for node in paragraph.iter(_q("t"), _q("delText")):
            if gone in {a.tag for a in node.iterancestors()}:
                continue
            if node.tag == _q("t") or view == "reject":
                parts.append(node.text or "")
        text = "".join(parts)
        if text.strip():
            texts.append(text)
    return texts


def _inserted(root: etree._Element) -> list[str]:
    return ["".join(t.text or "" for t in e.iter(_q("t"))) for e in root.iter(_q("ins")) if len(e)]


def _deleted(root: etree._Element) -> list[str]:
    return ["".join(t.text or "" for t in e.iter(_q("delText"))) for e in root.iter(_q("del")) if len(e)]


def test_a_word_change_is_marked_inside_the_revised_paragraph() -> None:
    root = _redline(
        _document(("The fee is 100 pounds.", "Heading 1")), _document(("The fee is 200 pounds.", "Heading 1"))
    )

    assert _inserted(root) == ["200"]
    assert _deleted(root) == ["100"]
    # The paragraph is the revised document's own, style and all.
    assert [s.get(_q("val")) for s in root.iter(_q("pStyle"))] == ["Heading1"]


def test_run_formatting_is_kept_and_no_markup_is_printed() -> None:
    old = docx.Document()
    old_paragraph = old.add_paragraph()
    old_paragraph.add_run("Supplier").bold = True
    old_paragraph.add_run(" shall deliver monthly.")
    new = docx.Document()
    new_paragraph = new.add_paragraph()
    new_paragraph.add_run("Supplier").bold = True
    new_paragraph.add_run(" shall deliver weekly.")

    data = redline_in_place(_save(old), _save(new), date=DATE)
    root = _main(data)

    bold_runs = [r for r in root.iter(_q("r")) if r.find(f"{_q('rPr')}/{_q('b')}") is not None]
    assert ["".join(t.text or "" for t in r.iter(_q("t"))) for r in bold_runs] == ["Supplier"]
    assert b"**" not in etree.tostring(root)
    assert (_inserted(root), _deleted(root)) == (["weekly"], ["monthly"])


def test_an_inserted_paragraph_is_marked_down_to_its_paragraph_mark() -> None:
    old = _document("Clause one.", "Clause three.")
    new = _document("Clause one.", "Clause two.", "Clause three.")
    root = _redline(old, new)

    assert _inserted(root) == ["Clause two."]
    marks = [p for p in root.iter(_q("p")) if p.find(f"{_q('pPr')}/{_q('rPr')}/{_q('ins')}") is not None]
    assert len(marks) == 1
    assert _texts(root, "accept") == ["Clause one.", "Clause two.", "Clause three."]
    assert _texts(root, "reject") == ["Clause one.", "Clause three."]


def test_a_removed_paragraph_comes_back_as_a_deletion_with_its_own_style() -> None:
    old = _document("Clause one.", ("Termination", "Heading 2"), "Clause three.")
    new = _document("Clause one.", "Clause three.")
    root = _redline(old, new)

    deleted = [p for p in root.iter(_q("p")) if p.find(f"{_q('pPr')}/{_q('rPr')}/{_q('del')}") is not None]
    assert len(deleted) == 1
    style = deleted[0].find(f"{_q('pPr')}/{_q('pStyle')}")
    assert style is not None and style.get(_q("val")) == "Heading2"
    assert _deleted(root) == ["Termination"]
    assert _texts(root, "accept") == ["Clause one.", "Clause three."]
    assert _texts(root, "reject") == ["Clause one.", "Termination", "Clause three."]


def test_a_removed_table_row_is_a_deleted_row_in_the_revised_table() -> None:
    old, new = docx.Document(), docx.Document()
    _table(old, [["Fee", "100"], ["Term", "12 months"], ["Law", "Ireland"]])
    _table(new, [["Fee", "100"], ["Law", "Ireland"]])
    root = _redline(old, new)

    rows = list(root.iter(_q("tr")))
    assert len(rows) == 3
    assert [r.find(f"{_q('trPr')}/{_q('del')}") is not None for r in rows] == [False, True, False]
    assert _texts(root, "accept") == ["Fee", "100", "Law", "Ireland"]
    assert _texts(root, "reject") == ["Fee", "100", "Term", "12 months", "Law", "Ireland"]


def test_a_removed_column_is_kept_as_deleted_cells() -> None:
    old, new = docx.Document(), docx.Document()
    _table(old, [["Item", "Optional Extras"], ["Training", "None"]])
    _table(new, [["Item"], ["Training"]])
    root = _redline(old, new)

    for row in root.iter(_q("tr")):
        cells = row.findall(_q("tc"))
        assert len(cells) == 2
        assert cells[1].find(f"{_q('tcPr')}/{_q('cellDel')}") is not None
    grid = root.find(f".//{_q('tblGrid')}")
    assert grid is not None and len(grid) == 2
    assert _texts(root, "reject") == ["Item", "Optional Extras", "Training", "None"]


def test_pending_changes_in_the_inputs_are_accepted_before_comparing() -> None:
    new = _document("The term is twelve months.")
    paragraph = new.paragraphs[0]._p
    paragraph.append(
        parse_xml(f'<w:ins xmlns:w="{W}" w:id="90" w:author="Dan"><w:r><w:t> Renewal is automatic.</w:t></w:r></w:ins>')
    )
    paragraph.append(
        parse_xml(f'<w:del xmlns:w="{W}" w:id="91" w:author="Dan"><w:r><w:delText> Struck.</w:delText></w:r></w:del>')
    )
    old = _document("The term is twelve months.")

    root = _redline(old, new)

    authors = {e.get(_q("author")) for e in root.iter(_q("ins"), _q("del"))}
    assert "Dan" not in authors
    # Dan's insertion is part of the revised document, so it is what we marked.
    assert _inserted(root) == [" Renewal is automatic."]
    assert _deleted(root) == []


def test_a_comment_on_text_a_pending_deletion_removes_keeps_both_ends() -> None:
    new = docx.Document()
    paragraph = new.add_paragraph()
    paragraph.add_run("Keep this. ")
    doomed = paragraph.add_run("Drop this.")
    new.add_comment(doomed, text="Why is this here?", author="Dan")
    # Move the commented run, the range end and the reference into a pending deletion.
    p = paragraph._p
    wrapper = parse_xml(f'<w:del xmlns:w="{W}" w:id="80" w:author="Dan"/>')
    for element in [e for e in p if e.tag in (_q("r"), _q("commentRangeEnd"))][1:]:
        wrapper.append(element)
    for t in wrapper.iter(_q("t")):
        t.tag = _q("delText")
    p.append(wrapper)
    old = _document("Keep this. Drop this.")

    root = _redline(old, new)

    starts = {e.get(_q("id")) for e in root.iter(_q("commentRangeStart"))}
    ends = {e.get(_q("id")) for e in root.iter(_q("commentRangeEnd"))}
    references = {e.get(_q("id")) for e in root.iter(_q("commentReference"))}
    assert starts == ends == references != set()


def test_prefixes_named_in_mc_ignorable_stay_declared() -> None:
    root = _redline(_document("Fee is 100."), _document("Fee is 200."))

    ignorable = root.get(MC_IGNORABLE, "").split()
    assert ignorable
    assert all(prefix in root.nsmap for prefix in ignorable)


def test_revision_ids_are_unique_and_clear_of_existing_annotations() -> None:
    old = _document("A.", "B.", "C.")
    new = _document("A!", "C.", "D.")
    comment_target = new.paragraphs[0].runs[0]
    new.add_comment(comment_target, text="note", author="Dan")
    root = _redline(old, new)

    revision_ids = [e.get(_q("id")) for e in root.iter(_q("ins"), _q("del"), _q("cellIns"), _q("cellDel"))]
    other_ids = {e.get(_q("id")) for e in root.iter(_q("commentRangeStart"), _q("commentReference"))}
    assert len(revision_ids) == len(set(revision_ids))
    assert not set(revision_ids) & other_ids


def test_identical_documents_produce_no_revisions() -> None:
    root = _redline(_document("Same.", "Text."), _document("Same.", "Text."))

    assert list(root.iter(_q("ins"), _q("del"))) == []


def test_the_result_opens_as_a_word_document() -> None:
    data = redline_in_place(_save(_document("Old wording.")), _save(_document("New wording.")), date=DATE)

    # python-docx's paragraph.text skips runs inside w:ins, so count paragraphs.
    assert len(docx.Document(io.BytesIO(data)).paragraphs) == 1


def test_input_that_is_not_a_docx_is_refused() -> None:
    with pytest.raises(InPlaceRedlineError):
        redline_in_place(b"%PDF-1.7 not a zip", _save(_document("Text.")))


# --------------------------------------------------------------------------
# keep_pending: the revised document was edited directly from the base, with
# the base's pending revisions still in it, by a tool that keeps tracked changes
# through a save but writes its own edits untracked.
# --------------------------------------------------------------------------


def _pending(paragraph: Any, kind: str, text: str, revision_id: int, author: str = "Dan") -> None:
    """Append ``text`` to a python-docx paragraph as ``author``'s pending insertion or deletion."""
    tag = "t" if kind == "ins" else "delText"
    paragraph._p.append(
        parse_xml(
            f'<w:{kind} xmlns:w="{W}" w:id="{revision_id}" w:author="{author}">'
            f'<w:r><w:{tag} xml:space="preserve">{text}</w:{tag}></w:r></w:{kind}>'
        )
    )


def _pending_paragraph(document: Any, text: str, revision_id: int, author: str = "Dan") -> None:
    """A whole paragraph, mark included, that ``author`` has pending as inserted."""
    paragraph = document.add_paragraph()
    _pending(paragraph, "ins", text, revision_id, author)
    paragraph._p.get_or_add_pPr().append(
        parse_xml(f'<w:rPr xmlns:w="{W}"><w:ins w:id="{revision_id + 1}" w:author="{author}"/></w:rPr>')
    )


def _keep(old: Any, new: Any) -> etree._Element:
    data = redline_in_place(_save(old), _save(new), author="Agent", date=DATE, keep_pending=True)
    return _main(data)


def _resolved(root: etree._Element, accepted: set[str]) -> list[str]:
    """Paragraph texts with ``accepted`` authors' revisions accepted and everyone else's rejected.

    Read by author rather than by revision id, so it shares nothing with the
    module's own per-author check.
    """

    def gone(marker: etree._Element) -> bool:
        taken = marker.get(_q("author")) in accepted
        return (marker.tag == _q("ins") and not taken) or (marker.tag == _q("del") and taken)

    texts: list[str] = []
    carried = ""
    for paragraph in root.iter(_q("p")):
        parts = []
        for node in paragraph.iter(_q("t"), _q("delText")):
            markers = [a for a in node.iterancestors(_q("ins"), _q("del")) if paragraph in a.iterancestors()]
            if not any(gone(m) for m in markers):
                parts.append(node.text or "")
        text = carried + "".join(parts)
        mark = paragraph.find(f"{_q('pPr')}/{_q('rPr')}")
        if mark is not None and any(gone(m) for m in mark if m.tag in (_q("ins"), _q("del"))):
            carried = text
            continue
        carried = ""
        if text.strip():
            texts.append(text)
    return texts


def _by(root: etree._Element, kind: str, author: str) -> list[str]:
    tag = "t" if kind == "ins" else "delText"
    return [
        "".join(t.text or "" for t in e.iter(_q(tag)))
        for e in root.iter(_q(kind))
        if e.get(_q("author")) == author and len(e)
    ]


def test_keep_pending_marks_only_the_new_edit_and_leaves_pending_changes_as_they_were() -> None:
    def document(term: str) -> Any:
        result = _document(f"The term is {term} months.")
        _pending(result.paragraphs[0], "ins", " Renewal is automatic.", 90)
        _pending(result.paragraphs[0], "del", " Struck.", 91)
        return result

    root = _keep(document("twelve"), document("six"))

    assert (_by(root, "ins", "Agent"), _by(root, "del", "Agent")) == (["six"], ["twelve"])
    assert (_by(root, "ins", "Dan"), _by(root, "del", "Dan")) == ([" Renewal is automatic."], [" Struck."])
    assert _resolved(root, {"Agent", "Dan"}) == ["The term is six months. Renewal is automatic."]
    assert _resolved(root, {"Dan"}) == ["The term is twelve months. Renewal is automatic."]
    assert _resolved(root, set()) == ["The term is twelve months. Struck."]


def test_keep_pending_writes_an_edit_inside_a_pending_insertion_around_it() -> None:
    def document(period: str) -> Any:
        result = _document("The term is twelve months.")
        _pending(result.paragraphs[0], "ins", f" Renewal is automatic for {period} at a time.", 90)
        return result

    root = _keep(document("one year"), document("two years"))

    # An insertion cannot hold another, so Dan's is split around ours...
    assert not [e for e in root.iter(_q("ins")) if (up := e.getparent()) is not None and up.tag == _q("ins")]
    assert _by(root, "ins", "Agent") == ["two years"]
    # ...and deleting text Dan inserted is a deletion inside an insertion of his.
    deletions = [e for e in root.iter(_q("del")) if e.get(_q("author")) == "Agent"]
    assert [up.get(_q("author")) for d in deletions if (up := d.getparent()) is not None] == ["Dan"]
    assert _resolved(root, {"Agent", "Dan"}) == [
        "The term is twelve months. Renewal is automatic for two years at a time."
    ]
    assert _resolved(root, {"Dan"}) == ["The term is twelve months. Renewal is automatic for one year at a time."]
    assert _resolved(root, set()) == ["The term is twelve months."]


def test_keep_pending_deletes_a_paragraph_with_its_pending_revisions_kept_inside() -> None:
    def document(*, with_second: bool) -> Any:
        result = _document("Clause one.")
        if with_second:
            second = result.add_paragraph("Clause two.")
            _pending(second, "ins", " As amended.", 90)
            _pending(second, "del", " As drafted.", 91)
        result.add_paragraph("Clause three.")
        return result

    root = _keep(document(with_second=True), document(with_second=False))

    assert _resolved(root, {"Agent", "Dan"}) == ["Clause one.", "Clause three."]
    assert _resolved(root, {"Dan"}) == ["Clause one.", "Clause two. As amended.", "Clause three."]
    assert _resolved(root, set()) == ["Clause one.", "Clause two. As drafted.", "Clause three."]
    revision_ids = [e.get(_q("id")) for e in root.iter(_q("ins"), _q("del"))]
    assert len(revision_ids) == len(set(revision_ids))


def test_keep_pending_deletes_a_pending_inserted_paragraph_with_both_marks_in_schema_order() -> None:
    old = _document("Clause one.")
    _pending(old.paragraphs[0], "ins", " Agreed.", 70, author="Ann")
    _pending_paragraph(old, "Clause two.", 90)
    old.add_paragraph("Clause three.")
    # The two documents number their revisions independently: Ann's is 91 here,
    # the id Dan's paragraph mark has in the base.
    new = _document("Clause one.", "Clause three.")
    _pending(new.paragraphs[0], "ins", " Agreed.", 91, author="Ann")

    root = _keep(old, new)

    marks = [p.find(f"{_q('pPr')}/{_q('rPr')}") for p in root.iter(_q("p"))]
    tags = [[(m.tag, m.get(_q("author"))) for m in mark] for mark in marks if mark is not None and len(mark)]
    assert tags == [[(_q("ins"), "Dan"), (_q("del"), "Agent")]]
    assert _resolved(root, {"Agent", "Dan", "Ann"}) == ["Clause one. Agreed.", "Clause three."]
    assert _resolved(root, {"Dan", "Ann"}) == ["Clause one. Agreed.", "Clause two.", "Clause three."]
    assert _resolved(root, set()) == ["Clause one.", "Clause three."]
    revision_ids = [e.get(_q("id")) for e in root.iter(_q("ins"), _q("del"))]
    assert len(revision_ids) == len(set(revision_ids))


def test_keep_pending_with_no_edit_adds_no_revisions() -> None:
    def document() -> Any:
        result = _document("The term is twelve months.")
        _pending(result.paragraphs[0], "ins", " Renewal is automatic.", 90)
        return result

    root = _keep(document(), document())

    assert {e.get(_q("author")) for e in root.iter(_q("ins"), _q("del"))} == {"Dan"}


def test_without_keep_pending_the_pending_revisions_are_still_accepted_first() -> None:
    old = _document("The term is twelve months.")
    new = _document("The term is twelve months.")
    _pending(new.paragraphs[0], "ins", " Renewal is automatic.", 90)

    root = _main(redline_in_place(_save(old), _save(new), author="Agent", date=DATE))

    assert {e.get(_q("author")) for e in root.iter(_q("ins"), _q("del"))} == {"Agent"}


def _plain(*runs: str) -> Any:
    """One paragraph of plain runs, as such a tool saves a segment it edited."""
    document = docx.Document()
    paragraph = document.add_paragraph()
    for text in runs:
        paragraph.add_run(text)
    return document


def test_keep_pending_puts_back_an_insertion_the_revised_paragraph_dropped() -> None:
    """The tool writes an edited segment as plain text, so an edit inside Dan's
    insertion comes back with the rest of his words no longer marked as his."""
    old = _document("Theta iota.")
    _pending(old.paragraphs[0], "ins", " Kappa lambda mu.", 90)
    old.paragraphs[0].add_run(" Nu xi.")

    root = _keep(old, _plain("Theta iota.", " Kappa LAMBDA mu.", " Nu xi."))

    assert _by(root, "ins", "Agent") == ["LAMBDA"]
    assert "".join(_by(root, "ins", "Dan")) == " Kappa  mu."
    assert _resolved(root, {"Agent", "Dan"}) == ["Theta iota. Kappa LAMBDA mu. Nu xi."]
    assert _resolved(root, {"Dan"}) == ["Theta iota. Kappa lambda mu. Nu xi."]
    assert _resolved(root, set()) == ["Theta iota. Nu xi."]


@pytest.mark.parametrize(
    "revised",
    [("Alpha beta gamma.", " Struck DELTA.", " Epsilon zeta."), ("Alpha beta gamma.", " Epsilon zeta.")],
    ids=["edited", "removed"],
)
def test_keep_pending_puts_back_a_deletion_the_revised_paragraph_dropped(revised: tuple[str, ...]) -> None:
    """The tool shows Dan's deleted text as a segment of its own, so an edit to it
    brings it back as plain text, and deleting that segment removes it outright."""
    old = _document("Alpha beta gamma.")
    _pending(old.paragraphs[0], "del", " Struck delta.", 91)
    old.paragraphs[0].add_run(" Epsilon zeta.")

    root = _keep(old, _plain(*revised))

    assert _by(root, "del", "Dan") == [" Struck delta."]
    assert _resolved(root, {"Agent", "Dan"}) == ["".join(revised)]
    assert _resolved(root, {"Dan"}) == ["Alpha beta gamma. Epsilon zeta."]
    assert _resolved(root, set()) == ["Alpha beta gamma. Struck delta. Epsilon zeta."]


def _paragraph_of(*parts: str | tuple[str, str, int]) -> Any:
    """One paragraph from plain runs and ("ins" | "del", text, id) pending revisions by Dan."""
    document = docx.Document()
    paragraph = document.add_paragraph()
    for part in parts:
        if isinstance(part, str):
            paragraph.add_run(part)
        else:
            _pending(paragraph, *part)
    return document


def test_keep_pending_puts_a_lost_deletion_back_inside_the_run_it_belongs_in() -> None:
    """The tool saved "3." + Dan's "8" + Dan's deleted "5" as one plain run "3.9". Put back
    at a position no run boundary marked, Dan's "5" landed at the paragraph's end."""
    old = _paragraph_of("3.", ("ins", "8", 90), ("del", "5", 91), " Phase Overview")

    root = _keep(old, _plain("3.9", " Phase Overview"))

    assert _resolved(root, {"Agent", "Dan"}) == ["3.9 Phase Overview"]
    assert _resolved(root, {"Dan"}) == ["3.8 Phase Overview"]
    assert _resolved(root, set()) == ["3.5 Phase Overview"]


def test_keep_pending_orders_our_deletion_before_a_pending_deletion_that_followed_it() -> None:
    """Both show nothing at the same place, so only the base can say which comes first."""
    old = _paragraph_of(
        "Onboarding workshops, ",
        ("ins", "limited to ten days", 90),
        ("del", "limited to [N] days", 91),
        " at the agreed rates.",
    )
    new = _paragraph_of(("del", "limited to [N] days", 91), " at the agreed rates.")

    root = _keep(old, new)

    assert _resolved(root, {"Agent", "Dan"}) == [" at the agreed rates."]
    assert _resolved(root, {"Dan"}) == ["Onboarding workshops, limited to ten days at the agreed rates."]
    assert _resolved(root, set()) == ["Onboarding workshops, limited to [N] days at the agreed rates."]


def test_keep_pending_splits_our_deletion_around_a_pending_deletion_inside_it() -> None:
    """The word diff took a space from each side of Dan's deletion, so the deleted
    range starts before it and ends after it."""
    old = _paragraph_of(
        "The system will run in line with ",
        ("ins", "the manual, provided that it is configured as described.", 90),
        ("del", "the manual.", 91),
        " If any component breaks",
    )
    new = _paragraph_of("The system will run in line with ", ("del", "the manual.", 91), " If any component breaks")

    root = _keep(old, new)

    assert _resolved(root, {"Agent", "Dan"}) == ["The system will run in line with  If any component breaks"]
    assert _resolved(root, {"Dan"}) == [
        "The system will run in line with the manual, provided that it is configured as described. "
        "If any component breaks"
    ]
    assert _resolved(root, set()) == ["The system will run in line with the manual. If any component breaks"]


def test_keep_pending_keeps_one_copy_of_a_pending_deletion_in_a_paragraph_it_emptied() -> None:
    """The tool removed Dan's inserted text and kept his deletion beside it, so the
    revised paragraph shows nothing and pairs with none of the base's: it is written
    as deleted and inserted again, and the deletion stood in both."""

    def document(*, inserted: bool) -> Any:
        result = _document("Clause one.")
        paragraph = result.add_paragraph()
        if inserted:
            _pending(paragraph, "ins", "The figures are estimates, not commitments of supply.", 90)
        _pending(paragraph, "del", "Proposals for revisions", 91)
        result.add_paragraph("Stage five.")
        return result

    root = _keep(document(inserted=True), document(inserted=False))

    assert _by(root, "del", "Dan") == ["Proposals for revisions"]
    assert _resolved(root, {"Agent", "Dan"}) == ["Clause one.", "Stage five."]
    assert _resolved(root, {"Dan"}) == [
        "Clause one.",
        "The figures are estimates, not commitments of supply.",
        "Stage five.",
    ]
    assert _resolved(root, set()) == ["Clause one.", "Proposals for revisions", "Stage five."]


def test_keep_pending_keeps_one_copy_of_a_pending_deletion_found_in_another_paragraph() -> None:
    """The same duplicate through a paired paragraph: the deletion is put back where the
    base had it, so the revised document's copy elsewhere is a second one."""
    old = _document("Alpha beta gamma.")
    second = old.add_paragraph()
    _pending(second, "del", "Gone ", 91)
    second.add_run("Delta epsilon.")
    new = _document("Alpha beta gamma.", "Delta epsilon.")
    _pending(new.paragraphs[0], "del", "Gone ", 91)

    root = _keep(old, new)

    assert _by(root, "del", "Dan") == ["Gone "]
    assert _resolved(root, {"Agent", "Dan"}) == ["Alpha beta gamma.", "Delta epsilon."]
    assert _resolved(root, set()) == ["Alpha beta gamma.", "Gone Delta epsilon."]


# --------------------------------------------------------------------------
# redline_docx's modes, and the alignment both renderers share
# --------------------------------------------------------------------------


def test_redline_docx_writes_in_place_by_default() -> None:
    old = _document(("The fee is 100 pounds.", "Heading 1"))
    new = _document(("The fee is 200 pounds.", "Heading 1"))

    root = _main(redline_docx(_save(old), _save(new), date=DATE))

    # The new document's own paragraph, style and all, carries the change.
    assert [s.get(_q("val")) for s in root.iter(_q("pStyle"))] == ["Heading1"]
    assert (_inserted(root), _deleted(root)) == (["200"], ["100"])


def test_rebuild_mode_is_the_text_renderer() -> None:
    old, new = _save(_document("The fee is 100 pounds.")), _save(_document("The fee is 200 pounds."))

    rebuilt = redline_docx(old, new, date=DATE, mode="rebuild")

    assert read_tracked_changes(rebuilt) == read_tracked_changes(
        build_redline_docx(extract_text(old), extract_text(new), date=DATE)
    )


def _refuse(*_args: object, **_kwargs: object) -> bytes:
    raise InPlaceRedlineError("refused for the test")


def test_auto_mode_falls_back_to_rebuild_and_says_so(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(inplace, "redline_in_place", _refuse)
    old, new = _save(_document("The fee is 100 pounds.")), _save(_document("The fee is 200 pounds."))

    with pytest.warns(InPlaceFallbackWarning, match="refused for the test"):
        data = redline_docx(old, new, date=DATE)

    assert read_tracked_changes(data) == (["200"], ["100"])


def test_in_place_mode_raises_instead_of_falling_back(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(inplace, "redline_in_place", _refuse)
    old, new = _save(_document("The fee is 100 pounds.")), _save(_document("The fee is 200 pounds."))

    with pytest.raises(InPlaceRedlineError):
        redline_docx(old, new, date=DATE, mode="in-place")


def test_align_paragraphs_pairs_edited_paragraphs_and_leaves_the_rest_one_sided() -> None:
    steps = align_paragraphs(
        ["Clause one.", "The fee is 100 pounds.", "Clause three."],
        ["Clause one.", "The fee is 200 pounds.", "A new clause.", "Clause three."],
    )

    assert [(s.old_index, s.new_index) for s in steps] == [(0, 0), (1, 1), (None, 2), (2, 3)]
    edited = steps[1].segments
    assert edited is not None
    assert [(s.kind, s.text) for s in edited if s.kind != "equal"] == [("delete", "100"), ("insert", "200")]
    assert steps[0].segments is None and steps[3].segments is None


def test_a_deleted_row_whose_cells_also_appear_in_other_rows_rejects_back_to_the_base() -> None:
    """A margin grid where the middle row goes and the rest shift: the paragraph
    alignment, which does not know about rows, pairs old row 1's rates with new row
    2's identical ones, and old row 2's "Tier B" with new row 3's "Tier C". Placed
    by those pairs, the deleted copy of old row 2 landed after its own first cell,
    so rejecting every change no longer gave the base table."""
    old, new = docx.Document(), docx.Document()
    _table(
        old,
        [
            ["Tier A", "Leverage at or above 3.00 to 1.00", "2.50%", "1.50%"],
            ["Tier B", "", "2.50%", "1.50%"],
            ["Tier B", "Leverage below 3.00 to 1.00", "2.25%", "1.25%"],
        ],
    )
    _table(
        new,
        [
            ["Tier A", "Leverage at or above 3.00 to 1.00", "2.75%", "1.75%"],
            ["Tier B", "Leverage above 3.50 to 1.00", "2.50%", "1.50%"],
            ["Tier C", "Leverage above 2.75 to 1.00", "2.25%", "1.25%"],
        ],
    )

    root = _redline(old, new)

    assert _texts(root, "reject") == [
        "Tier A", "Leverage at or above 3.00 to 1.00", "2.50%", "1.50%",
        "Tier B", "2.50%", "1.50%",
        "Tier B", "Leverage below 3.00 to 1.00", "2.25%", "1.25%",
    ]  # fmt: skip
    assert _texts(root, "accept") == [
        "Tier A", "Leverage at or above 3.00 to 1.00", "2.75%", "1.75%",
        "Tier B", "Leverage above 3.50 to 1.00", "2.50%", "1.50%",
        "Tier C", "Leverage above 2.75 to 1.00", "2.25%", "1.25%",
    ]  # fmt: skip


def test_a_row_added_to_a_table_of_recurring_values_is_one_inserted_row() -> None:
    """Every row holds "25,000", a row is added in the middle, and an edit above the
    table stops the line diff's common prefix early, so its common suffix runs back
    through the table and pairs the old second row's value with the added row's. Kept
    as it was, the added row was not marked as an inserted row at all; split without
    pairing again, the value left in its own cell came out deleted and reinserted.
    Only the opening figure and the new row are changes."""
    head = [["Vesting date", "Shares vesting"], ["First anniversary", "25,000"], ["Second anniversary", "25,000"]]
    tail = [["Third anniversary", "25,000"]]
    old, new = _document("The Founder holds 100,000 Shares."), _document("The Founder holds 125,000 Shares.")
    _table(old, [*head, *tail])
    _table(new, [*head, ["Interim date", "25,000"], *tail])

    root = _redline(old, new)

    inserted_rows = [r for r in root.iter(_q("tr")) if r.find(f"{_q('trPr')}/{_q('ins')}") is not None]
    assert ["".join(t.text or "" for t in r.iter(_q("t"))) for r in inserted_rows] == ["Interim date25,000"]
    assert sorted(_inserted(root)) == ["125", "25,000", "Interim date"]
    assert _deleted(root) == ["100"]
    assert _texts(root, "reject") == [
        "The Founder holds 100,000 Shares.",
        *[cell for row in [*head, *tail] for cell in row],
    ]
