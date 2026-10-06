"""Tests for kyralaw_redline.docx_io: the reader, the writer, and the CLI."""

from __future__ import annotations

import io
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from docx import Document

from kyralaw_redline import (
    DiffSegment,
    added_text,
    build_redline_docx,
    compute_redline_segments,
    extract_paragraphs,
    extract_text,
    read_tracked_changes,
    redline_docx,
    removed_text,
    render_segments_docx,
)
from kyralaw_redline.cli import main as cli_main

FIXED = datetime(2026, 1, 1, tzinfo=timezone.utc)


def _docx(*paragraphs: str) -> bytes:
    doc = Document()
    for text in paragraphs:
        doc.add_paragraph(text)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _with_body_xml(body_inner: str) -> bytes:
    """A .docx whose body is exactly ``body_inner``, for markup python-docx
    cannot author (pending tracked changes)."""
    base = _docx("placeholder")
    out = io.BytesIO()
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{w}"><w:body>{body_inner}</w:body></w:document>'
    )
    with zipfile.ZipFile(io.BytesIO(base)) as src, zipfile.ZipFile(out, "w") as dst:
        for item in src.infolist():
            data = xml.encode() if item.filename == "word/document.xml" else src.read(item.filename)
            dst.writestr(item, data)
    return out.getvalue()


def test_produces_a_valid_openable_docx() -> None:
    out = build_redline_docx("Old paragraph one.\n\nOld paragraph two.", "New paragraph one.\n\nOld paragraph two.")
    assert len(Document(io.BytesIO(out)).paragraphs) == 2


def test_single_word_change_produces_one_ins_and_one_del() -> None:
    out = build_redline_docx("This Agreement is dated 4 September 2023.", "This Agreement is dated 4 September 2026.")
    assert read_tracked_changes(out) == (["2026"], ["2023"])


def test_whole_new_paragraph_is_a_pure_insertion() -> None:
    out = build_redline_docx("Clause 1. Unchanged.", "Clause 1. Unchanged.\n\nClause 2. Brand new clause.")
    ins, dele = read_tracked_changes(out)
    assert dele == []
    assert ins == ["Clause 2. Brand new clause."]


def test_no_changes_produces_no_tracked_changes() -> None:
    text = "Clause 1. Nothing changed here at all."
    assert read_tracked_changes(build_redline_docx(text, text)) == ([], [])


def test_rendered_document_carries_exactly_the_segments_it_was_built_from() -> None:
    """The .docx must contain the diff compute_redline_segments computed and
    no other. The reordered clauses are what make this bite: a renderer with
    a block-level diff of its own agrees on plain edits and parts company
    when a paragraph moves."""
    old = "1. This Agreement is dated 4 September 2023.\n\nAlpha clause about delivery.\n\nBeta clause about payment."
    new = "1. This Agreement is dated 4 September 2026.\n\nBeta clause about payment.\n\nAlpha clause about delivery."
    ins, dele = read_tracked_changes(build_redline_docx(old, new))
    segments = compute_redline_segments(old, new)
    assert " ".join(ins).split() == added_text(segments).split()
    assert " ".join(dele).split() == removed_text(segments).split()


def test_fixed_date_and_author_are_stamped_on_every_revision() -> None:
    out = build_redline_docx("a b c", "a x c", author="Reviewer", date=FIXED)
    with zipfile.ZipFile(io.BytesIO(out)) as z:
        xml = z.read("word/document.xml").decode()
    assert xml.count('w:author="Reviewer"') == 2
    assert xml.count('w:date="2026-01-01T00:00:00Z"') == 2


def test_line_break_inside_a_paragraph_round_trips_as_a_break() -> None:
    old = "Row one\nRow two"
    new = "Row one\nRow 2"
    out = build_redline_docx(old, new)
    assert extract_text(out) == new  # accepted view of the redline is the new text
    with zipfile.ZipFile(io.BytesIO(out)) as z:
        assert b"<w:br/>" in z.read("word/document.xml")


def test_extract_text_joins_paragraphs_and_skips_empty_ones() -> None:
    assert extract_text(_docx("First.", "", "  ", "Second.")) == "First.\n\nSecond."


def test_extract_text_accepts_pending_tracked_changes() -> None:
    body = (
        '<w:p><w:r><w:t xml:space="preserve">The fee is </w:t></w:r>'
        '<w:del w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>£500</w:delText></w:r></w:del>'
        '<w:ins w:id="2" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:t>£750</w:t></w:r></w:ins>'
        "<w:r><w:t>.</w:t></w:r></w:p>"
    )
    assert extract_text(_with_body_xml(body)) == "The fee is £750."


def test_extract_text_reads_table_cells_as_paragraphs() -> None:
    doc = Document()
    doc.add_paragraph("Before.")
    table = doc.add_table(rows=1, cols=2)
    table.cell(0, 0).text = "Left cell"
    table.cell(0, 1).text = "Right cell"
    buf = io.BytesIO()
    doc.save(buf)
    assert extract_text(buf.getvalue()) == "Before.\n\nLeft cell\n\nRight cell"


def test_redline_docx_compares_two_files() -> None:
    out = redline_docx(_docx("Clause 1. The fee is £500."), _docx("Clause 1. The fee is £750."), date=FIXED)
    assert read_tracked_changes(out) == (["750"], ["500"])


def test_redline_of_a_document_with_its_own_tracked_changes_diffs_the_accepted_text() -> None:
    body = (
        '<w:p><w:r><w:t xml:space="preserve">Notice is </w:t></w:r>'
        '<w:ins w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:t>thirty</w:t></w:r></w:ins>'
        '<w:del w:id="2" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>ten</w:delText></w:r></w:del>'
        '<w:r><w:t xml:space="preserve"> days.</w:t></w:r></w:p>'
    )
    out = redline_docx(_with_body_xml(body), _docx("Notice is sixty days."))
    assert read_tracked_changes(out) == (["sixty"], ["thirty"])


def test_cli_writes_a_redline(tmp_path: Path) -> None:
    old, new, out = tmp_path / "old.docx", tmp_path / "new.docx", tmp_path / "out.docx"
    old.write_bytes(_docx("The term is one year."))
    new.write_bytes(_docx("The term is two years."))
    cli_main([str(old), str(new), "-o", str(out), "--author", "CLI"])
    ins, dele = read_tracked_changes(out.read_bytes())
    assert ins and dele


def test_original_view_rejects_pending_changes() -> None:
    body = (
        '<w:p><w:r><w:t xml:space="preserve">The fee is </w:t></w:r>'
        '<w:del w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>£500</w:delText></w:r></w:del>'
        '<w:ins w:id="2" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:t>£750</w:t></w:r></w:ins>'
        "<w:r><w:t>.</w:t></w:r></w:p>"
    )
    assert extract_text(_with_body_xml(body), view="original") == "The fee is £500."


def test_accepting_and_rejecting_a_redline_reproduce_both_inputs() -> None:
    old = "Clause 1. The fee is £500.\n\nClause 2. Notice is thirty days.\n\nClause 3. Unchanged."
    new = "Clause 1. The fee is £750.\n\nClause 2a. New clause added here.\n\nClause 3. Unchanged."
    out = build_redline_docx(old, new)
    assert extract_text(out) == new
    assert extract_text(out, view="original") == old


def test_deleted_paragraph_leaves_no_empty_paragraph_when_accepted() -> None:
    old = "Clause 1. Kept.\n\nClause 2. Removed entirely.\n\nClause 3. Kept."
    new = "Clause 1. Kept.\n\nClause 3. Kept.\n\nClause 4. Added entirely."
    out = build_redline_docx(old, new)
    assert extract_paragraphs(out) == new.split("\n\n")
    assert extract_paragraphs(out, view="original") == old.split("\n\n")


def test_paragraph_with_a_deleted_mark_runs_on_into_the_next() -> None:
    body = (
        '<w:p><w:pPr><w:rPr><w:del w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr>'
        "<w:r><w:t>First half, </w:t></w:r></w:p>"
        "<w:p><w:r><w:t>second half.</w:t></w:r></w:p>"
    )
    docx = _with_body_xml(body)
    assert extract_paragraphs(docx) == ["First half, second half."]
    assert extract_paragraphs(docx, view="original") == ["First half, ", "second half."]


def test_replaced_paragraph_does_not_gain_a_line_break() -> None:
    old = "Clause 1. Kept.\n\nThe Tenant shall pay the rent quarterly.\n\nClause 3. Kept."
    new = "Clause 1. Kept.\n\nA counterpart clause about something else entirely.\n\nClause 3. Kept."
    out = build_redline_docx(old, new)
    assert extract_paragraphs(out) == new.split("\n\n")
    assert extract_paragraphs(out, view="original") == old.split("\n\n")


def _both_views(old: list[str], new: list[str]) -> tuple[list[str], list[str]]:
    out = build_redline_docx("\n\n".join(old), "\n\n".join(new))
    return extract_paragraphs(out), extract_paragraphs(out, view="original")


def test_a_lone_space_between_two_changes_does_not_split_the_paragraph() -> None:
    segs = [
        DiffSegment("equal", "Alpha."),
        DiffSegment("equal", "\n\n"),
        DiffSegment("insert", "Brand new clause."),
        DiffSegment("equal", "\n\n"),
        DiffSegment("equal", "The investors agree that the "),
        DiffSegment("delete", "Founder"),
        DiffSegment("insert", "Executive"),
        DiffSegment("equal", " "),
        DiffSegment("delete", "may"),
        DiffSegment("insert", "can"),
        DiffSegment("equal", " transfer."),
    ]
    out = render_segments_docx(segs)
    assert extract_paragraphs(out) == [
        "Alpha.",
        "Brand new clause.",
        "The investors agree that the Executive can transfer.",
    ]
    assert extract_paragraphs(out, view="original") == ["Alpha.", "The investors agree that the Founder may transfer."]


def test_a_break_removed_on_one_side_becomes_a_tracked_paragraph_mark() -> None:
    # One paragraph in the old text, split by a line break in the new one: the
    # old side's paragraph break and the new side's line break share a newline.
    old, new = ["First part.", "Second part."], ["First part.\nSecond part."]
    assert _both_views(old, new) == (new, old)
    assert _both_views(new, old) == (old, new)


def test_paragraph_split_and_join_round_trip() -> None:
    old = ["Clause one keeps going and then", "continues in a second paragraph.", "Clause two."]
    new = ["Clause one keeps going and then continues in a second paragraph.", "Clause two.", "Clause three."]
    assert _both_views(old, new) == (new, old)
    assert _both_views(new, old) == (old, new)


def test_random_edits_always_accept_and_reject_to_both_versions() -> None:
    """Accepting every change must give the new paragraphs and rejecting every
    change the old ones, for any edit. The earlier line-based grouping failed
    this for roughly a quarter of these pairs."""
    import random

    words = [
        "the",
        "Supplier",
        "shall",
        "pay",
        "Customer",
        "fee",
        "within",
        "thirty",
        "days",
        "of",
        "invoice",
        "and",
        "may",
        "terminate",
        "notice",
    ]

    def para(rng: random.Random) -> str:
        return " ".join(rng.choice(words) for _ in range(rng.randint(1, 8))) + rng.choice([".", "", ","])

    for seed in range(400):
        rng = random.Random(seed)
        old = [para(rng) for _ in range(rng.randint(1, 6))]
        new = list(old)
        for _ in range(rng.randint(1, 4)):
            op = rng.choice(["ins", "del", "rep", "edit", "br"])
            i = rng.randrange(len(new) + (op == "ins"))
            if op == "ins":
                new.insert(i, para(rng))
            elif op == "del" and len(new) > 1:
                del new[i]
            elif op == "rep" and new:
                new[i] = para(rng)
            elif op == "edit" and new:
                w = new[i].split(" ")
                w[rng.randrange(len(w))] = rng.choice(words)
                new[i] = " ".join(w)
            elif op == "br" and new:
                new[i] = new[i] + "\n" + para(rng)
        assert _both_views(old, new) == (new, old), f"seed {seed}"


def test_deleting_or_inserting_the_last_or_first_paragraph_leaves_no_empty_one() -> None:
    """Segment lists where the separator sits outside the changed paragraph,
    so one view ends (or starts) with newlines. Those are not a break."""
    cases = [
        (
            [DiffSegment("equal", "Keep."), DiffSegment("equal", "\n\n"), DiffSegment("delete", "Gone.")],
            ["Keep."],
            ["Keep.", "Gone."],
        ),
        (
            [DiffSegment("equal", "Keep."), DiffSegment("equal", "\n\n"), DiffSegment("insert", "New.")],
            ["Keep.", "New."],
            ["Keep."],
        ),
        (
            [DiffSegment("delete", "Gone."), DiffSegment("equal", "\n\n"), DiffSegment("equal", "Keep.")],
            ["Keep."],
            ["Gone.", "Keep."],
        ),
        (
            [DiffSegment("insert", "New."), DiffSegment("equal", "\n\n"), DiffSegment("equal", "Keep.")],
            ["New.", "Keep."],
            ["Keep."],
        ),
    ]
    for segs, accepted, original in cases:
        out = render_segments_docx(segs)
        assert extract_paragraphs(out) == accepted, segs
        assert extract_paragraphs(out, view="original") == original, segs


def test_any_valid_segment_list_renders_to_both_versions() -> None:
    """render_segments_docx is public, so it must handle segment lists that
    compute_redline_segments would never produce: separators as their own
    equal, deleted or inserted segment, or split across two segments."""
    import random

    for seed in range(600):
        rng = random.Random(seed)
        segs: list[DiffSegment] = []
        old: list[str] = []
        new: list[str] = []
        had_old = had_new = False
        for k in range(rng.randint(1, 6)):
            op = rng.choice(["keep", "del", "ins", "edit"])
            text = f"P{k} " + " ".join(rng.choice(["alpha", "beta", "gamma"]) for _ in range(rng.randint(1, 4)))
            in_old, in_new = op != "ins", op != "del"
            if (had_old and in_old) or (had_new and in_new):
                both = had_old and in_old and had_new and in_new
                kind = "equal" if both else ("delete" if had_old and in_old else "insert")
                if kind == "equal" and rng.random() < 0.3:
                    segs += [DiffSegment("equal", "\n"), DiffSegment("equal", "\n")]
                else:
                    segs.append(DiffSegment(kind, "\n\n"))  # type: ignore[arg-type]
            if op == "keep":
                segs.append(DiffSegment("equal", text))
                old.append(text)
                new.append(text)
            elif op == "del":
                segs.append(DiffSegment("delete", text))
                old.append(text)
            elif op == "ins":
                segs.append(DiffSegment("insert", text))
                new.append(text)
            else:
                segs += [DiffSegment("equal", text + " "), DiffSegment("delete", "old"), DiffSegment("insert", "new")]
                old.append(text + " old")
                new.append(text + " new")
            had_old, had_new = had_old or in_old, had_new or in_new
        out = render_segments_docx(segs)
        accepted, original = extract_paragraphs(out), extract_paragraphs(out, view="original")
        # A Word document always has one paragraph, so an empty version is [""].
        assert (accepted if accepted != [""] else []) == new, f"seed {seed}"
        assert (original if original != [""] else []) == old, f"seed {seed}"
