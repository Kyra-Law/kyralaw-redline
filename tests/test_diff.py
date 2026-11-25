"""Tests for kyralaw_redline.diff.

The regression tests each pin one heuristic in diff.py. Their text is
invented, and each was checked to fail with its heuristic disabled (see the
note on each).
"""

from __future__ import annotations

from kyralaw_redline.diff import (
    DiffSegment,
    added_text,
    align_paragraphs,
    compute_redline_segments,
    diff_paragraph,
    removed_text,
)


def _equal_text(segments: list[DiffSegment]) -> str:
    return "".join(s.text for s in segments if s.kind == "equal")


def test_identical_text_is_all_equal() -> None:
    segments = diff_paragraph("The rent is £1,000 per month.", "The rent is £1,000 per month.")
    assert all(s.kind == "equal" for s in segments)


def test_single_word_edit_only_redlines_that_word() -> None:
    segments = diff_paragraph(
        "This Agreement is dated 4 September 2023 between the parties.",
        "This Agreement is dated 4 September 2026 between the parties.",
    )
    assert removed_text(segments) == "2023"
    assert added_text(segments) == "2026"
    kinds = [s.kind for s in segments if s.text]
    assert kinds == ["equal", "delete", "insert", "equal"]


def test_pure_insertion_and_deletion() -> None:
    ins = diff_paragraph("", "Newly added clause.")
    assert added_text(ins) == "Newly added clause."
    assert removed_text(ins) == ""

    dele = diff_paragraph("Clause being removed entirely.", "")
    assert removed_text(dele) == "Clause being removed entirely."
    assert added_text(dele) == ""


def test_reconstructs_old_and_new_text_exactly() -> None:
    old_text = "Clause 1. The fee is £500.\n\nClause 2. Notice is thirty days.\n\nClause 3. Unchanged."
    new_text = "Clause 1. The fee is £750.\n\nClause 2a. New clause added here.\n\nClause 3. Unchanged."
    segments = compute_redline_segments(old_text, new_text)
    assert "".join(s.text for s in segments if s.kind != "insert") == old_text
    assert "".join(s.text for s in segments if s.kind != "delete") == new_text


def test_added_text_does_not_jam_together_non_adjacent_insertions() -> None:
    """Two insert segments separated by unchanged text are not adjacent in
    the document, so added_text() must not join them into one word."""
    segments = diff_paragraph(
        "the Licensor shall grant to the Founder 40 options",
        "the Director shall transfer to the Founder 25 options",
    )
    assert added_text(segments).split() == ["Director", "transfer", "25"]
    assert removed_text(segments).split() == ["Licensor", "grant", "40"]


def test_does_not_merge_a_genuinely_correct_short_match() -> None:
    """A short common word between two edits is often a real match: "shall"
    is unchanged in "Licensor shall grant" -> "Director shall transfer". A
    merge is kept only when it is smaller, which it is not here."""
    segments = diff_paragraph(
        "the Licensor shall grant to the Founder 40 options",
        "the Director shall transfer to the Founder 25 options",
    )
    assert " shall " in _equal_text(segments)
    assert removed_text(segments).split() == ["Licensor", "grant", "40"]


def test_small_edit_next_to_a_new_paragraph_does_not_confuse_the_small_edit() -> None:
    """Two small edits with a whole new paragraph between them land in one
    line-level replace opcode. Word-diffing that block as one paragraph lets
    the new paragraph's vocabulary match the wrong words. Aligning blocks
    first keeps each edit to its own paragraph.

    The merge-and-trim heuristics hide most of this damage downstream, so the
    visible symptom here is small: without alignment the first paragraph's
    full stop is reported deleted although it was kept. Fails if diff_blocks
    word-diffs the concatenated block instead of aligning.
    """
    old_text = (
        "Alpha clause about delivery of goods to the Buyer.\n"
        "The Supplier shall deliver the Goods to the Buyer within 14 days (or 21 days) of the order.\n"
    )
    new_text = (
        "Alpha clause about delivery of goods to the Buyer, as amended.\n"
        "New clause: the Buyer shall inspect the Goods within 14 days of delivery and notify the Supplier "
        "of any defect.\n"
        "The Supplier shall deliver the Goods to the Buyer within 14 days of the order.\n"
    )
    segments = compute_redline_segments(old_text, new_text)
    assert removed_text(segments) == "(or 21 days) "
    assert ", as amended" in added_text(segments)
    assert "inspect the Goods" in added_text(segments)


def test_merges_across_a_wrongly_anchored_common_word() -> None:
    """'the "Licensor" or the "Owner"' -> 'the "Owner"' (plus a curly to
    straight quote change). A third "the" later in the sentence lets
    SequenceMatcher anchor the middle "the" to the wrong occurrence; the
    merge-and-trim in _render_run collapses it back to one edit.

    Fails if _group_runs never groups across a sandwiched equal.
    """
    segments = diff_paragraph(
        "Northwind Ltd shall be referred to as the “Licensor” or the “Owner”, "
        "and Sam Doe shall be referred to as the “Licensee”.",
        'Northwind Ltd shall be referred to as the "Owner", and Sam Doe shall be referred to as the “Licensee”.',
    )
    assert removed_text(segments) == "“Licensor” or the “Owner”,"
    assert added_text(segments) == '"Owner",'


def test_merges_across_a_wrongly_anchored_repeated_phrase() -> None:
    """ "thirty (30) days" appears twice; the first is changed, the second is
    not. SequenceMatcher anchors the second occurrence, so without a wide
    enough sandwiched-equal cap the real edit renders as unchanged.

    Fails with _MAX_SANDWICHED_EQUAL_TOKENS = 1.
    """
    segments = diff_paragraph(
        "notice of termination of thirty (30) days to the Supplier.  On such termination, "
        "the Supplier shall refund the sums described in Clause 9.1 above within thirty (30) days of such "
        "termination.",
        "notice of termination of [NUMBER] days to the Customer.  On such termination, "
        "the Customer shall refund the sums described in Clause 9.2 above within thirty (30) days of such "
        "termination.",
    )
    assert (
        removed_text(segments)
        == "thirty (30) days to the Supplier.  On such termination, the Supplier shall refund the sums described in "
        "Clause 9.1"
    )
    assert (
        added_text(segments)
        == "[NUMBER] days to the Customer.  On such termination, the Customer shall refund the sums described in "
        "Clause 9.2"
    )
    assert "above within thirty (30) days of such termination" in _equal_text(segments)


def test_does_not_merge_across_a_genuinely_long_unchanged_clause() -> None:
    """ "(or Director)" is deleted twice, independently, with 13 unchanged
    tokens between. The cap on sandwiched equals stops the run before that
    long equal, keeping two clean deletions.

    Fails with an uncapped _MAX_SANDWICHED_EQUAL_TOKENS.
    """
    segments = diff_paragraph(
        "The Founder (or Director) shall be a Departing Holder if employment or service is terminated by the "
        "Company for Cause, or if the Founder (or Director) resigns and within 12 months of the termination date:\n",
        "The Founder  shall be a Departing Holder if employment or service is terminated by the "
        "Company for Cause, or if the Founder  resigns and within 12 months of the termination date:\n",
    )
    assert removed_text(segments).split() == ["(or", "Director)", "(or", "Director)"]
    assert added_text(segments).strip() == ""
    assert (
        "shall be a Departing Holder if employment or service is terminated by the Company for Cause, or if the Founder"
        in _equal_text(segments)
    )


def test_wrongly_anchored_sentence_final_period_still_gets_deleted() -> None:
    """The old paragraph's lone trailing "." matches the "." inside "3.8" in
    the much longer new sentence. Without the low-information checks the
    paragraph shows no deletion at all, although it lost its full stop.

    Fails if _trim_common_affix and _render_run drop their
    low-information checks.
    """
    segments = diff_paragraph(
        "(b) the Licensee's general commercial use or product development.\n",
        "(b) the Licensee's general commercial use or product development; provided, however, "
        "that nothing in this Section 3.8 shall prevent the Licensee from using generalised, "
        "non-customer-specific learnings.\n",
    )
    assert removed_text(segments) == ".\n"
    assert added_text(segments).startswith("; provided, however,")


def test_an_edited_paragraph_pairs_across_a_blank_one() -> None:
    """Old [clause, blank], new [blank, edited clause]: pairs keep their order, so the
    clauses and the blanks cannot both pair. The clauses must, or the edited clause is
    deleted whole and inserted whole. Fails if two blanks score as much as text."""
    old = ["The Supplier shall deliver the Goods to the Premises within thirty days of the Order.", ""]
    new = ["", "The Supplier shall deliver the Goods to the Premises within sixty days of the Order, at its own cost."]
    steps = align_paragraphs(old, new)
    assert [(s.old_index, s.new_index) for s in steps if s.old_index is not None and s.new_index is not None] == [
        (0, 1)
    ]


def test_short_trim_at_a_genuine_boundary_is_not_rejected() -> None:
    """A trim ending in a lone punctuation character is only suspect when it
    swallows the entire old side. "2023." -> "2026." keeps "." as equal."""
    segments = diff_paragraph(
        "This Agreement is dated 4 September 2023.",
        "This Agreement is dated 4 September 2026.",
    )
    assert removed_text(segments) == "2023"
    assert added_text(segments) == "2026"
    assert segments[-1] == DiffSegment("equal", ".")


def test_many_small_edits_around_short_equals_still_prefers_naive() -> None:
    """A long paragraph rewritten with several small edits contains short
    low-information equals (", ") between them purely by grammar. Preferring
    the merged rendering whenever any such equal exists would lose the long
    unchanged phrases inside the paragraph; the check has to be run-wide.

    Fails if _render_run prefers merged whenever the naive rendering has any
    low-information equal.
    """
    old = (
        "Hosting and location. The service runs on a managed cloud provider with workloads scheduled "
        "on containers, and live and test environments separated into distinct accounts. Customer data is "
        "held and processed in the region chosen at signup. The default region is Canada; Australian and "
        "Japanese hosting are offered on request, and Brazilian hosting may be offered for an extra fee "
        "where agreed. Records kept inside the service, including indexed attachments and audit "
        "trails, stay within the chosen region. For Linked Sources, records stay in the original "
        "Connected Repository.\n"
    )
    new = (
        "Hosting and location. Customer data is held and processed in New Zealand, including records "
        "kept inside the service, indexed attachments and audit trails. For Linked Sources, records "
        "stay in the original Connected Repository.\n"
    )
    segments = diff_paragraph(old, new)
    equal = _equal_text(segments)
    assert "indexed attachments and audit trails" in equal
    assert "For Linked Sources, records stay in the original Connected Repository" in equal


def test_whitespace_tokens_do_not_spuriously_match_across_unrelated_text() -> None:
    """Without whitespace as junk, single spaces in unrelated text anchor
    matches and shred one edit into alternating fragments, and the deleted
    text comes back with its words jammed together."""
    old_text = "[DATE OF UPDATE] [DATE OF BOARD RESOLUTION]"
    new_text = "some unrelated new text with spaces and words scattered around for testing"
    segments = diff_paragraph(old_text, new_text)
    assert removed_text(segments) == old_text
    assert added_text(segments) == new_text


# A short line that recurs, such as a signature label, must never anchor the
# line-level pass across the document. The shape that breaks it: the
# recurring line sits early in the old version and late in
# both, and every other matching line is unique and on its own, so the longest
# matches are all single lines and SequenceMatcher picks the earliest, the
# recurring one, pairing it with the late copy. Everything between then came
# out as a wholesale deletion plus a wholesale insertion (about 480 words each
# way here, against a true change of 20 words). Fails without _line_opcodes'
# unique-line anchoring.
_RECURRING = "Signed by:"


def _recurring_line_doc(edited: bool) -> str:
    paras = ["SUPPLY AGREEMENT"]
    if not edited:
        paras.append(_RECURRING)
    for k in range(1, 21):
        paras.append(f"{k}.1 The Supplier keeps the register for lot {k} at its principal office.")
        verb = "must" if edited else "shall"
        paras.append(f"{k}.2 The Buyer {verb} inspect lot {k} within the inspection window.")
        if not edited and k == 10:
            paras.append(_RECURRING)
    paras += [_RECURRING, "Execution page"]
    return "\n\n".join(paras) + "\n"


def test_a_recurring_line_does_not_anchor_across_the_document() -> None:
    old_text, new_text = _recurring_line_doc(False), _recurring_line_doc(True)
    segments = compute_redline_segments(old_text, new_text)
    assert "".join(s.text for s in segments if s.kind != "insert") == old_text
    assert "".join(s.text for s in segments if s.kind != "delete") == new_text
    assert added_text(segments).split() == ["must"] * 20
    assert removed_text(segments).split().count("shall") == 20
    assert len(removed_text(segments).split()) == 24  # the 20 verbs and the two dropped labels
    # Every unchanged clause stays unchanged.
    for k in range(1, 21):
        assert f"{k}.1 The Supplier keeps the register for lot {k} at its principal office." in _equal_text(segments)


def test_a_recurring_line_still_matches_within_its_own_region() -> None:
    # The same label recurring in both versions, with a one-word edit beside
    # each copy: no copy is unique, so none anchors, and each must still match
    # its own counterpart inside the gap it sits in.
    old_text = "Part A\n\nSigned by:\n\nThe Buyer shall pay.\n\nPart B\n\nSigned by:\n\nThe Seller shall deliver.\n"
    new_text = "Part A\n\nSigned by:\n\nThe Buyer must pay.\n\nPart B\n\nSigned by:\n\nThe Seller must deliver.\n"
    segments = compute_redline_segments(old_text, new_text)
    assert added_text(segments).split() == ["must", "must"]
    assert removed_text(segments).split() == ["shall", "shall"]
