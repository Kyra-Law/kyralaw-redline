"""Word-granularity text diff for legal redlines.

A line-level diff (``difflib.unified_diff`` and friends) redlines a whole
paragraph as deleted and reinserted whenever a single word inside it changes.
That is unreadable for contracts, where most edits are a word or a phrase
inside an otherwise unchanged clause. This module fixes the granularity with
a two-pass, pure-Python diff (standard library only):

1. A line-level pass first, so wholesale paragraph insertions and deletions
   stay whole, which is how Word's own Compare Documents reports them. It
   anchors on lines unique to both versions (the idea behind patience diff)
   and uses ``difflib.SequenceMatcher`` only between those anchors, so a
   recurring line cannot pair two distant places (see :func:`_line_opcodes`).
2. Within each line-level "replace" block, paragraphs are aligned by
   similarity and each aligned pair gets a token-level ``SequenceMatcher``
   pass that finds the words that actually changed, leaving the surrounding
   text as plain equal segments.

Tokenisation is whitespace-preserving (``"".join(_tokenize(text)) == text``),
so the segment list reconstructs both sides exactly and can be rendered as
``w:ins``/``w:del`` runs in a real .docx (see :mod:`kyralaw_redline.docx`).

``SequenceMatcher`` does not compute a minimal diff, and legal prose repeats
short words and phrases ("the", "thirty (30) days") often enough that it
regularly anchors a match to the wrong occurrence. Most of the code below is
a small number of bounded heuristics that correct for that. Each one was
introduced for a failure seen on a real agreement, and each docstring records
the failure with invented text of the same shape.
"""

from __future__ import annotations

import difflib
import re
from dataclasses import dataclass
from typing import Literal

SegmentKind = Literal["equal", "insert", "delete"]

_TOKEN_RE = re.compile(r"\s+|\w+|[^\s\w]+", re.UNICODE)

# The two-pass diff can hit large all-different blocks (a whole paragraph
# replaced by unrelated new content). Token-level SequenceMatcher.get_opcodes()
# is O(n*m) in the worst case, so a giant replace block is marked whole rather
# than paying for a token diff unlikely to find real overlap.
_MAX_TOKENS_FOR_WORD_DIFF = 4000


@dataclass(frozen=True)
class DiffSegment:
    """One span of the redline: unchanged, inserted, or deleted text.

    ``text`` is verbatim (whitespace included) so segments reconstruct
    losslessly: concatenating ``equal`` + ``delete`` segments reproduces the
    old text, and ``equal`` + ``insert`` reproduces the new text.
    """

    kind: SegmentKind
    text: str


def _tokenize(text: str) -> list[str]:
    return _TOKEN_RE.findall(text)


def _is_junk_token(token: str) -> bool:
    """Whitespace-only tokens are near-indistinguishable from each other.

    Without this, SequenceMatcher's matching aligns a space in one changed
    phrase with an unrelated space elsewhere in the paragraph, fragmenting one
    real edit into many alternating tiny replace/equal opcodes. Marking them
    as junk excludes them from anchoring a match, so a genuinely different
    span stays one clean replace block instead of shredding on every space.

    Deliberately whitespace-only, not punctuation too. Treating lone
    punctuation as junk was tried, to stop a period anchoring a spurious
    match, and it broke a different document: making "(" and ")" invisible
    to the matcher removed exactly the context that anchored "(or Director)"
    as one unit, so the matcher latched onto a bare "or" later in the
    sentence instead and scrambled two clean edits. Punctuation carries real
    structural signal for this pass; only whitespace is interchangeable
    enough to always ignore. The period-anchoring problem is handled at the
    trim stage instead (:func:`_trim_common_affix`), where there is more
    context to tell the cases apart.
    """
    return token.isspace()


def _is_junk_line(line: str) -> bool:
    """Same reasoning as :func:`_is_junk_token`, for blank lines in the
    line-level pass."""
    return line.strip() == ""


_Opcode = tuple[str, int, int, int, int]

# An "equal" run at or under this length (in real, non-whitespace tokens) is
# worth *trying* as part of a wider merge across the two edits it sits
# between; see _group_runs's docstring for the cases this catches. Above it, a
# sandwiched equal is a hard boundary between separate runs.
#
# This has to be a bounded cap, neither "any length" nor "1 token"; both were
# tried and both broke a real document:
#   - Uncapped: a long, genuinely unchanged clause (13 tokens) got swept into
#     the same run as the two independent edits either side of it, so the
#     run's naive-vs-merged comparison was between two bad options instead of
#     the clean "edit, long equal, edit" split a cap preserves.
#   - Capped at 1: too narrow to offer "thirty (30) days" (5 tokens) or "of
#     such" (2 tokens) as merge candidates, so a wrongly anchored multi-word
#     phrase was never reconsidered.
# 8 sits with margin on both sides of the known cases.
_MAX_SANDWICHED_EQUAL_TOKENS = 8


def _group_runs(opcodes: list[_Opcode], old_tokens: list[str]) -> list[_Opcode | list[_Opcode]]:
    """Split opcodes into standalone "equal" opcodes and "runs".

    A run is a maximal sequence of edit opcodes, possibly interspersed with
    short sandwiched equals, which :func:`diff_paragraph` considers merging
    into one wider replace.

    A sandwiched equal is often a common word or phrase that repeats nearby,
    and SequenceMatcher sometimes anchors it to the wrong occurrence. When one
    sits between two edits and is short enough
    (:data:`_MAX_SANDWICHED_EQUAL_TOKENS`), a wider merge across it is worth
    trying, but only kept if it is smaller once trimmed (see
    :func:`_render_run`), because the same shape (edit, short common text,
    edit) is just as often a genuinely correct match.

    Case 1, a repeated common word: 'the "Licensor" or the "Owner"' became
    'the "Owner"'. "the" appears a third time later in the sentence (before
    "Licensee"), so SequenceMatcher matched the wrong "the", spreading one
    small edit across a replace, a meaningless one-word equal, and a delete.
    Grouping the three into one run and trimming the merged span's common
    affix collapses it back to the four words that actually changed.

    Case 2, a repeated phrase: a termination clause has "thirty (30) days"
    twice, once genuinely changed ("thirty (30) days to the Supplier" became
    "[NUMBER] days to the Customer") and once, later, genuinely unchanged
    ("within thirty (30) days of such termination"). SequenceMatcher anchored
    the second occurrence, rendering the real edit as untouched and sweeping
    the unchanged occurrence into an unrelated delete. Offering "thirty (30)
    days" (five tokens) and "of such" (two tokens) as merge candidates lets
    the run merge and re-trim across both bridging equals.

    The flip side, which keeps this bounded: "The Founder (or Director) shall
    be treated as ... or if the Founder (or Director) resigns" deletes "(or
    Director)" twice, independently, with 13 genuinely unchanged tokens
    between the two deletions. Offering those as a merge candidate too would
    sweep both edits into one run whose merged rendering replaces the whole
    span, losing the clean two-edit split.
    """
    groups: list[_Opcode | list[_Opcode]] = []
    i, n = 0, len(opcodes)
    while i < n:
        if opcodes[i][0] == "equal":
            groups.append(opcodes[i])
            i += 1
            continue
        end = i + 1
        while end < n:
            if opcodes[end][0] != "equal":
                end += 1
                continue
            _, ei1, ei2, _, _ = opcodes[end]
            word_tokens = [t for t in old_tokens[ei1:ei2] if not t.isspace()]
            if len(word_tokens) <= _MAX_SANDWICHED_EQUAL_TOKENS and end + 1 < n and opcodes[end + 1][0] != "equal":
                end += 1
                continue
            break
        groups.append(opcodes[i:end])
        i = end
    return groups


_LONE_PUNCT_RE = re.compile(r"^[^\s\w]$", re.UNICODE)


def _is_low_information_token(token: str) -> bool:
    """Whitespace, or a single punctuation character (".", ",", ";" ...).

    Unlike :func:`_is_junk_token`, this does not steer SequenceMatcher itself
    (punctuation carries real signal for matching). It is used only later, by
    :func:`_trim_common_affix` and :func:`_render_run`, to decide whether a
    span is genuine content or a near-certain coincidental match.
    """
    return token.isspace() or bool(_LONE_PUNCT_RE.match(token))


def _trim_common_affix(old_seq: list[str], new_seq: list[str]) -> tuple[list[str], list[str], list[str], list[str]]:
    """Shrink a replace pair to its minimal differing core.

    A merged run (see :func:`_group_runs`), or even a single SequenceMatcher
    replace opcode, can have both sides share a leading or trailing run of
    tokens. Trimming that shared prefix and suffix into "equal", rather than
    a redundant delete plus insert, is what reduces a merged span back down
    to the words that actually differ.

    The trim is rejected outright if it would consume the *entire* old side
    and everything in the old side is low-information (see
    :func:`_is_low_information_token`), because then the only thing anchoring
    the trim is coincidental punctuation or whitespace. Without this, an old
    remainder of ".\\n" gets read as "shared" with the tail of whatever new
    sentence replaced it, since most sentences end that way. Seen when
    "...or product development." became "...or product development; provided,
    however, that nothing in this Section 3.8 shall ... learnings.": the old
    full stop was explained away as matching the new one, producing zero
    deleted text for a paragraph that genuinely lost its full stop.

    This is an all-or-nothing check on the whole old side, not "was this trim
    low-information" in isolation: trimming the trailing "." off "2023." vs
    "2026." is just as low-information, but "2023" is left over as real
    content, so that trim is correct and must stand.
    """
    start = 0
    limit = min(len(old_seq), len(new_seq))
    while start < limit and old_seq[start] == new_seq[start]:
        start += 1

    end_old, end_new = len(old_seq), len(new_seq)
    while end_old > start and end_new > start and old_seq[end_old - 1] == new_seq[end_new - 1]:
        end_old -= 1
        end_new -= 1

    if start == end_old and old_seq and all(_is_low_information_token(t) for t in old_seq):
        start, end_old, end_new = 0, len(old_seq), len(new_seq)

    return old_seq[:start], old_seq[start:end_old], new_seq[start:end_new], old_seq[end_old:]


def _replace_segments(old_span: list[str], new_span: list[str]) -> list[DiffSegment]:
    """A trimmed replace: equal prefix, delete, insert, equal suffix (empty
    parts omitted)."""
    prefix, old_mid, new_mid, suffix = _trim_common_affix(old_span, new_span)
    segments = []
    if prefix:
        segments.append(DiffSegment("equal", "".join(prefix)))
    if old_mid:
        segments.append(DiffSegment("delete", "".join(old_mid)))
    if new_mid:
        segments.append(DiffSegment("insert", "".join(new_mid)))
    if suffix:
        segments.append(DiffSegment("equal", "".join(suffix)))
    return segments


def _changed_length(segments: list[DiffSegment]) -> int:
    return sum(len(s.text) for s in segments if s.kind in ("delete", "insert"))


def _render_run(run: list[_Opcode], old_tokens: list[str], new_tokens: list[str]) -> list[DiffSegment]:
    """Render a run two ways and keep whichever changes less text.

    The two renderings are opcode by opcode (keeping any sandwiched equals)
    and merged into one span then trimmed to its minimal differing core.
    Picking by size tells a genuine match sandwiched in an edit (naive wins:
    merging would only widen the diff) from a wrongly anchored one (merged
    wins: it collapses once freed from the bad anchor), without having to
    know in advance which case it is.

    Merged is preferred outright, skipping the size comparison, when the
    run's entire old-side span is low-information (whitespace or single
    punctuation characters). Then nothing on the old side is real content,
    so any "equal" the naive rendering found there is a coincidence. This is
    needed in addition to the check in :func:`_trim_common_affix` because a
    low-information equal can be a whole opcode of its own and never reach
    that function. Seen with the "Section 3.8" example there: an interior
    "equal '.'" let the naive rendering reuse the old paragraph's final
    period to plug the "." inside "3.8", splitting one coherent insert into
    two fragments that do not concatenate back to the inserted text.

    The check is over the whole run's old span, not "does the naive
    rendering contain a low-information equal": a long paragraph rewritten
    with many small edits legitimately contains short low-information equals
    (a lone ", ") between genuine edits, with hundreds of real words
    elsewhere in the run. A per-equal check always preferred merged there,
    discarding a naive rendering that was both smaller and correctly kept
    several long unchanged phrases as equal.
    """
    naive: list[DiffSegment] = []
    for tag, i1, i2, j1, j2 in run:
        if tag == "equal":
            naive.append(DiffSegment("equal", "".join(old_tokens[i1:i2])))
        elif tag == "delete":
            naive.append(DiffSegment("delete", "".join(old_tokens[i1:i2])))
        elif tag == "insert":
            naive.append(DiffSegment("insert", "".join(new_tokens[j1:j2])))
        elif tag == "replace":
            naive.extend(_replace_segments(old_tokens[i1:i2], new_tokens[j1:j2]))

    if len(run) == 1 and run[0][0] != "replace":
        return naive  # nothing to merge against

    i1, j1 = run[0][1], run[0][3]
    i2, j2 = run[-1][2], run[-1][4]
    run_old_span = old_tokens[i1:i2]
    merged = _replace_segments(run_old_span, new_tokens[j1:j2])

    if run_old_span and all(_is_low_information_token(t) for t in run_old_span):
        return merged
    return merged if _changed_length(merged) < _changed_length(naive) else naive


def diff_paragraph(old_text: str, new_text: str) -> list[DiffSegment]:
    """Word-level diff of two paragraph-sized texts."""
    old_tokens = _tokenize(old_text)
    new_tokens = _tokenize(new_text)
    if len(old_tokens) * len(new_tokens) > _MAX_TOKENS_FOR_WORD_DIFF * _MAX_TOKENS_FOR_WORD_DIFF:
        segments: list[DiffSegment] = []
        if old_text:
            segments.append(DiffSegment("delete", old_text))
        if new_text:
            segments.append(DiffSegment("insert", new_text))
        return segments

    sm = difflib.SequenceMatcher(_is_junk_token, old_tokens, new_tokens, autojunk=False)
    groups = _group_runs(list(sm.get_opcodes()), old_tokens)
    segments = []
    for group in groups:
        if isinstance(group, list):
            segments.extend(_render_run(group, old_tokens, new_tokens))
        else:
            _, i1, i2, _, _ = group
            segments.append(DiffSegment("equal", "".join(old_tokens[i1:i2])))
    return segments


def compute_redline_segments(old_text: str, new_text: str) -> list[DiffSegment]:
    """Hierarchical redline diff: line-level structure, word-level precision.

    Returns an ordered list of segments. ``"".join(s.text for s in segments
    if s.kind != "insert")`` reconstructs ``old_text``, and
    ``"".join(s.text for s in segments if s.kind != "delete")`` reconstructs
    ``new_text``.
    """
    old_lines = old_text.splitlines(keepends=True)
    new_lines = new_text.splitlines(keepends=True)

    segments: list[DiffSegment] = []
    for tag, i1, i2, j1, j2 in _line_opcodes(old_lines, new_lines):
        if tag == "equal":
            segments.append(DiffSegment("equal", "".join(old_lines[i1:i2])))
        elif tag == "delete":
            segments.append(DiffSegment("delete", "".join(old_lines[i1:i2])))
        elif tag == "insert":
            segments.append(DiffSegment("insert", "".join(new_lines[j1:j2])))
        elif tag == "replace":
            for block_segments in diff_blocks(old_lines[i1:i2], new_lines[j1:j2]):
                segments.extend(block_segments)
    return segments


def _line_opcodes(old_lines: list[str], new_lines: list[str]) -> list[_Opcode]:
    """Line-level opcodes, anchored only on lines that occur once on each side.

    ``SequenceMatcher`` finds the longest matching run first and recurses either
    side of it, so a short line that recurs ("Signed by:", "Schedule", a
    repeated defined-term heading) can become the anchor that splits a document
    in the wrong place. Seen on a heavily negotiated agreement (2026-11-19):
    two one-line matches between copies of a repeated line paired old line 76
    with new line 344, so 300 lines of edited clauses came out as one
    wholesale deletion followed by 300 lines of wholesale insertion, and the
    word-level pass never saw a single pair to compare (score 0.23).

    This is the anchoring idea of patience diff: take the lines that are
    unique in both ``old_lines`` and ``new_lines``, keep the longest chain of
    them that is in order on both sides, and treat those as fixed points. Each
    gap between two fixed points is diffed the same way, recursively, and a
    gap with no unique lines left falls back to ``SequenceMatcher``, now on a
    small, bounded region. A recurring line can then only ever match within
    the gap it sits in, never across the document.

    Adjacent deletes and inserts are merged into one ``replace`` so the
    paragraph alignment in :func:`diff_blocks` sees each edited region whole.
    """
    raw: list[_Opcode] = []
    _patience(old_lines, new_lines, 0, len(old_lines), 0, len(new_lines), raw)
    merged: list[tuple[bool, int, int, int, int]] = []  # (is_equal, i1, i2, j1, j2)
    for tag, i1, i2, j1, j2 in raw:
        if i1 == i2 and j1 == j2:
            continue
        equal = tag == "equal"
        if merged and merged[-1][0] == equal:
            _, a1, _, b1, _ = merged[-1]
            merged[-1] = (equal, a1, i2, b1, j2)
        else:
            merged.append((equal, i1, i2, j1, j2))
    out: list[_Opcode] = []
    for equal, i1, i2, j1, j2 in merged:
        if equal:
            out.append(("equal", i1, i2, j1, j2))
        elif i1 == i2:
            out.append(("insert", i1, i2, j1, j2))
        elif j1 == j2:
            out.append(("delete", i1, i2, j1, j2))
        else:
            out.append(("replace", i1, i2, j1, j2))
    return out


def _patience(a: list[str], b: list[str], alo: int, ahi: int, blo: int, bhi: int, out: list[_Opcode]) -> None:
    """Append opcodes for ``a[alo:ahi]`` against ``b[blo:bhi]`` to ``out`` (see :func:`_line_opcodes`)."""
    # Common prefix and suffix are equal outright, whatever repeats inside them.
    while alo < ahi and blo < bhi and a[alo] == b[blo]:
        out.append(("equal", alo, alo + 1, blo, blo + 1))
        alo, blo = alo + 1, blo + 1
    tail: list[_Opcode] = []
    while alo < ahi and blo < bhi and a[ahi - 1] == b[bhi - 1]:
        tail.append(("equal", ahi - 1, ahi, bhi - 1, bhi))
        ahi, bhi = ahi - 1, bhi - 1

    anchors = _unique_anchors(a, b, alo, ahi, blo, bhi)
    if not anchors:
        sm = difflib.SequenceMatcher(_is_junk_line, a[alo:ahi], b[blo:bhi], autojunk=False)
        out.extend((tag, alo + i1, alo + i2, blo + j1, blo + j2) for tag, i1, i2, j1, j2 in sm.get_opcodes())
    else:
        i, j = alo, blo
        for ai, bj in anchors:
            _patience(a, b, i, ai, j, bj, out)
            out.append(("equal", ai, ai + 1, bj, bj + 1))
            i, j = ai + 1, bj + 1
        _patience(a, b, i, ahi, j, bhi, out)
    out.extend(reversed(tail))


def _unique_anchors(a: list[str], b: list[str], alo: int, ahi: int, blo: int, bhi: int) -> list[tuple[int, int]]:
    """Lines occurring exactly once in each range, longest chain in order on both sides."""
    count_a: dict[str, int] = {}
    for line in a[alo:ahi]:
        count_a[line] = count_a.get(line, 0) + 1
    count_b: dict[str, int] = {}
    where_b: dict[str, int] = {}
    for k in range(blo, bhi):
        count_b[b[k]] = count_b.get(b[k], 0) + 1
        where_b[b[k]] = k
    pairs = [
        (k, where_b[a[k]])
        for k in range(alo, ahi)
        if not _is_junk_line(a[k]) and count_a[a[k]] == 1 and count_b.get(a[k]) == 1
    ]
    if len(pairs) < 2:
        return pairs
    # Longest increasing subsequence of the new-side positions (patience sorting).
    tails: list[int] = []  # index into pairs of the smallest tail for each length
    back: list[int] = [-1] * len(pairs)
    for idx, (_, bj) in enumerate(pairs):
        lo, hi = 0, len(tails)
        while lo < hi:
            mid = (lo + hi) // 2
            if pairs[tails[mid]][1] < bj:
                lo = mid + 1
            else:
                hi = mid
        if lo > 0:
            back[idx] = tails[lo - 1]
        if lo == len(tails):
            tails.append(idx)
        else:
            tails[lo] = idx
    chain: list[tuple[int, int]] = []
    k = tails[-1]
    while k != -1:
        chain.append(pairs[k])
        k = back[k]
    chain.reverse()
    return chain


# Below this similarity two blocks are different content, not an edited version
# of one another, so they are reported as a plain delete plus a plain insert
# rather than word-diffed against each other. Word-diffing unrelated paragraphs
# produces churn: every token of the old one comes out deleted and every token
# of the new one inserted, plus spurious "equal" runs wherever the two happen
# to share common words.
#
# Chosen from a sweep over a private corpus of 32 real document pairs, scored
# against Word's Compare output. The score was flat within 0.15 of a point
# from 0.05 to 0.35 and fell sharply above that (86 at 0.7), so the gain comes
# from aligning at all rather than from this exact number; 0.3 is the round
# middle of the flat band.
_ALIGN_MIN_RATIO = 0.3

# Guard on the O(n*m) alignment below. A replace opcode covering a wholesale
# rewrite can hold hundreds of paragraphs on each side, and the pairwise
# ratios are the expensive part. Past this many cells we fall back to
# positional pairing: a worse alignment, not a wrong one.
_ALIGN_MAX_CELLS = 40_000


def _block_similarity(old_block: str, new_block: str) -> float:
    """How alike two blocks are, 0..1, cheap-to-expensive with early outs.

    real_quick_ratio and quick_ratio are upper bounds, so a value below the
    floor rules the pair out without running the full comparison.
    """
    if not old_block.strip() or not new_block.strip():
        return 1.0 if old_block.strip() == new_block.strip() else 0.0
    sm = difflib.SequenceMatcher(None, old_block, new_block, autojunk=False)
    if sm.real_quick_ratio() < _ALIGN_MIN_RATIO or sm.quick_ratio() < _ALIGN_MIN_RATIO:
        return 0.0
    return sm.ratio()


def _align_blocks(old_blocks: list[str], new_blocks: list[str]) -> list[tuple[int | None, int | None]]:
    """Pair blocks by similarity, in order, allowing gaps on either side.

    Returns one entry per output block: ``(i, j)`` for a matched pair,
    ``(i, None)`` for a block only in the old text, ``(None, j)`` for one only
    in the new.

    Positional pairing, the obvious alternative, fails when a paragraph is
    inserted or removed in the *middle* of a replaced region: every later
    pairing shifts by one, so each subsequent paragraph is diffed against its
    neighbour rather than itself, and every mispairing becomes a whole
    paragraph delete plus a whole paragraph insert.

    Ordering is preserved deliberately: documents are sequences, so an
    alignment that lets pairs cross would match a clause to a later
    restatement of it and report the intervening text as changed. This is
    therefore a longest-common-subsequence over similarity rather than a free
    assignment.
    """
    n, m = len(old_blocks), len(new_blocks)
    if n * m > _ALIGN_MAX_CELLS:
        paired = min(n, m)
        out: list[tuple[int | None, int | None]] = [(k, k) for k in range(paired)]
        out.extend((k, None) for k in range(paired, n))
        out.extend((None, k) for k in range(paired, m))
        return out

    # score[i][j] = best total similarity aligning old_blocks[i:] with new_blocks[j:].
    score = [[0.0] * (m + 1) for _ in range(n + 1)]
    for i in range(n - 1, -1, -1):
        for j in range(m - 1, -1, -1):
            ratio = _block_similarity(old_blocks[i], new_blocks[j])
            match = (ratio + score[i + 1][j + 1]) if ratio >= _ALIGN_MIN_RATIO else -1.0
            score[i][j] = max(match, score[i + 1][j], score[i][j + 1])

    pairs: list[tuple[int | None, int | None]] = []
    i = j = 0
    while i < n and j < m:
        ratio = _block_similarity(old_blocks[i], new_blocks[j])
        match = (ratio + score[i + 1][j + 1]) if ratio >= _ALIGN_MIN_RATIO else -1.0
        if match >= score[i + 1][j] and match >= score[i][j + 1]:
            pairs.append((i, j))
            i, j = i + 1, j + 1
        elif score[i + 1][j] >= score[i][j + 1]:
            pairs.append((i, None))
            i += 1
        else:
            pairs.append((None, j))
            j += 1
    pairs.extend((k, None) for k in range(i, n))
    pairs.extend((None, k) for k in range(j, m))
    return pairs


def diff_blocks(old_blocks: list[str], new_blocks: list[str]) -> list[list[DiffSegment]]:
    """Diff two runs of blocks (paragraphs or lines), one segment list per
    output block.

    A line-level "replace" opcode is not always one old paragraph to one new
    one: a small edit and a wholesale new paragraph inserted right next to it
    land in the same opcode. Concatenating the block and word-diffing it as
    one paragraph lets the new paragraph's vocabulary confuse the search for
    the real, small edit (shared words spuriously "match"). Aligning blocks
    (see :func:`_align_blocks`) and diffing each pair on its own keeps the two
    changes independent; unaligned blocks are pure deletes or inserts.

    Seen on a signature block that gained a heading partway through:
    positional pairing matched "EXAMPLE HOLDINGS LIMITED:" with the new
    "COMPANY" heading rather than with the later "Signed for and on behalf of
    EXAMPLE HOLDINGS LIMITED:" it actually corresponds to.
    """
    result: list[list[DiffSegment]] = []
    for i, j in _align_blocks(old_blocks, new_blocks):
        if i is not None and j is not None:
            result.append(diff_paragraph(old_blocks[i], new_blocks[j]))
        elif i is not None:
            if old_blocks[i]:
                result.append([DiffSegment("delete", old_blocks[i])])
        elif j is not None and new_blocks[j]:
            result.append([DiffSegment("insert", new_blocks[j])])
    return result


@dataclass(frozen=True)
class ParagraphAlignment:
    """One step of a paragraph-level alignment.

    ``old_index`` and ``new_index`` point into the two paragraph lists; either is
    ``None`` for a paragraph that exists on one side only. ``segments`` is the
    word diff of a paired paragraph whose text changed, and ``None`` for an
    unchanged pair or a one-sided paragraph.
    """

    old_index: int | None
    new_index: int | None
    segments: list[DiffSegment] | None


def align_paragraphs(old_paragraphs: list[str], new_paragraphs: list[str]) -> list[ParagraphAlignment]:
    """Align two documents paragraph by paragraph, with a word diff per edited pair.

    The same passes as :func:`compute_redline_segments`, over paragraphs rather
    than the lines of one string: :func:`_line_opcodes`, anchored on paragraphs
    unique to both sides, then :func:`_align_blocks` to pair the paragraphs of
    each replaced region by similarity, then :func:`diff_paragraph` within each
    pair. It returns which paragraph is which instead of a flat stream, so the
    in-place renderer (:mod:`kyralaw_redline.inplace`), which writes into the
    documents themselves, can put each change in the paragraph it belongs to.
    Sharing the passes, rather than a second alignment there, keeps the in-place
    and rebuild renderers measuring the same diff.
    """
    steps: list[ParagraphAlignment] = []
    for tag, i1, i2, j1, j2 in _line_opcodes(old_paragraphs, new_paragraphs):
        if tag == "equal":
            steps.extend(ParagraphAlignment(i1 + k, j1 + k, None) for k in range(i2 - i1))
        elif tag == "delete":
            steps.extend(ParagraphAlignment(i, None, None) for i in range(i1, i2))
        elif tag == "insert":
            steps.extend(ParagraphAlignment(None, j, None) for j in range(j1, j2))
        else:
            for i, j in _align_blocks(old_paragraphs[i1:i2], new_paragraphs[j1:j2]):
                old_index = None if i is None else i1 + i
                new_index = None if j is None else j1 + j
                segments = None
                if old_index is not None and new_index is not None:
                    diff = diff_paragraph(old_paragraphs[old_index], new_paragraphs[new_index])
                    segments = None if all(s.kind == "equal" for s in diff) else diff
                steps.append(ParagraphAlignment(old_index, new_index, segments))
    return steps


def added_text(segments: list[DiffSegment]) -> str:
    """Inserted text across all segments, space-joined.

    Not a literal substring of the new text: two insert segments separated by
    an equal or delete segment ("the {ins:Director} shall {ins:transfer}")
    are not adjacent in the document, so joining them with ``""`` would jam
    them together ("Directortransfer") and corrupt any downstream word
    tokenisation. A space is always a safe separator here, since this is for
    reading or scoring what changed, not for reconstructing exact spacing.
    """
    return " ".join(s.text for s in segments if s.kind == "insert" and s.text)


def removed_text(segments: list[DiffSegment]) -> str:
    """Deleted text across all segments, space-joined. See :func:`added_text`."""
    return " ".join(s.text for s in segments if s.kind == "delete" and s.text)
