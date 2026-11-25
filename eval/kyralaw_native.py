# SPDX-License-Identifier: MIT
# Copyright (c) 2026 Kyra Law Limited
"""kyralaw-native: the first version of the in-place renderer, frozen as a baseline.

It is today's in-place writer (:func:`kyralaw_redline.inplace._render`) set back to
what it had then:

* its own paragraph alignment (:func:`align_paragraphs_v1` below): a plain
  ``SequenceMatcher`` over paragraphs, where today's anchors on paragraphs unique to
  both sides, and a word-level gate on block similarity that today's does not have;
* rows matched by unweighted votes, and no guard keeping pairs inside corresponding
  table rows (``respect_tables=False``).

Like the in-place renderer it refuses a pair whose read-back fails. It used to fall
back to a text renderer, which is not part of this repository, so here a refused
pair is reported as unscored. Kept so the results show what the current renderer
changed; use :class:`kyralaw_redline.engine.InPlaceEngine` for real work.
"""

from __future__ import annotations

import difflib
from datetime import datetime

from kyralaw_redline import inplace
from kyralaw_redline.diff import (
    _ALIGN_MAX_CELLS,
    _ALIGN_MIN_RATIO,
    ParagraphAlignment,
    _is_junk_line,
    _is_junk_token,
    _tokenize,
    diff_paragraph,
)
from kyralaw_redline.engine import RedlineEngine

# Below this share of matching words, two blocks are unrelated, and the character
# comparison is not run for them at all. The character ratio is the expensive part
# of an alignment and a poor judge of unrelated prose: two paragraphs sharing no
# wording still match on "the " and " of ", and can clear 0.3 that way. A word-level
# ratio is far cheaper and says plainly that such a pair has nothing in common.
_ALIGN_MIN_WORD_RATIO = 0.2


def _block_similarity(old_block: str, new_block: str, old_tokens: list[str], new_tokens: list[str]) -> float:
    """How alike two blocks are, 0..1, cheap-to-expensive with early outs.

    ``real_quick_ratio`` and ``quick_ratio`` are upper bounds, so a value below the
    floor rules the pair out without running the full comparison. The word-level
    check runs first, being the cheaper of the two.
    """
    if not old_block.strip() or not new_block.strip():
        return 1.0 if old_block.strip() == new_block.strip() else 0.0
    words = difflib.SequenceMatcher(_is_junk_token, old_tokens, new_tokens, autojunk=False)
    if (
        words.real_quick_ratio() < _ALIGN_MIN_WORD_RATIO
        or words.quick_ratio() < _ALIGN_MIN_WORD_RATIO
        or words.ratio() < _ALIGN_MIN_WORD_RATIO
    ):
        return 0.0
    sm = difflib.SequenceMatcher(None, old_block, new_block, autojunk=False)
    if sm.real_quick_ratio() < _ALIGN_MIN_RATIO or sm.quick_ratio() < _ALIGN_MIN_RATIO:
        return 0.0
    return sm.ratio()


def _align_blocks(old_blocks: list[str], new_blocks: list[str]) -> list[tuple[int | None, int | None]]:
    """Pair blocks by similarity, in order, allowing gaps on either side: a
    longest common subsequence over similarity, so pairs never cross."""
    n, m = len(old_blocks), len(new_blocks)
    if n * m > _ALIGN_MAX_CELLS:
        paired = min(n, m)
        out: list[tuple[int | None, int | None]] = [(k, k) for k in range(paired)]
        out.extend((k, None) for k in range(paired, n))
        out.extend((None, k) for k in range(paired, m))
        return out
    old_tokens = [_tokenize(b) for b in old_blocks]
    new_tokens = [_tokenize(b) for b in new_blocks]
    ratios = [[0.0] * m for _ in range(n)]
    score = [[0.0] * (m + 1) for _ in range(n + 1)]
    for i in range(n - 1, -1, -1):
        for j in range(m - 1, -1, -1):
            ratio = _block_similarity(old_blocks[i], new_blocks[j], old_tokens[i], new_tokens[j])
            ratios[i][j] = ratio
            match = (ratio + score[i + 1][j + 1]) if ratio >= _ALIGN_MIN_RATIO else -1.0
            score[i][j] = max(match, score[i + 1][j], score[i][j + 1])
    pairs: list[tuple[int | None, int | None]] = []
    i = j = 0
    while i < n and j < m:
        ratio = ratios[i][j]
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


def align_paragraphs_v1(old_paragraphs: list[str], new_paragraphs: list[str]) -> list[ParagraphAlignment]:
    """The first paragraph alignment: plain ``SequenceMatcher`` opcodes over the
    paragraphs, then :func:`_align_blocks` within each replaced region and
    :func:`kyralaw_redline.diff.diff_paragraph` within each pair."""
    sm = difflib.SequenceMatcher(_is_junk_line, old_paragraphs, new_paragraphs, autojunk=False)
    steps: list[ParagraphAlignment] = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
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


class NativeEngine(RedlineEngine):
    """The first in-place renderer, frozen: see the module docstring."""

    name = "kyralaw-native"

    def __init__(self, *, date: datetime | None = None) -> None:
        self.date = date

    def redline(self, old_docx: bytes, new_docx: bytes) -> bytes:
        return inplace._render(old_docx, new_docx, date=self.date, align=align_paragraphs_v1, respect_tables=False)
