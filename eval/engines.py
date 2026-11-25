# SPDX-License-Identifier: MIT
# Copyright (c) 2026 Kyra Law Limited
"""The algorithms the benchmark scores by default, each a RedlineEngine.

Every algorithm is a :class:`kyralaw_redline.engine.RedlineEngine`: two .docx files
in, one redline .docx out. ``eval/bench.py`` scores every engine in ``ENGINES``
unless told otherwise, and adds the optional ones it defines itself
(python-redlines, word-compare). To add one, subclass RedlineEngine and list an
instance in ``ENGINES``.
"""

from __future__ import annotations

import difflib
from collections.abc import Callable
from datetime import datetime, timezone

from kyralaw_native import NativeEngine

from kyralaw_redline import DiffSegment, extract_text, render_segments_docx
from kyralaw_redline.diff import _tokenize
from kyralaw_redline.engine import InPlaceEngine, RebuildEngine, RedlineEngine

FIXED_DATE = datetime(2026, 1, 1, tzinfo=timezone.utc)


def _opcode_segments(old: list[str], new: list[str], isjunk: Callable[[str], bool] | None) -> list[DiffSegment]:
    out: list[DiffSegment] = []
    sm = difflib.SequenceMatcher(isjunk, old, new, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            out.append(DiffSegment("equal", "".join(old[i1:i2])))
            continue
        if i2 > i1:
            out.append(DiffSegment("delete", "".join(old[i1:i2])))
        if j2 > j1:
            out.append(DiffSegment("insert", "".join(new[j1:j2])))
    return out


class LineDiffEngine(RedlineEngine):
    """What a line diff gives: any changed paragraph is deleted and reinserted whole."""

    name = "line-diff"

    def redline(self, old_docx: bytes, new_docx: bytes) -> bytes:
        old = extract_text(old_docx).splitlines(keepends=True)
        new = extract_text(new_docx).splitlines(keepends=True)
        return render_segments_docx(_opcode_segments(old, new, lambda s: not s.strip()), date=FIXED_DATE)


class TokenDiffEngine(RedlineEngine):
    """A plain token-level difflib over the whole document, with none of the
    paragraph alignment or anchoring heuristics. Isolates what those add."""

    name = "token-diff"

    def redline(self, old_docx: bytes, new_docx: bytes) -> bytes:
        old = _tokenize(extract_text(old_docx))
        new = _tokenize(extract_text(new_docx))
        return render_segments_docx(_opcode_segments(old, new, str.isspace), date=FIXED_DATE)


# kyralaw-redline writes in place or not at all: a pair it cannot redline in place
# is reported as unscored, never quietly rebuilt.
ENGINES: tuple[RedlineEngine, ...] = (
    InPlaceEngine(date=FIXED_DATE),
    RebuildEngine(date=FIXED_DATE),
    NativeEngine(date=FIXED_DATE),
    TokenDiffEngine(),
    LineDiffEngine(),
)
ALGORITHMS: dict[str, RedlineEngine] = {engine.name: engine for engine in ENGINES}
