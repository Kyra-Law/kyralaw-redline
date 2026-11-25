# SPDX-License-Identifier: MIT
# Copyright (c) 2026 Kyra Law Limited
"""RedlineEngine: the one interface every redline algorithm here implements.

A redline engine takes two Word documents and returns a third: the redline, with
the changes from the first to the second marked as tracked changes (``w:ins`` and
``w:del``)::

    engine = AutoEngine()
    redline = engine(old_docx, new_docx)   # bytes, bytes -> bytes

To add an algorithm, subclass :class:`RedlineEngine`, give it a ``name`` (lowercase
words joined by hyphens, checked when the class is defined) and implement
:meth:`RedlineEngine.redline`. The benchmark scorer (``eval/bench.py``) runs
engines and nothing else, so an engine is all it takes to be scored. Call an engine
rather than its ``redline`` method: the call checks that what came back is a .docx,
so a broken engine fails loudly instead of being scored as an empty redline.
"""

from __future__ import annotations

import re
from abc import ABC, abstractmethod
from datetime import datetime

from kyralaw_redline.docx_io import DEFAULT_AUTHOR, Mode, redline_docx

_NAME_RE = re.compile(r"[a-z0-9]+(?:-[a-z0-9]+)*")


class RedlineEngineError(RuntimeError):
    """An engine returned something that is not a .docx."""


def check_engine_name(name: object) -> str:
    """``name`` if it is lowercase words joined by hyphens, else ``TypeError``."""
    if not isinstance(name, str) or not _NAME_RE.fullmatch(name):
        raise TypeError(f"a redline engine's name must be lowercase words joined by hyphens, not {name!r}")
    return name


class RedlineEngine(ABC):
    """A redline algorithm: two .docx files in, one redline .docx out."""

    #: Lowercase words joined by hyphens, e.g. ``"kyralaw-redline-auto"``. Set on the
    #: class; an engine whose name is only known at run time sets it in ``__init__``.
    name: str
    #: Whether timing a call measures the algorithm. False for an engine that only
    #: reads redlines made elsewhere, whose time is a file read.
    measures_time: bool = True

    def __init_subclass__(cls, **kwargs: object) -> None:
        super().__init_subclass__(**kwargs)
        if "name" in cls.__dict__:
            check_engine_name(cls.__dict__["name"])

    @abstractmethod
    def redline(self, old_docx: bytes, new_docx: bytes) -> bytes:
        """The redline of ``old_docx`` against ``new_docx``, as a .docx with tracked changes."""

    def __call__(self, old_docx: bytes, new_docx: bytes) -> bytes:
        check_engine_name(getattr(self, "name", None))
        out = self.redline(old_docx, new_docx)
        if not isinstance(out, bytes) or not out.startswith(b"PK"):
            got = type(out).__name__ if not isinstance(out, bytes) else f"{len(out)} bytes that are not a zip package"
            raise RedlineEngineError(f"{self.name} returned {got}, not a .docx")
        return out

    def __repr__(self) -> str:
        return f"<{type(self).__name__} {getattr(self, 'name', '?')}>"


class _KyralawRedline(RedlineEngine):
    """kyralaw-redline through :func:`kyralaw_redline.docx_io.redline_docx`, in one mode."""

    _mode: Mode

    def __init__(self, *, author: str = DEFAULT_AUTHOR, date: datetime | None = None) -> None:
        self.author = author
        self.date = date

    def redline(self, old_docx: bytes, new_docx: bytes) -> bytes:
        return redline_docx(old_docx, new_docx, author=self.author, date=self.date, mode=self._mode)


class InPlaceEngine(_KyralawRedline):
    """The tracked changes written into the new document itself; raises when that is not possible."""

    name = "kyralaw-redline-in-place"
    _mode = "in-place"


class AutoEngine(_KyralawRedline):
    """In place, and rebuilt when that is not possible, with a warning: the library's default."""

    name = "kyralaw-redline-auto"
    _mode = "auto"


class RebuildEngine(_KyralawRedline):
    """A new, plain document built from the two texts."""

    name = "kyralaw-redline-rebuild"
    _mode = "rebuild"
