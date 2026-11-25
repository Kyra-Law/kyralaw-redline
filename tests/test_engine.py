"""RedlineEngine: the interface every algorithm implements, and what it enforces."""

from __future__ import annotations

import io
import sys
from pathlib import Path
from typing import Any

import docx
import pytest

from kyralaw_redline import read_tracked_changes, redline_docx
from kyralaw_redline.engine import (
    InPlaceEngine,
    RebuildEngine,
    RedlineEngine,
    RedlineEngineError,
    check_engine_name,
)


def _docx(text: str) -> bytes:
    document: Any = docx.Document()
    document.add_paragraph(text)
    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue()


OLD, NEW = _docx("The fee is 100 pounds."), _docx("The fee is 200 pounds.")


@pytest.mark.parametrize("name", ["Word Compare", "word_compare", "-leading", "trailing-", "UPPER", "", "a--b"])
def test_a_name_that_is_not_lowercase_words_joined_by_hyphens_is_refused(name: str) -> None:
    with pytest.raises(TypeError, match="lowercase words joined by hyphens"):
        type("Bad", (RedlineEngine,), {"name": name, "redline": lambda self, old, new: old})


def test_good_names_pass() -> None:
    for name in ("word-compare", "kyralaw-redline", "python-redlines", "v2"):
        assert check_engine_name(name) == name


def test_an_engine_without_redline_cannot_be_made() -> None:
    class Unfinished(RedlineEngine):
        name = "unfinished"

    with pytest.raises(TypeError):
        Unfinished()  # type: ignore[abstract]


def test_an_engine_that_returns_something_other_than_a_docx_fails_loudly() -> None:
    class Broken(RedlineEngine):
        name = "broken"

        def redline(self, old_docx: bytes, new_docx: bytes) -> bytes:
            return b""

    with pytest.raises(RedlineEngineError, match="broken returned 0 bytes"):
        Broken()(OLD, NEW)


def test_an_engine_whose_name_is_never_set_cannot_be_called() -> None:
    class Nameless(RedlineEngine):
        def redline(self, old_docx: bytes, new_docx: bytes) -> bytes:
            return new_docx

    with pytest.raises(TypeError):
        Nameless()(OLD, NEW)


@pytest.mark.parametrize(
    ("engine", "mode", "name"),
    [(InPlaceEngine(), "in-place", "kyralaw-redline"), (RebuildEngine(), "rebuild", "kyralaw-redline-rebuild")],
)
def test_the_library_engines_are_redline_docx_in_each_mode(engine: RedlineEngine, mode: Any, name: str) -> None:
    assert read_tracked_changes(engine(OLD, NEW)) == read_tracked_changes(redline_docx(OLD, NEW, mode=mode))
    assert engine.name == name


def test_every_benchmarked_algorithm_is_an_engine_with_its_own_name() -> None:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "eval"))
    import engines

    assert all(isinstance(engine, RedlineEngine) for engine in engines.ENGINES)
    assert list(engines.ALGORITHMS) == [engine.name for engine in engines.ENGINES]
    assert len(set(engines.ALGORITHMS)) == len(engines.ENGINES)
