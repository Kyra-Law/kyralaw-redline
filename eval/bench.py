# SPDX-License-Identifier: MIT
# Copyright (c) 2026 Kyra Law Limited
"""Score redline algorithms on kyralaw-redline-bench.

kyralaw-redline-bench is a set of pairs of Word documents (contracts and other legal
documents, synthetic, derived from Harvey LAB), each with Microsoft Word's own
Compare Documents of the two as the reference. This scores a redline for each pair
against that reference. The dataset is MIT licensed, like this repository.

Usage::

    python eval/bench.py                               # the built-in algorithms, test split
    python eval/bench.py --split dev                   # the 48 pairs meant for tuning
    python eval/bench.py -a kyralaw-redline            # one algorithm (repeatable)
    python eval/bench.py -a python-redlines            # needs: pip install python-redlines
    python eval/bench.py -a docxodus                   # needs: pip install docx-scalpel
    python eval/bench.py -a docxodus-checked           # the same, with its redlines checked
    python eval/bench.py -a word-compare               # Word's reference as a candidate: must give 100
    python eval/bench.py --redlines mine=out/          # score your own: out/<pair id>.docx
    python eval/bench.py --json results.json --save-redlines out/
    python eval/bench.py DATA ...                      # a copy of the dataset in folder DATA

Without ``DATA``, the dataset comes from Hugging Face at the revision
``eval/bench_data.py`` pins: downloaded once into the local Hugging Face cache, read from
there afterwards, and checked against its ``SHA256SUMS`` either way. ``DATA``, when
given, is a folder holding ``dev.jsonl``, ``test.jsonl`` and ``pairs/``.

**What is compared.** The text each redline marks inserted and deleted, read from its
``word/document.xml`` body (headers, footers and notes are not scored), against the
same read of Word's reference: as bags of tokens, giving precision, recall and F1 for
insertions and for deletions (``prf`` and ``score_tokens``). Moved text counts as
deleted where it was (``w:moveFrom``) and inserted where it went (``w:moveTo``), in
the reference and the candidate alike, so detecting moves is neither rewarded nor
penalised. Word marks moves in about one pair in seven.

**Fidelity**, 0 to 1, the mean of four checks that the redline is still a usable
Word document:

- ``markup_free``: no more ``**``, ``{{``, ``}}`` or pipe-table rows in the redline's
  accepted (original) text than in the new (old) version itself. The versions can
  contain such text; the check is that the redline adds none.
- ``tables``: as many tables as Word's reference.
- ``styles``: the same mix of paragraph styles as Word's reference.
- ``headers``: the new version's header and footer text kept, as a share of its
  tokens found in the redline's headers and footers. Text, not parts: Word's own
  Compare sometimes adds an empty header part the new version does not have.

**Score** = 100 x the mean of insert F1, delete F1 and fidelity, averaged over pairs.

**Integrity.** Accepting every change in a redline must give the new version's text,
and rejecting every change the old's, token for token. A redline that does not is
counted and reported, not scored differently: the score measures how a redline marks
the change, the integrity count whether it is the right document at all. A pair with
no redline (a missing file, or an algorithm that failed) is reported with its count
and left out of the averages, never scored as zero in silence.
"""

from __future__ import annotations

import argparse
import collections
import concurrent.futures
import functools
import hashlib
import importlib
import io
import json
import re
import statistics
import sys
import time
import zipfile
from collections.abc import Iterator
from dataclasses import asdict, dataclass, replace
from functools import cache
from pathlib import Path
from typing import Any, Literal
from xml.etree import ElementTree

sys.path.insert(0, str(Path(__file__).resolve().parent))

import bench_data
from engines import ALGORITHMS

from kyralaw_redline import extract_paragraphs
from kyralaw_redline.engine import RedlineEngine, check_engine_name

View = Literal["accepted", "original"]

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_INSERTING = frozenset({f"{_W}ins", f"{_W}moveTo"})
_DELETING = frozenset({f"{_W}del", f"{_W}moveFrom"})
_NOT_TEXT = frozenset({f"{_W}pPr", f"{_W}rPr", f"{_W}instrText", f"{_W}delInstrText"})
_MARKUP_RE = re.compile(r"\*\*|\{\{|\}\}")
_STORY_RE = re.compile(r"word/(header|footer)\d*\.xml")
CHECKS = ("markup_free", "tables", "styles", "headers")


# Scoring tokens --------------------------------------------------------------
#
# Lowercase words and single punctuation characters, compared as bags: indifferent
# to *where* an ambiguous edit is placed (deleting one of two adjacent "the"s is the
# same edit either way), while a redline that deletes and reinserts an unchanged
# word still loses precision on both sides.

SCORE_TOKEN_RE = re.compile(r"\w+|[^\w\s]", re.UNICODE)


def score_tokens(texts: list[str]) -> collections.Counter[str]:
    return collections.Counter(t.lower() for text in texts for t in SCORE_TOKEN_RE.findall(text))


@dataclass
class PRF:
    precision: float
    recall: float
    f1: float


def prf(predicted: collections.Counter[str], expected: collections.Counter[str]) -> PRF:
    """Bag-of-tokens precision, recall and F1. Empty against empty is a
    perfect score; anything against empty, or empty against anything, is an
    F1 of zero."""
    if not predicted and not expected:
        return PRF(1.0, 1.0, 1.0)
    hit = sum((predicted & expected).values())
    p = hit / sum(predicted.values()) if predicted else 1.0
    r = hit / sum(expected.values()) if expected else 1.0
    f1 = 2 * p * r / (p + r) if p + r else 0.0
    return PRF(p, r, f1)


# Reading a .docx -----------------------------------------------------------


def _package(docx: bytes) -> zipfile.ZipFile:
    return zipfile.ZipFile(io.BytesIO(docx))


def _body(docx: bytes) -> ElementTree.Element:
    body = ElementTree.fromstring(_package(docx).read("word/document.xml")).find(f"{_W}body")
    if body is None:
        raise ValueError("the document has no body")
    return body


def _paragraphs(el: ElementTree.Element) -> Iterator[ElementTree.Element]:
    """Every paragraph in document order, tables included, text boxes not: a
    paragraph inside another paragraph belongs to a drawing, not to the text."""
    for child in el:
        if child.tag == f"{_W}p":
            yield child
        else:
            yield from _paragraphs(child)


def _segments(p: ElementTree.Element) -> list[tuple[str, str]]:
    """A paragraph's text as (kind, text): "e" unchanged, "i" inserted or moved to,
    "d" deleted or moved from. Consecutive pieces of one kind are joined, so a word
    split across two revision elements stays one word."""
    out: list[list[str]] = []

    def add(kind: str, text: str) -> None:
        if not text:
            return
        if out and out[-1][0] == kind:
            out[-1][1] += text
        else:
            out.append([kind, text])

    def walk(el: ElementTree.Element, kind: str) -> None:
        for c in el:
            if c.tag in _NOT_TEXT or c.tag == f"{_W}p":
                continue
            if c.tag in _INSERTING:
                walk(c, "i")
            elif c.tag in _DELETING:
                walk(c, "d")
            elif c.tag in (f"{_W}t", f"{_W}delText"):
                add(kind, c.text or "")
            elif c.tag == f"{_W}tab":
                add(kind, "\t")
            elif c.tag in (f"{_W}br", f"{_W}cr"):
                add(kind, "\n")
            elif c.tag == f"{_W}noBreakHyphen":
                add(kind, "-")
            else:
                walk(c, kind)

    walk(p, "e")
    return [(k, t) for k, t in out]


def tracked_text(docx: bytes) -> tuple[list[str], list[str]]:
    """(inserted, deleted): the text of every inserted and every deleted stretch in the body."""
    inserted: list[str] = []
    deleted: list[str] = []
    for p in _paragraphs(_body(docx)):
        for kind, text in _segments(p):
            if kind == "i":
                inserted.append(text)
            elif kind == "d":
                deleted.append(text)
    return inserted, deleted


def _mark(p: ElementTree.Element) -> str | None:
    """ "i" if the paragraph's own mark was inserted (or moved to), "d" if deleted
    (or moved from), else None."""
    rpr = p.find(f"{_W}pPr/{_W}rPr")
    for c in rpr if rpr is not None else ():
        if c.tag in _INSERTING:
            return "i"
        if c.tag in _DELETING:
            return "d"
    return None


def _view(paragraphs: Iterator[ElementTree.Element], view: View) -> list[str]:
    drop = "d" if view == "accepted" else "i"
    out: list[str] = []
    carry = ""
    for p in paragraphs:
        text = carry + "".join(t for k, t in _segments(p) if k != drop)
        if _mark(p) == drop:
            # The mark does not exist in this view, so the paragraph runs on into the next.
            carry = text
            continue
        carry = ""
        out.append(text)
    if carry:
        out.append(carry)
    return out


def view_paragraphs(docx: bytes, view: View) -> list[str]:
    """The body as it reads with every change accepted or every change rejected, one
    string per paragraph, as Word shows it: a paragraph whose mark was inserted runs
    on into the next when changes are rejected, and one whose mark was deleted when
    they are accepted."""
    return _view(_paragraphs(_body(docx)), view)


def story_text(docx: bytes) -> list[str]:
    """The accepted text of every header and footer part."""
    package = _package(docx)
    out: list[str] = []
    for name in package.namelist():
        if _STORY_RE.fullmatch(name):
            out.extend(_view(_paragraphs(ElementTree.fromstring(package.read(name))), "accepted"))
    return out


def _tokens(paragraphs: list[str]) -> list[str]:
    return [t.lower() for text in paragraphs for t in SCORE_TOKEN_RE.findall(text)]


def _markup(paragraphs: list[str]) -> int:
    count = 0
    for text in paragraphs:
        count += len(_MARKUP_RE.findall(text))
        stripped = text.strip()
        if len(stripped) > 1 and stripped.startswith("|") and stripped.endswith("|"):
            count += 1
    return count


def _styles(docx: bytes) -> collections.Counter[str]:
    styles: collections.Counter[str] = collections.Counter()
    for p in _paragraphs(_body(docx)):
        style = p.find(f"{_W}pPr/{_W}pStyle")
        styles[(style.get(f"{_W}val") or "") if style is not None else ""] += 1
    return styles


def _share(a: int, b: int) -> float:
    return 1.0 if max(a, b) == 0 else min(a, b) / max(a, b)


def _kept(got: collections.Counter[str], expected: collections.Counter[str]) -> float:
    total = sum(expected.values())
    return 1.0 if total == 0 else sum((got & expected).values()) / total


# Scoring ---------------------------------------------------------------------


@dataclass
class PairResult:
    pair: str
    scored: bool
    error: str | None = None
    seconds: float | None = None
    insert_f1: float = 0.0
    delete_f1: float = 0.0
    insert_precision: float = 0.0
    insert_recall: float = 0.0
    delete_precision: float = 0.0
    delete_recall: float = 0.0
    fidelity: float = 0.0
    checks: dict[str, float] | None = None
    integrity: bool = True

    @property
    def score(self) -> float:
        return 100 * (self.insert_f1 + self.delete_f1 + self.fidelity) / 3


@dataclass
class Reference:
    """What every candidate for one pair is measured against, read once."""

    inserted: collections.Counter[str]
    deleted: collections.Counter[str]
    old_tokens: list[str]
    new_tokens: list[str]
    old_markup: int
    new_markup: int
    tables: int
    styles: collections.Counter[str]
    story_tokens: collections.Counter[str]


def read_reference(folder: Path) -> Reference:
    old, new, ref = ((folder / f"{name}.docx").read_bytes() for name in ("old", "new", "reference"))
    inserted, deleted = tracked_text(ref)
    old_view, new_view = view_paragraphs(old, "accepted"), view_paragraphs(new, "accepted")
    return Reference(
        inserted=score_tokens(inserted),
        deleted=score_tokens(deleted),
        old_tokens=_tokens(old_view),
        new_tokens=_tokens(new_view),
        old_markup=_markup(old_view),
        new_markup=_markup(new_view),
        tables=sum(1 for _ in _body(ref).iter(f"{_W}tbl")),
        styles=_styles(ref),
        story_tokens=score_tokens(story_text(new)),
    )


def score_redline(pair: str, redline: bytes, ref: Reference, seconds: float | None) -> PairResult:
    inserted, deleted = tracked_text(redline)
    ins, dele = prf(score_tokens(inserted), ref.inserted), prf(score_tokens(deleted), ref.deleted)
    accepted, original = view_paragraphs(redline, "accepted"), view_paragraphs(redline, "original")
    styles = _styles(redline)
    style_total = max(sum(styles.values()), sum(ref.styles.values()), 1)
    checks = {
        "markup_free": float(_markup(accepted) <= ref.new_markup and _markup(original) <= ref.old_markup),
        "tables": _share(sum(1 for _ in _body(redline).iter(f"{_W}tbl")), ref.tables),
        "styles": sum((styles & ref.styles).values()) / style_total,
        "headers": _kept(score_tokens(story_text(redline)), ref.story_tokens),
    }
    return PairResult(
        pair=pair,
        scored=True,
        seconds=seconds,
        insert_f1=ins.f1,
        delete_f1=dele.f1,
        insert_precision=ins.precision,
        insert_recall=ins.recall,
        delete_precision=dele.precision,
        delete_recall=dele.recall,
        fidelity=sum(checks.values()) / len(checks),
        checks=checks,
        integrity=_tokens(accepted) == ref.new_tokens and _tokens(original) == ref.old_tokens,
    )


# Engines ---------------------------------------------------------------------
#
# Every algorithm scored here is a RedlineEngine: two .docx files in, one redline
# .docx out. Beside those in eval/engines.py: docxodus (plain and checked), python-redlines,
# word-compare, and redlines
# made elsewhere, read from a folder.


class PythonRedlinesEngine(RedlineEngine):
    """Open-XML-PowerTools' WmlComparer through python-redlines, a third-party reference point."""

    name = "python-redlines"

    @cache  # noqa: B019 - one comparer per engine, and there is one engine per run
    def _comparer(self) -> Any:
        try:
            engines = importlib.import_module("python_redlines.engines")
        except ImportError:
            sys.exit("-a python-redlines needs the python-redlines package: pip install python-redlines")
        return engines.XmlPowerToolsEngine()

    def redline(self, old_docx: bytes, new_docx: bytes) -> bytes:
        out = self._comparer().run_redline("kyralaw-redline-bench", old_docx, new_docx)[0]
        if not out:
            raise RuntimeError("python-redlines returned no document")
        return bytes(out)


class DocxodusEngine(RedlineEngine):
    """Docxodus's comparer through docx-scalpel, its own Python client, a third-party reference point.

    The settings are the ones ``DocxCompare.Compare`` uses, the comparison Docxodus's
    ``redline`` command runs: the engine's defaults, which docx-scalpel mirrors field for
    field, plus ``pre_accept_input_revisions``. docx-scalpel sends every pair to one
    long-running Docxodus process, so the time per pair leaves out .NET's start-up, which
    python-redlines pays on every pair.
    """

    name = "docxodus"

    @cache  # noqa: B019 - one client per engine, and there is one engine per run
    def _client(self) -> tuple[Any, Any]:
        try:
            scalpel = importlib.import_module("docx_scalpel")
        except ImportError:
            sys.exit("-a docxodus needs the docx-scalpel package: pip install docx-scalpel")
        settings = scalpel.DocxDiffSettings(
            author_for_revisions="kyralaw-redline-bench", pre_accept_input_revisions=True
        )
        return scalpel, settings

    def redline(self, old_docx: bytes, new_docx: bytes) -> bytes:
        return self._compare(old_docx, new_docx, self._client()[1])

    def _compare(self, old_docx: bytes, new_docx: bytes, settings: Any) -> bytes:
        out = self._client()[0].docx_diff_compare(old_docx, new_docx, settings)
        if not out:
            raise RuntimeError("docx-scalpel returned no document")
        return bytes(out)


def _paragraph_words(docx: bytes, view: View) -> list[str]:
    """Each paragraph of ``docx`` in ``view``, its runs of whitespace as one space, trimmed."""
    return [" ".join(p.split()) for p in extract_paragraphs(docx, view=view)]


class DocxodusCheckedEngine(DocxodusEngine):
    """Docxodus with its redline checked, and compared again once if the check fails.

    Docxodus treats "-" and a few punctuation marks as word separators, so a paragraph of
    nothing else, such as "---", has no words and matches an empty paragraph. When an
    empty paragraph becomes "---" with a new paragraph after it, Docxodus can write the
    "---" as unchanged, and rejecting every change then keeps it. This engine checks that
    accepting every change gives the new document's paragraphs and rejecting every change
    the old one's, read by ``kyralaw_redline.extract_paragraphs`` with runs of whitespace
    treated as one space (as the scorer ignores whitespace, an empty paragraph and one of
    spaces are the same); if not, it compares again with the space as the only separator,
    and says so on stderr. A redline that fails both is refused, so the pair is reported
    rather than scored.
    """

    name = "docxodus-checked"

    def redline(self, old_docx: bytes, new_docx: bytes) -> bytes:
        settings = self._client()[1]
        old, new = _paragraph_words(old_docx, "accepted"), _paragraph_words(new_docx, "accepted")
        for attempt, attempt_settings in enumerate((settings, replace(settings, word_separators=" "))):
            out = self._compare(old_docx, new_docx, attempt_settings)
            if _paragraph_words(out, "original") == old and _paragraph_words(out, "accepted") == new:
                return out
            if attempt == 0:
                print(
                    "docxodus-checked: a redline does not give back both documents; comparing again with only "
                    "spaces as word separators",
                    file=sys.stderr,
                    flush=True,
                )
        raise RuntimeError("Docxodus's redline does not give back both documents, even with only spaces as separators")


class MissingRedline(LookupError):
    """There is no redline for this pair."""


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def pair_index(data: Path, rows: list[dict[str, Any]]) -> dict[tuple[str, str], str]:
    """Each pair's id, by the sha256 of its old and new documents."""
    index: dict[tuple[str, str], str] = {}
    for row in rows:
        folder = data / "pairs" / row["id"]
        key = (_sha256((folder / "old.docx").read_bytes()), _sha256((folder / "new.docx").read_bytes()))
        if key in index:
            raise ValueError(f"{row['id']} and {index[key]} have identical documents")
        index[key] = row["id"]
    return index


class PrebuiltRedlinesEngine(RedlineEngine):
    """Redlines made elsewhere, one ``<pair id>.docx`` per pair in a folder.

    For an algorithm that cannot run here: one that needs Word, a server or code that is
    not public. The pair is found from the two documents' contents, so it is called
    like any other engine. Its time is a file read, so none is reported.
    """

    measures_time = False

    def __init__(self, name: str, folder: Path, index: dict[tuple[str, str], str]) -> None:
        self.name = check_engine_name(name)
        self.folder = folder
        self.index = index

    def path(self, pair: str) -> Path:
        return self.folder / f"{pair}.docx"

    def redline(self, old_docx: bytes, new_docx: bytes) -> bytes:
        pair = self.index.get((_sha256(old_docx), _sha256(new_docx)))
        if pair is None:
            raise MissingRedline("these documents are not a pair of this dataset")
        path = self.path(pair)
        if not path.is_file():
            raise MissingRedline("no redline file")
        return path.read_bytes()


class WordCompareEngine(PrebuiltRedlinesEngine):
    """Word's own reference, scored as a candidate: a check on the scorer, which must give 100."""

    def __init__(self, data: Path, index: dict[tuple[str, str], str]) -> None:
        super().__init__("word-compare", data / "pairs", index)

    def path(self, pair: str) -> Path:
        return self.folder / pair / "reference.docx"


OPTIONAL = ("docxodus", "docxodus-checked", "python-redlines", "word-compare")


def run(
    data: Path,
    rows: list[dict[str, Any]],
    engine: RedlineEngine,
    save_to: Path | None,
    refs: dict[str, Reference],
    jobs: int = 1,
) -> list[PairResult]:
    """Score ``engine`` on every pair in ``rows``, in order.

    With ``jobs`` above 1 the pairs are spread over that many processes, each reading its
    own references, which makes a full test run several times faster; the times each pair
    reports are then taken under that load and are not comparable with a serial run.
    """
    if jobs > 1:
        with concurrent.futures.ProcessPoolExecutor(max_workers=jobs) as pool:
            return list(pool.map(functools.partial(_run_pair, data, engine=engine, save_to=save_to), rows, chunksize=4))
    return [_run_pair(data, row, engine=engine, save_to=save_to, refs=refs) for row in rows]


def _run_pair(
    data: Path,
    row: dict[str, Any],
    *,
    engine: RedlineEngine,
    save_to: Path | None,
    refs: dict[str, Reference] | None = None,
) -> PairResult:
    pair = row["id"]
    folder = data / "pairs" / pair
    if refs is None:
        reference = read_reference(folder)
    else:
        if pair not in refs:
            refs[pair] = read_reference(folder)
        reference = refs[pair]
    try:
        old, new = (folder / "old.docx").read_bytes(), (folder / "new.docx").read_bytes()
        started = time.perf_counter()
        redline = engine(old, new)
        seconds = time.perf_counter() - started if engine.measures_time else None
        result = score_redline(pair, redline, reference, seconds)
    except MissingRedline as exc:
        return PairResult(pair=pair, scored=False, error=str(exc))
    except Exception as exc:  # one pair's failure is reported, not fatal
        return PairResult(pair=pair, scored=False, error=f"{type(exc).__name__}: {exc}"[:300])
    if save_to is not None and engine.measures_time:
        save_to.mkdir(parents=True, exist_ok=True)
        (save_to / f"{pair}.docx").write_bytes(redline)
    return result


# Reporting -------------------------------------------------------------------


def _mean(xs: list[float]) -> float:
    return sum(xs) / len(xs) if xs else 0.0


def summarise(results: list[PairResult]) -> dict[str, Any]:
    scored = [r for r in results if r.scored]
    times = [r.seconds for r in scored if r.seconds is not None]
    return {
        "pairs": len(results),
        "scored": len(scored),
        "unscored": len(results) - len(scored),
        "score": _mean([r.score for r in scored]),
        "text_f1": 100 * _mean([(r.insert_f1 + r.delete_f1) / 2 for r in scored]),
        "insert_f1": _mean([r.insert_f1 for r in scored]),
        "delete_f1": _mean([r.delete_f1 for r in scored]),
        "fidelity": _mean([r.fidelity for r in scored]),
        "checks": {c: _mean([(r.checks or {}).get(c, 0.0) for r in scored]) for c in CHECKS},
        "integrity_failures": sum(1 for r in scored if not r.integrity),
        "median_seconds": statistics.median(times) if times else None,
    }


def report(summaries: dict[str, dict[str, Any]], split: str, pairs: int) -> str:
    lines = [
        f"kyralaw-redline-bench, {split} split, {pairs} pairs",
        "",
        "| Algorithm | Scored | Score | Text F1 | Insert F1 | Delete F1 | Fidelity | Integrity failures "
        "| Median time |",
        "| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |",
    ]
    for name, s in summaries.items():
        t = s["median_seconds"]
        lines.append(
            f"| {name} | {s['scored']} of {s['pairs']} | {s['score']:.2f} | {s['text_f1']:.2f} "
            f"| {s['insert_f1']:.3f} | "
            f"{s['delete_f1']:.3f} | {s['fidelity']:.3f} | {s['integrity_failures']} | "
            + ("-" if t is None else f"{t * 1000:.0f} ms" if t < 1 else f"{t:.2f} s")
            + " |"
        )
    lines.append("")
    lines.append(
        "Fidelity checks: "
        + "; ".join(
            f"{name} " + ", ".join(f"{c} {v:.3f}" for c, v in s["checks"].items()) for name, s in summaries.items()
        )
    )
    for name, s in summaries.items():
        if s["unscored"]:
            lines.append(
                f"WARNING: {name} has no redline for {s['unscored']} of {s['pairs']} pairs; "
                "they are left out of its averages"
            )
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument(
        "data", type=Path, nargs="?", help="a dataset folder (default: Hugging Face, cached locally; see above)"
    )
    parser.add_argument("--split", choices=("test", "dev"), default="test")
    parser.add_argument(
        "-a", "--algorithm", action="append", choices=sorted({*ALGORITHMS, *OPTIONAL}), help="repeatable"
    )
    parser.add_argument(
        "--redlines", action="append", default=[], metavar="NAME=DIR", help="score DIR/<pair id>.docx as NAME"
    )
    parser.add_argument("--limit", type=int, help="score only the first N pairs, for a quick look")
    parser.add_argument(
        "--jobs", type=int, default=1, help="score N pairs at once (times are then not comparable with a serial run)"
    )
    parser.add_argument("--json", type=Path, help="write every pair's scores here")
    parser.add_argument("--save-redlines", type=Path, help="keep the redlines the algorithms wrote, per algorithm")
    args = parser.parse_args()

    data: Path = args.data.expanduser().resolve() if args.data else bench_data.fetch()
    split_file = data / f"{args.split}.jsonl"
    if not split_file.is_file():
        sys.exit(f"{split_file} not found; DATA is the kyralaw-redline-bench folder")
    rows = [json.loads(line) for line in split_file.read_text().splitlines() if line.strip()]
    if args.limit:
        rows = rows[: args.limit]

    names = args.algorithm or ([] if args.redlines else list(ALGORITHMS))
    index = pair_index(data, rows)
    engines: list[RedlineEngine] = []
    for name in names:
        if name == "docxodus":
            engines.append(DocxodusEngine())
        elif name == "docxodus-checked":
            engines.append(DocxodusCheckedEngine())
        elif name == "python-redlines":
            engines.append(PythonRedlinesEngine())
        elif name == "word-compare":
            engines.append(WordCompareEngine(data, index))
        else:
            engines.append(ALGORITHMS[name])
    for spec in args.redlines:
        name, sep, folder = spec.partition("=")
        if not sep or not name or not folder:
            sys.exit(f"--redlines takes NAME=DIR, not {spec!r}")
        try:
            engines.append(PrebuiltRedlinesEngine(name, Path(folder).expanduser(), index))
        except TypeError as exc:
            sys.exit(f"--redlines {spec!r}: {exc}")
    refs: dict[str, Reference] = {}
    all_results: dict[str, list[PairResult]] = {}
    for engine in engines:
        save_to = args.save_redlines / engine.name if args.save_redlines else None
        all_results[engine.name] = run(data, rows, engine, save_to, refs, args.jobs)
        print(f"{engine.name}: done", file=sys.stderr, flush=True)

    summaries = {name: summarise(results) for name, results in all_results.items()}
    print(report(summaries, args.split, len(rows)))
    if args.jobs > 1:
        print(f"\nTimes were taken with --jobs {args.jobs}, under that load: compare them only with each other.")
    if args.json:
        args.json.write_text(
            json.dumps(
                {
                    "split": args.split,
                    "summaries": summaries,
                    "pairs": {name: [asdict(r) for r in results] for name, results in all_results.items()},
                },
                indent=1,
            )
        )


if __name__ == "__main__":
    main()
