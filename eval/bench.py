# SPDX-License-Identifier: MIT
# Copyright (c) 2026 Kyra Law Limited
"""Score redline algorithms on kyralaw-redline-bench.

kyralaw-redline-bench is a set of pairs of Word documents (contracts and other legal
documents, synthetic, derived from Harvey LAB), each with Microsoft Word's own
Compare Documents of the two as the reference. This scores a redline for each pair
against that reference. The dataset is MIT licensed, like this repository.

Usage::

    python eval/bench.py DATA                          # the built-in algorithms, test split
    python eval/bench.py DATA --split dev              # the 48 pairs meant for tuning
    python eval/bench.py DATA -a kyralaw-redline       # one algorithm (repeatable)
    python eval/bench.py DATA -a python-redlines       # needs: pip install python-redlines
    python eval/bench.py DATA --redlines mine=out/     # score your own: out/<pair id>.docx
    python eval/bench.py DATA --json results.json --save-redlines out/

``DATA`` is the dataset folder, holding ``dev.jsonl``, ``test.jsonl`` and ``pairs/``.

**What is compared.** The text each redline marks inserted and deleted, read from its
``word/document.xml`` body (headers, footers and notes are not scored), against the
same read of Word's reference: as bags of tokens, giving precision, recall and F1 for
insertions and for deletions (``prf`` and ``score_tokens``, shared with
``run_eval.py``). Moved text counts as deleted where it was (``w:moveFrom``) and
inserted where it went (``w:moveTo``), in the reference and the candidate alike, so
detecting moves is neither rewarded nor penalised. Word marks moves in about one pair
in seven.

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
import importlib
import io
import json
import re
import statistics
import sys
import time
import zipfile
from collections.abc import Callable, Iterator
from dataclasses import asdict, dataclass
from functools import cache
from pathlib import Path
from typing import Any, Literal
from xml.etree import ElementTree

sys.path.insert(0, str(Path(__file__).resolve().parent))

from generate import SCORE_TOKEN_RE, score_tokens
from run_eval import ALGORITHMS, prf

Algorithm = Callable[[bytes, bytes], bytes]
View = Literal["accepted", "original"]

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_INSERTING = frozenset({f"{_W}ins", f"{_W}moveTo"})
_DELETING = frozenset({f"{_W}del", f"{_W}moveFrom"})
_NOT_TEXT = frozenset({f"{_W}pPr", f"{_W}rPr", f"{_W}instrText", f"{_W}delInstrText"})
_MARKUP_RE = re.compile(r"\*\*|\{\{|\}\}")
_STORY_RE = re.compile(r"word/(header|footer)\d*\.xml")
CHECKS = ("markup_free", "tables", "styles", "headers")


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


# Algorithms ------------------------------------------------------------------


@cache
def _python_redlines_engine() -> Any:
    try:
        engines = importlib.import_module("python_redlines.engines")
    except ImportError:
        sys.exit("-a python-redlines needs the python-redlines package: pip install python-redlines")
    return engines.XmlPowerToolsEngine()


def _python_redlines(old: bytes, new: bytes) -> bytes:
    """Open-XML-PowerTools' WmlComparer through python-redlines, a third-party reference point."""
    out = _python_redlines_engine().run_redline("kyralaw-redline-bench", old, new)[0]
    if not out:
        raise RuntimeError("python-redlines returned no document")
    return bytes(out)


OPTIONAL: dict[str, Algorithm] = {"python-redlines": _python_redlines}


def _from_folder(folder: Path) -> Callable[[str], bytes | None]:
    def read(pair: str) -> bytes | None:
        path = folder / f"{pair}.docx"
        return path.read_bytes() if path.is_file() else None

    return read


def run(
    data: Path,
    rows: list[dict[str, Any]],
    algorithm: Algorithm | None,
    redlines: Callable[[str], bytes | None] | None,
    save_to: Path | None,
    refs: dict[str, Reference],
) -> list[PairResult]:
    results: list[PairResult] = []
    for row in rows:
        pair = row["id"]
        folder = data / "pairs" / pair
        if pair not in refs:
            refs[pair] = read_reference(folder)
        seconds: float | None = None
        redline: bytes | None
        try:
            if algorithm is not None:
                old, new = (folder / "old.docx").read_bytes(), (folder / "new.docx").read_bytes()
                started = time.perf_counter()
                redline = algorithm(old, new)
                seconds = time.perf_counter() - started
            else:
                assert redlines is not None
                redline = redlines(pair)
                if redline is None:
                    results.append(PairResult(pair=pair, scored=False, error="no redline file"))
                    continue
            result = score_redline(pair, redline, refs[pair], seconds)
        except Exception as exc:  # one pair's failure is reported, not fatal
            results.append(PairResult(pair=pair, scored=False, error=f"{type(exc).__name__}: {exc}"[:300]))
            continue
        if save_to is not None and algorithm is not None:
            save_to.mkdir(parents=True, exist_ok=True)
            (save_to / f"{pair}.docx").write_bytes(redline)
        results.append(result)
    return results


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
    parser.add_argument("data", type=Path, help="the dataset folder (dev.jsonl, test.jsonl, pairs/)")
    parser.add_argument("--split", choices=("test", "dev"), default="test")
    parser.add_argument(
        "-a", "--algorithm", action="append", choices=sorted({*ALGORITHMS, *OPTIONAL}), help="repeatable"
    )
    parser.add_argument(
        "--redlines", action="append", default=[], metavar="NAME=DIR", help="score DIR/<pair id>.docx as NAME"
    )
    parser.add_argument("--limit", type=int, help="score only the first N pairs, for a quick look")
    parser.add_argument("--json", type=Path, help="write every pair's scores here")
    parser.add_argument("--save-redlines", type=Path, help="keep the redlines the algorithms wrote, per algorithm")
    args = parser.parse_args()

    data: Path = args.data.expanduser().resolve()
    split_file = data / f"{args.split}.jsonl"
    if not split_file.is_file():
        sys.exit(f"{split_file} not found; DATA is the kyralaw-redline-bench folder")
    rows = [json.loads(line) for line in split_file.read_text().splitlines() if line.strip()]
    if args.limit:
        rows = rows[: args.limit]

    names = args.algorithm or ([] if args.redlines else list(ALGORITHMS))
    refs: dict[str, Reference] = {}
    all_results: dict[str, list[PairResult]] = {}
    for name in names:
        algorithm = ALGORITHMS.get(name) or OPTIONAL[name]
        save_to = args.save_redlines / name if args.save_redlines else None
        all_results[name] = run(data, rows, algorithm, None, save_to, refs)
        print(f"{name}: done", file=sys.stderr, flush=True)
    for spec in args.redlines:
        name, sep, folder = spec.partition("=")
        if not sep or not name or not folder:
            sys.exit(f"--redlines takes NAME=DIR, not {spec!r}")
        all_results[name] = run(data, rows, None, _from_folder(Path(folder).expanduser()), None, refs)

    summaries = {name: summarise(results) for name, results in all_results.items()}
    print(report(summaries, args.split, len(rows)))
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
