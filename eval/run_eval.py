"""Score redline algorithms against the synthetic corpus.

For every pair in ``eval/corpus/v1.jsonl`` this:

1. writes the old and new text as two real .docx files (python-docx, one
   paragraph per block);
2. runs each algorithm on those files, end to end, to a redline .docx;
3. reads the ``w:ins`` and ``w:del`` text back out of the file it wrote, and
   scores it against the recorded truth.

It scores the rendered file, not an intermediate diff, because the file is
what a user receives. A renderer that disagrees with the diff it was given is
a bug this design catches and a segment-level score cannot.

**Metric.** Inserted and deleted text are compared as bags of tokens
(lowercase words and single punctuation characters), giving precision, recall
and F1 for insertions and for deletions separately. Bags make the score
indifferent to *where* an ambiguous edit is placed (deleting one of two
adjacent "the"s is the same edit either way) while still penalising churn: a
redline that deletes and reinserts an unchanged word loses precision on both
sides. The headline score is 100 times the mean of the per-pair insert F1 and
delete F1, averaged over pairs.

**Integrity.** Each output is also checked to be a faithful redline: accepting
every change must give the new text exactly, and rejecting every change the
old text. A failure is reported and counted, never folded into the score.

**What the truth assumes.** A rewritten paragraph (``paragraph_replace``) is
recorded as a whole deletion plus a whole insertion, which is how Word reports
a paragraph replaced with unrelated content. An algorithm that keeps the odd
shared word ("the", "shall") as unchanged loses a little recall there; that is
a deliberate statement about readable legal redlines, not an accident.

Usage::

    python eval/run_eval.py                      # all algorithms, markdown report
    python eval/run_eval.py -a kyralaw-redline   # one algorithm
    python eval/run_eval.py --json results.json --save-redlines out/
"""

from __future__ import annotations

import argparse
import collections
import difflib
import io
import json
import statistics
import sys
import time
import warnings
import zipfile
from collections.abc import Callable
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Literal
from xml.etree import ElementTree

sys.path.insert(0, str(Path(__file__).resolve().parent))

import render
from generate import CORPUS_PATH, STYLES, score_tokens

from kyralaw_redline import (
    DiffSegment,
    InPlaceFallbackWarning,
    extract_paragraphs,
    extract_text,
    read_tracked_changes,
    redline_docx,
    render_segments_docx,
)
from kyralaw_redline.diff import _tokenize

FIXED_DATE = datetime(2026, 1, 1, tzinfo=timezone.utc)

Algorithm = Callable[[bytes, bytes], bytes]


# Algorithms ---------------------------------------------------------------


def _kyralaw_in_place(old_docx: bytes, new_docx: bytes) -> bytes:
    """The tracked changes written into the new document itself. Strict: a pair it
    cannot redline in place fails here rather than falling back unseen."""
    return redline_docx(old_docx, new_docx, date=FIXED_DATE, mode="in-place")


def _kyralaw_auto(old_docx: bytes, new_docx: bytes) -> bytes:
    """In place, rebuilt when in place is not possible: the library's default."""
    with warnings.catch_warnings():
        # kyralaw-redline-in-place refuses the same pairs, and that is where they are counted.
        warnings.simplefilter("ignore", InPlaceFallbackWarning)
        return redline_docx(old_docx, new_docx, date=FIXED_DATE, mode="auto")


def _kyralaw_rebuild(old_docx: bytes, new_docx: bytes) -> bytes:
    """The same diff written into a new, plain document."""
    return redline_docx(old_docx, new_docx, date=FIXED_DATE, mode="rebuild")


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


def _line_diff(old_docx: bytes, new_docx: bytes) -> bytes:
    """What a line diff gives: any changed paragraph is deleted and
    reinserted whole."""
    old = extract_text(old_docx).splitlines(keepends=True)
    new = extract_text(new_docx).splitlines(keepends=True)
    return render_segments_docx(_opcode_segments(old, new, lambda s: not s.strip()), date=FIXED_DATE)


def _token_diff(old_docx: bytes, new_docx: bytes) -> bytes:
    """A plain token-level difflib over the whole document, with none of the
    paragraph alignment or anchoring heuristics. Isolates what those add."""
    old = _tokenize(extract_text(old_docx))
    new = _tokenize(extract_text(new_docx))
    return render_segments_docx(_opcode_segments(old, new, str.isspace), date=FIXED_DATE)


ALGORITHMS: dict[str, Algorithm] = {
    "kyralaw-redline-in-place": _kyralaw_in_place,
    "kyralaw-redline-auto": _kyralaw_auto,
    "kyralaw-redline-rebuild": _kyralaw_rebuild,
    "token-diff": _token_diff,
    "line-diff": _line_diff,
}

# The kyralaw-redline algorithms the CI floor applies to.
OURS = ("kyralaw-redline-in-place", "kyralaw-redline-auto", "kyralaw-redline-rebuild")


# Scoring ------------------------------------------------------------------


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


@dataclass
class PairResult:
    id: str
    family: str
    edit_types: list[str]
    insert: PRF
    delete: PRF
    seconds: float
    integrity_errors: list[str] = field(default_factory=list)
    styles_kept: float = 0.0

    @property
    def score(self) -> float:
        return (self.insert.f1 + self.delete.f1) / 2


def _docx(
    blocks: list[str],
    styles: list[str] | None = None,
    cells: list[Any] | None = None,
    layout: dict[str, str] | None = None,
) -> bytes:
    """One version of a pair as a real document; see eval/render.py."""
    return render.render(blocks, styles, cells, layout)


def _version(pair: dict[str, object], side: str) -> bytes:
    """Render the old or new version and check it reads back as exactly the
    corpus text. A renderer that changed the text would change the ground
    truth under every algorithm, so it stops the run instead."""
    blocks = pair[side]
    assert isinstance(blocks, list)
    styles, cells, layout = pair.get(f"{side}_styles"), pair.get(f"{side}_cells"), pair.get("layout")
    assert styles is None or isinstance(styles, list)
    assert cells is None or isinstance(cells, list)
    assert layout is None or isinstance(layout, dict)
    data = _docx(blocks, styles, cells, layout)
    if extract_text(data) != "\n\n".join(b for b in blocks if b.strip()):
        raise RuntimeError(f"{pair['id']}: the rendered {side} version does not read back as the corpus text")
    return data


def _accepted_styles(docx_bytes: bytes) -> list[str]:
    """The style name of each paragraph in the accepted view, with the same
    merging as extract_paragraphs: a paragraph whose mark is a pending
    deletion runs on into the next one, which keeps the next one's style."""
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        root = ElementTree.fromstring(z.read("word/document.xml"))
        names: dict[str, str] = {}
        default = "Normal"
        if "word/styles.xml" in z.namelist():
            for st in ElementTree.fromstring(z.read("word/styles.xml")).iter(f"{w}style"):
                name = st.find(f"{w}name")
                sid = st.get(f"{w}styleId") or ""
                names[sid] = name.get(f"{w}val", sid) if name is not None else sid
                if st.get(f"{w}type") == "paragraph" and st.get(f"{w}default") == "1":
                    default = names[sid]
    out: list[str] = []
    body = root.find(f"{w}body")
    assert body is not None
    for p in body.iter(f"{w}p"):
        ppr = p.find(f"{w}pPr")
        if ppr is not None and ppr.find(f"{w}rPr/{w}del") is not None:
            continue
        pstyle = ppr.find(f"{w}pStyle") if ppr is not None else None
        out.append(names.get(pstyle.get(f"{w}val", ""), default) if pstyle is not None else default)
    return out


def styles_kept(docx_bytes: bytes, expected: list[str]) -> float:
    """Share of the new version's paragraphs whose style the redline keeps,
    position by position in the accepted view. Informational: not part of
    the score, which measures text."""
    if not expected:
        return 1.0
    got = _accepted_styles(docx_bytes)
    want = [STYLES[s] for s in expected]
    return sum(1 for g, e in zip(got, want, strict=False) if g == e) / len(want)


def _paragraphs(docx: bytes, view: Literal["accepted", "original"]) -> list[str]:
    """A view's paragraphs, empty ones included. The one exception: a Word
    document always holds at least one paragraph, so a view with no text at
    all (every paragraph inserted, or every one deleted) reads as one empty
    paragraph, which stands for an empty document rather than a defect."""
    paras = extract_paragraphs(docx, view=view)
    return [] if paras == [""] else paras


def score_pair(pair: dict[str, object], algorithm: Algorithm, save_to: Path | None) -> PairResult:
    old = pair["old"]
    new = pair["new"]
    assert isinstance(old, list) and isinstance(new, list)
    new_styles = pair.get("new_styles")
    assert new_styles is None or isinstance(new_styles, list)
    old_docx, new_docx = _version(pair, "old"), _version(pair, "new")

    start = time.perf_counter()
    out = algorithm(old_docx, new_docx)
    seconds = time.perf_counter() - start

    inserted, deleted = read_tracked_changes(out)
    expected_inserted = pair["expected_inserted"]
    expected_deleted = pair["expected_deleted"]
    assert isinstance(expected_inserted, list) and isinstance(expected_deleted, list)

    errors = []
    # Paragraph by paragraph, empty ones included: a deleted paragraph that
    # leaves an empty one behind when accepted is a defect, not a match.
    if _paragraphs(out, "accepted") != new:
        errors.append("accepting all changes does not give the new text")
    if _paragraphs(out, "original") != old:
        errors.append("rejecting all changes does not give the old text")

    if save_to is not None:
        target = save_to / f"{str(pair['id']).replace('/', '__')}.docx"
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(out)

    edit_types = pair["edit_types"]
    assert isinstance(edit_types, list)
    return PairResult(
        id=str(pair["id"]),
        family=str(pair["family"]),
        edit_types=[str(t) for t in edit_types],
        insert=prf(score_tokens(inserted), score_tokens(expected_inserted)),
        delete=prf(score_tokens(deleted), score_tokens(expected_deleted)),
        seconds=seconds,
        integrity_errors=errors,
        styles_kept=styles_kept(out, new_styles or ["normal"] * len(new)),
    )


# Reporting ----------------------------------------------------------------


def _mean(values: list[float]) -> float:
    return statistics.fmean(values) if values else float("nan")


def summarise(results: list[PairResult]) -> dict[str, float]:
    return {
        "score": 100 * _mean([r.score for r in results]),
        "insert_f1": _mean([r.insert.f1 for r in results]),
        "delete_f1": _mean([r.delete.f1 for r in results]),
        "insert_precision": _mean([r.insert.precision for r in results]),
        "delete_precision": _mean([r.delete.precision for r in results]),
        "pairs": float(len(results)),
        "integrity_failures": float(sum(1 for r in results if r.integrity_errors)),
        "styles_kept": _mean([r.styles_kept for r in results]),
        "seconds_total": sum(r.seconds for r in results),
    }


def report(all_results: dict[str, list[PairResult]], corpus: str) -> str:
    lines = [f"# Redline eval, corpus `{corpus}`", ""]
    lines += [
        "| Algorithm | Score | Insert F1 | Delete F1 | Insert P | Delete P | Integrity failures | Styles kept "
        "| Time (s) |",
        "| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |",
    ]
    for name, results in all_results.items():
        s = summarise(results)
        lines.append(
            f"| {name} | {s['score']:.2f} | {s['insert_f1']:.3f} | {s['delete_f1']:.3f} | "
            f"{s['insert_precision']:.3f} | {s['delete_precision']:.3f} | "
            f"{int(s['integrity_failures'])}/{int(s['pairs'])} | {s['styles_kept']:.1%} | {s['seconds_total']:.2f} |"
        )

    edit_types = sorted({t for rs in all_results.values() for r in rs if r.family == "single" for t in r.edit_types})
    names = list(all_results)
    lines += ["", "## Score by edit type (single-edit pairs)", ""]
    lines += ["| Edit type | " + " | ".join(names) + " |", "| --- |" + " ---: |" * len(names)]
    for t in edit_types:
        cells = [
            f"{summarise([r for r in all_results[n] if r.family == 'single' and r.edit_types == [t]])['score']:.2f}"
            for n in names
        ]
        lines.append(f"| {t} | " + " | ".join(cells) + " |")
    for family, label in (
        ("dense", "*dense (2 to 3 edits in each of 3 paragraphs)*"),
        ("mixed", "*mixed (4 to 8 edits)*"),
    ):
        cells = [f"{summarise([r for r in all_results[n] if r.family == family])['score']:.2f}" for n in names]
        lines.append(f"| {label} | " + " | ".join(cells) + " |")

    for name, results in all_results.items():
        worst = sorted(results, key=lambda r: r.score)[:5]
        lines += [
            "",
            f"## Lowest-scoring pairs: {name}",
            "",
            "| Pair | Score | Insert F1 | Delete F1 |",
            "| --- | ---: | ---: | ---: |",
        ]
        lines += [f"| `{r.id}` | {100 * r.score:.1f} | {r.insert.f1:.3f} | {r.delete.f1:.3f} |" for r in worst]
        failures = [r for r in results if r.integrity_errors]
        if failures:
            lines += ["", f"**Integrity failures ({len(failures)}):**", ""]
            lines += [f"- `{r.id}`: {'; '.join(r.integrity_errors)}" for r in failures]
    return "\n".join(lines) + "\n"


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument(
        "-a", "--algorithm", action="append", choices=sorted(ALGORITHMS), help="repeatable; default all"
    )
    parser.add_argument("--corpus", type=Path, default=CORPUS_PATH)
    parser.add_argument("--json", type=Path, help="write per-pair results here")
    parser.add_argument("--save-redlines", type=Path, help="write every redline .docx under this directory")
    parser.add_argument(
        "--min-score",
        type=float,
        help="exit 1 if either kyralaw-redline renderer scores below this, or has any integrity failure (for CI)",
    )
    args = parser.parse_args()

    pairs = [json.loads(line) for line in args.corpus.read_text(encoding="utf-8").splitlines() if line.strip()]
    if not pairs:
        sys.exit(f"{args.corpus} holds no pairs")

    names = args.algorithm or list(ALGORITHMS)
    all_results: dict[str, list[PairResult]] = {}
    for name in names:
        save_to = args.save_redlines / name if args.save_redlines else None
        all_results[name] = [score_pair(p, ALGORITHMS[name], save_to) for p in pairs]

    print(report(all_results, args.corpus.stem))

    if args.json:
        payload = {
            "corpus": args.corpus.stem,
            "summary": {n: summarise(rs) for n, rs in all_results.items()},
            "pairs": {n: [asdict(r) | {"score": r.score} for r in rs] for n, rs in all_results.items()},
        }
        args.json.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    if args.min_score is not None:
        ours = [name for name in OURS if name in all_results]
        if not ours:
            sys.exit(f"--min-score needs one of {', '.join(OURS)} among the algorithms")
        for name in ours:
            s = summarise(all_results[name])
            if s["integrity_failures"] or s["score"] < args.min_score:
                sys.exit(
                    f"{name} scored {s['score']:.2f} (minimum {args.min_score}) "
                    f"with {int(s['integrity_failures'])} integrity failures"
                )


if __name__ == "__main__":
    main()
