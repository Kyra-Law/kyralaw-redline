# kyralaw-redline

Word-level redlines for legal documents, rendered as native Word tracked changes, with an
[open evaluation set](https://huggingface.co/datasets/kyralaw/kyralaw-redline-bench) to measure them.

```bash
pip install kyralaw-redline   # once published; for now: pip install git+https://github.com/Kyra-Law/kyralaw-redline
kyralaw-redline old.docx new.docx -o redline.docx                  # in place, keeping the formatting
kyralaw-redline old.docx new.docx -o redline.docx --mode rebuild   # a clean, plain document
```

```python
from kyralaw_redline import redline_docx, compute_redline_segments

out = redline_docx(open("old.docx", "rb").read(), open("new.docx", "rb").read(), author="Reviewer")

for seg in compute_redline_segments("The fee is £500.", "The fee is £750."):
    print(seg.kind, repr(seg.text))
# equal 'The fee is £'
# delete '500'
# insert '750'
# equal '.'
```

Pure Python. The only dependencies are `python-docx` and `lxml`. No Word, no LibreOffice, no .NET.

## Two renderers, one diff

The diff is computed once ([`diff.py`](src/kyralaw_redline/diff.py)) and written one of two ways,
chosen with `redline_docx(..., mode=...)`:

| Mode | Writes | Keeps the document's formatting | When to use |
| --- | --- | --- | --- |
| `"in-place"` (default) | the tracked changes into the new document's own XML, the way Word's Compare does ([`inplace.py`](src/kyralaw_redline/inplace.py)) | yes: styles, numbering, tables, headers, footers, comments, sections | the redline someone will read |
| `"rebuild"` | a new, plain document built from the two texts ([`docx_io.py`](src/kyralaw_redline/docx_io.py)) | no | a content-only redline, or when in place is not possible |

Before returning, the in-place renderer reads its own result back: accepting every change must give
the new document and rejecting every change the old one. A result that fails is never returned:
`InPlaceRedlineError` is raised instead, and nothing is rebuilt behind your back. Ask for
`"rebuild"` if a plain redline will do.

## Why

A line diff redlines a whole paragraph as deleted and reinserted when one word in it changes.
For contracts, where most edits are a word or a phrase inside an otherwise unchanged clause,
that is unreadable. This library diffs in two passes: lines first, so whole-paragraph insertions
and deletions stay whole, then words within each changed paragraph, so a one-word edit is a
one-word redline.

`difflib.SequenceMatcher` does not compute a minimal diff, and legal prose repeats short words and
phrases ("the", "thirty (30) days") often enough that it regularly anchors a match to the wrong
occurrence. Most of [`diff.py`](src/kyralaw_redline/diff.py) is a handful of bounded heuristics that
correct for that. Each has a regression test that fails when the heuristic is removed.

## API

| Function | Does |
| --- | --- |
| `compute_redline_segments(old, new)` | the diff: an ordered list of `DiffSegment(kind, text)`, lossless both ways |
| `redline_docx(old_docx, new_docx, *, author, date, mode)` | redline two `.docx` files; `mode` is `"in-place"` (default) or `"rebuild"` (above) |
| `redline_in_place(old_docx, new_docx, *, author, date, keep_pending)` | the in-place renderer itself; `keep_pending` keeps both documents' pending revisions for a revised document edited directly from the base |
| `build_redline_docx(old, new, *, author, date)` | the rebuild renderer: diff two texts and render a new `.docx` with `w:ins` / `w:del` |
| `align_paragraphs(old, new)` | the paragraph alignment both renderers use, with a word diff per edited pair |
| `RedlineEngine` | the interface every algorithm implements: a `name` and `redline(old_docx, new_docx) -> bytes`; see below |
| `InPlaceEngine`, `RebuildEngine` | this library's two modes as engines, named `kyralaw-redline` and `kyralaw-redline-rebuild` |
| `render_segments_docx(segments, *, author, date)` | render any segment list through the same writer |
| `extract_text(docx, *, view)` | a `.docx`'s body text, pending changes accepted (`view="accepted"`) or rejected (`"original"`) |
| `extract_paragraphs(docx, *, view)` | the same, one string per paragraph, empty ones included |
| `read_tracked_changes(docx)` | the text of every `w:ins` and `w:del`, e.g. to read a Word Compare output |
| `diff_paragraph`, `diff_blocks`, `added_text`, `removed_text` | lower-level pieces |

Segments are lossless: joining the `equal` and `delete` segments reproduces the old text exactly, and
joining `equal` and `insert` reproduces the new text.

## Redline engines

Every algorithm, this library's and every one it is compared with, is a `RedlineEngine`: two
`.docx` files in, one redline `.docx` out.

```python
from kyralaw_redline import InPlaceEngine, RedlineEngine

redline = InPlaceEngine(author="Reviewer")(old_docx, new_docx)  # bytes, bytes -> bytes


class MyEngine(RedlineEngine):
    name = "my-engine"  # lowercase words joined by hyphens, checked when the class is defined

    def redline(self, old_docx: bytes, new_docx: bytes) -> bytes: ...
```

Calling an engine checks that what came back is a `.docx`, so a broken engine fails loudly rather
than being scored as an empty redline. The benchmark runs engines and nothing else.

## Limitations

Both renderers:
- A moved paragraph shows as a deletion plus an insertion.
- There is no character-level diff inside a word: "Supplier" to "Suppliers" redlines the word.
- Formatting-only changes are not marked.
- Headers, footers, footnotes and text boxes are not compared.

In place: the new document is kept whole, but inside *deleted* text only the words and their run
formatting come back; images, hyperlink targets, fields, bookmarks and comment anchors in it do not.

Rebuild: the output is a blank document built from text. Run formatting, styles, list numbering,
tables (cells become paragraphs), headers, footers, footnotes, comments, images and hyperlinks are
not carried over.

## The benchmark: kyralaw-redline-bench

**kyralaw-redline-bench** is 733 pairs of synthetic contracts and other legal documents derived
from [Harvey LAB](https://github.com/harveyai/harvey-labs), each with Microsoft Word's own Compare
Documents of the two versions as the reference. 48 pairs are a `dev` split for tuning; scores are
reported on the other 685, `test`. The dataset is MIT licensed and published on Hugging Face as
[`kyralaw/kyralaw-redline-bench`](https://huggingface.co/datasets/kyralaw/kyralaw-redline-bench); its
card explains how it was made and what was left out. [`eval/bench.py`](eval/bench.py) scores it, with no
Word needed:

```bash
uv run python eval/bench.py                          # built-in algorithms, test
uv run python eval/bench.py --split dev
uv run python eval/bench.py --redlines mine=out/     # your redlines: out/<id>.docx
uv run python eval/bench.py -a word-compare          # check the scorer: 100.00
uv run python eval/bench.py -a docxodus -a python-redlines   # the third-party comparers
uv run python eval/bench.py -a docxodus-checked     # Docxodus, its redlines checked (below)
```

The first run downloads the dataset from Hugging Face, at the revision
[`eval/bench_data.py`](eval/bench_data.py) pins, into the local Hugging Face cache (about 320 MB);
later runs read the cache and need no network. Every file is checked against the dataset's
`SHA256SUMS` before anything is scored. To score a copy of your own, pass its folder:
`eval/bench.py path/to/kyralaw-redline-bench`.

**Scoring.** Each algorithm runs end to end, `.docx` in and `.docx` out. The text a redline marks
inserted and deleted is read back from the file it wrote and compared with the same read of Word's
reference, as bags of tokens (lowercase words and single punctuation marks), giving precision,
recall and F1 for insertions and deletions separately. Moved text counts as deleted where it was
and inserted where it went, in the reference and the candidate alike, so detecting moves is neither
rewarded nor penalised.

- **Fidelity**, 0 to 1: whether the redline is still a usable Word document. No markup printed as
  text, as many tables and the same mix of paragraph styles as Word's reference, and the new
  version's header and footer text kept.
- **Score** = 100 x the mean of insert F1, delete F1 and fidelity, averaged over pairs.
- **Integrity:** accepting every change must give the new version's text, and rejecting every
  change the old one. Failures are counted beside the score. A pair with no redline is reported and
  left out of the averages, never scored as zero.

**Results** (test, 685 pairs, one Apple silicon Mac; the `docxodus` rows scored at [`0f74a70`](https://github.com/Kyra-Law/kyralaw-redline/tree/0f74a70)):

| Algorithm | What it is | Score | Text F1 | Fidelity | Integrity failures | Median time |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `word-compare` | Word's own reference, scored as a candidate: a check on the scorer | 100.00 | 100.00 | 1.000 | 0 | |
| `docxodus` | [Docxodus](https://github.com/JSv4/Docxodus) 17.0.0 through [docx-scalpel](https://pypi.org/project/docx-scalpel/) 0.14.0, the comparison its `redline` command runs; one process serves every pair | 96.33 | 94.58 | 0.998 | 7 | 165 ms |
| `docxodus-checked` | `docxodus`, each redline read back and, if it does not give back both documents, compared again with only spaces as word separators (7 pairs) | 96.33 | 94.58 | 0.998 | 0 | 298 ms |
| `python-redlines` | [python-redlines](https://github.com/JSv4/Python-Redlines) 0.3.0, Open XML PowerTools' comparer | 94.31 | 91.84 | 0.992 | 0 | 2.58 s |
| **`kyralaw-redline`** | this library's default, `mode="in-place"`: all 685 pairs written in place | **94.08** | 91.34 | 0.996 | 0 | 233 ms |
| `kyralaw-native` | the first in-place renderer, frozen as a baseline ([`eval/kyralaw_native.py`](eval/kyralaw_native.py)); refuses 3 pairs, left out | 94.05 | 91.30 | 0.995 | 0 | 122 ms |
| `kyralaw-redline-rebuild` | this library, `mode="rebuild"` | 78.72 | 92.39 | 0.514 | 0 | 353 ms |
| `token-diff` | a token-level difflib, no alignment or heuristics | 62.60 | 68.02 | 0.518 | 0 | 1.57 s |
| `line-diff` | a line diff | 49.68 | 48.56 | 0.519 | 0 | 220 ms |

- **Docxodus leads, by two points**, and is the only algorithm with integrity failures: 7 pairs,
  one bug. An empty paragraph that became "---", with a new paragraph after it, is written as
  unchanged, so rejecting every change keeps the "---". `docxodus-checked` reads each redline back
  and compares again with only spaces as word separators when it fails, which fixes all 7 without
  moving the score; the reading costs about 130 ms a pair.
- **`kyralaw-redline` and `kyralaw-native` mark the same changes.** On the 682 pairs
  `kyralaw-native` writes in place they score 94.07 and 94.05.
- **Only `kyralaw-native` refuses.** On 3 pairs, tables where a row is deleted or added and the
  same value recurs in other rows, its own accept/reject check fails and it produces nothing.
  `kyralaw-redline` keeps every pair inside corresponding table rows and writes those 3 in place
  too (about 96 each).
- **`kyralaw-redline-rebuild` marks the text about a point better than `kyralaw-redline`**,
  though they share one diff. They read the documents differently: rebuild diffs the text
  `extract_text` gives, in-place the paragraphs of the documents themselves, table cells included.
  Which reading costs what is worth a closer look.
- **Fidelity decides the ranking.** Rebuild writes into a blank document, so tables, headers and
  footers do not survive; in place they all do.
- **Times** are from producing each redline from the two files, on one machine. Treat them as
  relative.

To score your own algorithm, subclass `RedlineEngine` (above) and add an instance to `ENGINES` in
[`eval/engines.py`](eval/engines.py), or write its redlines to a folder and pass `--redlines`.

**Not compared: the commercial comparers.** The comparers law firms buy, such as Litera Compare and
Draftable, are not in these results, because their licences stand in the way of publishing an
independent evaluation. Litera's forbids releasing the results of any performance or functional
evaluation of its software without Litera's prior written approval
([Master Terms](https://www.litera.com/sites/default/files/file/2025-05/Litera-Master-Terms.pdf),
clause 3.3(vii)). Draftable's has no clause on evaluations, but limits its software to the
customer's internal business purposes and a trial to internal evaluation
([Standard Terms](https://www.draftable.com/legal/product-terms), clauses 1 and 5), so publishing
its scores would need Draftable's permission as well.

## Building the benchmark

[`bench/`](bench) is how kyralaw-redline-bench was made, start to finish: Word, driven out of
sight, rejects and accepts each Harvey LAB markup and compares the two versions, every reference
is checked to read back exactly, and `release.py` writes the dataset as published. It needs macOS
with Microsoft Word. See [`bench/README.md`](bench/README.md).

## Development

```bash
uv sync
uv run pytest
uv run ruff format --check . && uv run ruff check . && uv run mypy
```

## Verify in a clean container

The [`Dockerfile`](Dockerfile) builds a Linux machine holding nothing but this repository, its
locked dependencies and kyralaw-redline-bench, and [`scripts/verify.sh`](scripts/verify.sh) checks
everything this README claims, with no network:

```bash
docker build -t kyralaw-redline .
docker run --rm --network none kyralaw-redline                              # everything, ~45 min
docker run --rm --network none -e KYRALAW_BENCH_SPLIT=dev kyralaw-redline   # the 48 dev pairs, a few minutes
```

- **The dataset:** the build downloads it from Hugging Face at the pinned revision, checks every
  file against its `SHA256SUMS` and puts it in the image at `/data`. Docker's build cache keeps the
  download, so a rebuild fetches it again only when the pinned revision changes. To score another
  copy, mount it over `/data`.
- **Isolation:** no other Kyra Law codebase is installed or referred to anywhere in the code.
- **Everything CI checks:** the commands under Development.
- **The benchmark:** every engine scores the 685 test pairs, in parallel, and each score must equal
  the one in [`bench/results.md`](bench/results.md), which was produced on macOS. Word's own
  reference must score 100.00, and none of this repository's engines may have an integrity failure;
  a third-party engine's are compared with the table, like its scores. On a 10-core machine this
  takes about 45 minutes, most of it python-redlines. The dev split runs without the comparison;
  `-e KYRALAW_BENCH=skip` skips scoring altogether, and says so.

CI builds the image on every push, on x86-64 Linux, which fetches and checks the dataset, and runs
everything except the scoring.

## Licence

MIT. See [LICENSE](LICENSE). Copyright (c) 2026 Kyra Law Limited.
