# kyralaw-redline

Word-level redlines for legal documents, rendered as native Word tracked changes, with an open
evaluation set to measure them.

```bash
pip install kyralaw-redline   # once published; for now: pip install git+https://github.com/Kyra-Law/kyralaw-redline
kyralaw-redline old.docx new.docx -o redline.docx                  # in place, rebuilding if it must
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
| `"in-place"` | the tracked changes into the new document's own XML, the way Word's Compare does ([`inplace.py`](src/kyralaw_redline/inplace.py)) | yes: styles, numbering, tables, headers, footers, comments, sections | the redline someone will read |
| `"rebuild"` | a new, plain document built from the two texts ([`docx_io.py`](src/kyralaw_redline/docx_io.py)) | no | a content-only redline; the fallback |
| `"auto"` (default) | in place, or rebuilt when in place is not possible, with an `InPlaceFallbackWarning` | when it can | the default |

Before returning, the in-place renderer reads its own result back: accepting every change must give
the new document and rejecting every change the old one. A result that fails is never returned:
`"in-place"` raises `InPlaceRedlineError`, and `"auto"` rebuilds instead and warns.

## Why

A line diff redlines a whole paragraph as deleted and reinserted when one word in it changes.
For contracts, where most edits are a word or a phrase inside an otherwise unchanged clause,
that is unreadable. This library diffs in two passes: lines first, so whole-paragraph insertions
and deletions stay whole, then words within each changed paragraph, so a one-word edit is a
one-word redline.

`difflib.SequenceMatcher` does not compute a minimal diff, and legal prose repeats short words and
phrases ("the", "thirty (30) days") often enough that it regularly anchors a match to the wrong
occurrence. Most of [`diff.py`](src/kyralaw_redline/diff.py) is a handful of bounded heuristics that
correct for that. Each was introduced for a failure seen on a real agreement, and each has a
regression test that fails when the heuristic is removed.

## API

| Function | Does |
| --- | --- |
| `compute_redline_segments(old, new)` | the diff: an ordered list of `DiffSegment(kind, text)`, lossless both ways |
| `redline_docx(old_docx, new_docx, *, author, date, mode)` | redline two `.docx` files; `mode` is `"auto"`, `"in-place"` or `"rebuild"` (above) |
| `redline_in_place(old_docx, new_docx, *, author, date, keep_pending)` | the in-place renderer itself; `keep_pending` keeps both documents' pending revisions for a revised document edited directly from the base |
| `build_redline_docx(old, new, *, author, date)` | the rebuild renderer: diff two texts and render a new `.docx` with `w:ins` / `w:del` |
| `align_paragraphs(old, new)` | the paragraph alignment both renderers use, with a word diff per edited pair |
| `RedlineEngine` | the interface every algorithm implements: a `name` and `redline(old_docx, new_docx) -> bytes`; see below |
| `InPlaceEngine`, `AutoEngine`, `RebuildEngine` | this library's three modes as engines |
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
from kyralaw_redline import RedlineEngine, AutoEngine

redline = AutoEngine(author="Reviewer")(old_docx, new_docx)  # bytes, bytes -> bytes


class MyEngine(RedlineEngine):
    name = "my-engine"  # lowercase words joined by hyphens, checked when the class is defined

    def redline(self, old_docx: bytes, new_docx: bytes) -> bytes: ...
```

Calling an engine checks that what came back is a `.docx`, so a broken engine fails loudly rather
than being scored as an empty redline. The evaluations run engines and nothing else.

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

## The evaluation

[`eval/`](eval) is a synthetic benchmark that anyone can run, with no client data and no Word
licence:

```bash
uv sync
uv run python eval/run_eval.py
```

**Corpus.** Seven invented agreements in [`eval/base/`](eval/base): an NDA, a services
agreement, SaaS terms, a commercial lease, a shareholder side letter, a software licence and board
resolutions. There are 265 edited copies of them in [`eval/corpus/v4.jsonl`](eval/corpus/v4.jsonl),
generated by [`eval/generate.py`](eval/generate.py) from scripted edits.

Because the edits are scripted, the ground truth is exact by construction: every edit records the
text it deleted and the text it inserted, and the generator checks that the recorded truth accounts
for every token. The corpus file is committed, and CI checks it is exactly what the generator
produces, so the benchmark never changes by accident.

**Documents that look like real contracts.** The styling of the versions was modelled on a private
set of 33 real negotiated agreements. Most of those have tables, half use multi-level legal
numbering, most are justified with explicit spacing, many carry comments, and a third have headers
and footers with page fields. [`eval/render.py`](eval/render.py) gives the synthetic documents the
same features:

| Feature | How it appears |
| --- | --- |
| Paragraph styles | Title (all caps), Subtitle, Heading 1 to 3, Clause, Sub-clause, Sub-sub-clause, List Bullet |
| Numbering | direct multi-level numbering on each paragraph, legal (1. / 1.1 / (a) / (i)) or simple (1. / (a) / (i)); numbers are never part of the text |
| Lists | bulleted lists at two levels (• and –), and itemised lists numbered 1) with a nested a) level that restart in each list |
| Tables | parties, fee schedules, particulars, service levels, directors; header rows repeat |
| Run formatting | defined terms bold; Acts and Latin phrases italic; placeholders such as `[NAME]` highlighted |
| Fields and structure | clause cross-references as REF fields to bookmarked headings; the agreement date in a content control; emails as hyperlinks; a page break before signatures |
| Page furniture | header, footer with PAGE and NUMPAGES fields, Word comments; font, size and layout set per document |

None of it changes the text a reader extracts. The eval checks that for every document it renders,
and stops if a rendered version does not read back as exactly the corpus text.

| Family | Pairs | What it holds |
| --- | ---: | --- |
| `single` | 195 | one edit type per pair, for a per-type breakdown |
| `dense` | 28 | two or three edits in each of three paragraphs, possibly adjacent |
| `mixed` | 42 | four to eight edits of random types |

Edit types:
- word replacement, number change, punctuation change;
- phrase insertion and deletion;
- paragraph insertion, deletion and replacement;
- a new section (heading plus clause);
- a table row inserted or deleted;
- a list item inserted or deleted. Later items renumber, but numbers are not text, so a correct
  redline shows nothing there; a tool that treats list numbers as text reports a change on every
  later item.
- a defined-term rename across the document.

Earlier versions stay committed and can be scored with `--corpus`; scores compare only within a
version:
- `v1`: 139 pairs, five unstyled documents;
- `v2`: 209 pairs, paragraph styles only;
- `v3`: 238 pairs, styled documents with tables but no itemised lists.

**Scoring.** Each algorithm runs end to end, `.docx` in and `.docx` out. The `w:ins` and `w:del`
text is read back from the file it wrote and compared with the truth as bags of tokens
(lowercase words and single punctuation marks), giving precision, recall and F1 for insertions and
deletions separately. The headline score is 100 times the mean of insert F1 and delete F1, averaged
over pairs. Two further checks are reported beside it:

- **Integrity:** accepting every change must give the new document paragraph for paragraph, and
  rejecting every change the old one.
- **Styles kept:** the share of the new document's paragraphs whose style survives in the
  redline's accepted view. This is informational and not part of the score, which measures text.

**Results** (v4 corpus):

| Algorithm | Score | Insert F1 | Delete F1 | Integrity failures | Styles kept |
| --- | ---: | ---: | ---: | ---: | ---: |
| `kyralaw-redline-in-place` | 98.92 | 0.988 | 0.990 | 0/265 | 83.9% |
| `kyralaw-redline-auto` | 98.92 | 0.988 | 0.990 | 0/265 | 83.9% |
| `kyralaw-redline-rebuild` | 98.88 | 0.988 | 0.990 | 0/265 | 8.9% |
| `kyralaw-native`: the first in-place renderer, frozen (refuses 1 pair, left out) | 99.10 | 0.990 | 0.992 | 0/264 | 83.9% |
| `token-diff`: a token-level difflib, no alignment or heuristics | 91.45 | 0.894 | 0.935 | 0/265 | 8.9% |
| `line-diff`: a line diff | 52.68 | 0.530 | 0.523 | 0/265 | 8.9% |

The two kyralaw-redline renderers mark the same changes, since they share one diff; the in-place
one never had to fall back on this corpus, so `kyralaw-redline-auto` is the same. `kyralaw-native`
marks slightly better here: its first paragraph alignment has a word-level similarity check the
current one lacks, which is worth bringing across. It also refuses a table row deletion that the
current renderer handles.

The weak spot is **dense edits (92.50)**: adjacent edits in one paragraph still sometimes make the
merge heuristics replace a whole span instead of making the small edits within it.

Table row deletions (`row_delete`) used to be a second one, at 79.78: in a table of short, similar
rows, deleting one row paired each remaining row with its neighbour. The cause was the same as a
failure on a real negotiated agreement, a recurring line anchoring the line-level diff in the wrong
place, and anchoring on lines unique to both versions fixed both (100.00 since).

**Styles kept** is 8.9% for everything that writes into a blank document: only the plain
paragraphs match, and tables come out as paragraphs. The in-place renderer keeps the new document's
own paragraphs, so most of their styles survive (83.9%).

**What the scores don't tell you.** The corpus is synthetic: invented agreements with scripted
edits. Real negotiated documents contain moves, renumbering, formatting-only changes and editors'
idiosyncrasies that it does not model. Treat a score here as a regression signal and a way to
compare algorithms, not as the accuracy you will see on your own documents. On a private set of
32 real document pairs, scored against Word's Compare output, the same code scores in the low to
mid 90s.

To add an algorithm, subclass `RedlineEngine` (above) and add an instance to `ENGINES` in
[`eval/run_eval.py`](eval/run_eval.py); `eval/bench.py` picks it up too.

## The Word benchmark: kyralaw-redline-bench

The evaluation above uses documents generated here. [`eval/bench.py`](eval/bench.py) scores
redlines on documents Word itself compared: **kyralaw-redline-bench**, 733 pairs of synthetic
contracts and other legal documents derived from [Harvey LAB](https://github.com/harveyai/harvey-labs),
each with Microsoft Word's own Compare Documents of the two versions as the reference. 48 pairs are
a `dev` split for tuning; scores are reported on the other 685, `test`. The dataset is MIT licensed
and published separately; its card explains how it was made and what was left out.

```bash
uv run python eval/bench.py path/to/kyralaw-redline-bench                      # built-in algorithms, test
uv run python eval/bench.py path/to/kyralaw-redline-bench --split dev
uv run python eval/bench.py path/to/kyralaw-redline-bench --redlines mine=out/  # your redlines: out/<id>.docx
uv run python eval/bench.py path/to/kyralaw-redline-bench -a word-compare       # check the scorer: 100.00
```

It scores the same way as `run_eval.py`, with three differences that real Word documents need.
Moved text counts as deleted where it was and inserted where it went, for Word's reference and the
candidate alike. A paragraph whose mark was inserted or deleted runs on into the next, as in Word.
And the score includes **fidelity**, whether the redline is still a usable Word document: no markup
printed as text, tables and paragraph styles as in Word's reference, and the header and footer text
kept. Score = 100 x the mean of insert F1, delete F1 and fidelity.

**Results** (test, 685 pairs, one Apple silicon Mac):

| Algorithm | What it is | Score | Text F1 | Fidelity | Integrity failures | Median time |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `word-compare` | Word's own reference, scored as a candidate: a check on the scorer | 100.00 | 100.00 | 1.000 | 0 | |
| `python-redlines` | [python-redlines](https://github.com/JSv4/Python-Redlines) 0.3.0, Open XML PowerTools' comparer | 94.31 | 91.84 | 0.992 | 0 | 2.58 s |
| **`kyralaw-redline-auto`** | this library's default: in place, rebuilt only if that fails (it never did here) | **94.08** | 91.34 | 0.996 | 0 | 234 ms |
| `kyralaw-redline-in-place` | this library, `mode="in-place"`: all 685 pairs written in place | 94.08 | 91.34 | 0.996 | 0 | 233 ms |
| `kyralaw-native` | the first in-place renderer, frozen as a baseline ([`eval/kyralaw_native.py`](eval/kyralaw_native.py)); refuses 3 pairs, left out | 94.05 | 91.30 | 0.995 | 0 | 122 ms |
| `kyralaw-redline-rebuild` | this library, `mode="rebuild"` | 78.72 | 92.39 | 0.514 | 0 | 353 ms |
| `token-diff` | a token-level difflib, no alignment or heuristics | 62.60 | 68.02 | 0.518 | 0 | 1.57 s |
| `line-diff` | a line diff | 49.68 | 48.56 | 0.519 | 0 | 220 ms |

- **`kyralaw-redline-in-place` and `kyralaw-native` mark the same changes.** On the 682 pairs
  `kyralaw-native` writes in place they score 94.07 and 94.05.
- **Only `kyralaw-native` refuses.** On 3 pairs, tables where a row is deleted or added and the
  same value recurs in other rows, its own accept/reject check fails and it produces nothing.
  `kyralaw-redline-in-place` keeps every pair inside corresponding table rows and writes those 3
  in place too (about 96 each), so `kyralaw-redline-auto` never needed its fallback here.
- **`kyralaw-redline-rebuild` marks the text about a point better than `kyralaw-redline-in-place`**,
  though they share one diff. They read the documents differently: rebuild diffs the text
  `extract_text` gives, in-place the paragraphs of the documents themselves, table cells included.
  Which reading costs what is worth a closer look.
- **Fidelity decides the ranking.** Rebuild writes into a blank document, so tables, headers and
  footers do not survive; in place they all do.
- **Times** are from producing each redline from the two files, on one machine. Treat them as
  relative.

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
uv run python eval/generate.py --check
```

## Licence

MIT. See [LICENSE](LICENSE). Copyright (c) 2026 Kyra Law Limited.
