# CLAUDE.md

Guidance for Claude Code working in this repository. This repository is intended to be public.

## Rules

1. **No client or confidential material, ever.** No document text, party names, matter
   references or figures from real agreements, in code, comments, docstrings, tests, fixtures
   or commit messages, and no mention of private document sets or what was measured on them.
   The one deliberate exception is the dataset card's note (`bench/release.py`, "What it does
   not test") that an internal set of real contracts ranked the algorithms in the same order:
   keep it to a ranking and a range, with no counts, names or document detail.
   When a real document exposes a bug, reproduce it with invented text, and check that the new
   test fails with the fix removed.
2. **Do not push without an explicit instruction.** Committing locally is fine.
3. **No em dashes** in any output: code, comments, docs, commit messages.
4. **The renderers compute no diff of their own.** `docx_io.py` (rebuild) renders exactly the
   segments `compute_redline_segments` produces, and `inplace.py` (in-place) exactly what
   `align_paragraphs` produces; both live in `diff.py`. A second diff in either renderer makes the
   shipped file a different algorithm from the one the tests and the evaluation measure, and the
   two renderers would stop marking the same changes. The one liberty `inplace.py` takes is with
   tables (`_respect_tables`): a pair that crosses table rows is shown as a deletion plus an
   insertion, and paragraphs then left without a partner in corresponding cells are aligned again
   by `align_paragraphs` itself. Every word diff still comes from `diff.py`.
5. **Self-contained.** Everything needed to run every algorithm in the benchmark, and to build
   kyralaw-redline-bench (`bench/`), is in this repository. No code imports, calls or depends on
   any other Kyra Law codebase, and an algorithm that cannot be run from here is not in the
   published results.
6. **Never fail silently.** A skipped pair, an integrity failure or a missing input is reported
   with its magnitude, never folded into a score.

## The benchmark

- `bench/` builds kyralaw-redline-bench (macOS with Word); `eval/bench.py` scores it. A change to
  `bench/views.py` or `bench/build.py` must leave every existing pair's row unchanged, or be a new
  dataset version.
- Every algorithm the benchmark runs is a `RedlineEngine` (`src/kyralaw_redline/engine.py`): a
  lowercase-hyphenated `name` and `redline(old_docx, new_docx) -> bytes`. A new one is a subclass,
  listed in `ENGINES` in `eval/engines.py`.

## Checks

`scripts/verify.sh` runs everything below, and more, in the container (see the README). A change
that moves a kyralaw-redline-bench score must update `bench/results.md` in the same change: the
container compares every score with it.

**CI scores no redlines.** The dataset lives on Hugging Face, not here, so CI checks lint, types and tests
but scores nothing. A change to `diff.py` or `inplace.py` is scored by hand before it lands: `verify.sh` with the
dataset mounted, or `eval/bench.py` on the test split, compared with `bench/results.md`. Use the
`dev` split for tuning, never `test`.

```bash
uv run pytest
uv run ruff format --check . && uv run ruff check . && uv run mypy
```
