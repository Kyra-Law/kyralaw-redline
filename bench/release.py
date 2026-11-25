# SPDX-License-Identifier: MIT
# Copyright (c) 2026 Kyra Law Limited
"""Turn the pairs ``build.py`` made into kyralaw-redline-bench, as published.

    python bench/release.py --lab harvey-labs --pairs work/pairs --out work/release/v1.0.0 \
        --results bench/results.md

writes:

    README.md        the dataset card (Hugging Face front matter included)
    LICENSE          MIT, Harvey AI's notice for the source documents and Kyra Law's
    dev.jsonl        48 pairs for tuning, one JSON row each
    test.jsonl       the rest, for reporting
    refused.jsonl    every LAB markup that did not become a pair, and why
    pairs/krb-NNN/   old.docx, new.docx, reference.docx (Word's Compare of the two)
    SHA256SUMS       every file above

**Ids and splits never move.** ``release_ids.json`` gives each source markup, by
sha256, its ``krb-NNN`` id and its split, the first time it is released. A later
release adds new pairs at the end, always to ``test``, and never renumbers, reassigns
or reuses an id, so a score on ``dev`` or ``test`` means the same pairs in every
version. The first release drew its 48 dev pairs at random with a fixed seed.

**Each document is republished, not copied.** Who last saved it is set to
``kyralaw-redline-bench`` and its created date to the same fixed date as its modified
date, so the files say nothing about when or by whom the build ran. Before that,
every part of every file is checked for the account's user name and full name, for
"kyra" (other than the publisher name itself), and for an Office sensitivity label;
any hit stops the release. Nothing else in a document is changed.

The release is written to a temporary folder and moved into place only when complete;
an existing release folder is never overwritten.
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
import os
import pwd
import random
import re
import shutil
import statistics
import subprocess
import sys
import zipfile
from collections import Counter
from pathlib import Path
from typing import Any

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import build  # noqa: E402

NAME = "kyralaw-redline-bench"
VERSION = "1.0.0"
IDS = HERE / "release_ids.json"
DEV_SIZE = 48
SPLIT_SEED = 20261120
FIXED_DATE = "2026-01-01T00:00:00Z"
LAB_REPOSITORY = "https://github.com/harveyai/harvey-labs"
SCORER_REPOSITORY = "https://github.com/Kyra-Law/kyralaw-redline"
# The tag on SCORER_REPOSITORY holding the exact code that produced this version.
CODE_TAG = f"bench-v{VERSION}"

# The card shows kyralaw-redline once, in place (its default), and Docxodus as it ships, beside the
# other algorithms.
# bench/results.md keeps every variant: the repository's README shows them all, and the
# container compares each one's score.
CARD_HIDDEN = frozenset({"docxodus-checked", "kyralaw-native", "kyralaw-redline-rebuild"})
CARD_NOTES = f"""- **Keeping the document matters.** `token-diff` and `line-diff` write into a blank document, so
  tables, headers and footers do not survive, and their fidelity is about half. `docxodus`,
  `python-redlines` and `kyralaw-redline` write into the new document itself, the way Word does.
- **Deletions are harder than insertions** for every algorithm, by 4 to 8 points of F1 at the top.
- **Integrity failures** are redlines whose changes, all accepted or all rejected, do not give back
  the two documents. Only `docxodus` has any: 7, all one bug, where an empty paragraph became "---".
  Checking each redline and comparing again fixes them without moving the score; see
  [the repository's README]({SCORER_REPOSITORY}#the-benchmark-kyralaw-redline-bench).
- **Times** are for producing each redline from the two files, on one machine. Treat them as
  relative.
- kyralaw-redline's rebuild renderer, and a frozen earlier version of it, are scored in
  [its README]({SCORER_REPOSITORY}/tree/{CODE_TAG}#the-benchmark-kyralaw-redline-bench)."""
FILES = {"old.docx": "old.docx", "new.docx": "new.docx", "old_to_new.docx": "reference.docx"}

_MIT = """MIT License

Copyright (c) 2026 Harvey AI
Copyright (c) 2026 Kyra Law Limited

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
"""


def _forbidden() -> list[bytes]:
    """What no released file may contain: who ran the build, and who we are."""
    entry = pwd.getpwuid(os.getuid())
    names = [entry.pw_name, entry.pw_gecos.split(",")[0].strip(), "kyra", "MSIP_Label"]
    return [n.lower().encode() for n in names if n]


def _republish(docx: bytes, forbidden: list[bytes], where: str) -> bytes:
    """``docx`` with the publisher and dates fixed; refuses one that names the build's owner."""
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(docx)) as src, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            # The publisher name build.py writes is ours to keep; everything else is checked.
            scanned = data.lower().replace(NAME.encode(), b"")
            hit = next((f for f in forbidden if f in scanned), None)
            if hit is not None:
                raise SystemExit(f"{where}: {info.filename} contains {hit.decode()!r}; nothing released")
            if info.filename == "docProps/core.xml":
                text = data.decode("utf-8")
                text = re.sub(
                    r"<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>",
                    f"<cp:lastModifiedBy>{NAME}</cp:lastModifiedBy>",
                    text,
                )
                text = re.sub(r"(<dcterms:(?:created|modified)[^>]*>)[^<]*(<)", rf"\g<1>{FIXED_DATE}\g<2>", text)
                data = text.encode("utf-8")
            elif info.filename.startswith("word/") and info.filename.endswith(".xml"):
                # Every revision's real time, which pairs built before word.scrub fixed it
                # still carry.
                data = re.sub(rb'w16du:dateUtc="[^"]*"', f'w16du:dateUtc="{FIXED_DATE}"'.encode(), data)
            entry = zipfile.ZipInfo(info.filename, date_time=(2026, 1, 1, 0, 0, 0))
            entry.compress_type = zipfile.ZIP_DEFLATED
            dst.writestr(entry, data)
    return out.getvalue()


def _assign_ids(built: list[tuple[str, dict[str, Any]]]) -> dict[str, dict[str, str]]:
    """{source sha256: {"id", "split"}}, extending release_ids.json without moving anything in it."""
    ids: dict[str, dict[str, str]] = json.loads(IDS.read_text())["pairs"] if IDS.exists() else {}
    first_release = not ids
    taken = [int(v["id"].removeprefix("krb-")) for v in ids.values()]
    next_n = max(taken, default=0) + 1
    new = [entry["sha256"] for _, entry in built if entry["sha256"] not in ids]
    for sha in new:
        ids[sha] = {"id": f"krb-{next_n:03d}", "split": "test"}
        next_n += 1
    if first_release:
        for sha in random.Random(SPLIT_SEED).sample(sorted(new), DEV_SIZE):
            ids[sha]["split"] = "dev"
    IDS.write_text(json.dumps({"pairs": ids}, indent=2, sort_keys=True) + "\n")
    return ids


def _refusal_kind(reason: str) -> str:
    if "unreadable content" in reason:
        return "unreadable"
    if "comparison does not give" in reason:
        return "reference_not_exact"
    return "other"


def card_results(results: str) -> str:
    """The card's results: the intro and table of ``bench/results.md`` without the rows in
    CARD_HIDDEN, then CARD_NOTES in place of its notes, which compare those rows."""
    lines = results.strip().splitlines()
    table = [i for i, line in enumerate(lines) if line.startswith("|")]
    if not table:
        raise SystemExit("the results file has no table")
    rows = {m[1]: i for i in table if (m := re.match(r"\| `([a-z0-9-]+)` \|", lines[i]))}
    missing = sorted(CARD_HIDDEN - set(rows))
    if missing:
        raise SystemExit(f"the results table has no row for {missing}; CARD_HIDDEN is out of date")
    hidden = {rows[name] for name in CARD_HIDDEN}
    kept = [lines[i] for i in range(table[0], table[-1] + 1) if i not in hidden]
    return "\n".join([*lines[: table[0]], *kept, "", CARD_NOTES])


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--out", type=Path, required=True, help="the release folder to write; never overwritten")
    parser.add_argument("--pairs", type=Path, required=True, help="the folder build.py wrote the pairs to")
    parser.add_argument(
        "--lab", type=Path, required=True, help="the Harvey LAB clone the pairs were built from, to count its documents"
    )
    parser.add_argument("--results", type=Path, help="a markdown results table to put in the card")
    args = parser.parse_args()

    source: Path = args.pairs.expanduser().resolve()
    target: Path = args.out.expanduser().resolve()
    if target.exists():
        raise SystemExit(f"{target} exists; a release is never overwritten")
    manifest: dict[str, Any] = json.loads(build.MANIFEST.read_text())["pairs"]
    commits = {e["lab_commit"] for e in manifest.values()}
    if len(commits) != 1:
        raise SystemExit(f"the pairs come from {len(commits)} LAB commits; a release is built from one")
    lab_commit = commits.pop()
    order = sorted(manifest, key=lambda pid: int(pid.removeprefix("lab-")))
    built = [(pid, manifest[pid]) for pid in order if manifest[pid].get("status") == "built"]
    refused = [(pid, manifest[pid]) for pid in order if manifest[pid].get("status") == "refused"]
    unfinished = [pid for pid in order if manifest[pid].get("status") not in ("built", "refused")]
    if unfinished:
        raise SystemExit(f"{len(unfinished)} markups are neither built nor refused; finish the build first")
    ids = _assign_ids(built)

    staging = target.with_name(target.name + ".partial")
    if staging.exists():
        shutil.rmtree(staging)
    (staging / "pairs").mkdir(parents=True)
    forbidden = _forbidden()
    splits: dict[str, list[dict[str, Any]]] = {"dev": [], "test": []}
    move_pairs = format_pairs = 0
    for pid, entry in built:
        krb = ids[entry["sha256"]]
        row = json.loads((source / pid / "row.json").read_text())
        folder = staging / "pairs" / krb["id"]
        folder.mkdir()
        for name, published in FILES.items():
            data = _republish((source / pid / name).read_bytes(), forbidden, f"{pid}/{name}")
            (folder / published).write_bytes(data)
            if published == "reference.docx":
                body = zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml")
                move_pairs += b"<w:moveFrom " in body
                format_pairs += b"<w:rPrChange " in body or b"<w:pPrChange " in body
        splits[krb["split"]].append(
            {
                "id": krb["id"],
                "split": krb["split"],
                "files": {k: f"pairs/{krb['id']}/{k}.docx" for k in ("old", "new", "reference")},
                "old": row["old"],
                "new": row["new"],
                "expected_inserted": row["expected_inserted"],
                "expected_deleted": row["expected_deleted"],
                "stats": row["stats"],
                "reference": {
                    "tool": "Microsoft Word Compare Documents",
                    "word_version": row["reference"]["word_version"],
                },
                "source": {
                    "repository": LAB_REPOSITORY,
                    "commit": lab_commit,
                    "path": entry["source"],
                    "sha256": entry["sha256"],
                },
            }
        )
    for split, rows in splits.items():
        rows.sort(key=lambda r: r["id"])
        (staging / f"{split}.jsonl").write_text("".join(json.dumps(r, ensure_ascii=False) + "\n" for r in rows))
    (staging / "refused.jsonl").write_text(
        "".join(
            json.dumps(
                {
                    "source": {
                        "repository": LAB_REPOSITORY,
                        "commit": lab_commit,
                        "path": e["source"],
                        "sha256": e["sha256"],
                    },
                    "kind": _refusal_kind(e["reason"]),
                    "reason": e["reason"],
                }
            )
            + "\n"
            for _, e in refused
        )
    )
    (staging / "LICENSE").write_text(_MIT)
    every = splits["dev"] + splits["test"]
    docx_total, tracked_files = _survey(args.lab.expanduser().resolve(), lab_commit)
    facts = {
        "docx_total": docx_total,
        "tracked_files": tracked_files,
        "results": card_results(args.results.read_text()) if args.results else None,
        "lab_commit": lab_commit,
        "markups": len(manifest),
        "pairs": len(every),
        "dev": len(splits["dev"]),
        "test": len(splits["test"]),
        "refused": Counter(_refusal_kind(e["reason"]) for _, e in refused),
        "word_versions": sorted({r["reference"]["word_version"] for r in every}),
        "words": sorted(r["stats"]["words_total"] for r in every),
        "tables": sum(r["stats"]["tables"] > 0 for r in every),
        "insert_share": statistics.median(
            r["stats"]["words_inserted"] / (r["stats"]["words_inserted"] + r["stats"]["words_deleted"])
            for r in every
            if r["stats"]["words_inserted"] + r["stats"]["words_deleted"]
        ),
        "move_pairs": move_pairs,
        "format_pairs": format_pairs,
        "areas": Counter(r["source"]["path"].split("/")[1] for r in every),
    }
    (staging / "README.md").write_text(_card(facts))
    sums = sorted(p for p in staging.rglob("*") if p.is_file())
    (staging / "SHA256SUMS").write_text(
        "".join(f"{hashlib.sha256(p.read_bytes()).hexdigest()}  {p.relative_to(staging)}\n" for p in sums)
    )
    staging.rename(target)
    size = sum(p.stat().st_size for p in target.rglob("*") if p.is_file()) // 2**20
    print(
        f"{NAME} v{VERSION}: {len(every)} pairs ({facts['dev']} dev, {facts['test']} test), "
        f"{len(refused)} refused, {size} MB in {target}"
    )


def _survey(lab: Path, lab_commit: str) -> tuple[int, int]:
    """(every .docx in LAB's task folders, those with tracked insertions or deletions)."""
    head = subprocess.run(["git", "-C", str(lab), "rev-parse", "HEAD"], capture_output=True, text=True, check=True)
    if head.stdout.strip() != lab_commit:
        raise SystemExit(f"{lab} is at {head.stdout.strip()[:9]}, but the pairs were built from {lab_commit[:9]}")
    total = tracked = 0
    for path in (lab / "tasks").rglob("*.docx"):
        total += 1
        try:
            xml = zipfile.ZipFile(path).read("word/document.xml")
        except (zipfile.BadZipFile, KeyError):
            continue
        tracked += bool(re.search(rb"<w:(?:ins|del)\b", xml))
    return total, tracked


def _refusal_table(f: dict[str, Any], refused: Counter[str]) -> str:
    """The card's table of what was left out; a reason with no cases gets no row."""
    rows = [
        (
            "Duplicates",
            f["tracked_files"] - f["markups"],
            "The same file in more than one task folder; each was kept once.",
        ),
        (
            "Word calls the file damaged",
            refused.get("unreadable", 0),
            'On opening, Word reports "unreadable content" and offers to recover the document. A recovered '
            "document is Word's reconstruction of what it could read, not the document LAB published, so the "
            "offer was declined and the file left out. Other tools may open these files, but without Word "
            "there is no reference to score against.",
        ),
        (
            "Word's comparison did not reproduce the versions",
            refused.get("reference_not_exact", 0),
            "The check in step 4 failed: the reference, with its changes accepted or rejected, did not read "
            "exactly as `new.docx` or `old.docx`.",
        ),
        ("Other", refused.get("other", 0), "See `refused.jsonl`."),
    ]
    lines = ["| Left out | Markups | Why |", "| --- | ---: | --- |"]
    lines += [f"| {what} | {n} | {why} |" for what, n, why in rows if n]
    return "\n".join(lines)


def _card(f: dict[str, Any]) -> str:
    words = f["words"]
    refused = f["refused"]
    areas = ", ".join(f"{area} ({n})" for area, n in f["areas"].most_common(6))
    return f"""---
license: mit
pretty_name: kyralaw-redline-bench
language:
- en
tags:
- legal
- redline
- document-comparison
- tracked-changes
- docx
- benchmark
size_categories:
- n<1K
configs:
- config_name: default
  data_files:
  - split: dev
    path: dev.jsonl
  - split: test
    path: test.jsonl
---

# kyralaw-redline-bench

A benchmark for **redlines**: given two versions of a contract as Word documents, mark what changed
as tracked changes. Each of its {f["pairs"]} pairs holds the two versions and **Microsoft Word's own
Compare Documents** of them, which is the reference an algorithm is measured against.

Published by [Kyra Law](https://kyralaw.co.uk), version {VERSION}. Derived from
[Harvey LAB]({LAB_REPOSITORY}) (MIT, Copyright (c) 2026 Harvey AI); not affiliated with or endorsed by
Harvey.

## Quick start

Score your own redlines, one Word document per pair with tracked changes, using the scorer in
[kyralaw-redline]({SCORER_REPOSITORY}). It needs [uv](https://docs.astral.sh/uv/), and no Word:

```bash
git clone {SCORER_REPOSITORY} && cd kyralaw-redline
uv run python eval/bench.py --redlines mine=out/   # scores out/krb-NNN.docx, the test split
uv run python eval/bench.py --split dev            # the built-in algorithms, on the dev split
```

The first run downloads this dataset into the local Hugging Face cache, at the revision the scorer
pins, and checks every file against `SHA256SUMS`; later runs need no network. To get just the files:

```bash
uvx --from huggingface_hub hf download kyralaw/kyralaw-redline-bench --repo-type dataset --local-dir kyralaw-redline-bench
```

## Contents

| File | What |
| --- | --- |
| `pairs/krb-NNN/old.docx` | the earlier version |
| `pairs/krb-NNN/new.docx` | the later version |
| `pairs/krb-NNN/reference.docx` | Word's Compare of `old.docx` against `new.docx`, with tracked changes |
| `dev.jsonl`, `test.jsonl` | one row per pair: the file paths, both versions as paragraph lists, the inserted and deleted text Word marked, statistics, and the source |
| `refused.jsonl` | every source markup that did not become a pair, and why |
| `SHA256SUMS` | a checksum for every file |

**Splits.** `dev` ({f["dev"]} pairs, drawn at random with a fixed seed) is for tuning and inspection;
report scores on `test` ({f["test"]} pairs). Ids and splits are fixed: a later version adds pairs to
`test` and never renumbers or moves one.

## How it was made

### Where the documents come from

[Harvey LAB]({LAB_REPOSITORY}) is a benchmark for legal AI agents. Each of its tasks comes with a
folder of matter documents: contracts, term sheets, policies, disclosure schedules, written as
synthetic documents with fictional parties, generated in batches under the review of lawyers.
Some of those documents are a counterparty's markup of a draft, saved as a Word file with tracked
changes. A markup holds both versions of a document at once: the draft (every change rejected) and
the counterparty's version (every change accepted). That makes each one a redline problem with a
known answer.

At commit `{f["lab_commit"][:9]}`, LAB holds {f["docx_total"]:,} `.docx` files. {f["tracked_files"]} of
them carry tracked insertions or deletions. Some tasks share a document, so those
{f["tracked_files"]} are {f["markups"]} distinct files; each distinct file was used once.

### From a markup to a pair

Every step is done by Microsoft Word itself (version {", ".join(f["word_versions"])} for macOS),
because the benchmark measures agreement with Word, and Word is the authority on what a markup
means: which paragraphs a deleted paragraph mark joins, what a deleted table row takes with it.

1. **`old.docx`**: Word opens the markup and rejects every change.
2. **`new.docx`**: Word opens the markup again and accepts every change.
3. **`reference.docx`**: Word's Compare Documents of `old.docx` against `new.docx`, the answer an
   algorithm is scored against.
4. **The check.** Accepting every change in `reference.docx` must give exactly the text of
   `new.docx`, paragraph by paragraph, and rejecting every change exactly the text of `old.docx`.
   A reference that fails this would score algorithms against something other than the two
   documents they were given, so a pair that fails it is left out.
5. **Clean-up.** Word writes the name of whoever ran it, and the time, into every file it saves.
   "Last modified by" is set to `kyralaw-redline-bench` and every date to a fixed one, so the files
   say nothing about how or when the build ran. Nothing else in a document is changed after Word
   saves it.

**Why not use the markup itself as the reference?** LAB's markups were written by a program, not
by Word, and mark changes more coarsely: a changed figure comes out as the whole phrase deleted and
reinserted, where Word marks only the words that changed. An algorithm that matched Word exactly
would lose points against them. Word's own comparison of the two versions is the reference a
lawyer would recognise.

### What was left out, and why

{f["pairs"]} of the {f["markups"]} distinct markups became pairs. `refused.jsonl` lists the
{sum(refused.values())} that did not, each with its source path and the reason.

{_refusal_table(f, refused)}

## What it tests

- **Length.** A median of {statistics.median(words):,.0f} words per pair across both versions (10th
  to 90th percentile: {words[len(words) // 10]:,} to {words[9 * len(words) // 10]:,}), so alignment
  is tested over real contract lengths.
- **Realistic edit balance.** Insertions make up a median {f["insert_share"]:.0%} of the changed
  words.
- **Tables.** {f["tables"]} pairs have at least one.
- **Moves and formatting.** Word marks text as moved in {f["move_pairs"]} references and formatting
  changes in {f["format_pairs"]}. The source markups contain only insertions and deletions; Word found
  these itself.
- **Documents:** mostly {areas}.

What it does not test: comments, and negotiations by real parties. The edits are synthetic, written
under lawyer review rather than negotiated. As a check, five algorithms were scored on an earlier,
smaller build of this benchmark and on an internal set of real negotiated contracts, which cannot be
published. They ranked in the same order on both, and each algorithm's two scores were within four
points of each other. That suggests a score here carries over to real agreements; it does not prove
it.

## Scoring

The scorer (`eval/bench.py`) and the code that built this dataset (`bench/`) are in
[kyralaw-redline]({SCORER_REPOSITORY}); tag [`{CODE_TAG}`]({SCORER_REPOSITORY}/tree/{CODE_TAG}) is the
exact code that built this version's pairs and scored the results below. Each redline
is read for the text it marks inserted and deleted, through the same parser as the reference, and
compared as bags of words: precision, recall and F1 for insertions and for deletions. Moved text
counts as deleted where it was and inserted where it went, in the reference and in the candidate
alike, so an algorithm is neither rewarded nor penalised for detecting moves. A fidelity score
checks that the redline is still a usable Word document (tables, paragraph styles, headers and
footers kept; no markup printed as text).

## Results

{f["results"] or "Results are added with the first published version."}

## Licence and citation

MIT, see `LICENSE`. The source documents are Copyright (c) 2026 Harvey AI under the MIT licence; the
versions, references and metadata derived from them are Copyright (c) 2026 Kyra Law Limited, under
the same licence.

Please cite this dataset with its version, and Harvey LAB with the commit above:

```bibtex
@misc{{kyralawredlinebench2026,
  title  = {{kyralaw-redline-bench: a document-comparison benchmark for legal redlines}},
  author = {{{{Kyra Law}}}},
  year   = {{2026}},
  note   = {{Version {VERSION}. Derived from Harvey LAB, commit {f["lab_commit"][:9]}}}
}}

@misc{{harveylab2026,
  title  = {{{{Harvey LAB}}: The Legal Agent Benchmark}},
  author = {{{{Harvey AI}}}},
  year   = {{2026}},
  url    = {{{LAB_REPOSITORY}}}
}}
```
"""


if __name__ == "__main__":
    main()
