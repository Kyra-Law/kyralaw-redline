# SPDX-License-Identifier: MIT
# Copyright (c) 2026 Kyra Law Limited
"""Build kyralaw-redline-bench's pairs from Harvey LAB's tracked-change markups.

Harvey LAB (https://github.com/harveyai/harvey-labs, MIT, Copyright (c) 2026 Harvey
AI) is a legal-agent benchmark whose task folders hold synthetic matter documents,
generated under lawyer review with fictional parties. Some of them carry tracked
changes: a counterparty's markup of a draft. Each distinct one is a redline pair:

- ``old.docx``: the markup with every change rejected, by Word itself;
- ``new.docx``: the markup with every change accepted, by Word itself;
- ``old_to_new.docx``: Word's own Compare of those two, the reference. LAB's own markup
  is not the reference: it was written by a program and is coarser than Word's (a
  changed figure comes out as the whole phrase deleted and reinserted), and the
  benchmark measures agreement with Word.

Every reference is checked: accepting all its changes must give exactly the new
version and rejecting them all the old one (read with ``views.py``), or the pair is
refused. Word stamps the user's name into files it saves, so all three are scrubbed
(``word.scrub``).

The pairs are written to ``--out`` (about 140 MB, so not committed); this folder keeps
``manifest.json`` (each markup's LAB path, sha256, commit and outcome) and Harvey's
licence. A pair already built from the same markup is kept, so a build resumes where
it stopped and is otherwise a no-op. ``release.py`` turns the pairs into the dataset.

Needs macOS with Microsoft Word, and a clone of LAB:

    git clone --depth 1 https://github.com/harveyai/harvey-labs.git
    python bench/build.py --lab harvey-labs --out work/pairs
    python bench/build.py --lab harvey-labs --out work/pairs --limit 5   # a quick trial

``bench/expand.sh start`` runs it in the background with Word hidden throughout, and
``status`` / ``stop`` report on and end it. One build at a time: a second exits at once
rather than race the first over ``manifest.json``. SIGTERM lets the pair in progress
finish, then stops.
"""

from __future__ import annotations

import argparse
import fcntl
import hashlib
import json
import re
import signal
import subprocess
import sys
import time
import zipfile
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))

import views
import word

HERE = Path(__file__).resolve().parent
MANIFEST = HERE / "manifest.json"
LICENCE = HERE / "LICENSE-harvey-labs"
# Pairs between restarts of Word, which slows down over a long session
# (word.recycle_word): 5s a pair fresh, over 20s some 200 pairs later.
RECYCLE_WORD_EVERY = 50
_TRACKED_RE = re.compile(rb"<w:(?:ins|del)\b")


def find_markups(lab: Path) -> list[tuple[str, str]]:
    """(path relative to the clone, sha256) for every distinct .docx with tracked insertions or deletions."""
    seen: dict[str, str] = {}
    for path in sorted((lab / "tasks").rglob("*.docx")):
        try:
            xml = zipfile.ZipFile(path).read("word/document.xml")
        except (zipfile.BadZipFile, KeyError):
            continue
        if _TRACKED_RE.search(xml):
            sha = hashlib.sha256(path.read_bytes()).hexdigest()
            seen.setdefault(sha, str(path.relative_to(lab)))
    return sorted(((rel, sha) for sha, rel in seen.items()), key=lambda x: x[0])


def _paragraphs(docx: bytes, view: views.View) -> tuple[list[str], list[views.Block]]:
    blocks = views.read_blocks(docx)
    return [p.strip() for p in views.view_paragraphs(blocks, view) if p.strip()], blocks


def build_pair(markup: bytes) -> tuple[dict[str, bytes], dict[str, Any]]:
    """The three documents and the corpus row for one markup, or ValueError."""
    old = word.resolve(markup, "reject")
    new = word.resolve(markup, "accept")
    _versions(old, new)
    ref = word.compare(old, new)
    return {"old.docx": old, "new.docx": new, "old_to_new.docx": ref}, pair_row(old, new, ref, word.word_version())


def _versions(old: bytes, new: bytes) -> tuple[list[str], list[str]]:
    """The two versions' paragraphs; ValueError when there is nothing to compare."""
    old_paras, _ = _paragraphs(old, "new")
    new_paras, _ = _paragraphs(new, "new")
    if not old_paras or not new_paras:
        raise ValueError("one version is empty, so this is a full draft, not a revision")
    if old_paras == new_paras:
        raise ValueError("accepting and rejecting give the same text; nothing to compare")
    return old_paras, new_paras


def pair_row(old: bytes, new: bytes, ref: bytes, word_version: str) -> dict[str, Any]:
    """A pair's corpus row, after checking ``ref`` reads back as exactly the two versions.

    Pure: no Word. ValueError when the check fails.
    """
    old_paras, new_paras = _versions(old, new)
    ref_new, blocks = _paragraphs(ref, "new")
    ref_old, _ = _paragraphs(ref, "old")
    if ref_new != new_paras:
        raise ValueError("Word's comparison does not give the new version when its changes are accepted")
    if ref_old != old_paras:
        raise ValueError("Word's comparison does not give the old version when its changes are rejected")
    segments = [p.segs for p, _ in views.paragraphs(blocks)]
    inserted = [t for segs in segments for k, t in segs if k == "i"]
    deleted = [t for segs in segments for k, t in segs if k == "d"]
    row = {
        "old": old_paras,
        "new": new_paras,
        "expected_inserted": inserted,
        "expected_deleted": deleted,
        "family": "real" if inserted or deleted else "control",
        "stats": {
            "paragraphs_old": len(old_paras),
            "paragraphs_new": len(new_paras),
            "tables": sum(1 for b in blocks if isinstance(b, views.Table)),
            "insertions": len(inserted),
            "deletions": len(deleted),
            "words_inserted": sum(len(views.words(t)) for t in inserted),
            "words_deleted": sum(len(views.words(t)) for t in deleted),
            "words_total": len(views.words(" ".join(old_paras + new_paras))),
        },
        "reference": {"source": "word", "word_version": word_version},
    }
    return row


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--lab", type=Path, required=True, help="a clone of github.com/harveyai/harvey-labs")
    parser.add_argument("--out", type=Path, required=True, help="where the pairs are built, one folder each")
    parser.add_argument("--limit", type=int, default=None, help="build at most this many new pairs (a trial)")
    args = parser.parse_args()

    lab: Path = args.lab.expanduser().resolve()
    out: Path = args.out.expanduser().resolve()
    commit = subprocess.run(["git", "-C", str(lab), "rev-parse", "HEAD"], capture_output=True, text=True, check=True)
    lab_commit = commit.stdout.strip()

    out.mkdir(parents=True, exist_ok=True)
    lock = (out / ".build.lock").open("w")
    try:
        fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except BlockingIOError:
        raise SystemExit(f"another build of {out} is running; one at a time, they share manifest.json") from None
    # SIGTERM (expand.sh stop) lets the pair in progress finish, then stops. Cutting
    # a Word call off mid-operation is what left Word, on 2026-11-20, behind a hidden
    # "save changes?" prompt for a staged file, after which it opened nothing. A second
    # SIGTERM stops at once, as SystemExit, so the staged files are still removed.
    stop_requested = False

    def _on_term(*_: object) -> None:
        nonlocal stop_requested
        if stop_requested:
            sys.exit(143)
        stop_requested = True
        print(f"{_now()} stop requested; finishing the pair in progress first", flush=True)

    signal.signal(signal.SIGTERM, _on_term)
    word_was_running = word.word_running()
    if stale := word.clear_stale_staging():
        print(f"{_now()} removed {stale} staged file(s) an earlier, killed run left behind", flush=True)

    manifest: dict[str, Any] = json.loads(MANIFEST.read_text()) if MANIFEST.exists() else {"pairs": {}}
    by_sha = {entry["sha256"]: pid for pid, entry in manifest["pairs"].items()}
    next_n = max((int(pid.removeprefix("lab-")) for pid in manifest["pairs"]), default=0) + 1

    markups = find_markups(lab)
    todo = sum(1 for _, sha in markups if manifest["pairs"].get(by_sha.get(sha, ""), {}).get("status") is None)
    if args.limit is not None:
        todo = min(todo, args.limit)
    print(
        f"{_now()} {len(markups)} distinct tracked-change markups in LAB at {lab_commit[:9]}; {todo} to build",
        flush=True,
    )
    canary = _canary(out, manifest)
    built = failed = skipped = 0
    try:
        for rel, sha in markups:
            if stop_requested:
                print(f"{_now()} stopped on request", flush=True)
                break
            pid = by_sha.get(sha)
            if pid is None:
                pid = f"lab-{next_n:03d}"
                next_n += 1
                by_sha[sha] = pid
            entry = manifest["pairs"].setdefault(pid, {})
            entry.update({"source": rel, "sha256": sha, "lab_commit": entry.get("lab_commit", lab_commit)})
            target = out / pid
            if (target / "row.json").exists() and entry.get("status") == "built":
                skipped += 1
                continue
            if entry.get("status") == "refused":
                skipped += 1
                continue
            if args.limit is not None and built + failed >= args.limit:
                continue
            started = time.monotonic()
            try:
                files, row = _build_checked((lab / rel).read_bytes(), canary)
            except (ValueError, word.WordCompareError) as exc:
                entry.update({"status": "refused", "reason": str(exc)[:200]})
                failed += 1
                outcome = f"refused, {exc}"
            else:
                target.mkdir(parents=True, exist_ok=True)
                for name, data in files.items():
                    (target / name).write_bytes(data)
                (target / "row.json").write_text(json.dumps({"id": pid, **row}, ensure_ascii=False))
                entry.update({"status": "built", "words_total": row["stats"]["words_total"]})
                entry.pop("reason", None)
                built += 1
                canary = canary or files["new.docx"]
                outcome = "built"
            MANIFEST.write_text(json.dumps(manifest, indent=2, sort_keys=True) + "\n")
            done = built + failed
            print(f"{_now()} {pid} {outcome} in {time.monotonic() - started:.0f}s ({done} of {todo})", flush=True)
            if done % RECYCLE_WORD_EVERY == 0:
                memory = word.word_memory_gb()
                if memory is not None:
                    recycled = word.recycle_word()
                    print(
                        f"{_now()} Word at {memory:.1f} GB after {RECYCLE_WORD_EVERY} pairs: "
                        + ("restarted" if recycled else "not restarted, you are using it or it would not quit"),
                        flush=True,
                    )
    finally:
        rows = [
            json.loads((out / pid / "row.json").read_text())
            for pid in sorted(manifest["pairs"])
            if manifest["pairs"][pid].get("status") == "built" and (out / pid / "row.json").exists()
        ]
        (out / "corpus.jsonl").write_text("".join(json.dumps(r, ensure_ascii=False) + "\n" for r in rows))
        (out / "LICENSE-harvey-labs").write_text(LICENCE.read_text())
        refused = sum(1 for e in manifest["pairs"].values() if e.get("status") == "refused")
        print(
            f"{_now()} built {built}, kept {skipped}, refused now {failed}; "
            f"{len(rows)} pairs in {out}, {refused} refused in all",
            flush=True,
        )
        if not word_was_running:
            word.quit_word()


def _canary(out: Path, manifest: dict[str, Any]) -> bytes | None:
    """The smallest built pair's new version: a document Word is known to handle."""
    built = [out / pid / "new.docx" for pid, e in manifest["pairs"].items() if e.get("status") == "built"]
    present = [p for p in built if p.is_file()]
    return min(present, key=lambda p: p.stat().st_size).read_bytes() if present else None


def _word_handles(canary: bytes) -> bool:
    try:
        word.resolve(canary, "accept")
    except word.WordCompareError:
        return False
    return True


def _build_checked(markup: bytes, canary: bytes | None) -> tuple[dict[str, bytes], dict[str, Any]]:
    """``build_pair``, refusing a file only for what is wrong with the file.

    A Word error is first checked against ``canary``, a document Word has built
    before. If Word fails that too, the fault is Word's state, not the file's: Word is
    restarted and the file tried once more. If Word still fails the canary after a
    restart, the build stops (WordUnavailableError) rather than refuse every
    remaining file. On 2026-11-20 a Word stuck behind a hidden dialog had 33 good files
    refused in 16 seconds before this check existed.
    """
    try:
        return build_pair(markup)
    except word.UnreadableDocumentError:
        raise  # Word named the file as the problem; nothing to check
    except word.WordCompareError as exc:
        if canary is None or _word_handles(canary):
            # Word is fine now, but may not have been a moment ago: lab-753 failed to
            # open once and built in 4s on a retry (2026-11-20). One more try decides.
            print(f"{_now()} retrying once after: {str(exc)[:120]}", flush=True)
            return build_pair(markup)
    print(f"{_now()} Word failed a document it has built before; restarting it and retrying", flush=True)
    word.restart_word()
    if not _word_handles(canary):
        raise word.WordUnavailableError(
            "Word cannot handle a document it built before, even after a restart; stopping rather than "
            "refusing files for Word's state"
        )
    return build_pair(markup)


def _now() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S")


if __name__ == "__main__":
    main()
