# SPDX-License-Identifier: MIT
# Copyright (c) 2026 Kyra Law Limited
"""Where kyralaw-redline-bench comes from: Hugging Face, at a pinned revision, cached locally.

Usage::

    python eval/bench_data.py               # download once into the Hugging Face cache; print the folder
    python eval/bench_data.py --to /data    # also copy it to a plain folder (the Docker image does this)
    python eval/bench_data.py --check DIR   # check a folder against its own SHA256SUMS

The download goes through the standard Hugging Face cache (``~/.cache/huggingface``,
or ``$HF_HOME``), so it happens once per revision: later calls read the cache, and work
offline. Every file is then checked against the dataset's ``SHA256SUMS``; a missing or
altered file stops everything rather than being scored.

``REVISION`` is the Hugging Face commit ``bench/results.md`` was scored on. Move it only
together with that table, or the comparison in ``scripts/verify.sh`` compares against
different pairs.
"""

from __future__ import annotations

import argparse
import hashlib
import shutil
import sys
from pathlib import Path

REPO_ID = "kyralaw/kyralaw-redline-bench"
REVISION = "15e05ec5f189476ac8a4ea31e1f99b54ae438e7f"


def verify(folder: Path) -> int:
    """Check every file ``folder/SHA256SUMS`` lists; return how many. Stops on any mismatch."""
    sums = folder / "SHA256SUMS"
    if not sums.is_file():
        sys.exit(f"{sums} not found: {folder} is not a kyralaw-redline-bench release")
    problems: list[str] = []
    count = 0
    for line in sums.read_text().splitlines():
        if not line.strip():
            continue
        digest, _, name = line.partition("  ")
        path = folder / name
        count += 1
        if not path.is_file():
            problems.append(f"missing: {name}")
        elif hashlib.sha256(path.read_bytes()).hexdigest() != digest:
            problems.append(f"altered: {name}")
    if problems:
        shown = "\n  ".join(problems[:10])
        more = f"\n  ... and {len(problems) - 10} more" if len(problems) > 10 else ""
        sys.exit(f"{folder}: {len(problems)} of {count} files fail SHA256SUMS:\n  {shown}{more}")
    return count


def fetch() -> Path:
    """The dataset at REVISION, from the local Hugging Face cache, downloading it first if needed."""
    try:
        from huggingface_hub import snapshot_download
    except ImportError:
        sys.exit("huggingface_hub is not installed: run `uv sync`, or pass the dataset folder yourself")
    folder = Path(snapshot_download(repo_id=REPO_ID, repo_type="dataset", revision=REVISION))
    count = verify(folder)
    print(f"kyralaw-redline-bench {REVISION[:7]}: {count} files verified, cached in {folder}", file=sys.stderr)
    return folder


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--to", type=Path, help="copy the dataset to this folder, which must not exist yet")
    parser.add_argument("--check", type=Path, metavar="DIR", help="only check DIR against its SHA256SUMS")
    args = parser.parse_args()

    if args.check:
        print(f"{args.check}: {verify(args.check)} files verified", file=sys.stderr)
        return
    folder = fetch()
    if args.to:
        # The cache holds symlinks into its blob store; the copy holds the files themselves.
        shutil.copytree(folder, args.to, symlinks=False, ignore=shutil.ignore_patterns(".gitattributes"))
        print(f"{args.to}: {verify(args.to)} files verified", file=sys.stderr)
        folder = args.to
    print(folder)


if __name__ == "__main__":
    main()
