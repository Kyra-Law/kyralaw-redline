"""``kyralaw-redline OLD.docx NEW.docx -o REDLINE.docx``"""

from __future__ import annotations

import argparse
from pathlib import Path

from kyralaw_redline.docx_io import DEFAULT_AUTHOR, redline_docx


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(
        prog="kyralaw-redline",
        description="Compare two .docx files and write a redline with native Word tracked changes.",
    )
    parser.add_argument("old", type=Path, help="the earlier version")
    parser.add_argument("new", type=Path, help="the later version")
    parser.add_argument("-o", "--output", type=Path, required=True, help="where to write the redline .docx")
    parser.add_argument("--author", default=DEFAULT_AUTHOR, help=f"revision author (default: {DEFAULT_AUTHOR})")
    parser.add_argument(
        "--mode",
        choices=("auto", "in-place", "rebuild"),
        default="auto",
        help="in-place: write into the new document, keeping its formatting; rebuild: a new, plain document; "
        "auto (default): in-place, rebuilding when that is not possible",
    )
    args = parser.parse_args(argv)

    out = redline_docx(args.old.read_bytes(), args.new.read_bytes(), author=args.author, mode=args.mode)
    args.output.write_bytes(out)
    print(f"wrote {args.output} ({len(out):,} bytes)")


if __name__ == "__main__":
    main()
