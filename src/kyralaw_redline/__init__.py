"""Word-level redlines for legal documents, rendered as native Word tracked changes."""

from kyralaw_redline.diff import (
    DiffSegment,
    ParagraphAlignment,
    added_text,
    align_paragraphs,
    compute_redline_segments,
    diff_blocks,
    diff_paragraph,
    removed_text,
)
from kyralaw_redline.docx_io import (
    InPlaceFallbackWarning,
    build_redline_docx,
    extract_paragraphs,
    extract_text,
    read_tracked_changes,
    redline_docx,
    render_segments_docx,
)
from kyralaw_redline.inplace import InPlaceRedlineError, redline_in_place

__all__ = [
    "DiffSegment",
    "InPlaceFallbackWarning",
    "InPlaceRedlineError",
    "ParagraphAlignment",
    "added_text",
    "align_paragraphs",
    "build_redline_docx",
    "compute_redline_segments",
    "diff_blocks",
    "diff_paragraph",
    "extract_paragraphs",
    "extract_text",
    "read_tracked_changes",
    "redline_docx",
    "redline_in_place",
    "removed_text",
    "render_segments_docx",
]
