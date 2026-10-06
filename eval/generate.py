"""Generate the synthetic redline corpus from the base documents.

Each pair is a base document (``eval/base/*.txt``, blocks separated by blank
lines) and a copy of it with scripted edits applied. Because the edits are
scripted, the ground truth needs no Word Compare and no human labelling:
every edit records exactly the text it deleted and the text it inserted.

Edits never overlap, so that recorded truth stays exact. Each paragraph
takes at most one in-paragraph edit; a defined-term rename runs first and
protects the words it changed from later edits; a paragraph inserted by one
edit is never edited again; a deleted paragraph's earlier in-paragraph edits
are discarded, since the whole of its original text is what was deleted.

Each base block may start with a style tag (``{h1} Definitions``; see STYLES),
so pairs are styled documents and edits keep each paragraph's style.

Three families of pairs:

- ``single``: exactly one edit type per pair, so scores can be broken down by
  edit type.
- ``dense``: two or three in-paragraph edits in each of three paragraphs,
  possibly adjacent, which is where SequenceMatcher's anchoring goes wrong
  and the merge heuristics have to earn their keep.
- ``mixed``: four to eight edits of random types, closer to a real round of
  negotiation.

The output is committed (``eval/corpus/v2.jsonl``; ``v1.jsonl``, five
unstyled documents, stays as the frozen first version) and the evaluation reads
only that file, so a change to this generator never silently changes the
benchmark. ``--check`` verifies that the committed file is what this
generator produces, which CI runs.

Usage::

    python eval/generate.py            # rewrite eval/corpus/v2.jsonl
    python eval/generate.py --check    # exit 1 if it would change
"""

from __future__ import annotations

import argparse
import collections
import json
import random
import re
import sys
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path

CORPUS_VERSION = "v2"
EVAL_DIR = Path(__file__).resolve().parent
BASE_DIR = EVAL_DIR / "base"
CORPUS_PATH = EVAL_DIR / "corpus" / f"{CORPUS_VERSION}.jsonl"
EXTRA_CLAUSES = "_extra_clauses"

EDIT_TYPES = (
    "word_replace",
    "number_change",
    "punctuation_change",
    "phrase_insert",
    "phrase_delete",
    "paragraph_insert",
    "paragraph_delete",
    "paragraph_replace",
    "term_rename",
    "section_insert",
)
SINGLE_SEEDS_PER_TYPE = 2
MIXED_PAIRS_PER_DOCUMENT = 6
MIXED_EDITS = (4, 8)
DENSE_PAIRS_PER_DOCUMENT = 4
DENSE_PARAGRAPHS = 3
DENSE_EDITS_PER_PARAGRAPH = (2, 3)

# Single-word substitutions a reviewer might plausibly make. Keys are
# matched case-sensitively against a word's alphabetic core. No value may
# share a scoring token with its key ("exclusive" -> "non-exclusive" would),
# or the minimal diff is smaller than the recorded truth.
WORD_SUBSTITUTIONS = {
    "shall": "must",
    "may": "can",
    "reasonable": "proper",
    "reasonably": "properly",
    "written": "prior",
    "all": "any",
    "any": "every",
    "immediate": "prompt",
    "material": "serious",
    "solely": "only",
    "monthly": "quarterly",
    "annually": "quarterly",
    "lower": "higher",
    "higher": "lower",
    "daily": "weekly",
    "equal": "unequal",
    "secret": "confidential",
    "adequate": "sufficient",
    "fair": "reasonable",
}

PHRASES = (
    "acting reasonably",
    "promptly",
    "in writing",
    "at its own cost",
    "without undue delay",
    "(acting in good faith)",
    "to the extent permitted by law",
    "from time to time",
    "in all material respects",
    "subject to clause 12",
)

# Defined-term renames. Applied to every occurrence of the term as a whole
# word, including its possessive.
TERM_RENAMES = {
    "Consultant": "Contractor",
    "Client": "Customer",
    "Tenant": "Lessee",
    "Landlord": "Lessor",
    "Supplier": "Provider",
    "Customer": "Subscriber",
    "Founder": "Executive",
    "Platform": "Service",
    "Purpose": "Project",
    "Property": "Premises",
    "Licensor": "Vendor",
    "Licensee": "Purchaser",
    "Company": "Business",
}

# Scoring tokens: lowercase words and single punctuation characters. Shared
# with run_eval.py so the invariant checked here is the one scored there.
SCORE_TOKEN_RE = re.compile(r"\w+|[^\w\s]", re.UNICODE)

_WORD_CORE_RE = re.compile(r"^(\W*)(\w+)(\W*)$", re.UNICODE)
# A number that is a whole scoring token: "30", "(30)", "4%", never the "8"
# in "B8", which the scorer reads as one token.
_NUMBER_RE = re.compile(r"^(\W*)(\d+)(\W*)$")


# Paragraph styles a base document can mark a block with, as "{tag} text",
# and the Word style each renders as. Untagged blocks are "normal".
STYLES = {
    "normal": "Normal",
    "title": "Title",
    "subtitle": "Subtitle",
    "h1": "Heading 1",
    "h2": "Heading 2",
    "num": "List Number",
    "num2": "List Number 2",
    "bullet": "List Bullet",
}
_STYLE_TAG_RE = re.compile(r"^\{([a-z0-9]+)\} ")
BODY_STYLES = {"normal", "num", "num2", "bullet"}

# Headings a section_insert adds, each paired with a clause from the pool.
SECTION_HEADINGS = (
    "Anti-Bribery",
    "Force Majeure",
    "Notices",
    "Waiver",
    "Severance",
    "Counterparts",
    "Third Party Rights",
    "Non-Solicitation",
    "Business Continuity",
    "Dispute Resolution",
    "Costs",
    "Audit",
)


def split_style(block: str) -> tuple[str, str]:
    """``("h1", "Confidentiality")`` from ``"{h1} Confidentiality"``."""
    m = _STYLE_TAG_RE.match(block)
    if not m:
        return "normal", block
    tag = m.group(1)
    if tag not in STYLES:
        raise ValueError(f"unknown style tag {{{tag}}} in {block[:60]!r}")
    return tag, block[m.end() :]


@dataclass
class Paragraph:
    """One block of the document being edited.

    ``old`` is the original text, or None for a paragraph an edit inserted.
    ``edits`` holds the (deleted, inserted) pairs of in-paragraph edits.
    ``style`` is the paragraph's role (a key of STYLES); edits never change it.
    """

    old: str | None
    words: list[str]
    style: str = "normal"
    edits: list[tuple[str, str]] = field(default_factory=list)
    protected: set[int] = field(default_factory=set)
    locked: bool = False
    deleted: bool = False

    @property
    def editable(self) -> bool:
        return self.old is not None and not self.locked and not self.deleted

    @property
    def text(self) -> str:
        return " ".join(self.words)


def read_blocks(path: Path) -> list[str]:
    blocks = [b.strip() for b in path.read_text(encoding="utf-8").split("\n\n")]
    return [b for b in blocks if b]


def _core(regex: re.Pattern[str], word: str) -> str:
    m = regex.match(word)
    if m is None:
        raise AssertionError(f"{word!r} was selected but does not match {regex.pattern}")
    return m.group(2)


class EditFailed(Exception):
    """No valid target for this edit type in the current document state."""


class Editor:
    def __init__(self, blocks: list[str], extra: list[str], rng: random.Random) -> None:
        self.paras = []
        for block in blocks:
            style, text = split_style(block)
            self.paras.append(Paragraph(old=text, words=text.split(" "), style=style))
        texts = {p.old for p in self.paras}
        self.extra = [c for c in extra if c not in texts]
        self.rng = rng
        self.log: list[dict[str, object]] = []

    # Selection helpers ---------------------------------------------------

    def _editable_body(self) -> list[int]:
        """Editable paragraphs, excluding the title (block 0)."""
        return [i for i, p in enumerate(self.paras) if i > 0 and p.editable]

    def _word_targets(self, predicate: Callable[[str], bool]) -> list[tuple[int, int]]:
        out = []
        for i in self._editable_body():
            p = self.paras[i]
            for k, w in enumerate(p.words):
                if k not in p.protected and predicate(w):
                    out.append((i, k))
        return out

    def _pick_word(self, targets: list[tuple[int, int]]) -> tuple[int, int]:
        if not targets:
            raise EditFailed
        return self.rng.choice(targets)

    def _pick_paragraph(self, candidates: list[int]) -> int:
        if not candidates:
            raise EditFailed
        return self.rng.choice(candidates)

    def _record(self, kind: str, para: int, deleted: str, inserted: str) -> None:
        self.log.append({"type": kind, "paragraph": para, "deleted": deleted, "inserted": inserted})

    def _pop_extra(self) -> str:
        if not self.extra:
            raise EditFailed
        return self.extra.pop(self.rng.randrange(len(self.extra)))

    # Edit types ----------------------------------------------------------

    def _replace_core(self, kind: str, i: int, k: int, regex: re.Pattern[str], new_core: str) -> None:
        p = self.paras[i]
        m = regex.match(p.words[k])
        assert m
        p.words[k] = m.group(1) + new_core + m.group(3)
        p.edits.append((m.group(2), new_core))
        p.locked = True
        self._record(kind, i, m.group(2), new_core)

    def word_replace(self) -> None:
        def ok(w: str) -> bool:
            m = _WORD_CORE_RE.match(w)
            return m is not None and m.group(2) in WORD_SUBSTITUTIONS

        i, k = self._pick_word(self._word_targets(ok))
        core = _core(_WORD_CORE_RE, self.paras[i].words[k])
        self._replace_core("word_replace", i, k, _WORD_CORE_RE, WORD_SUBSTITUTIONS[core])

    def number_change(self) -> None:
        i, k = self._pick_word(self._word_targets(lambda w: bool(_NUMBER_RE.match(w))))
        old = _core(_NUMBER_RE, self.paras[i].words[k])
        new = old
        while new == old:
            factor = self.rng.choice((0.5, 1.5, 2, 3))
            new = str(max(1, round(int(old) * factor)))
        self._replace_core("number_change", i, k, _NUMBER_RE, new)

    def punctuation_change(self) -> None:
        i, k = self._pick_word(self._word_targets(lambda w: w.endswith(",") and len(w) > 1))
        p = self.paras[i]
        if k == len(p.words) - 1:
            raise EditFailed
        new = self.rng.choice((";", ""))
        p.words[k] = p.words[k][:-1] + new
        p.edits.append((",", new))
        p.locked = True
        self._record("punctuation_change", i, ",", new)

    def phrase_insert(self) -> None:
        i = self._pick_paragraph(
            [j for j in self._editable_body() if self.paras[j].style in BODY_STYLES and len(self.paras[j].words) >= 4]
        )
        p = self.paras[i]
        # After word k, never after the last word (that would follow the full
        # stop) and never after a word ending a sentence.
        slots = [k for k in range(len(p.words) - 1) if not p.words[k].endswith((".", ":"))]
        if not slots:
            raise EditFailed
        k = self.rng.choice(slots)
        phrase = self.rng.choice(PHRASES)
        p.words[k + 1 : k + 1] = phrase.split(" ")
        p.edits.append(("", phrase))
        p.locked = True
        self._record("phrase_insert", i, "", phrase)

    def phrase_delete(self) -> None:
        i = self._pick_paragraph([j for j in self._editable_body() if len(self.paras[j].words) >= 8])
        p = self.paras[i]
        length = self.rng.randint(2, 6)
        # Never the first or last word, so the paragraph keeps its opening and
        # its full stop, and never a protected word.
        starts = [s for s in range(1, len(p.words) - length) if not any(k in p.protected for k in range(s, s + length))]
        if not starts:
            raise EditFailed
        s = self.rng.choice(starts)
        deleted = " ".join(p.words[s : s + length])
        del p.words[s : s + length]
        p.edits.append((deleted, ""))
        p.locked = True
        self._record("phrase_delete", i, deleted, "")

    def dense_paragraph(self) -> None:
        """Two or three in-paragraph edits in one paragraph.

        Targets are chosen on the original word list so they cannot overlap,
        then applied right to left so earlier indices stay valid. Edits may
        be adjacent: the recorded truth stays exact, since a bag of tokens is
        indifferent to how a diff splits adjacent edits.
        """
        i = self._pick_paragraph([j for j in self._editable_body() if len(self.paras[j].words) >= 20])
        p = self.paras[i]
        used: set[int] = set(p.protected)
        planned: list[tuple[int, str, int]] = []  # (position, kind, length)
        for _ in range(self.rng.randint(*DENSE_EDITS_PER_PARAGRAPH)):
            kind = self.rng.choice(("word_replace", "number_change", "phrase_insert", "phrase_delete"))
            length = self.rng.randint(2, 5) if kind == "phrase_delete" else 1
            options = []
            for k in range(1, len(p.words) - length):
                span = range(k, k + length)
                if any(j in used for j in span):
                    continue
                w = p.words[k]
                if kind == "word_replace":
                    m = _WORD_CORE_RE.match(w)
                    if not (m and m.group(2) in WORD_SUBSTITUTIONS):
                        continue
                elif kind == "number_change" and not _NUMBER_RE.match(w):
                    continue
                elif kind == "phrase_insert" and w.endswith((".", ":")):
                    continue
                options.append(k)
            if not options:
                continue
            k = self.rng.choice(options)
            used.update(range(k, k + length))
            planned.append((k, kind, length))
        if len(planned) < 2:
            raise EditFailed
        for k, kind, length in sorted(planned, reverse=True):
            if kind == "word_replace":
                core = _core(_WORD_CORE_RE, p.words[k])
                self._replace_core("dense_paragraph", i, k, _WORD_CORE_RE, WORD_SUBSTITUTIONS[core])
            elif kind == "number_change":
                old = _core(_NUMBER_RE, p.words[k])
                self._replace_core("dense_paragraph", i, k, _NUMBER_RE, str(int(old) + self.rng.randint(1, 9)))
            elif kind == "phrase_insert":
                phrase = self.rng.choice(PHRASES)
                p.words[k + 1 : k + 1] = phrase.split(" ")
                p.edits.append(("", phrase))
                self._record("dense_paragraph", i, "", phrase)
            else:
                deleted = " ".join(p.words[k : k + length])
                del p.words[k : k + length]
                p.edits.append((deleted, ""))
                self._record("dense_paragraph", i, deleted, "")
        p.locked = True

    def _body_style_at(self, at: int) -> str:
        """The style a new clause at position ``at`` takes: that of the body
        paragraph before it, so a clause added inside a numbered list is
        numbered too."""
        for k in range(at - 1, -1, -1):
            if self.paras[k].style in BODY_STYLES:
                return self.paras[k].style
            if self.paras[k].style in ("h1", "h2"):
                break
        return "normal"

    def section_insert(self) -> None:
        """A new heading and a clause under it, inserted just before an
        existing heading of the same level (so it is a section of its own),
        or at the end when the document has no headings."""
        headings = [i for i, p in enumerate(self.paras) if i > 0 and p.style in ("h1", "h2") and not p.deleted]
        level = self.paras[headings[0]].style if headings else "h1"
        candidates = [i for i in headings if self.paras[i].style == level] or [len(self.paras)]
        at = self.rng.choice(candidates)
        clause = self._pop_extra()
        heading = self.rng.choice(SECTION_HEADINGS)
        body = self._body_style_at(at) if at < len(self.paras) else "normal"
        self.paras.insert(at, Paragraph(old=None, words=clause.split(" "), locked=True, style=body))
        self.paras.insert(at, Paragraph(old=None, words=heading.split(" "), locked=True, style=level))
        self._record("section_insert", at, "", heading)
        self._record("section_insert", at + 1, "", clause)

    def paragraph_insert(self) -> None:
        text = self._pop_extra()
        at = self.rng.randint(1, len(self.paras))
        self.paras.insert(at, Paragraph(old=None, words=text.split(" "), locked=True, style=self._body_style_at(at)))
        self._record("paragraph_insert", at, "", text)

    def paragraph_delete(self) -> None:
        i = self._pick_paragraph(self._editable_body())
        p = self.paras[i]
        assert p.old is not None
        p.deleted = True
        self._record("paragraph_delete", i, p.old, "")

    def paragraph_replace(self) -> None:
        i = self._pick_paragraph(self._editable_body())
        text = self._pop_extra()
        p = self.paras[i]
        assert p.old is not None
        p.deleted = True
        self.paras.insert(i + 1, Paragraph(old=None, words=text.split(" "), locked=True, style=p.style))
        self._record("paragraph_replace", i, p.old, text)

    def term_rename(self) -> None:
        """Rename one defined term everywhere. Must run before other edits."""
        if any(p.edits or p.deleted or p.old is None for p in self.paras):
            raise RuntimeError("term_rename must be the first edit applied")
        present = [t for t in TERM_RENAMES if sum(p.old.count(t) for p in self.paras if p.old) >= 3]
        if not present:
            raise EditFailed
        term = self.rng.choice(sorted(present))
        new = TERM_RENAMES[term]
        word_re = re.compile(rf"^(\W*){term}((?:'s)?\W*)$")
        count = 0
        for i, p in enumerate(self.paras):
            for k, w in enumerate(p.words):
                m = word_re.match(w)
                if m:
                    p.words[k] = m.group(1) + new + m.group(2)
                    p.edits.append((term, new))
                    p.protected.add(k)
                    count += 1
                    self._record("term_rename", i, term, new)
        if not count:
            raise EditFailed

    # Output --------------------------------------------------------------

    def styles(self) -> tuple[list[str], list[str]]:
        """The style of each old block and each new block, parallel to result()."""
        return (
            [p.style for p in self.paras if p.old is not None],
            [p.style for p in self.paras if not p.deleted],
        )

    def result(self) -> tuple[list[str], list[str], list[str], list[str]]:
        """(old blocks, new blocks, expected deleted, expected inserted)."""
        old = [p.old for p in self.paras if p.old is not None]
        new = [p.text for p in self.paras if not p.deleted]
        deleted: list[str] = []
        inserted: list[str] = []
        for p in self.paras:
            if p.old is None:
                inserted.append(p.text)
            elif p.deleted:
                deleted.append(p.old)
            else:
                deleted.extend(d for d, _ in p.edits if d)
                inserted.extend(s for _, s in p.edits if s)
        return old, new, deleted, inserted


def apply_edits(editor: Editor, types: list[str]) -> list[str]:
    """Apply edits of the given types, at most one rename and that one first.
    Returns the types that were applied; a type with no valid target in this
    document is skipped."""
    renames = [t for t in types if t == "term_rename"][:1]
    ordered = renames + [t for t in types if t != "term_rename"]
    applied = []
    for kind in ordered:
        try:
            getattr(editor, kind)()
        except EditFailed:
            continue
        applied.append(kind)
    return applied


def make_pair(
    pair_id: str, base: str, family: str, blocks: list[str], extra: list[str], types: list[str], seed: int
) -> dict[str, object] | None:
    editor = Editor(blocks, extra, random.Random(seed))
    applied = apply_edits(editor, types)
    if not applied:
        return None
    old, new, deleted, inserted = editor.result()
    if old == new:
        return None
    _check_truth(pair_id, old, new, deleted, inserted)
    return {
        "id": pair_id,
        "base": base,
        "family": family,
        "seed": seed,
        "edit_types": sorted(set(applied)),
        "edits": editor.log,
        "old": old,
        "old_styles": editor.styles()[0],
        "new_styles": editor.styles()[1],
        "new": new,
        "expected_deleted": deleted,
        "expected_inserted": inserted,
    }


def score_tokens(texts: list[str]) -> collections.Counter[str]:
    return collections.Counter(t.lower() for text in texts for t in SCORE_TOKEN_RE.findall(text))


def _check_truth(pair_id: str, old: list[str], new: list[str], deleted: list[str], inserted: list[str]) -> None:
    """The recorded truth must account for every token: old minus deleted
    plus inserted is new, exactly. A generator bug that records the wrong
    text fails here instead of quietly mis-scoring every algorithm."""
    old_bag, deleted_bag = score_tokens(old), score_tokens(deleted)
    if deleted_bag - old_bag:
        raise AssertionError(f"{pair_id}: expected deletions not in the old text: {deleted_bag - old_bag}")
    rebuilt = old_bag - deleted_bag + score_tokens(inserted)
    if rebuilt != score_tokens(new):
        raise AssertionError(
            f"{pair_id}: truth does not reconcile; off by {rebuilt - score_tokens(new)} / {score_tokens(new) - rebuilt}"
        )


def generate() -> list[dict[str, object]]:
    extra = read_blocks(BASE_DIR / f"{EXTRA_CLAUSES}.txt")
    bases = sorted(p for p in BASE_DIR.glob("*.txt") if p.stem != EXTRA_CLAUSES)
    pairs: list[dict[str, object]] = []
    skipped: list[str] = []
    for path in bases:
        blocks = read_blocks(path)
        for kind in EDIT_TYPES:
            for n in range(SINGLE_SEEDS_PER_TYPE):
                pair_id = f"{path.stem}/single/{kind}/{n}"
                seed = _seed(pair_id)
                pair = make_pair(pair_id, path.stem, "single", blocks, extra, [kind], seed)
                if pair is None:
                    skipped.append(pair_id)
                else:
                    pairs.append(pair)
        for n in range(DENSE_PAIRS_PER_DOCUMENT):
            pair_id = f"{path.stem}/dense/{n}"
            pair = make_pair(
                pair_id, path.stem, "dense", blocks, extra, ["dense_paragraph"] * DENSE_PARAGRAPHS, _seed(pair_id)
            )
            if pair is None:
                skipped.append(pair_id)
            else:
                pairs.append(pair)
        for n in range(MIXED_PAIRS_PER_DOCUMENT):
            pair_id = f"{path.stem}/mixed/{n}"
            seed = _seed(pair_id)
            rng = random.Random(seed)
            types = [rng.choice(EDIT_TYPES) for _ in range(rng.randint(*MIXED_EDITS))]
            pair = make_pair(pair_id, path.stem, "mixed", blocks, extra, types, seed + 1)
            if pair is None:
                skipped.append(pair_id)
            else:
                pairs.append(pair)
    if skipped:
        print(f"skipped {len(skipped)} pairs with no applicable edit: {', '.join(skipped)}", file=sys.stderr)
    return pairs


def _seed(pair_id: str) -> int:
    """A seed that depends only on the pair id, not on Python's hash salt or
    on the order pairs are generated in."""
    return sum((k + 1) * ord(c) for k, c in enumerate(pair_id))


def serialise(pairs: list[dict[str, object]]) -> str:
    return "".join(json.dumps(p, ensure_ascii=False, sort_keys=True) + "\n" for p in pairs)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument(
        "--check", action="store_true", help="exit 1 if the committed corpus differs from a regeneration"
    )
    args = parser.parse_args()

    content = serialise(generate())
    if args.check:
        current = CORPUS_PATH.read_text(encoding="utf-8") if CORPUS_PATH.exists() else ""
        if current != content:
            sys.exit(f"{CORPUS_PATH} is out of date with eval/generate.py; run it without --check and commit")
        print(f"{CORPUS_PATH} is up to date")
        return
    CORPUS_PATH.parent.mkdir(parents=True, exist_ok=True)
    CORPUS_PATH.write_text(content, encoding="utf-8")
    print(f"wrote {content.count(chr(10))} pairs to {CORPUS_PATH}")


if __name__ == "__main__":
    main()
