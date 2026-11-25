Scored with [`eval/bench.py`](https://github.com/Kyra-Law/kyralaw-redline/blob/bench-v1.0.0/eval/bench.py) on the
`test` split (685 pairs), on one Apple silicon Mac. The two `docxodus` rows were added later, scored by the same
scorer at [`0f74a70`](https://github.com/Kyra-Law/kyralaw-redline/tree/0f74a70). Every algorithm here runs from that repository.

| Algorithm | What it is | Score | Text F1 | Insert F1 | Delete F1 | Fidelity | Integrity failures | Median time per pair |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `word-compare` | Word's own reference, scored as a candidate: a check on the scorer | 100.00 | 100.00 | 1.000 | 1.000 | 1.000 | 0 | |
| `docxodus` | [Docxodus](https://github.com/JSv4/Docxodus) 17.0.0 through [docx-scalpel](https://pypi.org/project/docx-scalpel/) 0.14.0, the comparison its `redline` command runs; one process serves every pair | 96.33 | 94.58 | 0.973 | 0.918 | 0.998 | 7 | 165 ms |
| `docxodus-checked` | `docxodus`, each redline read back and, if it does not give back both documents, compared again with only spaces as word separators (7 pairs) | 96.33 | 94.58 | 0.973 | 0.918 | 0.998 | 0 | 298 ms |
| `python-redlines` | [python-redlines](https://github.com/JSv4/Python-Redlines) 0.3.0, Open XML PowerTools' comparer | 94.31 | 91.84 | 0.959 | 0.878 | 0.992 | 0 | 2.58 s |
| `kyralaw-redline` | [kyralaw-redline](https://github.com/Kyra-Law/kyralaw-redline), writing into the new document: all 685 pairs in place | 94.08 | 91.34 | 0.956 | 0.870 | 0.996 | 0 | 233 ms |
| `kyralaw-native` | kyralaw-redline's first in-place renderer, frozen as a baseline; refuses 3 pairs, left out | 94.05 | 91.30 | 0.956 | 0.870 | 0.995 | 0 | 122 ms |
| `kyralaw-redline-rebuild` | kyralaw-redline, writing a new, plain document | 78.72 | 92.39 | 0.964 | 0.884 | 0.514 | 0 | 353 ms |
| `token-diff` | a token-level difflib, no alignment or heuristics | 62.60 | 68.02 | 0.774 | 0.587 | 0.518 | 0 | 1.57 s |
| `line-diff` | a line diff | 49.68 | 48.56 | 0.600 | 0.371 | 0.519 | 0 | 220 ms |

- **`docxodus` leads, by two points.** Its insertions and deletions are both the closest to Word's, and
  it keeps the document as well as `kyralaw-redline` does. Its time leaves out .NET's start-up, which
  `python-redlines` pays on every pair.
- **Marking the change is close at the top; keeping the document is not.**
  `kyralaw-redline-rebuild` marks the text best of kyralaw-redline's renderers but writes into a
  blank document, so tables, headers and footers do not survive, and fidelity decides its score.
  `kyralaw-redline` and `kyralaw-native` write into the new document itself, the way Word does.
- **`kyralaw-redline` and `kyralaw-native` mark the same changes.** On the 682 pairs
  `kyralaw-native` writes, they score 94.07 and 94.05. On the other 3, tables where a row is
  deleted or added and a value recurs in other rows, `kyralaw-native`'s own check fails and it
  writes nothing; `kyralaw-redline` writes them in place (about 96 each).
- **Deletions are harder than insertions** for every algorithm, by 5 to 9 points of F1 at the top.
- **Integrity failures** are redlines whose changes, all accepted or all rejected, do not give back
  the two documents. Only `docxodus` has any: 7, all one bug. An empty paragraph that became "---",
  with a new paragraph after it, is written as unchanged, so rejecting every change keeps the "---".
  `docxodus-checked` catches each one by reading the redline back, and compares again with only
  spaces as word separators, which fixes all 7 without moving the score.
- **Times** are for producing each redline from the two files, on one machine. Treat them as
  relative.
