Scored with [`eval/bench.py`](https://github.com/Kyra-Law/kyralaw-redline/blob/main/eval/bench.py) on the
`test` split (685 pairs), on one Apple silicon Mac. Every algorithm here runs from that repository.

| Algorithm | What it is | Score | Text F1 | Insert F1 | Delete F1 | Fidelity | Integrity failures | Median time per pair |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `word-compare` | Word's own reference, scored as a candidate: a check on the scorer | 100.00 | 100.00 | 1.000 | 1.000 | 1.000 | 0 | |
| `python-redlines` | [python-redlines](https://github.com/JSv4/Python-Redlines) 0.3.0, Open XML PowerTools' comparer | 94.31 | 91.84 | 0.959 | 0.878 | 0.992 | 0 | 2.58 s |
| `kyralaw-redline-auto` | [kyralaw-redline](https://github.com/Kyra-Law/kyralaw-redline)'s default: in place, rebuilt only if that fails (it never did here) | 94.08 | 91.34 | 0.956 | 0.870 | 0.996 | 0 | 234 ms |
| `kyralaw-redline-in-place` | kyralaw-redline, writing into the new document: all 685 pairs in place | 94.08 | 91.34 | 0.956 | 0.870 | 0.996 | 0 | 233 ms |
| `kyralaw-native` | kyralaw-redline's first in-place renderer, frozen as a baseline; refuses 3 pairs, left out | 94.05 | 91.30 | 0.956 | 0.870 | 0.995 | 0 | 122 ms |
| `kyralaw-redline-rebuild` | kyralaw-redline, writing a new, plain document | 78.72 | 92.39 | 0.964 | 0.884 | 0.514 | 0 | 353 ms |
| `token-diff` | a token-level difflib, no alignment or heuristics | 62.60 | 68.02 | 0.774 | 0.587 | 0.518 | 0 | 1.57 s |
| `line-diff` | a line diff | 49.68 | 48.56 | 0.600 | 0.371 | 0.519 | 0 | 220 ms |

- **Marking the change is close at the top; keeping the document is not.**
  `kyralaw-redline-rebuild` marks the text best of kyralaw-redline's renderers but writes into a
  blank document, so tables, headers and footers do not survive, and fidelity decides its score.
  The in-place renderers write into the new document itself, the way Word does.
- **`kyralaw-redline-in-place` and `kyralaw-native` mark the same changes.** On the 682 pairs
  `kyralaw-native` writes, they score 94.07 and 94.05. On the other 3, tables where a row is
  deleted or added and a value recurs in other rows, `kyralaw-native`'s own check fails and it
  writes nothing; `kyralaw-redline-in-place` writes them in place (about 96 each).
- **Deletions are harder than insertions** for every algorithm, by about 8 points of F1 at the top.
- **Integrity failures** are redlines whose changes, all accepted or all rejected, do not give back
  the two documents. None of these algorithms has any.
- **Times** are for producing each redline from the two files, on one machine. Treat them as
  relative.
