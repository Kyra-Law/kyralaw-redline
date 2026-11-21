# Building kyralaw-redline-bench

This folder is how [kyralaw-redline-bench](../README.md#the-word-benchmark-kyralaw-redline-bench) was
made, start to finish. The scorer is [`eval/bench.py`](../eval/bench.py).

| File | Does |
| --- | --- |
| [`build.py`](build.py) | turns each tracked-change markup in Harvey LAB into a pair: Word rejects its changes (`old.docx`), accepts them (`new.docx`) and compares the two (`old_to_new.docx`, the reference), and the pair is kept only if the reference reads back as exactly the two versions |
| [`word.py`](word.py) | drives Microsoft Word through AppleScript, hidden, answering the prompts Word raises for files it calls damaged and restarting it when a file hangs it |
| [`views.py`](views.py) | reads a document the way Word shows it, with every change accepted or rejected: the read-back check |
| [`release.py`](release.py) | turns the pairs into the dataset as published: ids, splits, the dataset card, checksums |
| [`expand.sh`](expand.sh) | runs `build.py` in the background, logged and kept awake |
| `manifest.json` | each markup's LAB path, sha256, commit, and whether it was built or refused and why |
| `release_ids.json` | each pair's published id and split, fixed once released |
| `results.md` | the results table the card carries |
| `LICENSE-harvey-labs` | Harvey LAB's MIT licence, which travels with every copy of the pairs |

## Running it

Needs macOS with Microsoft Word, and a clone of Harvey LAB at the commit the manifest names.

```bash
git clone https://github.com/harveyai/harvey-labs.git
git -C harvey-labs checkout 7f9b1ee19f9224576a5a5143c4e6f75842ac7020

export KYRALAW_BENCH_LAB=harvey-labs KYRALAW_BENCH_PAIRS=work/pairs
bench/expand.sh start            # every markup, in the background; resumes where it stopped
bench/expand.sh status
uv run python bench/release.py --lab harvey-labs --pairs work/pairs \
    --out work/release/v1.0.0 --results bench/results.md
```

With this manifest, a build finds every markup already decided, so it only rebuilds pairs
that are missing from `work/pairs`. Word's output can vary between versions of Word: the
manifest and the dataset card record the version used (16.115.1).

## Word, out of sight

A build of hundreds of pairs runs while you work:

- Word is started hidden and behind your current app, and hidden again before and during every
  call, so no window or dialog appears and focus is never taken.
- While you use Word yourself (a document of yours is open, or Word is the frontmost app) the
  build waits, and resumes once you are done.
- About one LAB markup in 30 is a file Word calls damaged. Word asks "Word found unreadable
  content in ... Do you want to recover the contents?"; the build answers No (a recovered document
  is Word's guess at the content) and refuses the pair, in a few seconds.
- Word is restarted every 50 pairs, because it slows down over a long session, and whenever it
  stops answering for 45 seconds.
- A file is refused only for what is wrong with the file: when Word fails, the build first checks
  Word on a document it has built before, and restarts Word and retries if Word is at fault.
- `expand.sh stop` lets the pair in progress finish; cutting Word off mid-operation can leave it
  behind a hidden prompt.

Word writes the user's name and the time into every file it saves. `word.py` scrubs both, and
`release.py` refuses to publish a file that still names the account that built it.
