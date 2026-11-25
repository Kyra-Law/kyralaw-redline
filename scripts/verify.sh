#!/bin/bash
# Check everything this repository claims, from scratch. Meant for the container (see the
# Dockerfile), where nothing but this repository is installed; it runs anywhere with the
# dev and bench dependency groups installed.
#
#   scripts/verify.sh                  isolation, format, lint, types, tests
#   KYRALAW_BENCH_DATA=/data           also score kyralaw-redline-bench (default /data when that
#                                      folder exists) and compare every score with
#                                      bench/results.md
#   KYRALAW_BENCH_SPLIT=dev            the 48 dev pairs instead of the 685 test pairs: quicker,
#                                      but bench/results.md is for test, so nothing is compared
#
# Every step must pass; the first failure stops the run with a non-zero exit.

set -euo pipefail
cd "$(dirname "$0")/.."

step() { printf '\n== %s\n' "$*"; }

step "Isolation: no other Kyra Law codebase is installed or referenced"
python - <<'PY'
import importlib.util, sys
found = [name for name in ("clauze_core", "clauze", "kyralaw_evals") if importlib.util.find_spec(name) is not None]
if found:
    sys.exit(f"importable here: {found}; this environment is not isolated")
print("no other Kyra Law package is importable")
PY
if grep -rInE "clauze|kyralaw-evals|monorepo|PATH_DIR_CLAUZE" src eval bench tests scripts README.md CLAUDE.md \
    pyproject.toml --exclude=verify.sh; then
    echo "the lines above refer to another codebase"
    exit 1
fi
echo "no reference to another codebase in the code or the docs"

step "Format, lint, types"
ruff format --check .
ruff check .
mypy

step "Tests"
pytest -q -p no:cacheprovider

data="${KYRALAW_BENCH_DATA:-/data}"
if [ ! -f "$data/test.jsonl" ]; then
    step "kyralaw-redline-bench: not mounted at $data, skipped"
    echo "to score it: docker run --rm --network none -v /path/to/kyralaw-redline-bench:/data:ro kyralaw-redline"
    printf '\nAll checks passed (benchmark not scored).\n'
    exit 0
fi

split="${KYRALAW_BENCH_SPLIT:-test}"
engines=(word-compare python-redlines $(python -c "import sys; sys.path.insert(0, 'eval'); import engines; print(' '.join(engines.ALGORITHMS))"))
step "kyralaw-redline-bench, $split split: ${#engines[@]} engines, in parallel"
out="$(mktemp -d)"
pids=()
for engine in "${engines[@]}"; do
    python eval/bench.py "$data" --split "$split" -a "$engine" --json "$out/$engine.json" >"$out/$engine.md" 2>&1 &
    pids+=("$!")
done
failed=0
for i in "${!pids[@]}"; do
    if ! wait "${pids[$i]}"; then
        echo "${engines[$i]} failed:"; tail -n 20 "$out/${engines[$i]}.md"; failed=1
    fi
done
[ "$failed" -eq 0 ]

python - "$out" "$split" <<'PY'
import json, re, sys
from pathlib import Path
out, split = Path(sys.argv[1]), sys.argv[2]
got = {}
for path in sorted(out.glob("*.json")):
    got.update(json.loads(path.read_text())["summaries"])
print("| Engine | Scored | Score | Text F1 | Fidelity | Integrity failures | Median time |")
print("| --- | ---: | ---: | ---: | ---: | ---: | ---: |")
for name, s in got.items():
    t = s["median_seconds"]
    print(f"| {name} | {s['scored']} of {s['pairs']} | {s['score']:.2f} | {s['text_f1']:.2f} | {s['fidelity']:.3f} | "
          f"{s['integrity_failures']} | {'-' if t is None else f'{t * 1000:.0f} ms'} |")

problems = []
if f"{got['word-compare']['score']:.2f}" != "100.00":
    problems.append(f"word-compare scored {got['word-compare']['score']:.2f}; the scorer must give Word's own reference 100.00")
for name, s in got.items():
    if s["integrity_failures"]:
        problems.append(f"{name} has {s['integrity_failures']} integrity failures")
if split == "test":
    # Every score in bench/results.md was produced on another machine; on this one each
    # must come out the same, to the precision it is published at. Times are not compared.
    expected = {}
    for line in Path("bench/results.md").read_text().splitlines():
        m = re.match(r"\| `([a-z0-9-]+)` \| [^|]+ \| ([\d.]+) \| ([\d.]+) \| [\d.]+ \| [\d.]+ \| ([\d.]+) \| (\d+) \|", line)
        if m:
            expected[m[1]] = (m[2], m[3], m[4], int(m[5]))
    missing = sorted(set(expected) - set(got))
    if missing:
        problems.append(f"bench/results.md lists engines this run did not score: {missing}")
    for name, (score, text, fidelity, integrity) in expected.items():
        s = got.get(name)
        if s is None:
            continue
        mine = (f"{s['score']:.2f}", f"{s['text_f1']:.2f}", f"{s['fidelity']:.3f}", s["integrity_failures"])
        if mine != (score, text, fidelity, integrity):
            problems.append(f"{name}: this machine gives {mine}, bench/results.md says {(score, text, fidelity, integrity)}")
    print(f"\ncompared {len(expected)} engines with bench/results.md")
if problems:
    print("\n".join(["", "FAILED:", *problems]))
    sys.exit(1)
print("every score matches" if split == "test" else "word-compare is 100.00 and nothing has an integrity failure")
PY
printf '\nAll checks passed.\n'
