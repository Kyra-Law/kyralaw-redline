#!/bin/bash
# Build kyralaw-redline-bench's pairs in the background, with Word kept out of sight.
#
#   bench/expand.sh start [--limit N]   start (or resume) the build, return at once
#   bench/expand.sh status              progress, and the last lines of the log
#   bench/expand.sh stop                stop after the pair in progress; stop twice to stop at once
#
# Set first:
#   KYRALAW_BENCH_LAB     a clone of https://github.com/harveyai/harvey-labs
#   KYRALAW_BENCH_PAIRS   where the pairs are built (about 140 MB); the log goes there too
#   PYTHON                optional; default: uv run python
#
# What "silent" means (see word._ready_word):
# - Word is started hidden and behind your current app, and re-hidden before and
#   during every call, so no window or dialog appears and focus is never taken.
# - While you use Word yourself (a document of yours is open, or Word is the
#   frontmost app) the build waits, and resumes once you are done.
# - The Mac is kept from idle sleep while the build runs (caffeinate); the display
#   may still sleep.
# - Everything goes to $KYRALAW_BENCH_PAIRS/build.log, nothing to the terminal after
#   start returns. Word is quit at the end if the build started it.
#
# Needs macOS with Microsoft Word. The builder resumes where it stopped, so start is
# safe to repeat.

set -euo pipefail

here="$(cd "$(dirname "$0")" && pwd)"
: "${KYRALAW_BENCH_LAB:?set KYRALAW_BENCH_LAB to a clone of https://github.com/harveyai/harvey-labs}"
: "${KYRALAW_BENCH_PAIRS:?set KYRALAW_BENCH_PAIRS to the folder the pairs are built in}"
lab="${KYRALAW_BENCH_LAB/#\~/$HOME}"
pairs="${KYRALAW_BENCH_PAIRS/#\~/$HOME}"
log="$pairs/build.log"
pidfile="$pairs/build.pid"
read -r -a python <<<"${PYTHON:-uv run --quiet --project $here/.. python}"

running() {
    [ -f "$pidfile" ] && kill -0 "$(cat "$pidfile")" 2>/dev/null
}

case "${1:-}" in
start)
    shift
    if running; then
        echo "already running (pid $(cat "$pidfile")); see: $0 status"
        exit 0
    fi
    if [ ! -d "$lab/tasks" ]; then
        echo "no Harvey LAB clone at $lab; make one with:"
        echo "  git clone --depth 1 https://github.com/harveyai/harvey-labs.git \"$lab\""
        exit 1
    fi
    mkdir -p "$pairs"
    printf '\n=== %s start %s\n' "$(date '+%Y-%m-%d %H:%M:%S')" "$*" >>"$log"
    nohup nice -n 10 "${python[@]}" -u "$here/build.py" --lab "$lab" --out "$pairs" "$@" >>"$log" 2>&1 </dev/null &
    pid=$!
    echo "$pid" >"$pidfile"
    # Hold off idle sleep for exactly as long as the build lives.
    nohup caffeinate -i -s -w "$pid" >/dev/null 2>&1 </dev/null &
    echo "started (pid $pid); log: $log"
    echo "check on it with: $0 status"
    ;;
status)
    if running; then
        echo "running (pid $(cat "$pidfile"))"
    else
        echo "not running"
    fi
    "${python[@]}" - "$here/manifest.json" "$log" <<'EOF'
import json, re, sys
from pathlib import Path
pairs = json.loads(Path(sys.argv[1]).read_text())["pairs"] if Path(sys.argv[1]).exists() else {}
built = sum(e.get("status") == "built" for e in pairs.values())
refused = sum(e.get("status") == "refused" for e in pairs.values())
log = Path(sys.argv[2]).read_text().splitlines() if Path(sys.argv[2]).exists() else []
total = next((int(m[1]) for line in reversed(log) if (m := re.search(r"(\d+) distinct tracked-change markups", line))), None)
left = f", {total - built - refused} left" if total else ""
print(f"{built} built, {refused} refused{left}" + (f" of {total} markups" if total else ""))
times = [int(m[1]) for line in log[-200:] if (m := re.search(r" in (\d+)s \(\d+ of \d+\)$", line))]
if times and total:
    per = sorted(times)[len(times) // 2]  # median: a pause or a hang is not the pace
    print(f"recent pairs take {per:.0f}s (median); about {(total - built - refused) * per / 3600:.1f} h to go")
EOF
    echo "--- last lines of $log"
    tail -n 8 "$log" 2>/dev/null || true
    ;;
stop)
    if running; then
        kill -TERM "$(cat "$pidfile")"
        echo "stopping pid $(cat "$pidfile") once the pair in progress is done (up to a minute); stop again to end it at once"
    else
        echo "not running"
    fi
    ;;
*)
    echo "usage: $0 start [--limit N] | status | stop"
    exit 2
    ;;
esac
