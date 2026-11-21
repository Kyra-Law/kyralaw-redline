# SPDX-License-Identifier: MIT
# Copyright (c) 2026 Kyra Law Limited
"""Microsoft Word, driven through AppleScript: resolve a markup, and compare two documents.

kyralaw-redline-bench measures agreement with Word, so Word makes every version and
every reference (``bench/build.py``): it rejects a markup's changes for ``old.docx``,
accepts them for ``new.docx``, and compares the two for ``reference.docx``.

**macOS with Microsoft Word only.** Word is sandboxed, so each file is staged in its
own container (``~/Library/Containers/com.microsoft.Word/Data/Documents``), where it
may read and write without a file-access prompt. The first run asks once to let the
terminal control Word.

**Word never shows a dialog while a script runs.** Each script sets Word's ``display
alerts`` to none, and restores the user's own setting before it returns, on failure as
on success, and a failing script closes only the documents it opened itself.

**Word stays out of sight.** It is started hidden and behind the current app, kept
hidden before and during every call, and never brought forward, so a build of
hundreds of pairs can run while the user works (``_ready_word``). Whenever the user is
using Word themselves (a document of theirs is open, or Word is the frontmost app) the
build waits for them instead. Word failing to start or answer at all raises
``WordUnavailableError``, which a builder must not mistake for a bad file.

**The result is scrubbed before it is returned.** Word stamps ``cp:lastModifiedBy``
with the macOS user's name and dates every revision with the time of the run. Both are
set to fixed values (``scrub``), and a document is refused outright if the account's
full name still appears in any part, since these files are published.
"""

from __future__ import annotations

import io
import os
import plistlib
import pwd
import re
import subprocess
import time
import uuid
import zipfile
from pathlib import Path

WORD_APP = Path("/Applications/Microsoft Word.app")
STAGING = Path.home() / "Library/Containers/com.microsoft.Word/Data/Documents/kyralaw-redline-bench"

# Who the revisions are attributed to, and who the document says last saved it.
AUTHOR = "Word Compare"
PUBLISHER = "kyralaw-redline-bench"
FIXED_DATE = "2026-01-01T00:00:00Z"

_SCRIPT = """
on run argv
    set oldPath to item 1 of argv
    set newPath to item 2 of argv
    set outPath to item 3 of argv
    set oldName to item 4 of argv
    set authorName to item 5 of argv
    tell application "Microsoft Word"
        set previousAlerts to display alerts
        set display alerts to alerts none
        set resultDoc to missing value
        try
            open (POSIX file oldPath) add to recent files false
            set oldDoc to missing value
            try
                set oldDoc to document oldName
            end try
            if oldDoc is missing value then error "Word could not open the document"
            compare oldDoc path newPath author name authorName target compare target new ¬
                ignore all comparison warnings true add to recent files false
            set resultDoc to active document
            if name of resultDoc is oldName then error "Compare produced no new document"
            save as resultDoc file name outPath file format format document add to recent files false
            close resultDoc saving no
            close oldDoc saving no
        on error msg
            try
                if resultDoc is not missing value then close resultDoc saving no
            end try
            try
                close document oldName saving no
            end try
            set display alerts to previousAlerts
            error msg
        end try
        set display alerts to previousAlerts
    end tell
    return "ok"
end run
"""


class WordCompareError(RuntimeError):
    """This file could not be compared or resolved: a fact about the file."""


class UnreadableDocumentError(WordCompareError):
    """Word called the file damaged ("unreadable content") and could not open it.

    About 1.5% of Harvey LAB's markups. Word offers to recover them; the offer is
    declined, because a recovered document is Word's guess at the content.
    """


class WordUnavailableError(RuntimeError):
    """Word itself could not be reached: a fact about the machine, not the file.

    Deliberately not a WordCompareError, so a builder that refuses a pair on
    WordCompareError stops instead of refusing every later pair for a reason none
    of them has.
    """


# How long one Word call may take before Word is taken to be stuck. The largest
# LAB pair (53,000 words) took a few seconds per call; a stuck Word never answers,
# and at 120s the wait dominated a build where hangs are common.
WORD_TIMEOUT_SECONDS = 45
# Every file this module stages starts with a 12-hex-digit stem.
_STAGED_RE = re.compile(r"^[0-9a-f]{12}-")
_WORD_PROCESS = "Microsoft Word"


def _word_documents() -> list[str] | None:
    """Names of the documents Word has open, or None when Word does not answer."""
    try:
        proc = subprocess.run(
            ["osascript", "-e", 'tell application "Microsoft Word" to get name of every document'],
            capture_output=True,
            text=True,
            timeout=30,
            check=False,
        )
    except subprocess.TimeoutExpired:
        return None
    if proc.returncode != 0:
        return None
    out = proc.stdout.strip()
    return [] if out in ("", "missing value") else [n.strip() for n in out.split(", ")]


def word_running() -> bool:
    return subprocess.run(["pgrep", "-x", _WORD_PROCESS], capture_output=True, check=False).returncode == 0


def _launch_hidden() -> None:
    """Start Word hidden and behind the current app, and return once it answers.

    ``open -g -j``: not brought forward, and launched hidden, so not even Word's
    start window shows. Started by an Apple Event instead, Word shows that window
    for about a second (measured 2026-11-20).
    """
    subprocess.run(["open", "-g", "-j", "-a", _WORD_PROCESS], check=False)
    for _ in range(90):
        if _word_documents() is not None:
            _settle_hidden()
            return
        time.sleep(1)
    raise WordUnavailableError("Word did not start within 90s")


# How long after starting Word it is watched and hidden again whenever it shows.
_SETTLE_SECONDS = 12


def _settle_hidden() -> None:
    """Keep a Word this module just started hidden through its first seconds.

    ``open -g -j`` is not the last word: relaunched after it was killed hung, Word
    brought itself to the front a few seconds later, in front of whatever the user
    was doing, and the build then waited on it as though the user had opened it
    (seen 2026-11-20). Hiding it here also hands focus back to the app the user was in.
    Anything frontmost after this is the user's doing, and ``_ready_word`` waits on it.
    """
    end = time.monotonic() + _SETTLE_SECONDS
    hid_frontmost = False
    while time.monotonic() < end:
        state = _word_state()
        if state is not None and state[1] and (not state[0] or not hid_frontmost):
            # Hiding the frontmost app hands focus to another window, so that is done
            # at most once: repeated, it walked focus through the user's windows
            # (2026-11-20). Hiding a Word behind the user's app moves nothing.
            hid_frontmost = hid_frontmost or state[0]
            _hide_word()
        time.sleep(0.5)


def _osascript(script: str, what: str) -> subprocess.CompletedProcess[str]:
    """One short System Events script; a hang there is the machine's problem, not the file's."""
    try:
        return subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=30, check=False)
    except subprocess.TimeoutExpired:
        raise WordUnavailableError(f"System Events did not answer within 30s ({what})") from None


_STATE_SCRIPT = """
tell application "System Events"
    if not (exists process "Microsoft Word") then return "absent"
    set p to process "Microsoft Word"
    return ((frontmost of p) as text) & "|" & ((visible of p) as text)
end tell
"""


def _word_state() -> tuple[bool, bool] | None:
    """(frontmost, visible) for Word's process, or None when it is not running."""
    proc = _osascript(_STATE_SCRIPT, "report on Word")
    out = proc.stdout.strip()
    if proc.returncode != 0:
        raise WordUnavailableError(f"System Events would not report on Word: {proc.stderr.strip()}")
    if out == "absent":
        return None
    frontmost, visible = out.split("|")
    return frontmost == "true", visible == "true"


def _hide_word() -> None:
    """Hide Word's process, as Cmd-H would; its documents stay open and scriptable."""
    proc = _osascript(
        'tell application "System Events" to set visible of process "Microsoft Word" to false', "hide Word"
    )
    if proc.returncode != 0:
        raise WordUnavailableError(f"could not hide Word: {proc.stderr.strip()}")


def clear_stale_staging(max_age_seconds: float = 3600) -> int:
    """Remove staged files an earlier, killed run left behind; returns how many.

    A call removes its own files on the way out, but not when its process is killed
    outright or Word is restarted holding them. Only files of this module's naming and
    older than ``max_age_seconds`` go, so a concurrent call's files are safe.
    """
    if not STAGING.is_dir():
        return 0
    cutoff = time.time() - max_age_seconds
    stale = [p for p in STAGING.iterdir() if _STAGED_RE.match(p.name) and p.stat().st_mtime < cutoff]
    for path in stale:
        path.unlink(missing_ok=True)
    return len(stale)


def word_memory_gb() -> float | None:
    """Word's resident memory in GB, or None when it is not running."""
    proc = subprocess.run(["pgrep", "-x", _WORD_PROCESS], capture_output=True, text=True, check=False)
    if proc.returncode != 0:
        return None
    rss = subprocess.run(
        ["ps", "-o", "rss=", "-p", proc.stdout.split()[0]], capture_output=True, text=True, check=False
    ).stdout.strip()
    return int(rss) / 1048576 if rss.isdigit() else None


def recycle_word() -> bool:
    """Quit Word between files so the next call starts a fresh, hidden one.

    Word slows down over a long session: on 2026-11-20 its time per LAB pair rose
    from 5s to over 20s within about 200 pairs, at 3.5 GB resident, while the
    documents stayed the same size. Skipped, returning False, while the user is
    using Word (frontmost, or holding a document of theirs).
    """
    state = _word_state()
    if state is None or state[0]:
        return False
    docs = _word_documents()
    if docs is None or any(not _STAGED_RE.match(d) for d in docs):
        return False
    quit_word()
    for _ in range(30):
        if not word_running():
            return True
        time.sleep(1)
    return False


def restart_word() -> None:
    """Restart Word for a builder that suspects Word's state rather than a file.

    Refuses, with WordUnavailableError, while Word holds a document this module did
    not stage: those are the user's, and a restart would lose their unsaved work.
    """
    docs = _word_documents()
    if docs is not None and any(not _STAGED_RE.match(d) for d in docs):
        raise WordUnavailableError("Word holds documents of yours, so it was not restarted")
    _restart_word()


def quit_word() -> None:
    """Quit Word if it holds nothing but this module's documents; for a build that started it."""
    docs = _word_documents()
    if docs is None or any(not _STAGED_RE.match(d) for d in docs):
        return
    try:
        subprocess.run(
            ["osascript", "-e", 'tell application "Microsoft Word" to quit saving no'],
            capture_output=True,
            timeout=60,
            check=False,
        )
    except subprocess.TimeoutExpired:
        print("Word did not quit within 60s; left running", flush=True)


def _restart_word() -> None:
    """End Word and start it again in the background, alerts back on.

    For a Word stuck behind an error it shows for some files no matter what
    ``display alerts`` says (seen on 2026-11-19 with Harvey LAB markups): it then
    blocks every later call until someone dismisses the dialog or restarts it.
    """
    print(f"{time.strftime('%Y-%m-%d %H:%M:%S')} Word is not answering; restarting it", flush=True)
    subprocess.run(["pkill", "-TERM", "-x", _WORD_PROCESS], check=False)
    for _ in range(30):
        if subprocess.run(["pgrep", "-x", _WORD_PROCESS], capture_output=True, check=False).returncode != 0:
            break
        time.sleep(1)
    else:
        subprocess.run(["pkill", "-KILL", "-x", _WORD_PROCESS], check=False)
        time.sleep(2)
    # Owner files ("~$...") Word leaves beside documents it was killed holding.
    for lock in STAGING.glob("~$*"):
        lock.unlink(missing_ok=True)
    subprocess.run(["open", "-g", "-j", "-a", _WORD_PROCESS], check=False)
    for _ in range(60):
        try:
            proc = subprocess.run(
                ["osascript", "-e", 'tell application "Microsoft Word" to set display alerts to alerts all'],
                capture_output=True,
                timeout=30,
                check=False,
            )
        except subprocess.TimeoutExpired:
            continue
        if proc.returncode == 0:
            _settle_hidden()
            return
        time.sleep(1)
    raise WordUnavailableError("Word did not come back after a restart")


def _ready_word() -> None:
    """Return once Word answers, hidden, and the user is not using it.

    Every call goes through here, which is what keeps a long build out of sight:

    - Word not running: started hidden (``_launch_hidden``), never by the Apple
      Event that follows, which would show its start window.
    - The user is using Word, meaning it holds a document this build did not stage
      or it is the frontmost app: the build waits, so a user's document is never
      put at risk and nothing opens in front of them.
    - Otherwise Word is hidden before the call. A user who opened Word and closed
      their documents again finds it hidden once the build resumes.

    A Word that does not answer at all is restarted.
    """
    waiting_since: float | None = None
    while True:
        if not word_running():
            _launch_hidden()
        docs = _word_documents()
        if docs is None:
            _restart_word()
            continue
        staged = [d for d in docs if _STAGED_RE.match(d)]
        stuck = False
        for name in staged:
            # Left open by a call that failed, or reopened by Word's own document
            # recovery after a restart; ours either way.
            try:
                subprocess.run(
                    ["osascript", "-e", f'tell application "Microsoft Word" to close document "{name}" saving no'],
                    capture_output=True,
                    timeout=30,
                    check=False,
                )
            except subprocess.TimeoutExpired:
                stuck = True
        if stuck:
            _restart_word()
            continue
        others = [d for d in docs if not _STAGED_RE.match(d)]
        state = _word_state()
        frontmost = state is not None and state[0]
        if others or frontmost:
            if waiting_since is None:
                waiting_since = time.monotonic()
                why = (
                    f"Word has {len(others)} document(s) open that this build did not open"
                    if others
                    else "Word is the frontmost app"
                )
                print(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {why}; waiting until you are done with it", flush=True)
            time.sleep(30)
            continue
        if waiting_since is not None:
            minutes = (time.monotonic() - waiting_since) / 60
            print(f"{time.strftime('%Y-%m-%d %H:%M:%S')} resuming after {minutes:.0f} min", flush=True)
            waiting_since = None
        if state is not None and state[1]:
            _hide_word()
        return


# A call still running after this long is watched (_watch_call), about twice a second: a
# damaged-file prompt shows behind the user's windows until it is answered.
_WATCH_AFTER_SECONDS = 0.5

_WATCH_SCRIPT = """
on run argv
    set answered to ""
    tell application "System Events"
        if not (exists process "Microsoft Word") then return ""
        tell process "Microsoft Word"
            repeat with w in (every window)
                set t to ""
                try
                    set AppleScript's text item delimiters to " "
                    set t to (value of every static text of w) as text
                end try
                set ours to false
                repeat with stem in argv
                    if t contains (stem as text) then set ours to true
                end repeat
                if ours and t contains "unreadable content" then
                    click button "No" of w
                    set answered to "unreadable"
                else if ours and t contains "experienced an error trying to open" then
                    click button "OK" of w
                    if answered is "" then set answered to "error"
                end if
            end repeat
            if visible and not frontmost then set visible to false
        end tell
    end tell
    return answered
end run
"""


def _watch_call(stems: list[str]) -> str:
    """One look at Word during a slow call: answer its damaged-file prompts for this call's files.

    For a file Word calls damaged it sits idle for about 25 to 30 seconds, then asks
    "Word found unreadable content in <file>. Do you want to recover the contents?"
    and, after No, says "Word experienced an error trying to open the file." Neither
    is covered by ``display alerts``, both are readable and answerable while Word is
    hidden, and both name the staged file, so only this call's prompts are touched.
    Unanswered, they held Word until the call timed out and Word was killed, and a
    relaunched Word comes to the front (seen 2026-11-20). Word also unhides itself
    while opening such a file; it is hidden again here unless it is the frontmost app.

    Returns "unreadable" or "error" when it answered a prompt, else "".
    """
    try:
        proc = subprocess.run(
            ["osascript", "-", *stems], input=_WATCH_SCRIPT, capture_output=True, text=True, timeout=20, check=False
        )
    except subprocess.TimeoutExpired:
        return ""
    return proc.stdout.strip() if proc.returncode == 0 else ""


def _run_word(script: str, args: list[str], what: str) -> None:
    """Run one AppleScript against Word, restarting Word and failing this file if it hangs.

    While the call runs it is watched (``_watch_call``), so a file Word calls damaged
    is refused as such, quietly, instead of hanging until Word is killed.
    """
    _ready_word()
    stems = sorted({m.group(0)[:12] for a in args if (m := _STAGED_RE.match(Path(a).name))})
    proc = subprocess.Popen(
        ["osascript", "-", *args],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    started = time.monotonic()
    deadline = started + WORD_TIMEOUT_SECONDS
    unreadable = False
    pending_input: str | None = script
    while True:
        try:
            out, err = proc.communicate(input=pending_input, timeout=0.25)
            break
        except subprocess.TimeoutExpired:
            pending_input = None
        if time.monotonic() - started >= _WATCH_AFTER_SECONDS and stems:
            answered = _watch_call(stems)
            if answered:
                unreadable = unreadable or answered == "unreadable"
                deadline = time.monotonic() + WORD_TIMEOUT_SECONDS
        if time.monotonic() > deadline:
            proc.kill()
            proc.communicate()
            _restart_word()
            raise WordCompareError(f"{what}: Word hung for {WORD_TIMEOUT_SECONDS}s and was restarted")
    if unreadable:
        raise UnreadableDocumentError(f"{what}: Word found unreadable content in the file and could not open it")
    if proc.returncode != 0 or out.strip() != "ok":
        raise WordCompareError(f"{what}: {err.strip() or out.strip()}")


def word_version() -> str:
    """The installed Word's version, recorded beside every reference it produces."""
    info = WORD_APP / "Contents" / "Info.plist"
    if not info.is_file():
        raise WordUnavailableError(f"Microsoft Word is not installed at {WORD_APP}")
    with info.open("rb") as f:
        return str(plistlib.load(f)["CFBundleShortVersionString"])


def _full_name() -> str:
    return pwd.getpwuid(os.getuid()).pw_gecos.split(",")[0].strip()


def scrub(docx: bytes) -> bytes:
    """Word's output with the user's name and the run's time replaced by fixed values.

    Raises WordCompareError if the account's full name survives anywhere.
    """
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(docx)) as src, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "docProps/core.xml":
                text = data.decode("utf-8")
                text = re.sub(
                    r"<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>",
                    f"<cp:lastModifiedBy>{PUBLISHER}</cp:lastModifiedBy>",
                    text,
                )
                text = re.sub(r"(<dcterms:modified[^>]*>)[^<]*(</dcterms:modified>)", rf"\g<1>{FIXED_DATE}\g<2>", text)
                data = text.encode("utf-8")
            elif info.filename.startswith("word/") and info.filename.endswith(".xml"):
                data = re.sub(rb'w:date="[^"]*"', f'w:date="{FIXED_DATE}"'.encode(), data)
                # Word also stamps each revision with the real time of the run.
                data = re.sub(rb'w16du:dateUtc="[^"]*"', f'w16du:dateUtc="{FIXED_DATE}"'.encode(), data)
            entry = zipfile.ZipInfo(info.filename, date_time=(2026, 1, 1, 0, 0, 0))
            entry.compress_type = zipfile.ZIP_DEFLATED
            dst.writestr(entry, data)
    result = out.getvalue()
    name = _full_name()
    if name:
        with zipfile.ZipFile(io.BytesIO(result)) as check:
            leaked = [n for n in check.namelist() if name.encode() in check.read(n)]
        if leaked:
            raise WordCompareError(f"the user's name is still in {', '.join(leaked)} after scrubbing")
    return result


def compare(old: bytes, new: bytes) -> bytes:
    """Word's Compare of ``old`` against ``new``, as a scrubbed .docx."""
    word_version()  # fails clearly when Word is missing
    STAGING.mkdir(parents=True, exist_ok=True)
    stem = uuid.uuid4().hex[:12]
    old_path, new_path, out_path = (STAGING / f"{stem}-{part}.docx" for part in ("old", "new", "word"))
    old_path.write_bytes(old)
    new_path.write_bytes(new)
    try:
        _run_word(
            _SCRIPT, [str(old_path), str(new_path), str(out_path), old_path.name, AUTHOR], "Word's Compare failed"
        )
        if not out_path.is_file():
            raise WordCompareError("Word reported success but saved no comparison")
        return scrub(out_path.read_bytes())
    finally:
        for path in (old_path, new_path, out_path):
            path.unlink(missing_ok=True)


_RESOLVE_SCRIPT = """
on run argv
    set inPath to item 1 of argv
    set outPath to item 2 of argv
    set docName to item 3 of argv
    set mode to item 4 of argv
    set outName to item 5 of argv
    tell application "Microsoft Word"
        set previousAlerts to display alerts
        set display alerts to alerts none
        try
            open (POSIX file inPath) add to recent files false
            set d to missing value
            try
                set d to document docName
            end try
            if d is missing value then error "Word could not open the document"
            if mode is "reject" then
                reject all revisions d
            else
                accept all revisions d
            end if
            set track revisions of d to false
            save as d file name outPath file format format document add to recent files false
            close document outName saving no
        on error msg
            try
                close document docName saving no
            end try
            try
                close document outName saving no
            end try
            set display alerts to previousAlerts
            error msg
        end try
        set display alerts to previousAlerts
    end tell
    return "ok"
end run
"""


def resolve(markup: bytes, mode: str) -> bytes:
    """``markup`` with every tracked change rejected (``mode="reject"``) or accepted.

    Done by Word itself, so the result is the version Word shows, formatting and all;
    then scrubbed like a comparison (``scrub``). Used to turn a third-party markup into
    the two versions of a pair (bench/build.py). The document is closed by its new name,
    because ``save as`` renames it; never as "the active document", which could be
    one the user just opened.
    """
    if mode not in ("reject", "accept"):
        raise ValueError(f"mode must be reject or accept, not {mode!r}")
    word_version()
    STAGING.mkdir(parents=True, exist_ok=True)
    stem = uuid.uuid4().hex[:12]
    src, dst = STAGING / f"{stem}-in.docx", STAGING / f"{stem}-{mode}.docx"
    src.write_bytes(markup)
    try:
        _run_word(
            _RESOLVE_SCRIPT, [str(src), str(dst), src.name, mode, dst.name], f"Word could not {mode} all revisions"
        )
        if not dst.is_file():
            raise WordCompareError(f"Word reported success but saved nothing after {mode}")
        return scrub(dst.read_bytes())
    finally:
        for path in (src, dst):
            path.unlink(missing_ok=True)
