# --- agent-doc-sorter: watches Downloads and lets an AI agent sort each new file ---
import os                                           # process ID and API key
import subprocess                                   # checks whether another agent is running
import time                                         # waiting
from datetime import datetime                       # reads the dates of the log
from pathlib import Path                            # path handling

import anthropic                                    # Claude API
from watchdog.events import FileSystemEventHandler  # reacts to new files
from watchdog.observers import Observer             # watches a folder

from config import CATCHUP_DAYS, DL, INSTALLERS, LANGUAGE, LOG, LOG_MAX, LOG_OLD, PID, USER_NAME
from tools import find_similar, list_folders, log, move_file, notify, read_start


# --- The agent: the AI decides (ROCIF structured prompt) ---
def sort_file(name):
    log(f"NEW | {name}")
    client = anthropic.Anthropic(api_key=os.environ["DOC_SORTER_API_KEY"])  # key dedicated to this agent
    result = client.beta.messages.tool_runner(
        model="claude-haiku-5-5",                                       # fast, cheap model
        max_tokens=2000,
        tools=[list_folders, find_similar, read_start, move_file],      # its 4 tools
        messages=[{"role": "user", "content":
            f"ROLE: You are {USER_NAME}'s file sorting assistant.\n"
            f"OBJECTIVE: Move the downloaded file \"{name}\" to the right place.\n"
            f"CONTEXT: {USER_NAME} organizes files by organization or project (e.g. Documents/Bank) "
            "before organizing them by type. Specific subfolders are always preferred over generic folders.\n"
            "INSTRUCTIONS:\n"
            "1. List the existing folders and search for already sorted files with a similar name.\n"
            "2. If a similar file covers the same subject (another version, same series), put it in the same folder.\n"
            "3. Otherwise, if the name is not enough, read the beginning of the file.\n"
            "4. Put it in the most specific existing folder.\n"
            "5. If there is no folder for the organization or project, use a folder by document type "
            "if one exists (e.g. Documents/Invoices for an invoice). Put it in To_sort only if no "
            "folder fits, not even by type. Never guess.\n"
            f"FORMAT: A single move action, with a one-sentence reason written in {LANGUAGE}."}],
    ).until_done()                                                      # runs until the end
    if (DL / name).exists():                                            # the AI stopped without moving the file
        text = " ".join(b.text for b in result.content if b.type == "text").replace("\n", " ")
        log(f"NOT SORTED | {name} | stop: {result.stop_reason} | answer: {text[:300]}")
        move_file(name, "To_sort", "the AI did not choose a folder")    # safety net


# --- Waits until the file is fully written (stable size) ---
def wait_until_written(p, max_seconds=600):
    size = -1
    for _ in range(max_seconds):
        time.sleep(1)
        if not p.is_file():
            return False                                                # disappeared in the meantime
        if (DL / (p.name + ".part")).exists():                          # Firefox: empty file + .part in progress
            return False                                                # we'll be called again when the .part is renamed
        new_size = p.stat().st_size
        if new_size == size:
            return True                                                 # 2 identical readings: download finished
        size = new_size
    log(f"ERROR | {p.name} | size still changing after {max_seconds} s")
    notify(f"Error: {p.name}", "Download never finished, file not sorted", LOG)
    return False


# --- The trigger ---
class Handler(FileSystemEventHandler):
    def on_created(self, e):
        self.check(e.src_path)                                          # new file

    def on_moved(self, e):
        self.check(e.dest_path)                                         # renamed file (end of download)

    def check(self, p):
        p = Path(p)
        try:
            if p.parent == DL and p.suffix.lower() not in (".crdownload", ".tmp", ".part"):  # skips downloads in progress
                if p.suffix.lower() in INSTALLERS:
                    log(f"SKIPPED | {p.name} | installer, left in Downloads")  # no AI call: free
                elif wait_until_written(p):
                    sort_file(p.name)                                   # runs the agent
        except Exception as e:
            log(f"ERROR | {p.name} | {e}")                              # errors go to the log too
            notify(f"Error: {p.name}", str(e)[:200], LOG)               # e.g. expired API key


# --- Single instance: two agents would process every file twice ---
def other_agent():
    try:
        pid = int(PID.read_text())
    except (OSError, ValueError):
        return None                                                     # no PID file or damaged: nothing to check
    if pid == os.getpid():
        return None
    cmd = subprocess.run(["powershell.exe", "-NoProfile", "-Command",
                          f"(Get-CimInstance Win32_Process -Filter 'ProcessId={pid}').CommandLine"],
                         capture_output=True, text=True, errors="ignore",
                         creationflags=subprocess.CREATE_NO_WINDOW).stdout
    return pid if "doc_sorter.py" in cmd else None                      # checks the command line: the number may have been reused


# --- Catch-up: files that arrived while the agent was stopped ---
def already_seen():
    """Last time the agent processed each file name, according to the log."""
    seen = {}
    for f in (LOG_OLD, LOG):                                            # archive first: the current log has the latest dates
        if not f.exists():
            continue
        for line in f.read_text(encoding="utf-8", errors="ignore").splitlines():
            date, _, rest = line.partition(" | ")
            label, _, name = rest.partition(" | ")
            if label in ("NEW", "SKIPPED", "NOUVEAU"):                  # NOUVEAU: lines written before the English version
                if label == "SKIPPED":
                    name = name.partition(" | ")[0]                     # "name | installer, left in Downloads"
                try:
                    seen[name] = datetime.strptime(date, "%Y-%m-%d %H:%M:%S").timestamp()
                except ValueError:
                    pass                                                # damaged line: skipped
    return seen


def catch_up():
    cutoff = time.time() - CATCHUP_DAYS * 86400
    seen = already_seen()
    recent = [p for p in sorted(DL.iterdir()) if p.is_file() and p.stat().st_mtime > cutoff]
    pending = [p for p in recent if int(p.stat().st_mtime) > seen.get(p.name, 0)]  # already processed and unchanged since: left on purpose (log is to the second)
    log(f"CATCH-UP | {len(pending)} file(s) less than {CATCHUP_DAYS} days old"
        f" ({len(recent) - len(pending)} already processed, left in place)")
    for p in pending:
        Handler().check(p)                                              # same filters as a new download


# --- Startup ---
def main():
    other = other_agent()
    if other:
        log(f"ALREADY RUNNING | another agent is running (PID {other}), this one stops")
        return
    if LOG.exists() and LOG.stat().st_size > LOG_MAX:                   # log rotation: keeps the log from growing forever
        LOG.replace(LOG_OLD)                                            # the previous archive is overwritten
        log(f"LOG ROTATED | previous log moved to {LOG_OLD.name}")
    observer = Observer()
    observer.schedule(Handler(), str(DL))                               # watches Downloads
    observer.start()
    PID.write_text(str(os.getpid()))                                    # allows a targeted stop
    log(f"STARTUP | watching {DL} | PID {os.getpid()}")
    catch_up()
    while True:
        time.sleep(1)                                                   # Ctrl + C to stop


if __name__ == "__main__":
    main()
