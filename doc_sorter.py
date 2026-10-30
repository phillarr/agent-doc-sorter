# --- Libraries ---
import os, re, io, time, base64, shutil, filecmp, subprocess, anthropic  # processes, words, memory, time, encoding, files, notifications, Claude API
from datetime import datetime                       # date and time for the log
from pathlib import Path                            # path handling
from anthropic import beta_tool                     # turns a Python function into a tool for the AI
from pypdf import PdfReader, PdfWriter              # reads PDFs, extracts a page
from PIL import Image                               # shrinks images before sending
from docx import Document                           # reads Word files
from openpyxl import load_workbook                  # reads Excel files
from pptx import Presentation                       # reads PowerPoint files
from watchdog.observers import Observer             # watches a folder
from watchdog.events import FileSystemEventHandler  # reacts to new files

# --- Configuration ---
USER_NAME = "Phil"                                  # first name used in the prompt: change it
LANGUAGE = "French"                                 # language of the reasons (log, notifications): change it
HOME = Path.home()                                  # e.g. C:\Users\<name>
DL = HOME / "Downloads"                             # watched folder
TO_SORT = DL / "To_sort"                            # destination when in doubt
ROOTS = [HOME / "Documents", HOME / "Pictures",     # only allowed places (and their subfolders)
         HOME / "Music", HOME / "Videos"]
LOG = Path(__file__).parent / "journal.log"         # log file, next to the script
CATCHUP_DAYS = 7                                    # on startup, sorts files less than 7 days old
PID = Path(__file__).parent / "agent.pid"           # process ID, to stop only this agent
DEPTH = 4                                           # subfolder levels shown to the AI (Documents/Work/Client/Project/2026)
IMAGES = (".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".tif", ".tiff")  # shown to the AI, which can read them

# --- Log: written to the screen AND to the file ---
def log(message):
    line = f"{datetime.now():%Y-%m-%d %H:%M:%S} | {message}"
    print(line)
    with open(LOG, "a", encoding="utf-8") as f:                        # "a" = append without erasing
        f.write(line + "\n")

# --- Windows notification: warns when a file is not sorted ---
TOAST = (
    "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null;"
    "$x = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02);"
    "$t = $x.GetElementsByTagName('text');"
    "$t.Item(0).AppendChild($x.CreateTextNode($env:TOAST_TITLE)) | Out-Null;"
    "$t.Item(1).AppendChild($x.CreateTextNode($env:TOAST_TEXT)) | Out-Null;"
    "if ($env:TOAST_LINK) { $x.DocumentElement.SetAttribute('activationType', 'protocol');"  # click = opens the link
    " $x.DocumentElement.SetAttribute('launch', $env:TOAST_LINK) };"
    "[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("
    "'{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\\WindowsPowerShell\\v1.0\\powershell.exe')"
    ".Show([Windows.UI.Notifications.ToastNotification]::new($x))")

def notify(title, text, link=None):                                     # link: folder or file opened by a click
    try:                                                                # texts passed as variables: no injection through a file name
        subprocess.Popen(["powershell.exe", "-NoProfile", "-Command", TOAST],
                         env={**os.environ, "TOAST_TITLE": title, "TOAST_TEXT": text,
                              "TOAST_LINK": link.as_uri() if link else ""},
                         creationflags=subprocess.CREATE_NO_WINDOW)     # no window, no waiting
    except Exception as e:
        log(f"  notification failed: {e}")

# --- Security: the AI can only touch files placed directly in Downloads ---
def in_downloads(name):
    f = (DL / name).resolve()
    if f.parent != DL.resolve() or not f.is_file():                     # blocks "../Documents/..." and folders
        log(f"  refused: {name} (outside Downloads)")
        return None
    return f

# --- Tool 1: see the existing folders ---
@beta_tool
def list_folders() -> str:
    """Lists the existing folders and subfolders where a file can be sorted."""
    folders = []
    for root in ROOTS:
        for dirpath, subdirs, _ in os.walk(root):
            subdirs[:] = sorted(s for s in subdirs if not s.startswith("."))  # skips hidden folders
            if len(Path(dirpath).relative_to(root).parts) >= DEPTH:
                subdirs.clear()                                         # does not go deeper
            folders.append(Path(dirpath).relative_to(HOME).as_posix())  # e.g. Documents/Bank/Statements
    return "\n".join(folders[:500])                                     # capped to stay fast

# --- Tool 2: find similar files already sorted ---
@beta_tool
def find_similar(name: str) -> str:
    """Searches the sorted folders for files whose name resembles this one (same subject, another version, same series).
    A similar file already sorted often points to the right folder."""
    words = {w for w in re.split(r"[\W_]+", Path(name).stem.lower())   # words of the name, accents included
             if len(w) >= 4 and not w.isdigit()}                        # skips "ia", "(1)", "2026"…
    if not words:
        return "Name too short to search for similar files"
    found = []
    for root in ROOTS:
        for dirpath, subdirs, files in os.walk(root):
            subdirs[:] = [s for s in subdirs if not s.startswith(".")]
            for f in files:
                common = words & set(re.split(r"[\W_]+", Path(f).stem.lower()))
                if len(common) >= min(2, len(words)):                   # at least 2 words in common
                    found.append((len(common), (Path(dirpath) / f).relative_to(HOME).as_posix()))
    if not found:
        return "No similar file already sorted"
    found.sort(reverse=True)                                            # most similar first
    return "\n".join(f"{path} ({n} words in common)" for n, path in found[:10])

# --- Tool 3: read the beginning of a file ---
@beta_tool
def read_start(name: str) -> str | list:
    """Reads the beginning of a PDF, Word, Excel, PowerPoint, text or image file from the Downloads folder.
    For an image or a scanned PDF, returns the image of the first page."""
    f = in_downloads(name)
    if not f:
        return "File not found in Downloads"
    ext = f.suffix.lower()
    try:
        if ext == ".pdf":
            pdf = PdfReader(f)
            text = pdf.pages[0].extract_text() or ""                    # first page
            if len(text.strip()) < 20:                                  # no text: scanned PDF
                page = PdfWriter(); page.add_page(pdf.pages[0])         # first page only, to limit cost
                buffer = io.BytesIO(); page.write(buffer)
                log(f"  {name}: scanned PDF, first page sent to the AI")
                return [{"type": "document", "source": {"type": "base64", "media_type": "application/pdf",
                         "data": base64.standard_b64encode(buffer.getvalue()).decode()}},
                        {"type": "text", "text": "Scanned PDF: here is its first page."}]
        elif ext in IMAGES:
            img = Image.open(f); img.seek(0)                            # first frame of an animated GIF or multipage TIFF
            img = img.convert("RGB"); img.thumbnail((1568, 1568))       # shrunk: fewer tokens, still readable
            buffer = io.BytesIO(); img.save(buffer, "JPEG", quality=85)
            log(f"  {name}: image sent to the AI")
            return [{"type": "image", "source": {"type": "base64", "media_type": "image/jpeg",
                     "data": base64.standard_b64encode(buffer.getvalue()).decode()}},
                    {"type": "text", "text": "Here is the image."}]
        elif ext == ".docx":
            text = "\n".join(p.text for p in Document(f).paragraphs)    # paragraphs
        elif ext in (".xlsx", ".xlsm"):
            sheet = load_workbook(f, read_only=True, data_only=True).active
            text = "\n".join(" | ".join(str(c) for c in row if c is not None)
                             for row in sheet.iter_rows(max_row=20, values_only=True))  # 20 rows
        elif ext == ".pptx":
            text = "\n".join(shape.text_frame.text                      # slide text, in order
                             for slide in Presentation(f).slides
                             for shape in slide.shapes if shape.has_text_frame)
        else:
            with open(f, "rb") as h:
                head = h.read(8000)                                     # 8 KB only, even for a 2 GB video
            if b"\x00" in head:                                         # null bytes = binary (image, program, archive…)
                log(f"  {name}: binary file, not read")
                return "Binary file (image, program, archive…): content unreadable, decide from the name."
            text = head.decode("utf-8", errors="ignore")                # plain text
        log(f"  reading the content of {name}")
        return text[:1500]                                              # 1500 characters max
    except Exception as e:
        return f"Unreadable: {e}"

# --- Tool 4: move a file ---
@beta_tool
def move_file(name: str, destination: str, reason: str) -> str:
    """Moves a file to an existing folder (path as given by list_folders) or to 'To_sort'.
    reason: explain in one short sentence why you chose this folder."""
    src = in_downloads(name)
    if not src:                                                         # already moved or outside Downloads
        return "File not found in Downloads"
    name = src.name
    if destination == "To_sort":
        folder = TO_SORT
        folder.mkdir(exist_ok=True)                                     # the only folder it can create
    else:
        folder = (HOME / destination).resolve()
        allowed = any(folder == r or r in folder.parents for r in ROOTS)
        if not allowed or not folder.is_dir():                          # security: existing and allowed only
            log(f"  refused: {destination}")
            return "Destination not allowed or missing, pick one from the list"
    target = folder / name
    if target.is_file() and filecmp.cmp(src, target, shallow=False):    # same name AND same content, byte by byte
        src.unlink()                                                    # useless copy: the original is already sorted
        log(f"DUPLICATE | {name} already in {folder} (identical) -> downloaded copy deleted")
        notify(f"Duplicate deleted: {name}", f"Already in {folder.relative_to(HOME).as_posix()}", folder)
        return "already there, identical: copy deleted"
    if target.exists():                                                 # same name, different content: avoid overwriting
        target = folder / f"{src.stem}_{datetime.now():%Y-%m-%d_%H%M%S}{src.suffix}"
    shutil.move(src, target)                                            # actual move
    log(f"SORTED | {name} -> {target} | reason: {reason}")
    if folder == TO_SORT:
        notify(f"To sort: {name}", reason, folder)                      # the user has to decide
    else:
        notify(f"Sorted: {target.name}", folder.relative_to(HOME).as_posix(), folder)  # the move is too fast to see: shows where it went
    return "moved"

# --- The agent: the AI decides (ROCIF structured prompt) ---
def sort_file(name):
    log(f"NEW | {name}")
    result = anthropic.Anthropic(api_key=os.environ["DOC_SORTER_API_KEY"]).beta.messages.tool_runner(  # key dedicated to this agent
        model="claude-haiku-5-5", max_tokens=2000,                      # fast, cheap model
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
            f"FORMAT: A single move action, with a one-sentence reason written in {LANGUAGE}."}]
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
    def on_created(self, e): self.check(e.src_path)                     # new file
    def on_moved(self, e): self.check(e.dest_path)                      # renamed file (end of download)
    def check(self, p):
        p = Path(p)
        try:
            if p.parent == DL and p.suffix.lower() not in (".crdownload", ".tmp", ".part"):  # skips downloads in progress
                if wait_until_written(p):
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

if other := other_agent():
    log(f"ALREADY RUNNING | another agent is running (PID {other}), this one stops")
    raise SystemExit

# --- Startup ---
obs =Observer(); obs.schedule(Handler(), str(DL)); obs.start()         # watches Downloads
PID.write_text(str(os.getpid()))                                        # allows a targeted stop
log(f"STARTUP | watching {DL} | PID {os.getpid()}")

# --- Catch-up: files that arrived while the agent was stopped ---
def already_seen():
    """Last time the agent processed each file name, according to the log."""
    seen = {}
    if LOG.exists():
        for line in LOG.read_text(encoding="utf-8", errors="ignore").splitlines():
            date, _, rest = line.partition(" | ")
            label, _, name = rest.partition(" | ")
            if label in ("NEW", "NOUVEAU"):                             # NOUVEAU: lines written before the English version
                try:
                    seen[name] = datetime.strptime(date, "%Y-%m-%d %H:%M:%S").timestamp()
                except ValueError:
                    pass                                                # damaged line: skipped
    return seen

cutoff = time.time() - CATCHUP_DAYS * 86400
seen = already_seen()
recent = [p for p in sorted(DL.iterdir()) if p.is_file() and p.stat().st_mtime > cutoff]
pending = [p for p in recent if p.stat().st_mtime > seen.get(p.name, 0)]  # already processed and unchanged since: left on purpose
log(f"CATCH-UP | {len(pending)} file(s) less than {CATCHUP_DAYS} days old"
    f" ({len(recent) - len(pending)} already processed, left in place)")
for p in pending:
    Handler().check(p)                                                  # same filters as a new download

while True: time.sleep(1)                                               # Ctrl + C to stop
