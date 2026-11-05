# --- Automated tests: run with "python -m pytest" (no API call, no real file touched) ---
import os
import time

import pytest

import doc_sorter
import tools


# --- A fake home folder, so the tests never touch the real Downloads or Documents ---
@pytest.fixture
def notes():
    return []                                                           # notifications shown, instead of real toasts


@pytest.fixture
def home(tmp_path, monkeypatch, notes):
    home = tmp_path.resolve()
    dl = home / "Downloads"
    roots = [home / "Documents", home / "Pictures", home / "Music", home / "Videos"]
    for folder in [dl, *roots]:
        folder.mkdir()
    log = home / "journal.log"
    for module in (tools, doc_sorter):
        monkeypatch.setattr(module, "DL", dl)
        monkeypatch.setattr(module, "LOG", log)
        monkeypatch.setattr(module, "notify", lambda title, text, link=None: notes.append((title, text, link)))
    monkeypatch.setattr(tools, "HOME", home)
    monkeypatch.setattr(tools, "ROOTS", roots)
    monkeypatch.setattr(tools, "TO_SORT", dl / "To_sort")
    monkeypatch.setattr(doc_sorter, "LOG_OLD", home / "journal.old.log")
    return home


def move(name, destination):
    return tools.move_file.func(name, destination, "test")             # .func: the Python function behind the tool


# --- Security ---
def test_refuses_files_outside_downloads(home):
    (home / "Documents" / "secret.pdf").write_text("secret")
    assert move("../Documents/secret.pdf", "Documents") == "File not found in Downloads"
    assert (home / "Documents" / "secret.pdf").exists()


def test_refuses_folders(home):
    (home / "Downloads" / "folder").mkdir()
    assert tools.in_downloads("folder") is None


def test_refuses_destination_outside_allowed_folders(home):
    (home / "Desktop").mkdir()
    (home / "Downloads" / "a.txt").write_text("a")
    assert "not allowed" in move("a.txt", "Desktop")
    assert "not allowed" in move("a.txt", "Documents/../..")
    assert (home / "Downloads" / "a.txt").exists()


def test_refuses_missing_destination(home):
    (home / "Downloads" / "a.txt").write_text("a")
    assert "not allowed" in move("a.txt", "Documents/Missing")


# --- Moving ---
def test_moves_and_notifies(home, notes):
    (home / "Documents" / "Bank").mkdir()
    (home / "Downloads" / "statement.pdf").write_text("statement")
    assert move("statement.pdf", "Documents/Bank") == "moved"
    assert (home / "Documents" / "Bank" / "statement.pdf").exists()
    assert notes == [("Sorted: statement.pdf", "Documents/Bank", home / "Documents" / "Bank")]


def test_to_sort_is_created(home, notes):
    (home / "Downloads" / "mystery.bin").write_text("?")
    move("mystery.bin", "To_sort")
    assert (home / "Downloads" / "To_sort" / "mystery.bin").exists()
    assert notes[0][0] == "To sort: mystery.bin"


def test_identical_duplicate_is_deleted(home, notes):
    (home / "Documents" / "a.txt").write_text("same")
    (home / "Downloads" / "a.txt").write_text("same")
    assert "identical" in move("a.txt", "Documents")
    assert not (home / "Downloads" / "a.txt").exists()
    assert notes[0][0] == "Duplicate deleted: a.txt"


def test_same_name_different_content_gets_a_suffix(home):
    (home / "Documents" / "a.txt").write_text("old")
    (home / "Downloads" / "a.txt").write_text("new")
    move("a.txt", "Documents")
    assert (home / "Documents" / "a.txt").read_text() == "old"         # never overwritten
    copies = list((home / "Documents").glob("a_*.txt"))
    assert len(copies) == 1 and copies[0].read_text() == "new"


# --- The other tools ---
def test_list_folders_stops_at_depth(home):
    (home / "Documents" / "a" / "b" / "c" / "d" / "e").mkdir(parents=True)
    folders = tools.list_folders.func().splitlines()
    assert "Documents/a/b/c/d" in folders
    assert "Documents/a/b/c/d/e" not in folders


def test_find_similar(home):
    (home / "Documents" / "Bank").mkdir()
    (home / "Documents" / "Bank" / "bank_statement_march.pdf").write_text("x")
    assert "Documents/Bank/bank_statement_march.pdf" in tools.find_similar.func("bank_statement_april.pdf")
    assert tools.find_similar.func("holiday_photos.zip") == "No similar file already sorted"


def test_read_start_text_and_binary(home):
    (home / "Downloads" / "note.txt").write_text("Invoice number 42")
    (home / "Downloads" / "data.bin").write_bytes(b"\x00\x01\x02")
    assert tools.read_start.func("note.txt") == "Invoice number 42"
    assert tools.read_start.func("data.bin").startswith("Binary file")


# --- The agent ---
def test_installers_stay_in_downloads(home, monkeypatch):
    called = []
    monkeypatch.setattr(doc_sorter, "sort_file", called.append)
    (home / "Downloads" / "setup.exe").write_text("x")
    doc_sorter.Handler().check(home / "Downloads" / "setup.exe")
    assert called == []                                                 # no AI call
    assert "SKIPPED | setup.exe" in (home / "journal.log").read_text()


def test_downloads_in_progress_are_ignored(home, monkeypatch):
    called = []
    monkeypatch.setattr(doc_sorter, "sort_file", called.append)
    monkeypatch.setattr(doc_sorter.time, "sleep", lambda s: None)
    (home / "Downloads" / "video.mp4").write_text("")
    (home / "Downloads" / "video.mp4.part").write_text("")             # Firefox still downloading
    doc_sorter.Handler().check(home / "Downloads" / "video.mp4.part")
    doc_sorter.Handler().check(home / "Downloads" / "video.mp4")
    assert called == []


def test_catch_up_leaves_processed_files_in_place(home, monkeypatch):
    called = []
    monkeypatch.setattr(doc_sorter, "sort_file", called.append)
    monkeypatch.setattr(doc_sorter.time, "sleep", lambda s: None)
    old, new = home / "Downloads" / "kept.pdf", home / "Downloads" / "new.pdf"
    old.write_text("x")
    new.write_text("y")
    past = time.time() - 3600
    os.utime(old, (past, past))                                         # not modified since it was processed
    (home / "journal.old.log").write_text(f"{time.strftime('%Y-%m-%d %H:%M:%S')} | NOUVEAU | kept.pdf\n")
    doc_sorter.catch_up()
    assert called == ["new.pdf"]


def test_in_use(home):
    f = home / "Downloads" / "report.docx"
    f.write_text("x")
    assert not tools.in_use(f)
    with open(f):                                                       # Python, like Word, forbids moving an open file
        assert tools.in_use(f)
    assert not tools.in_use(f)
    assert not tools.in_use(home / "Downloads" / "missing.docx")


def test_open_file_waits_until_closed(home, monkeypatch):
    called = []
    monkeypatch.setattr(doc_sorter, "sort_file", called.append)
    monkeypatch.setattr(doc_sorter.time, "sleep", lambda s: None)
    monkeypatch.setattr(doc_sorter, "waiting", set())
    f = home / "Downloads" / "report.docx"
    f.write_text("x")
    with open(f):                                                       # opened in Word before the agent sorts it
        doc_sorter.Handler().check(f)
        doc_sorter.Handler().check(f)                                   # Word saves: new event, same file
        assert called == []                                             # no AI call while it is open
        assert doc_sorter.closed_files() == []
    log = (home / "journal.log").read_text()
    assert log.count("WAITING | report.docx") == 1
    for p in doc_sorter.closed_files():
        doc_sorter.Handler().check(p)
    assert called == ["report.docx"]                                    # closed: sorted
    assert doc_sorter.waiting == set()


def test_office_lock_files_are_ignored(home, monkeypatch):
    called = []
    monkeypatch.setattr(doc_sorter, "sort_file", called.append)
    monkeypatch.setattr(doc_sorter.time, "sleep", lambda s: None)
    (home / "Downloads" / "~$report.docx").write_text("x")
    doc_sorter.Handler().check(home / "Downloads" / "~$report.docx")
    assert called == []


# --- Monthly budget ---
class FakeAnthropic:
    """Instead of the API: 2 requests, the second one moves the file."""
    def __init__(self, api_key):
        self.beta = self
        self.messages = self

    def tool_runner(self, **kwargs):
        usage = type("U", (), {"input_tokens": 1000, "cache_creation_input_tokens": None,
                               "cache_read_input_tokens": 0, "output_tokens": 200})()
        def answers():
            yield type("M", (), {"usage": usage})()
            move("invoice.pdf", "Documents")
            yield type("M", (), {"usage": usage})()
        return answers()


def test_month_cost_sums_this_month_only(home):
    (home / "journal.old.log").write_text(
        "2026-11-01 08:00:00 | TOKENS | a.pdf | 2 requests | 10 in / 5 out | $0.0100\n"
        "2026-10-31 23:59:59 | TOKENS | b.pdf | 2 requests | 10 in / 5 out | $5.0000\n")
    (home / "journal.log").write_text(
        "2026-11-05 10:00:00 | TOKENS | c$.pdf | 3 requests | 10 in / 5 out | $0.0250\n"
        "2026-11-05 10:00:01 | SORTED | d.pdf -> Documents | reason: cost $9\n")
    now = doc_sorter.datetime(2026, 11, 20)
    assert doc_sorter.month_cost(now) == pytest.approx(0.035)


def test_sort_logs_its_cost(home, monkeypatch):
    monkeypatch.setattr(doc_sorter.anthropic, "Anthropic", FakeAnthropic)
    monkeypatch.setenv("DOC_SORTER_API_KEY", "test")
    (home / "Downloads" / "invoice.pdf").write_text("x")
    doc_sorter.sort_file("invoice.pdf")
    assert (home / "Documents" / "invoice.pdf").exists()
    assert "TOKENS | invoice.pdf | 2 requests | 2000 in / 400 out | $0.0004" in (home / "journal.log").read_text()


def test_budget_reached_leaves_file_without_calling_the_api(home, monkeypatch, notes):
    monkeypatch.setattr(doc_sorter, "month_cost", lambda: 1.0123)
    monkeypatch.setattr(doc_sorter.anthropic, "Anthropic", None)        # would crash if called
    (home / "Downloads" / "invoice.pdf").write_text("x")
    doc_sorter.sort_file("invoice.pdf")
    assert (home / "Downloads" / "invoice.pdf").exists()
    log = (home / "journal.log").read_text()
    assert "BUDGET REACHED | invoice.pdf | $1.0123 spent this month (limit $1.00)" in log
    assert "NEW | invoice.pdf" not in log                               # not marked as processed: sorted by a later catch-up
    assert len(notes) == 1
