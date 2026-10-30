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
