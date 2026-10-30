# agent-doc-sorter

AI agent that automatically sorts downloaded files on Windows.

When a file lands in Downloads, the agent looks at your existing folders, reads the beginning of the file if needed, then moves it to the right place in Documents, Pictures, Music or Videos. When in doubt, it puts it in `Downloads\To_sort` and shows a notification.

The reasons Claude gives (shown in the log and notifications) are written in the language set by `LANGUAGE`.

## How it works

- `doc_sorter.py` watches `~/Downloads` (watchdog). For each new file, a Claude agent (`claude-haiku-5-5`, via `tool_runner`) decides where to put it using 4 tools:
  - `list_folders`: existing folders under Documents, Pictures, Music, Videos (4 levels deep);
  - `find_similar`: already sorted files whose name shares at least 2 words with the new one (so a new version goes next to the old one);
  - `read_start`: beginning of the content (PDF, Word, Excel, PowerPoint, text; other binary files are not read). Images (resized to 1568 px) and the first page of scanned PDFs are sent as-is to Claude, which reads them;
  - `move_file`: moves the file to an existing, allowed folder, or to `Downloads/To_sort` when in doubt.
- Safeguards: the agent only touches files placed directly in Downloads, only moves them into existing folders under Documents, Pictures, Music and Videos, and never overwrites a file (a date-time suffix is added if the name already exists; an identical copy is deleted).
- Installers (`.exe`, `.msi`) are left in Downloads, without calling the AI: they are usually run right away.
- Every action is written to `journal.log` (next to the script, ignored by git). On startup, a log larger than 1 MB is renamed `journal.old.log` (replacing the previous one) and a new log is started.
- On startup, the agent sorts files less than 7 days old that arrived while it was stopped. A file already processed (its name appears in the log) and not modified since is assumed to have been left in Downloads on purpose, and is not moved.
- A Windows notification shows where each file went (or that it was a duplicate), each file sent to `To_sort` and each error. Clicking it opens the folder (or the log, for an error).
- Only one agent runs at a time: if another one is already running, a new instance writes `ALREADY RUNNING` to the log and stops.

**Privacy**: your folder names, the names of some already sorted files and the beginning of downloaded files' content are sent to the Anthropic API to decide where files go.

## Installation

Requirements: Windows 10/11, Python 3.10 or newer, an account on [console.anthropic.com](https://console.anthropic.com) (the API is billed per use, separately from any Claude subscription).

1. Clone the repository and install the dependencies:

   ```powershell
   git clone https://github.com/phillarr/agent-doc-sorter.git $HOME\agent-doc-sorter
   cd $HOME\agent-doc-sorter
   pip install -r requirements.txt
   ```

2. Create an API key in the Anthropic console (Settings → API Keys), then save it as a User environment variable, in PowerShell:

   ```powershell
   [Environment]::SetEnvironmentVariable("DOC_SORTER_API_KEY", "your-key", "User")
   ```

3. Set `USER_NAME` (your first name) and `LANGUAGE` (language of the reasons, e.g. `"English"`) at the top of `doc_sorter.py`.

4. Test it in the foreground (Ctrl + C to stop), in a new PowerShell window so the variable is picked up:

   ```powershell
   python $HOME\agent-doc-sorter\doc_sorter.py
   ```

5. Start the agent with Windows: create the file `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\agent-doc-sorter.bat` containing:

   ```bat
   start "" pyw "%USERPROFILE%\agent-doc-sorter\doc_sorter.py"
   ```

   `pyw` runs Python without a window. If it is not installed, replace `pyw` with the full path to `pythonw.exe`.

## After each code change

The running process keeps the old version: you need to stop it and start it again.

On startup, the agent writes its process ID to `agent.pid` (ignored by git), so you can stop only this process:

```powershell
Get-Process -Id (Get-Content $HOME\agent-doc-sorter\agent.pid) -ErrorAction SilentlyContinue | Where-Object Name -eq pythonw | Stop-Process -Force
explorer.exe "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\agent-doc-sorter.bat"
```

The `pythonw` filter avoids stopping another program if the ID has been reused. Going through `explorer.exe` ensures the agent inherits the `DOC_SORTER_API_KEY` user variable.

Then check in `journal.log` that a recent `STARTUP` line appears, and that a test file is `SORTED`. An authentication error only shows up on the first file.

## License

MIT, see [LICENSE](LICENSE).
