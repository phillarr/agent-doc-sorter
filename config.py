# --- Configuration: the settings to adapt ---
from pathlib import Path                            # path handling

USER_NAME = "Phil"                                  # first name used in the prompt: change it
LANGUAGE = "French"                                 # language of the reasons (log, notifications): change it

HOME = Path.home()                                  # e.g. C:\Users\<name>
DL = HOME / "Downloads"                             # watched folder
TO_SORT = DL / "To_sort"                            # destination when in doubt
ROOTS = [HOME / "Documents", HOME / "Pictures",     # only allowed places (and their subfolders)
         HOME / "Music", HOME / "Videos"]

HERE = Path(__file__).parent                        # project folder
LOG = HERE / "journal.log"                          # log file, next to the script
LOG_OLD = HERE / "journal.old.log"                  # previous log, kept after rotation
LOG_MAX = 1_000_000                                 # on startup, the log is rotated above 1 MB (about a year)
PID = HERE / "agent.pid"                            # process ID, to stop only this agent

CATCHUP_DAYS = 7                                    # on startup, sorts files less than 7 days old
DEPTH = 4                                           # subfolder levels shown to the AI (Documents/Work/Client/Project/2026)
INSTALLERS = (".exe", ".msi")                       # left in Downloads: they are usually run right away
IMAGES = (".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".tif", ".tiff")  # shown to the AI, which can read them
