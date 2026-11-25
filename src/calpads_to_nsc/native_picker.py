import subprocess
import sys
from pathlib import Path


MAC_DIALOG = """
on run argv
    set dialogTitle to item 1 of argv
    set startFolder to (POSIX file (item 2 of argv)) as alias
    set dialogKind to item 3 of argv
    activate
    try
        if dialogKind is "folder" then
            set picked to choose folder with prompt dialogTitle default location startFolder
        else if dialogKind is "save" then
            set fileName to item 4 of argv
            set picked to choose file name with prompt dialogTitle default location startFolder default name fileName
        else
            set allowedTypes to items 4 thru -1 of argv
            set picked to choose file with prompt dialogTitle of type allowedTypes default location startFolder
        end if
        return POSIX path of picked
    on error number -128
        return ""
    end try
end run
"""


class DialogUnavailable(Exception):
    pass


def pick_path(title, start, extensions=(), *, choose_folder=False, save_name=None):
    folder = Path(start)
    if not folder.is_dir():
        folder = Path.cwd()
    if sys.platform == "darwin":
        kind = "folder" if choose_folder else ("save" if save_name else "file")
        options = [save_name] if save_name else [ext.lstrip(".") for ext in extensions]
        command = ["/usr/bin/osascript", "-e", MAC_DIALOG, title, str(folder), kind, *options]
        try:
            result = subprocess.run(command, capture_output=True, text=True)
        except OSError:
            raise DialogUnavailable() from None
        if result.returncode:
            raise DialogUnavailable()
        selected = result.stdout.removesuffix("\n")
    else:
        try:
            import tkinter as tk
            from tkinter import filedialog
        except ImportError:
            raise DialogUnavailable() from None
        try:
            root = tk.Tk()
        except tk.TclError:
            raise DialogUnavailable() from None
        try:
            root.withdraw()
            options = dict(parent=root, title=title, initialdir=str(folder))
            if choose_folder:
                selected = filedialog.askdirectory(**options, mustexist=True)
            else:
                options["filetypes"] = [("Supported files", tuple("*" + ext for ext in extensions))]
                if save_name:
                    selected = filedialog.asksaveasfilename(**options, initialfile=save_name)
                else:
                    selected = filedialog.askopenfilename(**options)
        except tk.TclError:
            raise DialogUnavailable() from None
        finally:
            root.destroy()
    return Path(selected) if selected else None
