import time
import os
import sys
import shutil
import winshell
from pathlib import Path
from win32com.client import Dispatch
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from win10toast import ToastNotifier

# --- SETTINGS ---
DOWNLOADS_PATH = Path.home() / "Downloads"
toaster = ToastNotifier()

FILE_TYPES = {
    "Images": [".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp"],
    "Videos": [".mp4", ".mov", ".avi", ".mkv"],
    "Music": [".mp3", ".wav", ".flac"],
    "Installers": [".exe", ".msi", ".zip", ".rar", ".7z"],
    "Documents": [".pdf", ".txt", ".docx", ".xlsx", ".pptx", ".csv"]
}

def add_to_startup():
    """Creates a shortcut in the Windows Startup folder for onedir EXE."""
    try:
        startup_path = winshell.startup()
        
        # Get the absolute path of the running EXE
        # In onedir mode, sys.executable is the path to the EXE inside the folder
        exe_path = os.path.realpath(sys.executable)
        exe_name = os.path.basename(exe_path)
        
        # Define where the shortcut will live
        shortcut_path = os.path.join(startup_path, "DownloadOrganizer.lnk")

        # We always update the shortcut in case the folder was moved
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = exe_path
        # CRITICAL: Set the working directory to the folder containing the EXE
        shortcut.WorkingDirectory = os.path.dirname(exe_path)
        shortcut.IconLocation = exe_path
        shortcut.Description = "Automated Download Organizer"
        shortcut.save()
        return True
    except Exception as e:
        # If running as a script (not EXE), this might fail or behave differently
        return False

def move_files():
    moved_count = 0
    if not DOWNLOADS_PATH.exists():
        return 0

    for item in DOWNLOADS_PATH.iterdir():
        if item.is_dir() or item.name.startswith('.') or item.suffix.lower() in [".crdownload", ".tmp", ".part"]:
            continue

        file_ext = item.suffix.lower()
        for category, extensions in FILE_TYPES.items():
            if file_ext in extensions:
                dest_folder = DOWNLOADS_PATH / category
                dest_folder.mkdir(exist_ok=True)
                
                target_path = dest_folder / item.name
                if target_path.exists():
                    target_path = dest_folder / f"{item.stem}_{int(time.time())}{item.suffix}"
                
                try:
                    time.sleep(1) # Wait for file lock
                    shutil.move(str(item), str(target_path))
                    moved_count += 1
                except:
                    pass
    return moved_count

class DownloadHandler(FileSystemEventHandler):
    def on_modified(self, event):
        if not event.is_directory:
            move_files()

if __name__ == "__main__":
    # Ensure startup shortcut is set/updated
    add_to_startup()

    # Initial Clean
    initial_count = move_files()
    if initial_count > 0:
        toaster.show_toast("Organizer Active", f"Sorted {initial_count} files!", duration=5)

    # Observer setup
    event_handler = DownloadHandler()
    observer = Observer()
    observer.schedule(event_handler, str(DOWNLOADS_PATH), recursive=False)
    observer.start()
    
    try:
        while True:
            time.sleep(10)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()