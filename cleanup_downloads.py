"""
Downloads Folder Cleanup & Organizer
=====================================
Reusable script to clean and organize the Downloads folder.

What it does:
  1. Deletes empty folders
  2. Deletes failed downloads (.crdownload)
  3. Deletes temp files (.tmp, ~WRL*)
  4. Deletes old calendar invites (.ics)
  5. Optionally deletes installers (.exe, .msi, .msix) — prompts first
  6. Optionally deletes redundant ZIPs where extracted folder exists — prompts first
  7. Organizes remaining files into subfolders by type

Usage:
  python cleanup_downloads.py                    # defaults to your Downloads folder
  python cleanup_downloads.py "D:\\MyDownloads"  # custom path
  python cleanup_downloads.py --dry-run          # preview only, no changes
"""

import os
import sys
import shutil
import re
from collections import defaultdict

# ── Configuration ──────────────────────────────────────────────────────────────

# Default Downloads path (change if needed)
DEFAULT_DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")

# Folders to never touch (your projects, etc.)
PROTECTED_FOLDERS = {
    "extractemails.py",
}

# Folder mapping by extension
FOLDER_MAP = {
    # Videos
    ".mp4": "Videos", ".mov": "Videos", ".avi": "Videos",
    ".mkv": "Videos", ".wmv": "Videos", ".webm": "Videos",
    # Audio
    ".mp3": "Audio", ".wav": "Audio", ".flac": "Audio",
    ".aac": "Audio", ".m4a": "Audio", ".ogg": "Audio",
    # Images
    ".jpg": "Images", ".jpeg": "Images", ".png": "Images",
    ".gif": "Images", ".bmp": "Images", ".heic": "Images",
    ".webp": "Images", ".svg": "Images", ".tiff": "Images",
    ".ico": "Images",
    # PDFs
    ".pdf": "Case Files",
    # Documents
    ".docx": "Documents", ".doc": "Documents",
    ".pptx": "Documents", ".ppt": "Documents",
    ".txt": "Documents", ".rtf": "Documents",
    # Spreadsheets
    ".xlsx": "Spreadsheets", ".xls": "Spreadsheets",
    ".csv": "Spreadsheets",
    # Web & Code
    ".html": "Web and Code", ".htm": "Web and Code",
    ".php": "Web and Code", ".js": "Web and Code",
    ".json": "Web and Code", ".py": "Web and Code",
    ".md": "Web and Code", ".css": "Web and Code",
    ".xml": "Web and Code",
    # Compressed
    ".zip": "Compressed", ".rar": "Compressed",
    ".7z": "Compressed", ".tar": "Compressed",
    ".gz": "Compressed",
}

# Extensions to auto-delete (junk)
JUNK_EXTENSIONS = {".crdownload", ".tmp", ".ics", ".partial"}

# Installer extensions (prompt before deleting)
INSTALLER_EXTENSIONS = {".exe", ".msi", ".msix"}


# ── Helpers ────────────────────────────────────────────────────────────────────

def safe_print(text):
    """Print with ASCII fallback for emoji/unicode filenames."""
    try:
        print(text)
    except UnicodeEncodeError:
        print(text.encode("ascii", "replace").decode("ascii"))


def ask_yes_no(prompt, default_yes=False):
    """Ask a yes/no question. Returns True for yes."""
    suffix = " [Y/n]: " if default_yes else " [y/N]: "
    answer = input(prompt + suffix).strip().lower()
    if not answer:
        return default_yes
    return answer in ("y", "yes")


def get_folder_for_file(filename):
    """Determine which organized folder a file belongs in."""
    ext = os.path.splitext(filename)[1].lower()
    return FOLDER_MAP.get(ext)


# ── Cleanup Steps ──────────────────────────────────────────────────────────────

def delete_empty_folders(downloads_path, dry_run=False):
    """Delete all empty folders (excluding protected ones)."""
    deleted = 0
    for name in os.listdir(downloads_path):
        path = os.path.join(downloads_path, name)
        if not os.path.isdir(path):
            continue
        if name in PROTECTED_FOLDERS:
            continue
        # Check if truly empty (no files recursively)
        has_files = False
        for _, _, files in os.walk(path):
            if files:
                has_files = True
                break
        if not has_files:
            safe_print(f"  {'[DRY RUN] ' if dry_run else ''}Delete empty folder: {name}")
            if not dry_run:
                try:
                    shutil.rmtree(path)
                    deleted += 1
                except Exception as e:
                    safe_print(f"    ERROR: {e}")
            else:
                deleted += 1
    return deleted


def delete_junk_files(downloads_path, dry_run=False):
    """Delete temp files, failed downloads, calendar invites, etc."""
    deleted = 0
    for name in os.listdir(downloads_path):
        path = os.path.join(downloads_path, name)
        if not os.path.isfile(path):
            continue
        ext = os.path.splitext(name)[1].lower()
        is_junk = ext in JUNK_EXTENSIONS
        is_tmp = name.startswith("~WRL") or name.startswith("~$")
        # Files with no extension and GUID-like names
        is_blob = (not ext and re.match(r"^[0-9a-f]{8}-", name))

        if is_junk or is_tmp or is_blob:
            size_mb = os.path.getsize(path) / (1024 * 1024)
            safe_print(f"  {'[DRY RUN] ' if dry_run else ''}Delete: {name} ({size_mb:.1f} MB)")
            if not dry_run:
                try:
                    os.remove(path)
                    deleted += 1
                except Exception as e:
                    safe_print(f"    ERROR: {e}")
            else:
                deleted += 1
    return deleted


def delete_installers(downloads_path, dry_run=False):
    """Prompt to delete installer files."""
    installers = []
    for name in os.listdir(downloads_path):
        path = os.path.join(downloads_path, name)
        if not os.path.isfile(path):
            continue
        ext = os.path.splitext(name)[1].lower()
        if ext in INSTALLER_EXTENSIONS:
            size_mb = os.path.getsize(path) / (1024 * 1024)
            installers.append((name, path, size_mb))

    if not installers:
        return 0

    total_mb = sum(s for _, _, s in installers)
    print(f"\n  Found {len(installers)} installer(s) ({total_mb:.1f} MB total):")
    for name, _, size_mb in installers:
        safe_print(f"    {name} ({size_mb:.1f} MB)")

    if dry_run:
        print("  [DRY RUN] Would prompt to delete these.")
        return len(installers)

    if ask_yes_no(f"\n  Delete all {len(installers)} installers?"):
        deleted = 0
        for name, path, _ in installers:
            try:
                os.remove(path)
                deleted += 1
            except Exception as e:
                safe_print(f"    ERROR deleting {name}: {e}")
        return deleted
    return 0


def find_redundant_zips(downloads_path, dry_run=False):
    """Find ZIP files where an extracted folder with the same name exists."""
    redundant = []
    for name in os.listdir(downloads_path):
        path = os.path.join(downloads_path, name)
        if not os.path.isfile(path):
            continue
        if not name.lower().endswith(".zip"):
            continue
        folder_name = os.path.splitext(name)[0]
        folder_path = os.path.join(downloads_path, folder_name)
        if os.path.isdir(folder_path):
            size_mb = os.path.getsize(path) / (1024 * 1024)
            redundant.append((name, path, size_mb))

    if not redundant:
        return 0

    total_mb = sum(s for _, _, s in redundant)
    print(f"\n  Found {len(redundant)} redundant ZIP(s) ({total_mb:.1f} MB) — extracted folders exist:")
    for name, _, size_mb in redundant:
        safe_print(f"    {name} ({size_mb:.1f} MB)")

    if dry_run:
        print("  [DRY RUN] Would prompt to delete these.")
        return len(redundant)

    if ask_yes_no(f"\n  Delete all {len(redundant)} redundant ZIPs?", default_yes=True):
        deleted = 0
        for name, path, _ in redundant:
            try:
                os.remove(path)
                deleted += 1
            except Exception as e:
                safe_print(f"    ERROR deleting {name}: {e}")
        return deleted
    return 0


def organize_files(downloads_path, dry_run=False):
    """Move loose files into organized subfolders."""
    # Determine which target folders we need
    needed_folders = set()
    files_to_move = []

    for name in os.listdir(downloads_path):
        path = os.path.join(downloads_path, name)
        if not os.path.isfile(path):
            continue
        # Skip system files
        if name.lower() in ("desktop.ini", "thumbs.db", ".ds_store"):
            continue

        dest_folder = get_folder_for_file(name)
        if dest_folder:
            needed_folders.add(dest_folder)
            files_to_move.append((name, path, dest_folder))

    # Create folders
    for folder in sorted(needed_folders):
        folder_path = os.path.join(downloads_path, folder)
        if not os.path.exists(folder_path):
            if not dry_run:
                os.makedirs(folder_path, exist_ok=True)
            safe_print(f"  {'[DRY RUN] ' if dry_run else ''}Created folder: {folder}")

    # Move files
    moved = 0
    for name, src, dest_folder in files_to_move:
        dest_path = os.path.join(downloads_path, dest_folder, name)
        # Handle name collisions
        if os.path.exists(dest_path):
            base, ext = os.path.splitext(name)
            counter = 1
            while os.path.exists(dest_path):
                dest_path = os.path.join(downloads_path, dest_folder, f"{base} ({counter}){ext}")
                counter += 1

        if not dry_run:
            try:
                shutil.move(src, dest_path)
                moved += 1
            except Exception as e:
                safe_print(f"    ERROR moving {name}: {e}")
        else:
            safe_print(f"  [DRY RUN] Move: {name} -> {dest_folder}/")
            moved += 1

    return moved


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    # Parse arguments
    dry_run = "--dry-run" in sys.argv
    path_args = [a for a in sys.argv[1:] if a != "--dry-run"]
    downloads_path = path_args[0] if path_args else DEFAULT_DOWNLOADS

    if not os.path.isdir(downloads_path):
        print(f"ERROR: Path does not exist: {downloads_path}")
        sys.exit(1)

    mode = "[DRY RUN MODE]" if dry_run else ""
    print(f"\n{'='*60}")
    print(f"  Downloads Cleanup & Organizer {mode}")
    print(f"  Path: {downloads_path}")
    print(f"{'='*60}\n")

    # Count initial state
    initial_files = len([f for f in os.listdir(downloads_path) if os.path.isfile(os.path.join(downloads_path, f))])
    initial_folders = len([f for f in os.listdir(downloads_path) if os.path.isdir(os.path.join(downloads_path, f))])
    print(f"Current state: {initial_files} files, {initial_folders} folders at root\n")

    # Step 1: Empty folders
    print("Step 1: Removing empty folders...")
    n = delete_empty_folders(downloads_path, dry_run)
    print(f"  -> {n} empty folders {'would be ' if dry_run else ''}removed\n")

    # Step 2: Junk files
    print("Step 2: Removing junk files (temp, failed downloads, etc.)...")
    n = delete_junk_files(downloads_path, dry_run)
    print(f"  -> {n} junk files {'would be ' if dry_run else ''}removed\n")

    # Step 3: Installers
    print("Step 3: Checking for old installers...")
    n = delete_installers(downloads_path, dry_run)
    print(f"  -> {n} installers {'would be ' if dry_run else ''}removed\n")

    # Step 4: Redundant ZIPs
    print("Step 4: Checking for redundant ZIPs...")
    n = find_redundant_zips(downloads_path, dry_run)
    print(f"  -> {n} redundant ZIPs {'would be ' if dry_run else ''}removed\n")

    # Step 5: Organize
    if not dry_run:
        if ask_yes_no("Step 5: Organize remaining files into folders?", default_yes=True):
            n = organize_files(downloads_path, dry_run)
            print(f"  -> {n} files moved into organized folders\n")
        else:
            print("  Skipped organization.\n")
    else:
        print("Step 5: Organizing files (preview)...")
        n = organize_files(downloads_path, dry_run)
        print(f"  -> {n} files would be moved\n")

    # Final state
    final_files = len([f for f in os.listdir(downloads_path) if os.path.isfile(os.path.join(downloads_path, f))])
    final_folders = len([f for f in os.listdir(downloads_path) if os.path.isdir(os.path.join(downloads_path, f))])
    print(f"{'='*60}")
    print(f"  Done! Root now has {final_files} files, {final_folders} folders")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
