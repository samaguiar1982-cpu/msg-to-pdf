"""
Find and Remove Duplicate Files
================================
Scans one or more folders (recursively) for duplicate files based on content hash.
Shows a report and optionally moves duplicates to a "_duplicates" folder or deletes them.
"""

import os
import hashlib
import shutil
from collections import defaultdict
from datetime import datetime


# ── CONFIGURE THESE ──────────────────────────────────────────────────────────
FOLDERS_TO_SCAN = [
    r"C:\Users\SAguiar\Downloads",
]

# Where to move duplicates (set to None to delete permanently instead)
DUPLICATES_FOLDER = r"C:\Users\SAguiar\Desktop\_duplicates"

# Minimum file size to consider (in bytes). Set to 0 to include all files.
# 1 = skip completely empty files
MIN_FILE_SIZE = 1

# File extensions to SKIP (lowercase, with dot). Leave empty to scan everything.
SKIP_EXTENSIONS = set()
# Example: SKIP_EXTENSIONS = {".lnk", ".ini", ".tmp"}

# Folders to skip (exact folder names, case-insensitive)
SKIP_FOLDERS = {"node_modules", ".git", "__pycache__", ".venv", "venv"}
# ─────────────────────────────────────────────────────────────────────────────


def get_file_hash(filepath, chunk_size=8192):
    """Return SHA-256 hash of a file's contents."""
    h = hashlib.sha256()
    try:
        with open(filepath, "rb") as f:
            while True:
                chunk = f.read(chunk_size)
                if not chunk:
                    break
                h.update(chunk)
    except (PermissionError, OSError) as e:
        print(f"  [SKIP] Cannot read: {filepath}  ({e})")
        return None
    return h.hexdigest()


def collect_files(folders, min_size, skip_ext, skip_folders):
    """Walk the folders and group files by size (quick pre-filter)."""
    size_map = defaultdict(list)  # size -> [filepath, ...]
    file_count = 0

    for folder in folders:
        folder = os.path.abspath(folder)
        if not os.path.isdir(folder):
            print(f"  [WARN] Folder not found, skipping: {folder}")
            continue
        print(f"  Scanning: {folder}")
        for root, dirs, files in os.walk(folder):
            # Skip unwanted directories
            dirs[:] = [d for d in dirs if d.lower() not in skip_folders]
            for fname in files:
                fpath = os.path.join(root, fname)
                ext = os.path.splitext(fname)[1].lower()
                if skip_ext and ext in skip_ext:
                    continue
                try:
                    fsize = os.path.getsize(fpath)
                except OSError:
                    continue
                if fsize < min_size:
                    continue
                size_map[fsize].append(fpath)
                file_count += 1

    return size_map, file_count


def find_duplicates(size_map):
    """From size-grouped files, hash only potential duplicates and return groups."""
    hash_map = defaultdict(list)  # hash -> [filepath, ...]
    hashed = 0
    # Only hash files where 2+ share the same size
    candidates = {sz: paths for sz, paths in size_map.items() if len(paths) > 1}
    total_candidates = sum(len(p) for p in candidates.values())

    print(f"\n  Hashing {total_candidates} candidate files ...")
    for size, paths in candidates.items():
        for fpath in paths:
            fhash = get_file_hash(fpath)
            if fhash:
                hash_map[fhash].append(fpath)
            hashed += 1
            if hashed % 200 == 0:
                print(f"    ... hashed {hashed}/{total_candidates}")

    # Keep only groups with actual duplicates
    duplicates = {h: paths for h, paths in hash_map.items() if len(paths) > 1}
    return duplicates


def pick_original(paths):
    """
    From a list of duplicate paths, pick the 'original' to KEEP.
    Strategy: keep the file with the shortest path (likely the original name).
    """
    return min(paths, key=lambda p: (len(p), p))


def format_size(size_bytes):
    """Human-readable file size."""
    for unit in ("B", "KB", "MB", "GB"):
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"


def main():
    print("=" * 70)
    print("  DUPLICATE FILE FINDER")
    print("=" * 70)
    print()

    # Step 1: Collect files grouped by size
    print("[1/3] Collecting files ...")
    size_map, file_count = collect_files(
        FOLDERS_TO_SCAN, MIN_FILE_SIZE, SKIP_EXTENSIONS, SKIP_FOLDERS
    )
    print(f"  Found {file_count} files total.\n")

    if file_count == 0:
        print("No files to scan. Check your FOLDERS_TO_SCAN setting.")
        return

    # Step 2: Find duplicates by hashing
    print("[2/3] Finding duplicates ...")
    duplicates = find_duplicates(size_map)

    if not duplicates:
        print("\n  No duplicate files found! Your folders are clean.")
        return

    # Step 3: Report
    dup_count = sum(len(p) - 1 for p in duplicates.values())
    wasted = sum(
        os.path.getsize(p) * (len(paths) - 1)
        for paths in duplicates.values()
        for p in paths[:1]
        if os.path.exists(p)
    )

    print(f"\n{'=' * 70}")
    print(f"  RESULTS: {len(duplicates)} groups, {dup_count} duplicate files")
    print(f"  Wasted space: ~{format_size(wasted)}")
    print(f"{'=' * 70}\n")

    group_num = 0
    for fhash, paths in sorted(duplicates.items(), key=lambda x: -len(x[1])):
        group_num += 1
        original = pick_original(paths)
        size = format_size(os.path.getsize(paths[0])) if os.path.exists(paths[0]) else "?"
        print(f"  Group {group_num} ({len(paths)} copies, {size} each):")
        for p in paths:
            tag = " [KEEP]" if p == original else ""
            print(f"    {'>>>' if p == original else '   '} {p}{tag}")
        print()

    # Step 4: Ask user what to do
    print("-" * 70)
    if DUPLICATES_FOLDER:
        print(f"  Action: MOVE duplicates to {DUPLICATES_FOLDER}")
    else:
        print("  Action: PERMANENTLY DELETE duplicates")
    print(f"  Files to remove: {dup_count}")
    print("-" * 70)

    answer = input("\n  Proceed? (yes / no): ").strip().lower()
    if answer not in ("yes", "y"):
        print("\n  Cancelled. No files were changed.")
        return

    # Step 5: Remove duplicates
    print("\n[3/3] Removing duplicates ...\n")
    removed = 0
    errors = 0

    for fhash, paths in duplicates.items():
        original = pick_original(paths)
        for p in paths:
            if p == original:
                continue
            try:
                if DUPLICATES_FOLDER:
                    # Preserve relative structure inside the duplicates folder
                    rel = os.path.basename(p)
                    dest = os.path.join(DUPLICATES_FOLDER, rel)
                    # Handle name collisions in destination
                    base, ext = os.path.splitext(dest)
                    counter = 1
                    while os.path.exists(dest):
                        dest = f"{base}_{counter}{ext}"
                        counter += 1
                    os.makedirs(DUPLICATES_FOLDER, exist_ok=True)
                    shutil.move(p, dest)
                else:
                    os.remove(p)
                removed += 1
            except (PermissionError, OSError) as e:
                print(f"  [ERROR] {p}: {e}")
                errors += 1

    print(f"\n  Done! {removed} duplicates {'moved' if DUPLICATES_FOLDER else 'deleted'}.")
    if errors:
        print(f"  {errors} files could not be processed (permission errors, etc.).")
    if DUPLICATES_FOLDER and removed:
        print(f"  Review moved files at: {DUPLICATES_FOLDER}")


if __name__ == "__main__":
    main()
