"""
Remove all .bin files under the given folder (e.g. Extracted_Attachments).
If a folder becomes empty after removing .bin files, the folder is removed too.

  python remove_bin_files.py "Z:\\path\\to\\Extracted_Attachments"
"""
import os
import sys

if len(sys.argv) < 2:
    path = r"Z:\z-UPS Flight 2976 Cases\Investigation\ORR\Okolona Fire\Response\EMAILS\ORIGINAL 502 Emails from Okolona Download\Extracted_Attachments"
else:
    path = sys.argv[1]

if not os.path.isdir(path):
    print(f"Not a folder: {path}")
    sys.exit(1)

removed_files = 0
for root, _, files in os.walk(path):
    for f in files:
        if f.lower().endswith(".bin"):
            p = os.path.join(root, f)
            try:
                os.remove(p)
                removed_files += 1
            except OSError as e:
                print(f"  Skip {p}: {e}")

removed_dirs = 0
for root, dirs, _ in os.walk(path, topdown=False):
    for d in dirs:
        p = os.path.join(root, d)
        if os.path.isdir(p) and not os.listdir(p):
            try:
                os.rmdir(p)
                removed_dirs += 1
            except OSError:
                pass

print(f"Removed {removed_files} .bin file(s) and {removed_dirs} empty folder(s) under {path}")
