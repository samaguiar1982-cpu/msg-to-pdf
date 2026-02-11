"""
Remove empty folders under the given path. Run to clean up Extracted_Attachments.

  python clean_empty_folders.py "Z:\\path\\to\\Extracted_Attachments"
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

removed = 0
for root, dirs, _ in os.walk(path, topdown=False):
    for d in dirs:
        p = os.path.join(root, d)
        if os.path.isdir(p) and not os.listdir(p):
            try:
                os.rmdir(p)
                removed += 1
                safe = d.encode("ascii", "replace").decode("ascii")
                print(f"  Removed: {safe}")
            except OSError as e:
                print(f"  Skip {d}: {e}")

print(f"Removed {removed} empty folder(s) under {path}")
