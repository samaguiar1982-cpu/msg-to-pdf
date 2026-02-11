"""
Extract all attachments from .msg and .eml files into a folder structure that
identifies which email each attachment came from.

Each source email gets its own subfolder (named after the email file). All
attachments from that email are saved inside that subfolder with their
original filenames (sanitized). Duplicate names get _1, _2, etc.

Nested .msg/.eml attachments are SKIPPED here (they are already in
Extracted_Nested from find_and_extract_nested_msg.py), so you don't get
duplicates. Only the folder is created when an email has at least one
non-email attachment.

  python extract_attachments.py
  python extract_attachments.py "Z:\\path\\to\\folder_with_emails"
  python extract_attachments.py "Z:\\path\\to\\emails" "Z:\\path\\to\\output"

Requires: pip install extract-msg
"""
import os
import re
import sys
from email import policy
from email.parser import BytesParser

SOURCE_FOLDER = r"Z:\z-UPS Flight 2976 Cases\Investigation\ORR\Okolona Fire\Response\EMAILS\ORIGINAL 502 Emails from Okolona Download"
OUTPUT_FOLDER = r"Z:\z-UPS Flight 2976 Cases\Investigation\ORR\Okolona Fire\Response\EMAILS\ORIGINAL 502 Emails from Okolona Download\Extracted_Attachments"


def safe_name(name, max_len=100):
    """Make a string safe for use as a filename or folder name."""
    if not name or not str(name).strip():
        return "unnamed"
    s = re.sub(r'[<>:"/\\|?*\n\r]', "_", str(name).strip())
    s = s[:max_len].strip() or "unnamed"
    return s


def is_nested_email_attachment(att_name):
    """True if this attachment is a nested .msg or .eml (already in Extracted_Nested)."""
    if not att_name:
        return False
    low = att_name.lower().strip()
    return low.endswith(".msg") or low.endswith(".eml")


def get_attachment_data(attachment):
    """Get raw bytes from an extract_msg attachment (handles embedded Message)."""
    data = attachment.data
    if isinstance(data, bytes):
        return data
    if hasattr(data, "exportBytes"):
        return data.exportBytes()
    from email.generator import BytesGenerator
    import io
    buf = io.BytesIO()
    BytesGenerator(buf, policy=policy.default).flatten(data)
    return buf.getvalue()


def get_attachments_msg(filepath):
    """Get list of (filename, raw_bytes) for all attachments in a .msg file."""
    import extract_msg
    msg = extract_msg.Message(filepath)
    result = []
    try:
        for i, att in enumerate(msg.attachments):
            name = (
                getattr(att, "longFilename", None)
                or getattr(att, "shortFilename", None)
                or f"attachment_{i}"
            )
            if not name or not name.strip():
                name = f"attachment_{i}"
            name = safe_name(name, 200)
            if not os.path.splitext(name)[1]:
                name = name + ".bin"
            try:
                data = get_attachment_data(att)
                if data:
                    result.append((name, data))
            except Exception:
                pass
    finally:
        msg.close()
    return result


def get_attachments_eml(filepath):
    """Get list of (filename, raw_bytes) for all attachments in an .eml file."""
    result = []
    try:
        with open(filepath, "rb") as f:
            msg = BytesParser(policy=policy.default).parse(f)
        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            filename = part.get_filename()
            if not filename:
                continue
            filename = safe_name(filename, 200)
            if not os.path.splitext(filename)[1]:
                filename = filename + ".bin"
            payload = part.get_payload(decode=True)
            if payload and len(payload) > 0:
                result.append((filename, payload))
    except Exception:
        pass
    return result


def main():
    source = sys.argv[1] if len(sys.argv) > 1 else SOURCE_FOLDER
    out_root = sys.argv[2] if len(sys.argv) > 2 else OUTPUT_FOLDER

    if not os.path.isdir(source):
        print(f"Source folder not found: {source}")
        return

    os.makedirs(out_root, exist_ok=True)
    print("=" * 60)
    print("Extracting attachments from .msg and .eml files")
    print("=" * 60)
    print(f"Source: {source}")
    print(f"Output: {out_root}")
    print("(Each email gets a subfolder; attachments saved inside with original names.)\n")

    emails_processed = 0
    emails_with_attachments = 0
    total_attachments = 0

    for filename in sorted(os.listdir(source)):
        lower = filename.lower()
        if not (lower.endswith(".msg") or lower.endswith(".eml")):
            continue

        filepath = os.path.join(source, filename)
        if not os.path.isfile(filepath):
            continue

        try:
            if lower.endswith(".msg"):
                attachments = get_attachments_msg(filepath)
            else:
                attachments = get_attachments_eml(filepath)
        except Exception as e:
            print(f"  SKIP {filename}: {e}")
            continue

        # Skip nested .msg/.eml (already in Extracted_Nested) to avoid duplicates
        attachments = [(n, d) for n, d in attachments if not is_nested_email_attachment(n)]

        emails_processed += 1
        if not attachments:
            continue

        base = os.path.splitext(filename)[0]
        folder_name = safe_name(base, 100)
        email_folder = os.path.join(out_root, folder_name)
        if os.path.exists(email_folder) and not os.path.isdir(email_folder):
            n = 1
            while os.path.exists(os.path.join(out_root, f"{folder_name}_{n}")):
                n += 1
            email_folder = os.path.join(out_root, f"{folder_name}_{n}")
        os.makedirs(email_folder, exist_ok=True)

        emails_with_attachments += 1
        safe_fn = filename.encode("ascii", "replace").decode("ascii")
        print(f"  [{safe_fn}] -> {len(attachments)} attachment(s) in folder: {os.path.basename(email_folder).encode('ascii', 'replace').decode('ascii')}")

        for att_name, data in attachments:
            save_path = os.path.join(email_folder, att_name)
            n = 1
            while os.path.exists(save_path):
                stem, ext = os.path.splitext(att_name)
                save_path = os.path.join(email_folder, f"{stem}_{n}{ext}")
                n += 1
            try:
                with open(save_path, "wb") as f:
                    f.write(data)
                total_attachments += 1
                safe_print = att_name.encode("ascii", "replace").decode("ascii")
                print(f"      -> {safe_print}")
            except Exception as e:
                print(f"      -> FAILED {att_name}: {e}")

    # Remove empty folders (e.g. from a previous run that created folders before checking)
    removed = 0
    for root, dirs, _ in os.walk(out_root, topdown=False):
        for d in dirs:
            path = os.path.join(root, d)
            if os.path.isdir(path) and not os.listdir(path):
                try:
                    os.rmdir(path)
                    removed += 1
                except OSError:
                    pass
    if removed:
        print(f"Removed {removed} empty folder(s).")

    print("\n" + "=" * 60)
    print(f"Done. Processed {emails_processed} emails. {emails_with_attachments} had attachments. Extracted {total_attachments} files.")
    print(f"Nested .msg/.eml are skipped (see Extracted_Nested). Attachments are in: {out_root}")
    print("=" * 60)


if __name__ == "__main__":
    main()
