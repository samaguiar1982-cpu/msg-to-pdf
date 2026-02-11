"""
Find .msg and .eml files that contain nested email attachments (.msg or .eml), then extract them.
Run from any folder; uses the path below or pass a folder as first argument.

  python find_and_extract_nested_msg.py
  python find_and_extract_nested_msg.py "Z:\\path\\to\\EMAILS"
  python find_and_extract_nested_msg.py "Z:\\path\\to\\EMAILS" "Z:\\path\\to\\output"
"""
import extract_msg
import os
import sys
from email import policy
from email.parser import BytesParser

SOURCE_FOLDER = r"Z:\z-UPS Flight 2976 Cases\Investigation\ORR\Okolona Fire\Response\EMAILS\ORIGINAL 502 Emails from Okolona Download"
OUTPUT_FOLDER = r"Z:\z-UPS Flight 2976 Cases\Investigation\ORR\Okolona Fire\Response\EMAILS\ORIGINAL 502 Emails from Okolona Download\Extracted_Nested"

# File extensions we treat as email (nested attachments we want to find/extract)
EMAIL_EXTENSIONS = (".msg", ".eml")


def get_attachment_filename(attachment, index):
    """Get best available filename for an attachment (extract_msg)."""
    return (
        getattr(attachment, "longFilename", None)
        or getattr(attachment, "shortFilename", None)
        or f"embedded_{index}"
    )


def is_email_attachment(attachment, att_name, from_msg=True):
    """True if this attachment is an embedded email (.msg or .eml)."""
    if att_name:
        lower = att_name.lower()
        if lower.endswith(".msg") or lower.endswith(".eml"):
            return True
    if from_msg and hasattr(attachment, "type"):
        t = str(getattr(attachment, "type", "")).lower()
        if "msg" in t or "rfc822" in t or "outlook" in t:
            return True
    return False


def normalize_email_filename(name, prefer_ext=".msg"):
    """Ensure filename has .msg or .eml extension."""
    if not name:
        return f"embedded{prefer_ext}"
    lower = name.lower()
    if lower.endswith(".msg") or lower.endswith(".eml"):
        return name
    return name + prefer_ext


def get_nested_from_eml(filepath):
    """
    Parse an .eml file and return list of (filename, raw_bytes) for parts that are .msg/.eml attachments.
    """
    nested = []
    try:
        with open(filepath, "rb") as f:
            msg = BytesParser(policy=policy.default).parse(f)
        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            filename = part.get_filename()
            ctype = (part.get_content_type() or "").lower()
            # attachment or inline with filename; or content type suggests email
            is_attachment = part.get_content_disposition() in ("attachment", "inline") or filename
            if not is_attachment and "rfc822" not in ctype and "message" not in ctype:
                continue
            if not filename:
                if "rfc822" in ctype or "message" in ctype:
                    filename = f"embedded_{len(nested)}.eml"
                elif "ms-outlook" in ctype or "vnd.ms-outlook" in ctype:
                    filename = f"embedded_{len(nested)}.msg"
                else:
                    continue
            low = filename.lower()
            if not (low.endswith(".msg") or low.endswith(".eml")):
                if "rfc822" in ctype or (part.get_content_type() or "").startswith("message/"):
                    filename = filename + ".eml" if not low.endswith((".msg", ".eml")) else filename
                elif "ms-outlook" in ctype or "vnd.ms-outlook" in ctype:
                    filename = filename + ".msg" if not low.endswith(".msg") else filename
                else:
                    continue
            payload = part.get_payload(decode=True)
            if payload is not None and len(payload) > 0:
                nested.append((filename, payload))
    except Exception:
        pass
    return nested


def run_one_pass(source, out):
    """Scan source folder for .msg/.eml with nested email attachments; extract to out. Returns (scanned, with_nested, extracted)."""
    os.makedirs(out, exist_ok=True)
    scanned_count = 0
    files_with_nested = 0
    extracted_count = 0

    for filename in sorted(os.listdir(source)):
        lower = filename.lower()
        if not (lower.endswith(".msg") or lower.endswith(".eml")):
            continue

        scanned_count += 1
        filepath = os.path.join(source, filename)

        if lower.endswith(".msg"):
            try:
                msg = extract_msg.Message(filepath)
                nested = []
                for i, attachment in enumerate(msg.attachments):
                    att_name = get_attachment_filename(attachment, i)
                    if not is_email_attachment(attachment, att_name):
                        continue
                    att_name = normalize_email_filename(att_name, ".msg")
                    data = attachment.data
                    if not isinstance(data, bytes):
                        if hasattr(data, "exportBytes"):
                            data = data.exportBytes()
                        else:
                            from email.generator import BytesGenerator
                            import io
                            buf = io.BytesIO()
                            BytesGenerator(buf, policy=policy.default).flatten(data)
                            data = buf.getvalue()
                    nested.append((att_name, data))

                if nested:
                    files_with_nested += 1
                    print(f"  [.msg] {filename} -> {len(nested)} nested: {[n[0] for n in nested]}")
                    for att_name, data in nested:
                        safe_parent = filename.replace(".msg", "")[:80]
                        safe_att = att_name[:80]
                        save_name = f"{safe_parent}__nested__{safe_att}"
                        save_path = os.path.join(out, save_name)
                        counter = 1
                        while os.path.exists(save_path):
                            base, ext = os.path.splitext(save_name)
                            save_path = os.path.join(out, f"{base}_{counter}{ext}")
                            counter += 1
                        try:
                            with open(save_path, "wb") as f:
                                f.write(data)
                            extracted_count += 1
                            print(f"      -> saved: {os.path.basename(save_path)}")
                        except Exception as e:
                            print(f"      -> ERROR saving {att_name}: {e}")

                msg.close()
            except Exception as e:
                print(f"  SKIPPED {filename}: {e}")

        else:
            nested = get_nested_from_eml(filepath)
            if nested:
                files_with_nested += 1
                print(f"  [.eml] {filename} -> {len(nested)} nested: {[n[0] for n in nested]}")
                for att_name, data in nested:
                    safe_parent = filename.replace(".eml", "")[:80]
                    safe_att = att_name[:80]
                    save_name = f"{safe_parent}__nested__{safe_att}"
                    save_path = os.path.join(out, save_name)
                    counter = 1
                    while os.path.exists(save_path):
                        base, ext = os.path.splitext(save_name)
                        save_path = os.path.join(out, f"{base}_{counter}{ext}")
                        counter += 1
                    try:
                        with open(save_path, "wb") as f:
                            f.write(data)
                        extracted_count += 1
                        print(f"      -> saved: {os.path.basename(save_path)}")
                    except Exception as e:
                        print(f"      -> ERROR saving {att_name}: {e}")

    return scanned_count, files_with_nested, extracted_count


def main():
    source = sys.argv[1] if len(sys.argv) > 1 else SOURCE_FOLDER
    out = sys.argv[2] if len(sys.argv) > 2 else OUTPUT_FOLDER
    max_depth = 5  # max levels of nesting to unpack

    if not os.path.isdir(source):
        print(f"Source folder not found: {source}")
        return

    total_scanned = 0
    total_with_nested = 0
    total_extracted = 0
    level = 1

    while level <= max_depth:
        print("=" * 60)
        print(f"Level {level}: Scanning for nested email attachments")
        print("=" * 60)
        print(f"Source: {source}")
        print(f"Output: {out}\n")

        scanned, with_nested, extracted = run_one_pass(source, out)
        total_scanned += scanned
        total_with_nested += with_nested
        total_extracted += extracted

        print(f"\nLevel {level} -> Scanned {scanned}, found {with_nested} with nested, extracted {extracted}.")

        if extracted == 0:
            break
        source = out
        out = os.path.join(out, "Extracted_Nested")
        level += 1

    print("\n" + "=" * 60)
    print(f"Done. Total scanned: {total_scanned}. Total with nested: {total_with_nested}. Total extracted: {total_extracted}.")
    print("Output path(s) shown above for each level.")
    print("=" * 60)


if __name__ == "__main__":
    main()
