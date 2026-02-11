"""
Convert .msg and .eml email files to PDF in bulk.
Run from any folder; pass a folder path as first argument (default: Extracted_Nested).

  python emails_to_pdf.py
  python emails_to_pdf.py "Z:\\path\\to\\folder_with_emails"
  python emails_to_pdf.py "Z:\\path\\to\\emails" "Z:\\path\\to\\pdf_output"

Requires: pip install extract-msg xhtml2pdf (or pip install pdfkit for better rendering via wkhtmltopdf)
"""
import os
import re
import sys
from email import policy
from email.parser import BytesParser

SOURCE_FOLDER = r"Z:\z-UPS Flight 2976 Cases\Investigation\ORR\Okolona Fire\Response\EMAILS\ORIGINAL 502 Emails from Okolona Download\Extracted_Nested"
PDF_FOLDER = r"Z:\z-UPS Flight 2976 Cases\Investigation\ORR\Okolona Fire\Response\EMAILS\ORIGINAL 502 Emails from Okolona Download\Extracted_Nested\PDFs"

# wkhtmltopdf path for pdfkit (better PDF rendering). Set to None to use xhtml2pdf only.
WKHTMLTOPDF_PATH = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"


def safe_filename(name, max_len=120):
    """Make a string safe for use as a filename."""
    if not name or not name.strip():
        return "email"
    s = re.sub(r'[<>:"/\\|?*\n\r]', "_", name.strip())
    s = s[:max_len].strip() or "email"
    return s


def html_escape(s):
    if s is None:
        return ""
    s = str(s)
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def build_html(subject, from_addr, to_addr, date_str, body_html_or_plain):
    """Build a simple HTML document for the email."""
    body = (body_html_or_plain or "").strip() or "(no body)"
    # If it looks like HTML, embed in div (sanitize for PDF lib); else escape and use <pre>
    if body.startswith("<") and ">" in body[:50]:
        body = body.replace("currentcolor", "inherit").replace("currentColor", "inherit")
        body_block = f'<div class="body">{body}</div>'
    else:
        body_block = f'<pre class="body">{html_escape(body)}</pre>'
    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8"/>
<style>
  body {{ font-family: sans-serif; margin: 1in; }}
  .meta {{ color: #444; margin-bottom: 1em; border-bottom: 1px solid #ccc; padding-bottom: 0.5em; }}
  .meta p {{ margin: 0.2em 0; }}
  .body {{ white-space: pre-wrap; word-wrap: break-word; }}
  .body img {{ max-width: 100%; }}
</style>
</head>
<body>
  <div class="meta">
    <p><b>Subject:</b> {html_escape(subject)}</p>
    <p><b>From:</b> {html_escape(from_addr)}</p>
    <p><b>To:</b> {html_escape(to_addr)}</p>
    <p><b>Date:</b> {html_escape(date_str)}</p>
  </div>
  {body_block}
</body>
</html>"""


def get_email_content_msg(filepath):
    """Get (subject, from, to, date, body) from a .msg file. Body is HTML or plain."""
    import extract_msg
    msg = extract_msg.Message(filepath)
    try:
        subject = msg.subject or ""
        from_addr = msg.sender or ""
        to_addr = getattr(msg, "to", None) or ""
        date_str = str(msg.date) if msg.date else ""
        body = msg.htmlBody or msg.body or ""
        if body and hasattr(body, "decode"):
            body = body.decode("utf-8", errors="replace")
        return subject, from_addr, to_addr, date_str, body or ""
    finally:
        msg.close()


def get_email_content_eml(filepath):
    """Get (subject, from, to, date, body) from an .eml file."""
    with open(filepath, "rb") as f:
        msg = BytesParser(policy=policy.default).parse(f)
    subject = str(msg.get("subject", "") or "")
    from_addr = str(msg.get("from", "") or "")
    to_addr = str(msg.get("to", "") or "")
    date_str = str(msg.get("date", "") or "")
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ctype = (part.get_content_type() or "").lower()
            if ctype == "text/html":
                raw = part.get_payload(decode=True)
                if raw:
                    body = raw.decode("utf-8", errors="replace")
                break
            if ctype == "text/plain" and not body:
                raw = part.get_payload(decode=True)
                if raw:
                    body = raw.decode("utf-8", errors="replace")
    else:
        raw = msg.get_payload(decode=True)
        if raw:
            body = raw.decode("utf-8", errors="replace")
    return subject, from_addr, to_addr, date_str, body or ""


def html_to_pdf(html_string, pdf_path):
    """Convert HTML string to PDF. Uses pdfkit + wkhtmltopdf if available, else xhtml2pdf."""
    if WKHTMLTOPDF_PATH and os.path.isfile(WKHTMLTOPDF_PATH):
        try:
            import pdfkit
            config = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)
            pdfkit.from_string(html_string, pdf_path, configuration=config, options={"encoding": "UTF-8"})
            return
        except Exception:
            pass
    from xhtml2pdf import pisa
    with open(pdf_path, "wb") as pdf_file:
        pisa.CreatePDF(html_string.encode("utf-8"), pdf_file, encoding="utf-8")


def main():
    source = sys.argv[1] if len(sys.argv) > 1 else SOURCE_FOLDER
    out = sys.argv[2] if len(sys.argv) > 2 else PDF_FOLDER

    if not os.path.isdir(source):
        print(f"Source folder not found: {source}")
        return

    os.makedirs(out, exist_ok=True)
    print("=" * 60)
    print("Converting .msg and .eml files to PDF")
    print("=" * 60)
    print(f"Source: {source}")
    print(f"Output: {out}\n")

    converted = 0
    failed = 0

    for filename in sorted(os.listdir(source)):
        lower = filename.lower()
        if not (lower.endswith(".msg") or lower.endswith(".eml")):
            continue

        filepath = os.path.join(source, filename)
        if not os.path.isfile(filepath):
            continue

        base = os.path.splitext(filename)[0]
        pdf_name = safe_filename(base) + ".pdf"
        pdf_path = os.path.join(out, pdf_name)
        counter = 1
        while os.path.exists(pdf_path):
            pdf_path = os.path.join(out, f"{safe_filename(base)}_{counter}.pdf")
            counter += 1

        try:
            if lower.endswith(".msg"):
                subject, from_addr, to_addr, date_str, body = get_email_content_msg(filepath)
            else:
                subject, from_addr, to_addr, date_str, body = get_email_content_eml(filepath)

            html = build_html(subject, from_addr, to_addr, date_str, body)
            html_to_pdf(html, pdf_path)
            converted += 1
            safe_print = filename.encode("ascii", "replace").decode("ascii")
            print(f"  OK: {safe_print} -> {os.path.basename(pdf_path)}")
        except Exception as e:
            failed += 1
            safe_print = filename.encode("ascii", "replace").decode("ascii")
            print(f"  FAILED: {safe_print} - {e}")

    print("\n" + "=" * 60)
    print(f"Done. Converted {converted} to PDF. Failed: {failed}.")
    print(f"PDFs saved to: {out}")
    print("=" * 60)


if __name__ == "__main__":
    main()
