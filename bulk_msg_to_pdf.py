"""
Convert .msg files to PDF using extract_msg + pdfkit (wkhtmltopdf).

Properly handles Outlook HTML by:
  - Extracting <style> and <body> content from the email's full HTML document
  - Replacing cid: image references with base64 data URIs
  - Building a single flat HTML document (no nested <html>/<body>)
  - Converting via wkhtmltopdf for proper rendering

  python bulk_msg_to_pdf.py

Requires: pip install extract-msg pdfkit
Requires: wkhtmltopdf installed
"""
import os
import re
import sys
import base64
import mimetypes
import extract_msg
import pdfkit
from email import policy
from email.parser import BytesParser

# wkhtmltopdf configuration
WKHTMLTOPDF_PATH = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
config = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)

INPUT_FOLDER = r"Z:\z-UPS Flight 2976 Cases\Investigation\ORR\Okolona Fire\Response\EMAILS\ORIGINAL 502 Emails from Okolona Download"
OUTPUT_FOLDER = r"Z:\z-UPS Flight 2976 Cases\Investigation\ORR\Okolona Fire\Response\EMAILS\ORIGINAL 502 Emails from Okolona Download\PDFs"

PDFKIT_OPTIONS = {
    "encoding": "UTF-8",
    "enable-local-file-access": "",
    "no-stop-slow-scripts": "",
    "quiet": "",
}


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


def extract_body_and_styles(html):
    """Extract inner <body> content and <style> blocks from a full HTML document.
    Returns (body_inner_html, all_styles_concatenated)."""
    if not html or not html.strip():
        return html, ""

    # Extract all <style> blocks
    styles = []
    for m in re.finditer(r"<style[^>]*>(.*?)</style>", html, re.DOTALL | re.IGNORECASE):
        styles.append(m.group(1).strip())
    all_css = "\n".join(styles)

    # Extract inner <body> content
    body_match = re.search(r"<body[^>]*>(.*)</body>", html, re.DOTALL | re.IGNORECASE)
    if body_match:
        body_content = body_match.group(1).strip()
    else:
        # Not a full document; return as-is
        body_content = html

    return body_content, all_css


def resolve_cid_images(html, attachments):
    """Replace cid:xxx references in HTML with base64 data URIs from attachments."""
    if not html or not attachments:
        return html

    # Build a map: content_id -> data URI
    cid_map = {}
    for att in attachments:
        cid = getattr(att, "contentId", None) or getattr(att, "cid", None) or ""
        if not cid:
            continue
        # Strip angle brackets if present: <image001.png> -> image001.png
        cid = cid.strip("<>").strip()
        if not cid:
            continue

        data = att.data
        if not isinstance(data, bytes):
            continue

        # Guess MIME type from filename
        filename = (
            getattr(att, "longFilename", None)
            or getattr(att, "shortFilename", None)
            or ""
        )
        mime_type, _ = mimetypes.guess_type(filename)
        if not mime_type:
            mime_type = "application/octet-stream"

        b64 = base64.b64encode(data).decode("ascii")
        cid_map[cid] = f"data:{mime_type};base64,{b64}"

    if not cid_map:
        return html

    # Replace src="cid:xxx" with the data URI
    def replace_cid(match):
        ref = match.group(1).strip()
        if ref in cid_map:
            return f'src="{cid_map[ref]}"'
        return match.group(0)

    html = re.sub(r'src\s*=\s*["\']cid:([^"\']+)["\']', replace_cid, html, flags=re.IGNORECASE)
    return html


def sanitize_html(html):
    """Fix Outlook HTML quirks that break wkhtmltopdf."""
    if not html:
        return html
    html = re.sub(r'\bdir\s*=\s*["\']auto["\']', 'dir="ltr"', html, flags=re.IGNORECASE)
    html = html.replace("currentcolor", "inherit").replace("currentColor", "inherit")
    return html


def build_html(sender, to, cc, date_str, subject, body_content, extra_css):
    """Build a single flat HTML document: header table + email styles + body content."""
    extra_style_block = f"<style>\n{extra_css}\n</style>" if extra_css else ""
    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  body {{ font-family: sans-serif; margin: 0.75in; font-size: 11pt; }}
  .email-header {{ border-collapse: collapse; margin-bottom: 1em; width: 100%; }}
  .email-header th {{ text-align: left; padding: 4px 10px; background: #f0f0f0;
    border: 1px solid #ccc; width: 70px; font-size: 10pt; color: #333; }}
  .email-header td {{ padding: 4px 10px; border: 1px solid #ccc; font-size: 10pt; }}
  hr.divider {{ border: none; border-top: 1px solid #ccc; margin: 1em 0; }}
  .email-body img {{ max-width: 100%; height: auto; }}
  .email-body {{ word-wrap: break-word; overflow-wrap: break-word; }}
</style>
{extra_style_block}
</head>
<body>
<table class="email-header">
  <tr><th>From</th><td>{html_escape(sender)}</td></tr>
  <tr><th>To</th><td>{html_escape(to)}</td></tr>
  <tr><th>CC</th><td>{html_escape(cc)}</td></tr>
  <tr><th>Date</th><td>{html_escape(date_str)}</td></tr>
  <tr><th>Subject</th><td>{html_escape(subject)}</td></tr>
</table>
<hr class="divider">
<div class="email-body">
{body_content}
</div>
</body>
</html>"""


def build_plaintext_html(sender, to, cc, date_str, subject, plain_text):
    """Build HTML for a plain-text email."""
    return build_html(
        sender, to, cc, date_str, subject,
        f"<pre style=\"white-space:pre-wrap; word-wrap:break-word; font-family:sans-serif; font-size:11pt;\">{html_escape(plain_text)}</pre>",
        "",
    )


def convert_one_msg(msg_path, pdf_path):
    """Read a .msg file and convert to PDF."""
    msg = extract_msg.Message(msg_path)
    try:
        sender = msg.sender or ""
        to = getattr(msg, "to", None) or ""
        cc = getattr(msg, "cc", None) or ""
        date_str = str(msg.date) if msg.date else ""
        subject = msg.subject or ""

        html_body = msg.htmlBody or ""
        plain_body = msg.body or ""

        if html_body and hasattr(html_body, "decode"):
            html_body = html_body.decode("utf-8", errors="replace")
        if plain_body and hasattr(plain_body, "decode"):
            plain_body = plain_body.decode("utf-8", errors="replace")

        if html_body and html_body.strip():
            body_content, extra_css = extract_body_and_styles(html_body)
            body_content = resolve_cid_images(body_content, msg.attachments)
            body_content = sanitize_html(body_content)
            extra_css = sanitize_html(extra_css)
            final_html = build_html(sender, to, cc, date_str, subject, body_content, extra_css)
        else:
            final_html = build_plaintext_html(sender, to, cc, date_str, subject, plain_body)
    finally:
        msg.close()

    pdfkit.from_string(final_html, pdf_path, configuration=config, options=PDFKIT_OPTIONS)


def convert_one_eml(eml_path, pdf_path):
    """Read an .eml file and convert to PDF."""
    with open(eml_path, "rb") as f:
        msg = BytesParser(policy=policy.default).parse(f)

    sender = str(msg.get("from", "") or "")
    to = str(msg.get("to", "") or "")
    cc = str(msg.get("cc", "") or "")
    date_str = str(msg.get("date", "") or "")
    subject = str(msg.get("subject", "") or "")

    html_body = ""
    plain_body = ""
    cid_map = {}

    if msg.is_multipart():
        for part in msg.walk():
            ctype = (part.get_content_type() or "").lower()
            if ctype == "text/html" and not html_body:
                raw = part.get_payload(decode=True)
                if raw:
                    html_body = raw.decode("utf-8", errors="replace")
            elif ctype == "text/plain" and not plain_body:
                raw = part.get_payload(decode=True)
                if raw:
                    plain_body = raw.decode("utf-8", errors="replace")
            # Collect CID images
            cid = part.get("Content-ID", "")
            if cid:
                cid = cid.strip("<>").strip()
                payload = part.get_payload(decode=True)
                if cid and payload:
                    mime = part.get_content_type() or "application/octet-stream"
                    b64 = base64.b64encode(payload).decode("ascii")
                    cid_map[cid] = f"data:{mime};base64,{b64}"
    else:
        raw = msg.get_payload(decode=True)
        if raw:
            ctype = (msg.get_content_type() or "").lower()
            decoded = raw.decode("utf-8", errors="replace")
            if ctype == "text/html":
                html_body = decoded
            else:
                plain_body = decoded

    if html_body and html_body.strip():
        body_content, extra_css = extract_body_and_styles(html_body)
        # Replace CID references
        if cid_map:
            def replace_cid(match):
                ref = match.group(1).strip()
                if ref in cid_map:
                    return f'src="{cid_map[ref]}"'
                return match.group(0)
            body_content = re.sub(r'src\s*=\s*["\']cid:([^"\']+)["\']', replace_cid, body_content, flags=re.IGNORECASE)
        body_content = sanitize_html(body_content)
        extra_css = sanitize_html(extra_css)
        final_html = build_html(sender, to, cc, date_str, subject, body_content, extra_css)
    else:
        final_html = build_plaintext_html(sender, to, cc, date_str, subject, plain_body)

    pdfkit.from_string(final_html, pdf_path, configuration=config, options=PDFKIT_OPTIONS)


def main():
    source = sys.argv[1] if len(sys.argv) > 1 else INPUT_FOLDER
    out = sys.argv[2] if len(sys.argv) > 2 else OUTPUT_FOLDER

    if not os.path.isdir(source):
        print(f"Input folder not found: {source}")
        return

    os.makedirs(out, exist_ok=True)

    email_files = sorted(
        f for f in os.listdir(source)
        if (f.lower().endswith(".msg") or f.lower().endswith(".eml"))
        and os.path.isfile(os.path.join(source, f))
    )
    total = len(email_files)
    print(f"Found {total} .msg/.eml files in {source}", flush=True)
    print(f"Output folder: {out}\n", flush=True)

    converted = 0
    failed = 0

    for i, filename in enumerate(email_files, start=1):
        file_path = os.path.join(source, filename)
        pdf_name = os.path.splitext(filename)[0] + ".pdf"
        pdf_path = os.path.join(out, pdf_name)

        try:
            if filename.lower().endswith(".msg"):
                convert_one_msg(file_path, pdf_path)
            else:
                convert_one_eml(file_path, pdf_path)
            converted += 1
            safe = filename.encode("ascii", "replace").decode("ascii")
            print(f"Converted {i}/{total}: {safe}", flush=True)
        except Exception as e:
            failed += 1
            safe = filename.encode("ascii", "replace").decode("ascii")
            err = str(e).encode("ascii", "replace").decode("ascii")
            print(f"FAILED {i}/{total}: {safe} - {err}", flush=True)

    print("\n" + "=" * 60)
    print(f"Summary: {converted} converted, {failed} failed (total {total})")
    print("=" * 60)


if __name__ == "__main__":
    main()
