# Email & File Utilities Toolkit

A collection of Python scripts for processing Outlook emails, managing files, and system cleanup.

## Scripts

### Email Processing
| Script | Description |
|---|---|
| `bulk_msg_to_pdf.py` | Convert `.msg` and `.eml` files to PDF using `extract_msg` + `pdfkit` (wkhtmltopdf). Handles embedded images, HTML sanitization, and Outlook quirks. |
| `find_and_extract_nested_msg.py` | Scan `.msg` and `.eml` files for nested email attachments and extract them into organized folders. Supports recursive extraction. |
| `extract_attachments.py` | Extract all non-email attachments from `.msg` and `.eml` files into per-email subfolders. |
| `emails_to_pdf.py` | Earlier/simpler email-to-PDF converter (superseded by `bulk_msg_to_pdf.py`). |

### File Cleanup
| Script | Description |
|---|---|
| `cleanup_downloads.py` | **Reusable Downloads folder cleaner & organizer.** Deletes junk, empty folders, old installers, redundant ZIPs, and organizes files by type. |
| `find_duplicates.py` | Find duplicate files by content hash (MD5). Works on any folder. |
| `clean_empty_folders.py` | Remove empty folders recursively from a given path. |
| `remove_bin_files.py` | Remove `.bin` files (attachments without extensions) and clean up resulting empty folders. |

### Utilities
| Script | Description |
|---|---|
| `md_to_docx.py` | Convert Markdown files to Word (.docx) with tables, formatting, and styling. |

## Quick Start

### Clean up your Downloads folder
```bash
# Preview what would be cleaned (no changes made)
python cleanup_downloads.py --dry-run

# Run the cleanup
python cleanup_downloads.py

# Clean a custom path
python cleanup_downloads.py "D:\MyDownloads"
```

### Convert emails to PDF
```bash
# Convert all .msg/.eml files in a folder to PDF
python bulk_msg_to_pdf.py "Z:\path\to\EMAILS" "Z:\path\to\EMAILS\PDFs"
```

### Extract nested emails
```bash
# Find and extract emails within emails
python find_and_extract_nested_msg.py
# (edit SOURCE and OUTPUT paths at top of script)
```

### Find duplicate files
```bash
python find_duplicates.py "C:\path\to\scan"
```

## Requirements

```
extract-msg
pdfkit
python-docx
```

Also requires:
- [wkhtmltopdf](https://wkhtmltopdf.org/downloads.html) installed at `C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe`
- Python 3.10+
