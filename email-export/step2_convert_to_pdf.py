#!/usr/bin/env python3
"""
GAVEL — Step 2: Convert Emails to PDF
======================================
Converts .eml files (from Step 1) into clean, readable PDFs.
Attachments are placed immediately after their email in the output folder.

SETUP (run once):
    Mac:     /usr/local/bin/pip3 install reportlab
    Windows: pip install reportlab

USAGE:
    1. Run step1_extract_emails.py first
    2. Edit the JOBS section below to match your Step 1 output folders
    3. Run: python3 step2_convert_to_pdf.py

WHAT YOU GET:
    - One PDF per email, named: 0001_Subject_00_email.pdf
    - Attachments saved right after: 0001_Subject_att1_filename.pdf
    - Emails with attachments show "Attachments: filename" in the header
    - Emails without attachments show nothing extra
    - All text in black, clean and readable
"""

import os
import email
import re
import html as htmlmod
import shutil
import quopri
import platform
from email import policy

# ---------------------------------------------------------------
# CONFIGURATION — edit this section before running
# Must match the out_dir values from step1_extract_emails.py
# ---------------------------------------------------------------
JOBS = [
    {
        "label": "Person Name",
        "eml_dir": "./Person_Name_Emails/",   # Folder created by Step 1
        "out_dir": "./Person_Name_PDFs/",      # Where to save the PDFs
    },
    # Add more people by copying the block above:
    # {
    #     "label": "Another Person",
    #     "eml_dir": "./Another_Person_Emails/",
    #     "out_dir": "./Another_Person_PDFs/",
    # },
]
# ---------------------------------------------------------------


def check_reportlab():
    """Check if reportlab is installed and give clear instructions if not."""
    try:
        import reportlab
        return True
    except ImportError:
        print("\nERROR: The 'reportlab' library is not installed.")
        print("To install it, open Terminal (Mac) or Command Prompt (Windows) and run:")
        if platform.system() == 'Windows':
            print("    pip install reportlab")
        else:
            print("    /usr/local/bin/pip3 install reportlab")
        print("\nThen run this script again.")
        return False


def safe(t):
    """Escape HTML special characters for PDF rendering."""
    return str(t).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def decode_payload(part):
    """Decode email part, handling quoted-printable and different character sets."""
    cte = str(part.get('Content-Transfer-Encoding', '')).lower().strip()
    raw = part.get_payload(decode=True)
    if raw is None:
        return ''
    # Handle quoted-printable encoding (removes = artifacts)
    if cte == 'quoted-printable' or b'=\n' in raw or b'=3D' in raw:
        try:
            raw = quopri.decodestring(raw)
        except:
            pass
    charset = part.get_content_charset() or 'utf-8'
    try:
        return raw.decode(charset, 'ignore')
    except:
        return raw.decode('utf-8', 'ignore')


def clean_html(raw):
    """Strip HTML tags and return clean readable plain text."""
    # Remove style and script blocks entirely
    raw = re.sub(r'<(style|script)[^>]*>.*?</(style|script)>', '', raw, flags=re.DOTALL | re.IGNORECASE)
    # Replace line-break tags with newlines
    raw = re.sub(r'<(br|BR)\s*/?>', '\n', raw)
    raw = re.sub(r'</(p|div|tr|li|h[1-6]|blockquote)>', '\n', raw, flags=re.IGNORECASE)
    # Strip all remaining HTML tags
    raw = re.sub(r'<[^>]+>', '', raw)
    # Decode HTML entities (e.g. &amp; → &)
    raw = htmlmod.unescape(raw)
    # Clean up whitespace while preserving paragraph breaks
    lines = [l.strip() for l in raw.splitlines()]
    result = []
    prev_blank = False
    for l in lines:
        if not l:
            if not prev_blank:
                result.append('')
            prev_blank = True
        else:
            result.append(l)
            prev_blank = False
    return '\n'.join(result).strip()


def clean_plain(text):
    """Remove quoted-printable soft line breaks and encoding artifacts."""
    text = re.sub(r'=\r?\n', '', text)       # Remove soft line breaks
    text = re.sub(r'=[0-9A-Fa-f]{2}', '', text)  # Remove hex artifacts
    return text.strip()


def eml_to_pdf(eml_path, pdf_path):
    """Convert a single .eml file to a PDF."""
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.colors import black
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.units import inch

    with open(eml_path, 'rb') as f:
        msg = email.message_from_binary_file(f, policy=policy.default)

    # Extract email headers
    subject     = str(msg.get('Subject', 'No Subject'))
    sender      = str(msg.get('From', ''))
    to          = str(msg.get('To', ''))
    date        = str(msg.get('Date', ''))
    body        = ''
    attachments = []

    # Extract body and attachment names
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            cd = str(part.get('Content-Disposition', ''))
            fn = part.get_filename()
            if fn and 'attachment' in cd:
                attachments.append(fn)
            elif ct == 'text/plain' and 'attachment' not in cd and not body:
                try:
                    body = clean_plain(decode_payload(part))
                except:
                    pass
        # Fall back to HTML if no plain text found
        if not body:
            for part in msg.walk():
                ct = part.get_content_type()
                cd = str(part.get('Content-Disposition', ''))
                if ct == 'text/html' and 'attachment' not in cd:
                    try:
                        body = clean_html(decode_payload(part))
                        break
                    except:
                        pass
    else:
        try:
            raw = decode_payload(msg)
            if '<html' in raw.lower() or '<div' in raw.lower():
                body = clean_html(raw)
            else:
                body = clean_plain(raw)
        except:
            body = ''

    # Build the PDF
    header_style = ParagraphStyle('header', fontSize=10, leading=14, textColor=black)
    body_style   = ParagraphStyle('body',   fontSize=9,  leading=13, textColor=black)

    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        rightMargin=inch,
        leftMargin=inch,
        topMargin=inch,
        bottomMargin=inch
    )

    story = [
        Paragraph(f"<b>From:</b> {safe(sender)}",    header_style),
        Paragraph(f"<b>To:</b> {safe(to)}",          header_style),
        Paragraph(f"<b>Date:</b> {safe(date)}",       header_style),
        Paragraph(f"<b>Subject:</b> {safe(subject)}", header_style),
    ]

    # Only show attachments line if there are actual attachments
    if attachments:
        att_list = ', '.join(safe(a) for a in attachments)
        story.append(Paragraph(f"<b>Attachments:</b> {att_list}", header_style))

    story.append(Spacer(1, 0.2 * inch))

    # Add email body line by line
    for line in body.splitlines():
        line = line.strip()
        if line:
            story.append(Paragraph(safe(line), body_style))
        else:
            story.append(Spacer(1, 0.1 * inch))

    doc.build(story)


def run_job(label, eml_dir, out_dir):
    """Convert all EML files in a folder to PDFs."""
    eml_dir = os.path.normpath(eml_dir)
    out_dir = os.path.normpath(out_dir)

    if not os.path.exists(eml_dir):
        print(f"\n[{label}] ERROR: EML folder not found at:")
        print(f"         {eml_dir}")
        print(f"         Make sure you ran Step 1 first.")
        return

    # Wipe output folder completely to avoid duplicates from previous runs
    shutil.rmtree(out_dir, ignore_errors=True)
    os.makedirs(out_dir, exist_ok=True)

    eml_files = sorted([f for f in os.listdir(eml_dir) if f.endswith('.eml')])
    total = len(eml_files)

    if total == 0:
        print(f"\n[{label}] No .eml files found in {eml_dir}")
        print(f"         Make sure Step 1 completed successfully.")
        return

    print(f"\n[{label}] Converting {total} emails to PDF...")
    ok = bad = 0

    for fname in eml_files:
        try:
            base     = fname.replace('.eml', '')
            eml_path = os.path.join(eml_dir, fname)
            pdf_path = os.path.join(out_dir, base + '_00_email.pdf')

            # Convert email to PDF
            eml_to_pdf(eml_path, pdf_path)

            # Copy attachments immediately after their email, in sorted order
            for att in sorted(os.listdir(eml_dir)):
                if att.startswith(base + '_att') and not att.endswith('.eml'):
                    shutil.copy2(
                        os.path.join(eml_dir, att),
                        os.path.join(out_dir, att)
                    )

            ok += 1
            if ok % 25 == 0:
                print(f"  {ok}/{total} done...")

        except Exception as e:
            bad += 1
            print(f"  Warning: Could not convert {fname} — {e}")

    print(f"[{label}] Done! {ok} PDFs created, {bad} errors")
    print(f"[{label}] Output folder: {os.path.abspath(out_dir)}")


# ---------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------
if __name__ == '__main__':
    print("=" * 50)
    print("GAVEL — Step 2: Convert Emails to PDF")
    print(f"Running on: {platform.system()} {platform.release()}")
    print("=" * 50)

    if not check_reportlab():
        exit(1)

    for job in JOBS:
        run_job(job['label'], job['eml_dir'], job['out_dir'])

    print("\nAll jobs complete.")
    print("Check your output folders for the PDFs.")
