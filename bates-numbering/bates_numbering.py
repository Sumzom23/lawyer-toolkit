#!/usr/bin/env python3
"""
GAVEL — Bates Numbering
========================
Stamps sequential Bates numbers on every page of every PDF in a folder.
Output format: PREFIX00001 (e.g. SMITH00001, JONES00042)

SETUP (run once):
    Mac:     /usr/local/bin/pip3 install pypdf reportlab
    Windows: pip install pypdf reportlab

USAGE:
    1. Edit the CONFIGURATION section below
    2. Mac:     python3 bates_numbering.py
       Windows: python bates_numbering.py

OUTPUT:
    A new folder with all PDFs stamped — original files are never modified.
"""

import os
import platform

# ---------------------------------------------------------------
# CONFIGURATION — edit this section before running
# ---------------------------------------------------------------

# Folder containing the PDFs you want to Bates stamp
INPUT_FOLDER = "/Users/sumithmurthy/Desktop/pdfs_to_stamp/"

# Folder where stamped PDFs will be saved (created automatically)
OUTPUT_FOLDER = "/Users/sumithmurthy/Desktop/pdfs_stamped/"

# Prefix for the Bates number (e.g. "SMITH" produces SMITH00001)
PREFIX = "TEST"

# Starting number (change this if continuing from a previous batch)
START_NUMBER = 1

# How many digits to use (5 = 00001, 6 = 000001)
DIGITS = 5

# Position of the stamp on the page
# Options: "bottom-right", "bottom-left", "bottom-center", "top-right", "top-left"
POSITION = "bottom-right"

# Font size of the Bates number
FONT_SIZE = 10

# Margin from the edge of the page in points (72 points = 1 inch)
MARGIN = 30

# ---------------------------------------------------------------


def check_libraries():
    """Check required libraries are installed."""
    missing = []
    try:
        import pypdf
    except ImportError:
        missing.append('pypdf')
    try:
        import reportlab
    except ImportError:
        missing.append('reportlab')

    if missing:
        print("\nERROR: Missing required libraries: " + ", ".join(missing))
        print("Install them by running:")
        if platform.system() == 'Windows':
            print(f"    pip install {' '.join(missing)}")
        else:
            print(f"    /usr/local/bin/pip3 install {' '.join(missing)}")
        return False
    return True


def get_stamp_position(page_width, page_height, text_width, text_height):
    """Calculate x, y coordinates for the stamp based on POSITION setting."""
    if POSITION == "bottom-right":
        x = page_width - text_width - MARGIN
        y = MARGIN
    elif POSITION == "bottom-left":
        x = MARGIN
        y = MARGIN
    elif POSITION == "bottom-center":
        x = (page_width - text_width) / 2
        y = MARGIN
    elif POSITION == "top-right":
        x = page_width - text_width - MARGIN
        y = page_height - text_height - MARGIN
    elif POSITION == "top-left":
        x = MARGIN
        y = page_height - text_height - MARGIN
    else:
        x = page_width - text_width - MARGIN
        y = MARGIN
    return x, y


def create_stamp_pdf(bates_number, page_width, page_height):
    """Create a single-page PDF containing just the Bates stamp."""
    import io
    from reportlab.pdfgen import canvas
    from reportlab.lib.colors import black

    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_width, page_height))
    c.setFont("Helvetica-Bold", FONT_SIZE)
    c.setFillColor(black)

    text_width = c.stringWidth(bates_number, "Helvetica-Bold", FONT_SIZE)
    text_height = FONT_SIZE

    x, y = get_stamp_position(page_width, page_height, text_width, text_height)
    c.drawString(x, y, bates_number)
    c.save()

    packet.seek(0)
    return packet


def stamp_pdf(input_path, output_path, start_page_number):
    """Stamp all pages of a PDF with sequential Bates numbers."""
    from pypdf import PdfReader, PdfWriter
    import io

    reader = PdfReader(input_path)
    writer = PdfWriter()
    current_number = start_page_number

    for page in reader.pages:
        # Get page dimensions
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)

        # Format Bates number: PREFIX + zero-padded number
        bates_number = f"{PREFIX}{str(current_number).zfill(DIGITS)}"

        # Create stamp overlay
        stamp_packet = create_stamp_pdf(bates_number, page_width, page_height)

        from pypdf import PdfReader as PR
        stamp_reader = PR(stamp_packet)
        stamp_page = stamp_reader.pages[0]

        # Merge stamp onto original page
        page.merge_page(stamp_page)
        writer.add_page(page)
        current_number += 1

    # Save stamped PDF
    with open(output_path, 'wb') as f:
        writer.write(f)

    return current_number  # Return next available number


def run():
    """Main function — stamp all PDFs in the input folder."""
    input_folder  = os.path.normpath(INPUT_FOLDER)
    output_folder = os.path.normpath(OUTPUT_FOLDER)

    if not os.path.exists(input_folder):
        print(f"\nERROR: Input folder not found: {input_folder}")
        print("Check that INPUT_FOLDER is set correctly at the top of the script.")
        return

    os.makedirs(output_folder, exist_ok=True)

    # Get all PDFs sorted alphabetically
    pdf_files = sorted([f for f in os.listdir(input_folder) if f.lower().endswith('.pdf')])

    if not pdf_files:
        print(f"\nNo PDF files found in: {input_folder}")
        return

    print(f"\nFound {len(pdf_files)} PDF files.")
    print(f"Prefix: {PREFIX} | Starting number: {START_NUMBER} | Position: {POSITION}")
    print(f"Starting Bates stamp...\n")

    current_number = START_NUMBER
    ok = bad = 0

    for fname in pdf_files:
        input_path  = os.path.join(input_folder, fname)
        output_path = os.path.join(output_folder, fname)

        try:
            next_number = stamp_pdf(input_path, output_path, current_number)
            pages_stamped = next_number - current_number
            print(f"  ✓ {fname} — {pages_stamped} pages ({PREFIX}{str(current_number).zfill(DIGITS)} to {PREFIX}{str(next_number-1).zfill(DIGITS)})")
            current_number = next_number
            ok += 1
        except Exception as e:
            print(f"  ✗ {fname} — Error: {e}")
            bad += 1

    print(f"\nDone! {ok} files stamped, {bad} errors.")
    print(f"Last Bates number used: {PREFIX}{str(current_number - 1).zfill(DIGITS)}")
    print(f"Output folder: {os.path.abspath(output_folder)}")


# ---------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------
if __name__ == '__main__':
    print("=" * 50)
    print("GAVEL — Bates Numbering")
    print(f"Running on: {platform.system()} {platform.release()}")
    print("=" * 50)

    if not check_libraries():
        exit(1)

    run()
