#!/usr/bin/env python3
"""
GAVEL — PDF Merger
===================
Merges all PDFs in a folder into one single PDF.
Adds a table of contents at the front with page numbers.
Sorts files alphabetically or by Bates number.

SETUP (run once):
    Mac:     /usr/local/bin/pip3 install pypdf reportlab
    Windows: pip install pypdf reportlab

USAGE:
    Mac:     python3 pdf_merger.py
    Windows: python pdf_merger.py

    The script will ask you for the folder and options when you run it.
    No editing required.
"""

import os
import platform
from datetime import datetime

# ---------------------------------------------------------------
# OPTIONAL — customize defaults here
# These are just the starting values — you can change them each
# time you run the script when prompted
# ---------------------------------------------------------------

# Default sort order: "alphabetical" or "bates"
DEFAULT_SORT = "alphabetical"

# Default output filename
DEFAULT_OUTPUT_NAME = "merged_output.pdf"

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


def get_input():
    """Ask the user for the folder and options."""
    print("\n" + "-" * 60)

    # Get input folder
    while True:
        folder = input("Enter the full path to the folder containing your PDFs: ").strip()
        # Remove surrounding quotes if pasted from Finder
        folder = folder.strip('"').strip("'")
        folder = os.path.normpath(os.path.expanduser(folder))
        if os.path.exists(folder):
            break
        print(f"  Folder not found: {folder}")
        print("  Tip: In Finder, right-click the folder and hold Option to copy the path.")

    # Get sort order
    sort_input = input("Sort files alphabetically or by Bates number? (a=alphabetical, b=bates, press Enter for alphabetical): ").strip().lower()
    sort_order = "bates" if sort_input == "b" else "alphabetical"

    # Get output filename
    name_input = input(f"Output filename (press Enter for '{DEFAULT_OUTPUT_NAME}'): ").strip()
    output_name = name_input if name_input else DEFAULT_OUTPUT_NAME
    if not output_name.lower().endswith('.pdf'):
        output_name += '.pdf'

    # Add table of contents?
    toc_input = input("Add a table of contents at the front? (y/n, press Enter for yes): ").strip().lower()
    add_toc = toc_input != "n"

    return folder, sort_order, output_name, add_toc


def sort_files(pdf_files, sort_order):
    """Sort PDF files alphabetically or by Bates number."""
    if sort_order == "bates":
        # Extract trailing numbers from filename for Bates sorting
        import re
        def bates_key(f):
            numbers = re.findall(r'\d+', f)
            return int(numbers[-1]) if numbers else 0
        return sorted(pdf_files, key=bates_key)
    else:
        return sorted(pdf_files)


def create_toc_pdf(toc_entries, output_path):
    """Create a table of contents PDF page."""
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.colors import black, HexColor
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.units import inch
    from reportlab.lib import colors

    doc = SimpleDocTemplate(
        output_path,
        pagesize=letter,
        rightMargin=inch,
        leftMargin=inch,
        topMargin=inch,
        bottomMargin=inch
    )

    title_style = ParagraphStyle('title', fontSize=16, fontName='Helvetica-Bold',
                                  textColor=black, spaceAfter=20, leading=20)
    meta_style  = ParagraphStyle('meta',  fontSize=9,  fontName='Helvetica',
                                  textColor=HexColor('#666666'), spaceAfter=20)
    header_style = ParagraphStyle('header', fontSize=10, fontName='Helvetica-Bold',
                                   textColor=black)
    item_style   = ParagraphStyle('item',   fontSize=9,  fontName='Helvetica',
                                   textColor=black, leading=14)

    story = []
    story.append(Paragraph("Table of Contents", title_style))
    story.append(Paragraph(
        f"Generated: {datetime.now().strftime('%B %d, %Y at %I:%M %p')} | "
        f"{len(toc_entries)} documents",
        meta_style
    ))

    # Build table data
    table_data = [["#", "Document", "Pages", "Starting Page"]]
    for i, entry in enumerate(toc_entries, 1):
        table_data.append([
            str(i),
            entry['name'],
            str(entry['pages']),
            str(entry['start_page'])
        ])

    table = Table(table_data, colWidths=[0.4*inch, 4.2*inch, 0.8*inch, 1.0*inch])
    table.setStyle(TableStyle([
        # Header row
        ('BACKGROUND',   (0, 0), (-1, 0),  HexColor('#1a1a2e')),
        ('TEXTCOLOR',    (0, 0), (-1, 0),  colors.white),
        ('FONTNAME',     (0, 0), (-1, 0),  'Helvetica-Bold'),
        ('FONTSIZE',     (0, 0), (-1, 0),  9),
        ('BOTTOMPADDING',(0, 0), (-1, 0),  8),
        ('TOPPADDING',   (0, 0), (-1, 0),  8),
        # Data rows
        ('FONTNAME',     (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE',     (0, 1), (-1, -1), 8),
        ('TOPPADDING',   (0, 1), (-1, -1), 5),
        ('BOTTOMPADDING',(0, 1), (-1, -1), 5),
        ('ROWBACKGROUNDS',(0, 1),(-1, -1), [colors.white, HexColor('#f5f5f5')]),
        # Grid
        ('GRID',         (0, 0), (-1, -1), 0.5, HexColor('#dddddd')),
        ('ALIGN',        (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN',        (2, 0), (3, -1),  'CENTER'),
        ('VALIGN',       (0, 0), (-1, -1), 'MIDDLE'),
    ]))

    story.append(table)
    doc.build(story)


def merge_pdfs(folder, sort_order, output_name, add_toc):
    """Merge all PDFs in the folder into one file."""
    from pypdf import PdfReader, PdfWriter
    import io

    # Get all PDFs
    all_files = [f for f in os.listdir(folder) if f.lower().endswith('.pdf')]
    pdf_files = sort_files(all_files, sort_order)

    if not pdf_files:
        print(f"\nNo PDF files found in: {folder}")
        return

    print(f"\nFound {len(pdf_files)} PDF files. Merging...\n")

    # Build TOC entries and collect page counts
    toc_entries = []
    current_page = 2 if add_toc else 1  # Page 1 is TOC if enabled

    for fname in pdf_files:
        fpath = os.path.join(folder, fname)
        try:
            reader = PdfReader(fpath)
            page_count = len(reader.pages)
            clean_name = os.path.splitext(fname)[0]  # Remove .pdf extension
            toc_entries.append({
                'name': clean_name,
                'pages': page_count,
                'start_page': current_page,
                'path': fpath
            })
            current_page += page_count
            print(f"  ✓ {fname} ({page_count} pages)")
        except Exception as e:
            print(f"  ✗ {fname} — Error reading: {e}")

    # Create final merged PDF
    writer = PdfWriter()

    # Add TOC first if requested
    if add_toc:
        toc_path = os.path.join(folder, "_toc_temp.pdf")
        create_toc_pdf(toc_entries, toc_path)
        toc_reader = PdfReader(toc_path)
        for page in toc_reader.pages:
            writer.add_page(page)

    # Add all PDFs
    for entry in toc_entries:
        try:
            reader = PdfReader(entry['path'])
            for page in reader.pages:
                writer.add_page(page)
        except Exception as e:
            print(f"  ✗ Could not merge {entry['name']}: {e}")

    # Save output
    output_path = os.path.join(folder, output_name)
    with open(output_path, 'wb') as f:
        writer.write(f)

    # Clean up temp TOC file
    if add_toc and os.path.exists(toc_path):
        os.remove(toc_path)

    total_pages = sum(e['pages'] for e in toc_entries)
    print(f"\nDone! Merged {len(toc_entries)} files ({total_pages} total pages)")
    print(f"Output: {output_path}")


def run():
    """Main function."""
    folder, sort_order, output_name, add_toc = get_input()
    merge_pdfs(folder, sort_order, output_name, add_toc)

    print()
    again = input("Merge another folder? (y/n): ").strip().lower()
    if again == "y":
        run()


# ---------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------
if __name__ == '__main__':
    print("=" * 60)
    print("GAVEL — PDF Merger")
    print(f"Running on: {platform.system()} {platform.release()}")
    print("=" * 60)

    if not check_libraries():
        exit(1)

    run()
