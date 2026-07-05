#!/usr/bin/env python3
"""
GAVEL — Step 1: Extract Emails from MBOX
=========================================
Converts a .mbox file into individual .eml files and saves all attachments.

SETUP (run once):
    No extra libraries needed — uses Python built-ins only.

USAGE:
    1. Edit the JOBS section below with your mbox path and output folder
    2. Run: python3 step1_extract_emails.py

HOW TO FIND YOUR MBOX PATH:
    Mac: Right-click the .mbox folder in Finder → "New Terminal at Folder" → type pwd
    Windows: Shift + right-click the folder → "Copy as path"
    Then add /mbox at the end of the path.
"""

import os
import mailbox
import platform

# ---------------------------------------------------------------
# CONFIGURATION — edit this section before running
# ---------------------------------------------------------------
JOBS = [
    {
        "label": "Person Name",                          # Name shown in the terminal
        "mbox_path": "/path/to/Person Name.mbox/mbox",  # Full path to your mbox file
        "out_dir": "./Person_Name_Emails/",              # Where to save the output
    },
    # Add more people by copying the block above and changing the values
    # {
    #     "label": "Another Person",
    #     "mbox_path": "/path/to/Another Person.mbox/mbox",
    #     "out_dir": "./Another_Person_Emails/",
    # },
]
# ---------------------------------------------------------------


def sanitize(name, max_len=60):
    """Remove characters that are illegal in filenames on Mac and Windows."""
    illegal = r'\/:*?"<>|'
    return ''.join('_' if c in illegal else c for c in str(name))[:max_len].strip()


def run_job(label, mbox_path, out_dir):
    """Extract all emails and attachments from one mbox file."""

    # Normalize path separators for the current OS
    mbox_path = os.path.normpath(mbox_path)
    out_dir = os.path.normpath(out_dir)

    if not os.path.exists(mbox_path):
        print(f"\n[{label}] ERROR: mbox file not found at:")
        print(f"         {mbox_path}")
        print(f"         Check that the path is correct and the drive is connected.")
        return

    os.makedirs(out_dir, exist_ok=True)

    mbox = mailbox.mbox(mbox_path)
    total = len(mbox)
    print(f"\n[{label}] Found {total} emails. Starting extraction...")

    ok = bad = 0

    for i, msg in enumerate(mbox):
        try:
            subj = str(msg.get('subject', 'NoSubject') or 'NoSubject')
            base = f"{i+1:04d}_{sanitize(subj)}"

            # Save email as .eml
            eml_path = os.path.join(out_dir, base + '.eml')
            with open(eml_path, 'wb') as f:
                f.write(msg.as_bytes())

            # Save attachments
            att_num = 0
            if msg.is_multipart():
                for part in msg.walk():
                    fn = part.get_filename()
                    cd = str(part.get('Content-Disposition', ''))
                    if fn and 'attachment' in cd:
                        att_num += 1
                        fn_safe = sanitize(fn, max_len=80)
                        att_path = os.path.join(out_dir, f"{base}_att{att_num}_{fn_safe}")
                        payload = part.get_payload(decode=True)
                        if payload:
                            with open(att_path, 'wb') as f:
                                f.write(payload)

            ok += 1
            if ok % 50 == 0:
                print(f"  {ok}/{total} done...")

        except Exception as e:
            bad += 1
            print(f"  Warning: Could not process email {i+1} — {e}")

    print(f"[{label}] Done! {ok} emails saved, {bad} errors")
    print(f"[{label}] Output folder: {os.path.abspath(out_dir)}")


# ---------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------
if __name__ == '__main__':
    print("=" * 50)
    print("GAVEL — Step 1: Extract Emails from MBOX")
    print(f"Running on: {platform.system()} {platform.release()}")
    print("=" * 50)

    for job in JOBS:
        run_job(job['label'], job['mbox_path'], job['out_dir'])

    print("\nAll jobs complete.")
    print("Next step: run step2_convert_to_pdf.py")
