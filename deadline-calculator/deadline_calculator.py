#!/usr/bin/env python3
"""
GAVEL — Deadline Calculator
============================
Calculates legal deadlines from a filing or trigger date.
Automatically skips weekends and US federal holidays.

SETUP (run once):
    Mac:     /usr/local/bin/pip3 install holidays
    Windows: pip install holidays

USAGE:
    Mac:     python3 deadline_calculator.py
    Windows: python deadline_calculator.py

    The script will ask you for the date and state when you run it.
    No editing required.
"""

import os
import platform
from datetime import date, timedelta

# ---------------------------------------------------------------
# OPTIONAL — customize your deadline list here
# Add, remove, or rename deadlines as needed
# ---------------------------------------------------------------
DEADLINES = [
    ("Response due",       21),
    ("Reply due",          28),
    ("30-day deadline",    30),
    ("45-day deadline",    45),
    ("60-day deadline",    60),
    ("90-day deadline",    90),
    ("120-day deadline",  120),
    ("6-month deadline",  180),
    ("1-year deadline",   365),
]
# ---------------------------------------------------------------


def check_libraries():
    """Check required libraries are installed."""
    try:
        import holidays
        return True
    except ImportError:
        print("\nERROR: The 'holidays' library is not installed.")
        print("Install it by running:")
        if platform.system() == 'Windows':
            print("    pip install holidays")
        else:
            print("    /usr/local/bin/pip3 install holidays")
        return False


def get_input():
    """Ask the user for the trigger date and state."""
    print("\n" + "-" * 60)

    # Get trigger date
    while True:
        date_input = input("Enter the trigger date (YYYY-MM-DD, e.g. 2026-09-15): ").strip()
        try:
            parts = date_input.split("-")
            trigger_date = date(int(parts[0]), int(parts[1]), int(parts[2]))
            break
        except (ValueError, IndexError):
            print("  Invalid date. Please use the format YYYY-MM-DD (e.g. 2026-09-15)")

    # Get trigger label
    label_input = input("What does this date represent? (press Enter for 'Filing Date'): ").strip()
    trigger_label = label_input if label_input else "Filing Date"

    # Get state
    state_input = input("Enter your state abbreviation for state holidays (e.g. TX, CA, NY — or press Enter to skip): ").strip().upper()
    state = state_input if len(state_input) == 2 else ""

    # Save to CSV?
    save_input = input("Save results to CSV? (y/n, press Enter for yes): ").strip().lower()
    save_csv = save_input != "n"

    return trigger_date, trigger_label, state, save_csv


def get_holidays(year, state):
    """Get US federal + state holidays for a given year."""
    import holidays as hol
    return hol.US(state=state if state else None, years=year)


def is_business_day(d, all_holidays):
    """Return True if the date is a business day."""
    if d.weekday() >= 5:
        return False
    if d in all_holidays:
        return False
    return True


def add_business_days(start_date, days, state):
    """Add calendar days then push to next business day if needed."""
    result = start_date + timedelta(days=days)
    all_holidays = {}
    for year in range(start_date.year, result.year + 2):
        all_holidays.update(get_holidays(year, state))
    while not is_business_day(result, all_holidays):
        result += timedelta(days=1)
    return result


def format_date(d):
    return d.strftime("%B %d, %Y")


def day_of_week(d):
    return d.strftime("%A")


def save_to_csv(trigger_date, trigger_label, results, state):
    """Save results to a CSV file in the same folder as the script."""
    import csv

    script_dir = os.path.dirname(os.path.abspath(__file__))
    filename = f"deadlines_{trigger_date.strftime('%Y-%m-%d')}.csv"
    filepath = os.path.join(script_dir, filename)

    with open(filepath, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(["Deadline", "Calendar Days", "Due Date", "Day of Week"])
        writer.writerow([trigger_label, 0, format_date(trigger_date), day_of_week(trigger_date)])
        writer.writerow([])
        for r in results:
            writer.writerow([r['label'], r['days'], r['deadline'], r['day_of_week']])

    print(f"\nCSV saved to: {filepath}")


def run():
    """Main function."""
    trigger_date, trigger_label, state, save_csv = get_input()

    print(f"\n{trigger_label}: {format_date(trigger_date)} ({day_of_week(trigger_date)})")
    print(f"State holidays: {'Federal only' if not state else state}")
    print("-" * 65)
    print(f"{'Deadline':<30} {'Days':>5}  {'Due Date':<22} {'Day'}")
    print("-" * 65)

    results = []
    for label, days in DEADLINES:
        deadline = add_business_days(trigger_date, days, state)
        day_name = day_of_week(deadline)
        print(f"{label:<30} {days:>5}  {format_date(deadline):<22} {day_name}")
        results.append({
            "label": label,
            "days": days,
            "deadline": format_date(deadline),
            "day_of_week": day_name
        })

    print("-" * 65)
    print(f"\nNote: All deadlines adjusted to next business day if they")
    print(f"fall on a weekend or {'federal/' + state if state else 'US federal'} holiday.")

    if save_csv:
        save_to_csv(trigger_date, trigger_label, results, state)

    # Ask if they want to run again
    print()
    again = input("Calculate deadlines for another date? (y/n): ").strip().lower()
    if again == "y":
        run()


# ---------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------
if __name__ == '__main__':
    print("=" * 65)
    print("GAVEL — Deadline Calculator")
    print(f"Running on: {platform.system()} {platform.release()}")
    print("=" * 65)

    if not check_libraries():
        exit(1)

    run()
