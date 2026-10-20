# -*- coding: utf-8 -*-
"""
BAU — Introduction to Data Analytics (SDF 1)
Even split of 35 students into 5 variants (A–E) — exactly 7 students per variant.

Rule (SAME as index.html): variant = VARIANTS[(seat_no - 1) % 5]
  1, 6, 11, 16, 21, 26, 31 -> A    2, 7, 12, 17, 22, 27, 32 -> B    3, 8, 13, 18, 23, 28, 33 -> C
  4, 9, 14, 19, 24, 29, 34 -> D    5, 10, 15, 20, 25, 30, 35 -> E

Students type the seat number from this list into the portal.

Usage:
  python telebeler_35.py                  # template list (Student 1 ... Student 35)
  python telebeler_35.py list.txt         # one full name per line (35 lines)
  python telebeler_35.py list.txt --csv split.csv   # save the split for Excel
"""

import argparse
import csv
import sys

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

VARIANTS = ["A", "B", "C", "D", "E"]
GROUP_SIZE = 35


def variant_for(sira_no):
    """Deterministic variant for a seat number (starting from 1)."""
    n = int(sira_no)
    if n < 1:
        raise ValueError("Seat number must be 1 or greater")
    return VARIANTS[(n - 1) % len(VARIANTS)]


def build_roster(names=None):
    names = list(names or [f"Student {i}" for i in range(1, GROUP_SIZE + 1)])
    if len(names) != GROUP_SIZE:
        raise SystemExit(f"The list has {len(names)} names, {GROUP_SIZE} required.")
    return [{"seat": i, "name": name, "variant": variant_for(i)} for i, name in enumerate(names, start=1)]


def check_balance(roster):
    counts = {v: 0 for v in VARIANTS}
    for t in roster:
        counts[t["variant"]] += 1
    expected = len(roster) // len(VARIANTS)
    assert all(c == expected for c in counts.values()), f"Split is not even: {counts}"
    return counts


def print_distribution(roster):
    print("=" * 52)
    print(f"{'Seat':<6}{'Name':<32}{'Variant':<10}")
    print("-" * 52)
    for t in roster:
        print(f"{t['seat']:<6}{t['name']:<32}Variant {t['variant']}")
    print("=" * 52)

    counts = check_balance(roster)
    print("\nStudents per variant:")
    for v in VARIANTS:
        sira = ", ".join(str(t["seat"]) for t in roster if t["variant"] == v)
        print(f"  * Variant {v}: {counts[v]} students  (seats: {sira})")
    print(f"\n✅ Even split: {len(roster)} students / {len(VARIANTS)} variants = {counts['A']} students per variant")


def main():
    ap = argparse.ArgumentParser(description="Even split of 35 students into 5 variants")
    ap.add_argument("siyahi", nargs="?", help="Text file with one full name per line (UTF-8)")
    ap.add_argument("--csv", help="Write the split to a CSV file")
    args = ap.parse_args()

    names = None
    if args.siyahi:
        with open(args.siyahi, encoding="utf-8-sig") as f:
            names = [line.strip() for line in f if line.strip()]

    roster = build_roster(names)
    print_distribution(roster)

    if args.csv:
        with open(args.csv, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=["seat", "name", "variant"], delimiter=";")
            w.writeheader()
            w.writerows(roster)
        print(f"📄 CSV saved: {args.csv}")


if __name__ == "__main__":
    main()
