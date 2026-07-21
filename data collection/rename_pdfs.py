"""
PDF Renamer — renames PDFs in a directory to the convention:
  <BP/AR>_<Policy Number>_<Policy Title>_<County>_<District>_<Year Adopted>_<Year Revised>

Example:
  BP_3510_Green Schools Operations_Alameda_Alameda Unified_2019_2019

Usage:
  python rename_pdfs.py                        # interactive mode
  python rename_pdfs.py --dir ./policies       # process a specific directory
  python rename_pdfs.py --dir ./policies --dry-run  # preview renames without applying
"""

import os
import re
import argparse


# ── Helpers ──────────────────────────────────────────────────────────────────

def strip_county_suffix(county: str) -> str:
    """Remove trailing ' County' (case-insensitive) from a county string."""
    return re.sub(r"\s*county\s*$", "", county.strip(), flags=re.IGNORECASE)


def sanitize(value: str) -> str:
    """Remove characters that are illegal in most file-system names."""
    # Strip leading/trailing whitespace, then replace forbidden chars with a space
    value = value.strip()
    value = re.sub(r'[<>:"/\\|?*]', " ", value)
    # Collapse multiple spaces
    value = re.sub(r" {2,}", " ", value)
    return value


def build_filename(doc_type: str, policy_number: str, policy_title: str,
                   county: str, district: str,
                   year_adopted: str, year_revised: str) -> str:
    """Assemble and return the new filename (without extension)."""
    county = strip_county_suffix(county)
    parts = [
        sanitize(doc_type).upper(),
        sanitize(policy_number),
        sanitize(policy_title),
        sanitize(county),
        sanitize(district),
        sanitize(year_adopted),
        sanitize(year_revised),
    ]
    return "_".join(parts)


# ── Interactive prompt ────────────────────────────────────────────────────────

def prompt(label: str, options: list[str] | None = None,
           default: str | None = None) -> str:
    """Prompt the user for input, with optional validation and default."""
    hint = ""
    if options:
        hint = f" [{'/'.join(options)}]"
    if default:
        hint += f" (default: {default})"
    while True:
        value = input(f"  {label}{hint}: ").strip()
        if not value and default:
            return default
        if options and value.upper() not in [o.upper() for o in options]:
            print(f"    ✗ Please enter one of: {', '.join(options)}")
            continue
        if value:
            return value
        print("    ✗ This field cannot be empty.")


def collect_metadata_interactive(filename: str) -> dict | None:
    """
    Ask the user to supply metadata for a single PDF file.
    Returns a dict of field values, or None to skip the file.
    """
    print(f"\n{'─' * 60}")
    print(f"  File : {filename}")
    print(f"{'─' * 60}")
    skip = input("  Skip this file? [y/N]: ").strip().lower()
    if skip == "y":
        return None

    doc_type    = prompt("Document type", options=["BP", "AR"])
    pol_number  = prompt("Policy number (e.g. 3510)")
    pol_title   = prompt("Policy title  (e.g. Green Schools Operations)")
    county      = prompt('County        (e.g. "Alameda" or "Alameda County")')
    district    = prompt("District      (e.g. Alameda Unified)")
    year_adopt  = prompt("Year adopted  (e.g. 2019)")
    year_rev    = prompt("Year revised  (e.g. 2019)", default=year_adopt)

    return {
        "doc_type":     doc_type,
        "policy_number": pol_number,
        "policy_title":  pol_title,
        "county":        county,
        "district":      district,
        "year_adopted":  year_adopt,
        "year_revised":  year_rev,
    }


# ── Rename logic ──────────────────────────────────────────────────────────────

def rename_pdfs(directory: str, dry_run: bool = False) -> None:
    """
    Walk `directory`, prompt for metadata for each PDF, and rename it.
    If dry_run=True, print the proposed renames without touching the disk.
    """
    pdf_files = sorted(
        f for f in os.listdir(directory)
        if f.lower().endswith(".pdf") and os.path.isfile(os.path.join(directory, f))
    )

    if not pdf_files:
        print(f"\nNo PDF files found in: {directory}")
        return

    print(f"\nFound {len(pdf_files)} PDF(s) in: {os.path.abspath(directory)}")
    if dry_run:
        print("  *** DRY-RUN MODE — no files will be changed ***")

    renamed, skipped, errors = 0, 0, 0

    for original_name in pdf_files:
        metadata = collect_metadata_interactive(original_name)

        if metadata is None:
            print("  → Skipped.")
            skipped += 1
            continue

        new_stem = build_filename(**metadata)
        new_name = new_stem + ".pdf"

        src = os.path.join(directory, original_name)
        dst = os.path.join(directory, new_name)

        # Handle name collisions
        if os.path.exists(dst) and dst != src:
            base, ext = os.path.splitext(dst)
            counter = 1
            while os.path.exists(dst):
                dst = f"{base}_{counter}{ext}"
                counter += 1
            new_name = os.path.basename(dst)

        print(f"\n  Old : {original_name}")
        print(f"  New : {new_name}")

        if dry_run:
            print("  → (dry-run) Would rename.")
        else:
            try:
                os.rename(src, dst)
                print("  ✓ Renamed.")
                renamed += 1
            except OSError as exc:
                print(f"  ✗ Error renaming: {exc}")
                errors += 1

    # ── Summary ──────────────────────────────────────────────────────────────
    print(f"\n{'═' * 60}")
    print("  Summary")
    print(f"{'═' * 60}")
    if dry_run:
        print(f"  Would rename : {len(pdf_files) - skipped}")
    else:
        print(f"  Renamed  : {renamed}")
        if errors:
            print(f"  Errors   : {errors}")
    print(f"  Skipped  : {skipped}")
    print()


# ── Entry point ───────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Rename PDFs to <BP/AR>_<PolicyNum>_<Title>_<County>_<District>_<YrAdopt>_<YrRev>"
    )
    parser.add_argument(
        "--dir", "-d",
        default=None,
        help="Directory containing PDFs (prompted if omitted)",
    )
    parser.add_argument(
        "--dry-run", "-n",
        action="store_true",
        help="Preview renames without making any changes",
    )
    args = parser.parse_args()

    directory = args.dir
    if not directory:
        directory = input("\nEnter the path to the PDF directory: ").strip()

    directory = os.path.expanduser(directory)

    if not os.path.isdir(directory):
        print(f"\nError: '{directory}' is not a valid directory.")
        raise SystemExit(1)

    rename_pdfs(directory, dry_run=args.dry_run)


if __name__ == "__main__":
    main()