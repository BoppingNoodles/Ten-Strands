"""
Reorganize Ten Strands policy PDFs by policy instead of by district.

Source structure (district-based):
    Ten Strands/<District Name>/BP_3510_Green Schools Operations_Alameda_Alameda Unified_2019_2019.pdf

Target structure (policy-based), created under:
    Ten Strands/By Policy/<BP/AR> <Policy Number> <Policy Title>/<original filename>.pdf

Naming convention expected for each PDF:
    <BP/AR>_<Policy Number>_<Policy Title>_<County>_<District>_<Year Adopted>_<Year Revised>.pdf

Example:
    BP_3510_Green Schools Operations_Alameda_Alameda Unified_2019_2019.pdf
    -> folder: "BP 3510 Green Schools Operations"

Files are COPIED (not moved) — originals stay untouched in the district folders.
"""

import os
import re
import shutil

# ---- CONFIG ----
BASE_DIR = r"C:\Users\caden\Documents\Ten Strands\Ten-Strands\Ten Strands"
DEST_DIR = os.path.join(BASE_DIR, "By Policy")
# folders to skip when scanning for district folders (e.g. the destination itself)
SKIP_FOLDERS = {"By Policy"}
# -----------------

# Matches: <BP/AR>_<PolicyNumber>_<PolicyTitle>_<County>_<District>_<YearAdopted>_<YearRevised>.pdf
# Fields are underscore-delimited; title/county/district may contain spaces and periods but not
# underscores. Year fields may be a 4-digit year OR the literal "NA".
# The canonical title/folder comes from CANONICAL_POLICIES below, keyed strictly on
# (doctype, number), so title-wording differences across districts don't matter.
FILENAME_PATTERN = re.compile(
    r"^(?P<doctype>BP|AR)_(?P<number>\d+(?:\.\d+)?)_(?P<title>[^_]+)_(?P<county>[^_]+)_"
    r"(?P<district>[^_]+)_(?P<year_adopted>\d{4}|NA)_(?P<year_revised>\d{4}|NA)\.pdf$",
    re.IGNORECASE,
)

# Canonical (doctype, policy number) -> folder name, per the reference list provided.
CANONICAL_POLICIES = {
    ("AR", "3511.1"): "AR 3511.1 Integrated Waste Management",
    ("AR", "3514"): "AR 3514 Environmental Safety",
    ("AR", "3514.1"): "AR 3514.1 Hazardous Substances",
    ("AR", "3514.2"): "AR 3514.2 Integrated Pest Management",
    ("AR", "5142.2"): "AR 5142.2 Safe Routes to School",
    ("AR", "7110"): "AR 7110 Facilities Master Plan",
    ("BP", "3510"): "BP 3510 Green Schools Operations",
    ("BP", "3511"): "BP 3511 Energy and Water Management",
    ("BP", "3511.1"): "BP 3511.1 Integrated Waste Management",
    ("BP", "3514"): "BP 3514 Environmental Safety",
    ("BP", "3514.1"): "BP 3514.1 Hazardous Substances",
    ("BP", "5142.2"): "BP 5142.2 Safe Routes to School",
    ("BP", "6142.5"): "BP 6142.5 Environmental Education",
    ("BP", "7110"): "BP 7110 Facilities Master Plan",
}


def sanitize_folder_name(name: str) -> str:
    """Remove characters that are illegal in Windows folder names."""
    return re.sub(r'[<>:"/\\|?*]', "", name).strip()


def parse_filename(filename: str):
    """
    Try to extract (doctype, number) from a policy PDF filename, then look up
    the canonical folder name from CANONICAL_POLICIES.
    Returns the canonical folder name string, or None if:
      - the filename doesn't match the expected naming pattern, or
      - the (doctype, number) isn't in the canonical policy list.
    """
    match = FILENAME_PATTERN.match(filename)
    if not match:
        return None
    doctype = match.group("doctype").upper()
    number = match.group("number").strip()
    return CANONICAL_POLICIES.get((doctype, number))


def main():
    if not os.path.isdir(BASE_DIR):
        print(f"ERROR: Base directory not found:\n  {BASE_DIR}")
        return

    os.makedirs(DEST_DIR, exist_ok=True)

    copied_count = 0
    total_pdf_count = 0
    skipped_files = []

    # Walk each district folder directly under BASE_DIR
    for entry in os.scandir(BASE_DIR):
        if not entry.is_dir():
            continue
        if entry.name in SKIP_FOLDERS:
            continue

        district_folder = entry.path

        for root, _dirs, files in os.walk(district_folder):
            for filename in files:
                if not filename.lower().endswith(".pdf"):
                    continue

                total_pdf_count += 1

                folder_name = parse_filename(filename)
                if folder_name is None:
                    skipped_files.append(os.path.join(root, filename))
                    continue

                folder_name = sanitize_folder_name(folder_name)
                policy_folder = os.path.join(DEST_DIR, folder_name)
                os.makedirs(policy_folder, exist_ok=True)

                src_path = os.path.join(root, filename)
                dst_path = os.path.join(policy_folder, filename)

                # Avoid overwriting if an identically named file already exists
                if os.path.exists(dst_path):
                    base, ext = os.path.splitext(filename)
                    dup_index = 1
                    while os.path.exists(dst_path):
                        dst_path = os.path.join(policy_folder, f"{base} ({dup_index}){ext}")
                        dup_index += 1

                shutil.copy2(src_path, dst_path)
                copied_count += 1

    print(f"Done. Copied {copied_count}/{total_pdf_count} PDF(s) into '{DEST_DIR}'.")

    if skipped_files:
        print(f"\n{len(skipped_files)} file(s) were skipped:")
        for f in skipped_files:
            print(f"  - {f}")
        print(
            "\nA file is skipped if either:\n"
            "  1. Its filename doesn't match the naming convention:\n"
            "     <BP/AR>_<Policy Number>_<Policy Title>_<County>_<District>_<Year Adopted>_<Year Revised>.pdf\n"
            "  2. Its <BP/AR> + <Policy Number> isn't in CANONICAL_POLICIES at the top of this script.\n"
            "     Add it there if it's a valid policy that should have its own folder."
        )


if __name__ == "__main__":
    main()