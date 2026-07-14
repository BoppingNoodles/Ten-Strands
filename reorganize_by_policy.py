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

# Matches: <BP/AR>_<PolicyNumber>_<PolicyTitle>_...(rest ignored)
FILENAME_PATTERN = re.compile(
    r"^(?P<doctype>BP|AR)[ _]+(?P<number>[\w.]+)[ _]+(?P<title>.+?)_[^_]+_[^_]+_\d{4}_\d{4}\.pdf$",
    re.IGNORECASE,
)


def sanitize_folder_name(name: str) -> str:
    """Remove characters that are illegal in Windows folder names."""
    return re.sub(r'[<>:"/\\|?*]', "", name).strip()


def parse_filename(filename: str):
    """
    Try to extract (doctype, number, title) from a policy PDF filename.
    Returns None if the filename doesn't match the expected pattern.
    """
    match = FILENAME_PATTERN.match(filename)
    if not match:
        return None
    doctype = match.group("doctype").upper()
    number = match.group("number").strip()
    title = match.group("title").strip().replace("_", " ")
    return doctype, number, title


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

                parsed = parse_filename(filename)
                if parsed is None:
                    skipped_files.append(os.path.join(root, filename))
                    continue

                doctype, number, title = parsed
                folder_name = sanitize_folder_name(f"{doctype} {number} {title}")
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
        print(f"\n{len(skipped_files)} file(s) did not match the expected naming pattern and were skipped:")
        for f in skipped_files:
            print(f"  - {f}")
        print(
            "\nIf these are still valid policy PDFs, check that they follow the convention:\n"
            "  <BP/AR>_<Policy Number>_<Policy Title>_<County>_<District>_<Year Adopted>_<Year Revised>.pdf"
        )


if __name__ == "__main__":
    main()