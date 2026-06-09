"""
create_district_folders.py
"""

import os
import sys

DISTRICTS = [
    "Alameda Unified",
    "Albany City Unified",
    "Berkeley Unified",
    "Castro Valley Unified",
    "Emery Unified",
    "Fremont Unified",
    "Hayward Unified",
    "Livermore Valley Joint Unified",
    "Mountain House Elementary",
    "Newark Unified",
    "New Haven Unified",
    "Oakland Unified",
    "Piedmont City Unified",
    "San Leandro Unified",
    "San Lorenzo Unified",
    "Dublin Unified",
    "Pleasanton Unified",
    "Sunol Glen Unified",
]

TARGET_DIR = r"C:\Users\caden\Documents\Ten Strands\PDFs\Alameda"

if not os.path.isdir(TARGET_DIR):
    print(f"\nError: '{TARGET_DIR}' is not a valid directory.")
    sys.exit(1)

created, skipped = [], []

for district in DISTRICTS:
    folder_path = os.path.join(TARGET_DIR, district)
    if os.path.exists(folder_path):
        skipped.append(district)
    else:
        os.makedirs(folder_path)
        created.append(district)

print(f"\nTarget: {TARGET_DIR}\n")

if created:
    print(f"✓ Created ({len(created)}):")
    for name in created:
        print(f"    {name}")

if skipped:
    print(f"\n— Skipped / already exists ({len(skipped)}):")
    for name in skipped:
        print(f"    {name}")

print(f"\nDone. {len(created)} created, {len(skipped)} skipped.\n")