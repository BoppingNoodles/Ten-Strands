"""
create_district_folders.py

Creates one subfolder per district in the same directory as this script.
Already-existing folders are skipped automatically.
"""

import os

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

# Always resolve relative to this script's location (Ten-Strands folder)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

created, skipped = [], []

for district in DISTRICTS:
    folder_path = os.path.join(BASE_DIR, district)
    if os.path.exists(folder_path):
        skipped.append(district)
    else:
        os.makedirs(folder_path)
        created.append(district)

print(f"\nBase directory: {BASE_DIR}\n")

if created:
    print(f"✓ Created ({len(created)}):")
    for name in created:
        print(f"    {name}")

if skipped:
    print(f"\n— Skipped / already exists ({len(skipped)}):")
    for name in skipped:
        print(f"    {name}")

print(f"\nDone. {len(created)} created, {len(skipped)} skipped.\n")