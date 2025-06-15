import os
import re
import pandas as pd
from string import ascii_uppercase
from collections import defaultdict

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JASTINTECH_DIR = os.path.abspath(os.path.join(BASE_DIR, '..', 'Artwork', 'JastinTech'))
LEGEND_PATH = os.path.join(BASE_DIR, 'ItemCodeLEGEND.csv')

# Build Item Name → lowest Item Number mapping
legend_df = pd.read_csv(LEGEND_PATH, header=None, dtype=str, names=['ItemName', 'ItemCode'])
legend_df['ItemNum'] = legend_df['ItemCode'].str.extract(r'(\d+)').astype(int)
itemname_to_min_itemnum = legend_df.groupby('ItemName')['ItemNum'].min().to_dict()
item_names = set(legend_df['ItemName']) - {'Megaphone', 'Magnet'}

# Scan all folders in JastinTech
folders = [f for f in os.listdir(JASTINTECH_DIR) if os.path.isdir(os.path.join(JASTINTECH_DIR, f))]

# Group folders by 4-digit PO number found anywhere in the name
po_groups = defaultdict(list)
for folder in folders:
    match = re.search(r'(\d{4})', folder)
    if match:
        po_num = match.group(1)
        po_groups[po_num].append(folder)

# Process each PO group
for po_num, siblings in po_groups.items():
    if len(siblings) < 2:
        continue

    # Find folders with an Item Name (not Megaphone/Magnet) in the name
    item_folders = []
    for folder in siblings:
        for item in item_names:
            if re.search(rf'\b{re.escape(item)}\b', folder, re.IGNORECASE):
                item_folders.append((folder, item))
                break

    # Only proceed if at least one folder has an item name
    if len(item_folders) == 0:
        continue

    # Find used alphabetical indicators among non-item-name siblings
    used_letters = set()
    for folder in siblings:
        has_item = any(re.search(rf'\b{re.escape(item)}\b', folder, re.IGNORECASE) for item in item_names)
        after_po = re.search(rf'{po_num}\s+([A-Z])\b', folder)
        if not has_item and after_po:
            used_letters.add(after_po.group(1).upper())

    # Sort item folders by the lowest Item Number for their Item Name
    item_folders.sort(key=lambda x: itemname_to_min_itemnum.get(x[1], float('inf')))

    # Assign next available letters in order, skipping used ones
    letter_iter = (l for l in ascii_uppercase if l not in used_letters)
    for folder, item in item_folders:
        # If already has a letter after PO number, skip
        if re.search(rf'{po_num}\s+[A-Z]\b', folder):
            continue
        letter = next(letter_iter)
        item_pos = folder.lower().find(item.lower())
        if item_pos == -1:
            continue
        before = folder[:item_pos]
        after = folder[item_pos:]
        new_name = re.sub(rf'({po_num})(\s*)', rf'\1 {letter} ', before, count=1) + after
        src = os.path.join(JASTINTECH_DIR, folder)
        dst = os.path.join(JASTINTECH_DIR, new_name)
        if not os.path.exists(dst):
            os.rename(src, dst)