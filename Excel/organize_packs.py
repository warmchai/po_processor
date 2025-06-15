#!/usr/bin/env python3
import os
import re
import shutil
import sys
from string import ascii_uppercase

# Define your pack-type limits here:
MAX_QUANTITIES = {
    "12_pack": 4,
    "3_pack": 16,
}

def parse_qty(name):
    m = re.search(r'_Qty_(\d+)$', name)
    return int(m.group(1)) if m else 1

def chunk_items(items, limit):
    groups = []
    current = []
    total = 0
    for name, qty in items:
        if total + qty > limit:
            groups.append(current)
            current, total = [], 0
        current.append(name)
        total += qty
    if current:
        groups.append(current)
    return groups

def add_A_B_to_special_folders(po_path):
    parent = os.path.dirname(po_path)
    po_name = os.path.basename(po_path)
    po_match = re.search(r'(PO\s*\d+)', po_name, re.IGNORECASE)
    if not po_match:
        return 0
    po_number = po_match.group(1)
    siblings = []
    for sibling in os.listdir(parent):
        sibling_path = os.path.join(parent, sibling)
        if not os.path.isdir(sibling_path):
            continue
        if sibling == po_name:
            continue
        if po_number in sibling:
            if re.search(r'Magnet', sibling, re.IGNORECASE):
                siblings.append(('Magnet', sibling, sibling_path))
            elif re.search(r'Megaphone', sibling, re.IGNORECASE):
                siblings.append(('Megaphone', sibling, sibling_path))
    siblings.sort(key=lambda x: 0 if x[0] == 'Magnet' else 1)
    used_letters = []
    for idx, (keyword, sibling, sibling_path) in enumerate(siblings):
        letter = ascii_uppercase[idx]
        if not re.search(rf'{po_number}\s*{letter}\b', sibling):
            new_name = re.sub(
                rf'({re.escape(po_number)})\s*',
                rf'\1 ' + letter + ' ',
                sibling,
                count=1,
                flags=re.IGNORECASE
            )
            new_path = os.path.join(parent, new_name)
            if not os.path.exists(new_path):
                print(f"Renaming '{sibling}' to '{new_name}'")
                os.rename(sibling_path, new_path)
            used_letters.append(letter)
        else:
            used_letters.append(letter)
    return len(used_letters)

def organize_po(po_path):
    special_count = add_A_B_to_special_folders(po_path)
    po_name = os.path.basename(po_path)
    parent = os.path.dirname(po_path)
    all_groups = []
    for pack_type in ("12_pack", "3_pack"):
        for brand in sorted(os.listdir(po_path)):
            brand_dir = os.path.join(po_path, brand)
            orders_dir = os.path.join(brand_dir, pack_type)
            if not os.path.isdir(orders_dir):
                continue
            items = []
            for fname in sorted(os.listdir(orders_dir)):
                fpath = os.path.join(orders_dir, fname)
                if os.path.isdir(fpath):
                    qty = parse_qty(fname)
                    items.append((fname, qty))
            groups = chunk_items(items, MAX_QUANTITIES[pack_type])
            for group in groups:
                all_groups.append((brand, pack_type, group))
    start_idx = special_count
    for idx, (brand, pack_type, group) in enumerate(all_groups):
        letter = ascii_uppercase[idx + start_idx]
        new_po = f"{po_name} {letter}"
        new_po_path = os.path.join(parent, new_po)
        dest = os.path.join(new_po_path, brand, pack_type)
        os.makedirs(dest, exist_ok=True)
        for fname in group:
            src = os.path.join(po_path, brand, pack_type, fname)
            dst = os.path.join(dest, fname)
            shutil.move(src, dst)
        print(f"Created {new_po} with {len(group)} orders in {pack_type}")

def delete_empty_folders(root_dir):
    # Keep deleting empty folders until none are left
    removed = True
    while removed:
        removed = False
        for dirpath, dirnames, filenames in os.walk(root_dir, topdown=False):
            if not dirnames and not filenames:
                print(f"Deleting empty folder: {dirpath}")
                os.rmdir(dirpath)
                removed = True

if __name__ == "__main__":
    # Get the script's directory (Excel), then Artwork/JastinTech
    script_dir = os.path.dirname(os.path.abspath(__file__))
    artwork_dir = os.path.abspath(os.path.join(script_dir, "..", "Artwork", "JastinTech"))
    if not os.path.isdir(artwork_dir):
        print(f"Target directory not found: {artwork_dir}")
        sys.exit(1)
    for name in os.listdir(artwork_dir):
        po_dir = os.path.join(artwork_dir, name)
        if os.path.isdir(po_dir) and re.search(r'\bPO\s*\d+', name, re.IGNORECASE):
            print(f"Processing {name}…")
            organize_po(po_dir)
    # Delete empty folders after all processing
    delete_empty_folders(artwork_dir)
    print("Done.")