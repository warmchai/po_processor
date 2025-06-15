import os
import re
import shutil
import pandas as pd

def num_to_letter(num):
    try:
        n = int(num)
        if n < 1:
            return None
        return chr(ord('A') + n - 1)
    except:
        return None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_DIR = BASE_DIR
ARTWORK_DIR = os.path.abspath(os.path.join(BASE_DIR, '..', 'Artwork'))

zaz_csv = next((f for f in os.listdir(EXCEL_DIR) if f.lower().startswith('zaz') and f.lower().endswith('.csv')), None)
if not zaz_csv:
    print("No ZAZ CSV file found in Excel folder.")
    exit(1)

df = pd.read_csv(os.path.join(EXCEL_DIR, zaz_csv), dtype=str)
df['PO Description'] = df['PO Description'].ffill()

po_folders = {}
for po_desc in df['PO Description'].unique():
    if pd.isna(po_desc) or not po_desc.strip():
        continue
    folder_name = po_desc.strip()
    folder_path = os.path.join(ARTWORK_DIR, folder_name)
    po_folders[po_desc] = folder_path
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print(f"Created folder: {folder_path}")

order_to_po = dict(zip(df['Order#'], df['PO Description']))

for main_folder in ['4', '5', '8']:
    main_path = os.path.join(ARTWORK_DIR, main_folder)
    if not os.path.isdir(main_path):
        continue
    for root, dirs, files in os.walk(main_path):
        for d in dirs:
            for order_num, po_desc in order_to_po.items():
                if not isinstance(order_num, str) or not isinstance(po_desc, str):
                    continue
                parts = order_num.split('_')
                base_num = parts[0]
                suffix = parts[1] if len(parts) > 1 else None
                candidates = []
                # If there is a suffix, try to match lettered folders, base folder (if no siblings), and Qty folders (if no siblings)
                if suffix:
                    letter = num_to_letter(suffix)
                    if letter:
                        candidates.append(f"{base_num}_{letter}")
                        candidates.append(f"{base_num}{letter}")
                    # Check for base folder (only if no lettered siblings)
                    if d == base_num:
                        siblings = [name for name in dirs if (name.startswith(base_num + "_") and not name.startswith(base_num + "_Qty_")) or re.match(rf"^{re.escape(base_num)}[A-Z]$", name)]
                        if not siblings:
                            candidates.append(base_num)
                    # Check for Qty folder (only if no lettered siblings)
                    if re.fullmatch(rf"{re.escape(base_num)}_Qty_\d+", d):
                        siblings = [name for name in dirs if (name.startswith(base_num + "_") and not name.startswith(base_num + "_Qty_")) or re.match(rf"^{re.escape(base_num)}[A-Z]$", name)]
                        if not siblings:
                            candidates.append(d)
                else:
                    # No suffix: match base folder, and also any Qty folder if no lettered siblings
                    candidates.append(base_num)
                    if re.fullmatch(rf"{re.escape(base_num)}_Qty_\d+", d):
                        siblings = [name for name in dirs if (name.startswith(base_num + "_") and not name.startswith(base_num + "_Qty_")) or re.match(rf"^{re.escape(base_num)}[A-Z]$", name)]
                        if not siblings:
                            candidates.append(d)
                if d in candidates:
                    src = os.path.join(root, d)
                    dest = os.path.join(po_folders[po_desc], d)
                    if os.path.abspath(src) == os.path.abspath(dest):
                        continue
                    if os.path.exists(dest):
                        print(f"Destination already exists, skipping: {dest}")
                        continue
                    print(f"Moving {src} -> {dest}")
                    shutil.move(src, dest)
        dirs[:] = [d for d in dirs if os.path.exists(os.path.join(root, d))]

for main_folder in ['4', '5', '8']:
    main_path = os.path.join(ARTWORK_DIR, main_folder)
    if not os.path.isdir(main_path):
        continue
    empty = True
    for root, dirs, files in os.walk(main_path):
        if files:
            empty = False
            break
    if empty:
        print(f"Deleting empty folder tree: {main_path}")
        shutil.rmtree(main_path)