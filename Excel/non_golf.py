import os
import shutil
import pandas as pd
import re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ARTWORK_DIR = os.path.abspath(os.path.join(BASE_DIR, '..', 'Artwork', 'JastinTech'))
EXCEL_DIR = BASE_DIR

# --- FIND CSV FILES ---
jastin_csv = next((f for f in os.listdir(EXCEL_DIR) if f.lower().startswith('jastin') and f.lower().endswith('.csv')), None)
legend_csv = 'ItemCodeLEGEND.csv'

if not jastin_csv or not os.path.exists(os.path.join(EXCEL_DIR, legend_csv)):
    print("Required CSV files not found.")
    exit(1)

# --- LOAD DATA ---
jastin_df = pd.read_csv(os.path.join(EXCEL_DIR, jastin_csv), dtype=str, usecols=[6, 8, 9])  # G, I, J
jastin_df.columns = ['PO#', 'Order#', 'PTC Item#']
jastin_df['PO#'] = jastin_df['PO#'].ffill()  # Ensure all rows have a PO#
legend_df = pd.read_csv(os.path.join(EXCEL_DIR, legend_csv), dtype=str)
legend_dict = dict(zip(legend_df.iloc[:,1], legend_df.iloc[:,0]))  # {ItemCode: ItemName}

# --- GROUP BY PO ---
for po_num, group in jastin_df.groupby('PO#'):
    print(f"\n=== Processing PO#: {po_num} ===")
    # Extract just the 4 digits from the PO# (e.g., "PO 5373" -> "5373")
    match = re.search(r'(\d{4})$', str(po_num).strip())
    if not match:
        print(f"Could not extract 4-digit PO number from: {po_num}")
        continue
    po_digits = match.group(1)

    # Find PO folder whose name ends with those 4 digits
    po_folder = None
    for folder in os.listdir(ARTWORK_DIR):
        if folder.strip().endswith(po_digits):
            po_folder = os.path.join(ARTWORK_DIR, folder)
            break
    print(f"PO folder resolved: {po_folder}")
    if not po_folder or not os.path.isdir(po_folder):
        print(f"PO folder not found for PO#: {po_num} (looking for folder ending with {po_digits})")
        continue

    # Identify golf folders (contain subfolders ending with "_pack")
    golf_folders = set()
    for sub in os.listdir(po_folder):
        sub_path = os.path.join(po_folder, sub)
        if os.path.isdir(sub_path):
            for subsub in os.listdir(sub_path):
                if subsub.lower().endswith("_pack"):
                    golf_folders.add(sub)

    # --- PROCESS NON-GOLF ITEMS ---
    non_golf_rows = []
    for _, row in group.iterrows():
        item_code = str(row['PTC Item#']).strip()
        item_name = legend_dict.get(item_code, "")
        order_num = str(row['Order#']).strip()
        print(f"\nDEBUG: order_num={order_num}, item_code={item_code}, item_name={item_name}")
        if not item_name or item_name.lower().endswith("_pack"):
            print("DEBUG: Skipping row due to missing or _pack item_name")
            continue
        non_golf_rows.append((order_num, item_code, item_name))

        # Extract just the numbers before the first underscore in order_num
        order_num_prefix = order_num.split('_')[0] if '_' in order_num else order_num
        print(f"DEBUG: order_num_prefix={order_num_prefix}")

        # 1. Handle numeric subfolders (e.g., "7")
        for sub in os.listdir(po_folder):
            sub_path = os.path.join(po_folder, sub)
            if sub in golf_folders or not os.path.isdir(sub_path):
                continue
            if sub.isdigit():
                for f in os.listdir(sub_path):
                    print(f"DEBUG: [subfolder] Checking file: {f} for item_code: {item_code} or order_num_prefix: {order_num_prefix}")
                    if item_code in f or order_num_prefix in f:
                        new_sub_path = os.path.join(po_folder, item_name)
                        print(f"DEBUG: Renaming {sub_path} to {new_sub_path}")
                        if not os.path.exists(new_sub_path):
                            os.rename(sub_path, new_sub_path)
                            sub_path = new_sub_path
                        # Special: If Megaphone, delete .png inside
                        if item_name.lower() == "megaphone":
                            for file in os.listdir(sub_path):
                                if file.lower().endswith('.png'):
                                    print(f"DEBUG: Deleting {os.path.join(sub_path, file)}")
                                    os.remove(os.path.join(sub_path, file))
                        break

        # 2. Handle loose files in PO folder (look for PTC Item# or order_num_prefix)
        for f in os.listdir(po_folder):
            f_path = os.path.join(po_folder, f)
            print(f"DEBUG: [loose] Checking file: {f} for item_code: {item_code} or order_num_prefix: {order_num_prefix}")
            print(f"DEBUG: item_code in f? {item_code in f}")
            print(f"DEBUG: order_num_prefix in f? {order_num_prefix in f}")
            if os.path.isfile(f_path) and (item_code in f or order_num_prefix in f):
                item_folder = os.path.join(po_folder, item_name)
                os.makedirs(item_folder, exist_ok=True)
                print(f"DEBUG: Moving {f_path} to {item_folder}")
                shutil.move(f_path, item_folder)

        # 3. If Megaphone, also delete .png files in the item folder (for loose files)
        if item_name.lower() == "megaphone":
            item_folder = os.path.join(po_folder, item_name)
            if os.path.exists(item_folder):
                for file in os.listdir(item_folder):
                    if file.lower().endswith('.png'):
                        print(f"DEBUG: Deleting {os.path.join(item_folder, file)}")
                        os.remove(os.path.join(item_folder, file))

    # --- DECIDE IF WE NEED TO CREATE A NEW OUTER FOLDER OR JUST RENAME ---
    # Count non-golf folders (excluding golf folders and hidden/system folders)
    all_subfolders = [d for d in os.listdir(po_folder) if os.path.isdir(os.path.join(po_folder, d))]
    non_golf_subfolders = [d for d in all_subfolders if d not in golf_folders and not d.startswith('.') and not d.lower().endswith('_pack')]

    # If there are NO golf folders and only one non-golf folder, just rename the PO folder
    if not golf_folders and len(non_golf_subfolders) == 1:
        item_name = non_golf_rows[0][2] if non_golf_rows else ""
        if item_name:
            new_folder_name = f"{os.path.basename(po_folder)} {item_name}".rstrip()
            new_folder_path = os.path.join(ARTWORK_DIR, new_folder_name)
            print(f"DEBUG: Renaming {po_folder} to {new_folder_path}")
            if not os.path.exists(new_folder_path):
                os.rename(po_folder, new_folder_path)
            else:
                print(f"Target folder {new_folder_path} already exists, cannot rename {po_folder}.")
    else:
        # --- CREATE [original_folder_name] [ItemName] FOR EACH NON-GOLF ITEM ---
        for order_num, item_code, item_name in non_golf_rows:
            item_folder = os.path.join(po_folder, item_name)
            if not os.path.exists(item_folder):
                continue
            new_outer = os.path.join(ARTWORK_DIR, f"{os.path.basename(po_folder)} {item_name}".rstrip())
            if not os.path.exists(new_outer):
                os.makedirs(new_outer, exist_ok=True)
            dest = os.path.join(new_outer, item_name)
            print(f"DEBUG: Moving {item_folder} to {dest}")
            if not os.path.exists(dest):
                shutil.move(item_folder, dest)
            else:
                shutil.rmtree(item_folder)

print("Done.")