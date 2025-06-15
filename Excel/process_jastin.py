import os
import shutil
import glob
import re
from datetime import datetime
import pandas as pd

def prompt_date():
    while True:
        date_str = input("What's the PO Date? (e.g. 6/10/2025, 06-10-25, etc): ")
        try:
            for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%m.%d.%Y", "%m/%d/%y", "%m-%d-%y", "%m.%d.%y"):
                try:
                    dt = datetime.strptime(date_str, fmt)
                    return dt
                except ValueError:
                    continue
            match = re.match(r'(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})', date_str)
            if match:
                m, d, y = match.groups()
                y = int(y)
                if y < 100:
                    y += 2000
                return datetime(int(y), int(m), int(d))
        except Exception:
            pass
        print("Could not parse date. Please try again.", flush=True)

def prompt_po_numbers():
    while True:
        po_str = input("What are today's Jastin Zazzle PO numbers? (e.g. 2224 2225): ")
        po_numbers = re.findall(r'\b2\d{3}\b', po_str)
        if po_numbers:
            return sorted([int(po) for po in po_numbers])
        print("Please enter one or more 4-digit PO numbers starting with 2.", flush=True)

def move_and_cleanup_golf(dest_folder):
    golf_folder = os.path.join(dest_folder, "Golf")
    if os.path.isdir(golf_folder):
        for item in os.listdir(golf_folder):
            src = os.path.join(golf_folder, item)
            dst = os.path.join(dest_folder, item)
            shutil.move(src, dst)
        # Remove Golf folder if empty
        if not os.listdir(golf_folder):
            os.rmdir(golf_folder)

def main():
    # Set up paths
    script_dir = os.path.dirname(os.path.abspath(__file__))
    workspace_dir = os.path.abspath(os.path.join(script_dir, ".."))
    artwork_dir = os.path.join(workspace_dir, "Artwork")
    excel_dir = os.path.join(workspace_dir, "Excel")
    jastintech_dir = os.path.join(artwork_dir, "JastinTech")

    # 1. Prompt user
    po_date = prompt_date()
    po_numbers = prompt_po_numbers()
    date_str = f"{po_date.month}.{po_date.day}.{str(po_date.year)[2:]}"  # m.d.yy

    # 2. Create JastinTech subfolders for Zazzle POs
    os.makedirs(jastintech_dir, exist_ok=True)
    po_folder_names = []
    for po in po_numbers:
        folder_name = f"{date_str} PO {po}".rstrip()  # Ensure no trailing spaces
        po_folder_names.append(folder_name)
        os.makedirs(os.path.join(jastintech_dir, folder_name), exist_ok=True)

    # 3. Move golf folders (unchanged)
    golf_dir = os.path.join(artwork_dir, "golf")
    if os.path.isdir(golf_dir):
        tm_folders = [f for f in os.listdir(golf_dir) if f.lower().startswith("taylor_made_tp5")]
        titleist_folders = [f for f in os.listdir(golf_dir) if f.lower().startswith("titleist_pro_v1")]
        if len(po_numbers) >= 2 and tm_folders and titleist_folders:
            lower_po, higher_po = sorted(po_numbers)
            shutil.move(os.path.join(golf_dir, titleist_folders[0]), os.path.join(jastintech_dir, f"{date_str} PO {lower_po}".rstrip(), titleist_folders[0]))
            shutil.move(os.path.join(golf_dir, tm_folders[0]), os.path.join(jastintech_dir, f"{date_str} PO {higher_po}".rstrip(), tm_folders[0]))
        if not os.listdir(golf_dir):
            os.rmdir(golf_dir)

    # 4. Process Excel (CSV) file
    jastin_csvs = glob.glob(os.path.join(excel_dir, "Jastin*.csv"))
    if not jastin_csvs:
        print("No Jastin CSV file found in Excel folder.")
        return
    csv_path = jastin_csvs[0]
    df = pd.read_csv(csv_path, dtype=str, keep_default_na=False)

    # 5. Move and relabel folders referenced by PO# (Column G)
    artwork_folders = os.listdir(artwork_dir)
    for idx, row in df.iterrows():
        po_col = row.get('PO#', '')
        match = re.search(r'(\d{4})', str(po_col))
        if match:
            po_num = match.group(1)
            for folder in artwork_folders:
                if po_num in folder:
                    src_folder = os.path.join(artwork_dir, folder)
                    dest_folder = os.path.join(jastintech_dir, f"{date_str} PO {po_num}".rstrip())
                    # Only move/rename if not already at destination with correct name
                    if os.path.isdir(src_folder):
                        # If destination already exists, remove it first (to avoid error)
                        if os.path.exists(dest_folder):
                            shutil.rmtree(dest_folder)
                        shutil.move(src_folder, dest_folder)
                        # After moving and renaming, handle Golf subfolder
                        move_and_cleanup_golf(dest_folder)
    print("Artwork folders processed, moved, and relabeled based on PO# in CSV.")

if __name__ == "__main__":
    main()