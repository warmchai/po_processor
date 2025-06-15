import os
import re
from collections import defaultdict

MAX_PACK_QTY = {
    "12_pack": 4,
    "3_pack": 16
}
QTY_PATTERN = re.compile(r'_Qty_(\d+)$')

def check_pack_quantities(jastintech_dir):
    print(f"Checking pack quantities in: {jastintech_dir}\n")
    pack_summary = defaultdict(int)  # (brand, pack_type) -> total qty
    pack_limit_errors = []
    file_count_errors = []
    po_packtype_counts = defaultdict(int)  # (PO folder, pack_type) -> count

    for po_folder in os.listdir(jastintech_dir):
        po_path = os.path.join(jastintech_dir, po_folder)
        if not os.path.isdir(po_path):
            continue
        for brand in os.listdir(po_path):
            brand_path = os.path.join(po_path, brand)
            if not os.path.isdir(brand_path):
                continue
            for pack_type in MAX_PACK_QTY.keys():
                pack_dir = os.path.join(brand_path, pack_type)
                if not os.path.isdir(pack_dir):
                    continue
                qty_sum = 0
                for qty_folder in os.listdir(pack_dir):
                    qty_folder_path = os.path.join(pack_dir, qty_folder)
                    if not os.path.isdir(qty_folder_path):
                        continue
                    m = QTY_PATTERN.search(qty_folder)
                    qty = int(m.group(1)) if m else 1
                    qty_sum += qty
                    # Check for file count > 3
                    files = [f for f in os.listdir(qty_folder_path) if os.path.isfile(os.path.join(qty_folder_path, f))]
                    if len(files) > 3:
                        file_count_errors.append(f"ERROR: {qty_folder_path} has {len(files)} files (should be no more than 3)!")
                pack_summary[(brand, pack_type)] += qty_sum
                po_packtype_counts[(po_folder, pack_type)] += qty_sum
                if qty_sum > MAX_PACK_QTY[pack_type]:
                    pack_limit_errors.append(f"ERROR: {pack_dir} has {qty_sum} packs (limit: {MAX_PACK_QTY[pack_type]})")

    print("Pack summary per golf subfolder:")
    if pack_summary:
        for (brand, pack_type), count in sorted(pack_summary.items()):
            print(f"  {brand} - {pack_type}: {count} packs")
    else:
        print("  No packs found.")
    print("Pack quantity and order folder checks complete.\n")

    # --- Print the PO folder summary with pack type ---
    print("Pack count per JastinTech PO folder:")
    for (po_folder, pack_type), total_packs in sorted(po_packtype_counts.items()):
        print(f"{po_folder} -> {total_packs} Packs ({pack_type})")
    print()

    if not pack_limit_errors:
        print("All folders are under their limits! (3 pack - 16, 12 pack - 4)")
    else:
        print("Some folders exceed their limits:")
        for err in pack_limit_errors:
            print(err)

    print()
    if not file_count_errors:
        print("No folders containing more than 3 files!")
    else:
        print("Some folders contain more than 3 files:")
        for err in file_count_errors:
            print(err)

if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    jastintech_dir = os.path.abspath(os.path.join(base_dir, '..', 'Artwork', 'JastinTech'))
    if not os.path.isdir(jastintech_dir):
        print(f"JastinTech folder not found at: {jastintech_dir}")
    else:
        check_pack_quantities(jastintech_dir)