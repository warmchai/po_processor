import os
import re
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_DIR = BASE_DIR
ARTWORK_DIR = os.path.abspath(os.path.join(BASE_DIR, '..', 'Artwork'))

# --- Locate the ZAZ CSV file in the Excel folder ---
zaz_csv = next((f for f in os.listdir(EXCEL_DIR) if f.lower().startswith('zaz') and f.lower().endswith('.csv')), None)
if not zaz_csv:
    print("ERROR: ZAZ CSV file not found in Excel folder.")
    exit(1)

# --- Read and forward-fill the PO Description column ---
df = pd.read_csv(os.path.join(EXCEL_DIR, zaz_csv), dtype=str)
df['PO Description'] = df['PO Description'].ffill()

# --- Prepare Order# to PO Description mapping ---
order_to_po = dict(zip(df['Order#'], df['PO Description']))

# --- Prepare PO Description to list of orders mapping ---
po_to_orders = {}
for order_num, po_desc in order_to_po.items():
    if pd.isna(po_desc) or not po_desc.strip():
        continue
    if po_desc not in po_to_orders:
        po_to_orders[po_desc] = []
    po_to_orders[po_desc].append(order_num)

all_sorted = True

for po_desc, orders in po_to_orders.items():
    po_folder = os.path.join(ARTWORK_DIR, po_desc)
    if not os.path.isdir(po_folder):
        print(f"ERROR: No folder found for PO Description: {po_desc}")
        all_sorted = False
        continue
    po_subfolders = os.listdir(po_folder)
    for order_num in orders:
        base_num = order_num.split('_')[0]
        found = False
        for d in po_subfolders:
            # Use the exact same regex as the processor
            if re.match(rf"^{re.escape(base_num)}(_?[A-Z]+)?(_Qty_\d+)?$", d):
                found = True
                break
        if not found:
            print(f"ERROR: Order {order_num} not found in {po_desc}")
            all_sorted = False

print()
print("ALL ZAZZLE ORDERS PROPERLY SORTED")
if all_sorted:
    print()
    print("- All order folders are properly sorted.")
    print()
    for po_desc, orders in po_to_orders.items():
        print(f"{po_desc} -> {len(orders)} Orders")
else:
    print()
    print("- Zazzle order folder errors found. See above for details.")