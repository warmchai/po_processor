import os
import pandas as pd

def convert_all_excels_to_csv():
    folder_path = os.path.dirname(os.path.abspath(__file__))  # Always the Excel folder
    print(f"exceltocsv.py started in {folder_path}", flush=True)
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(('.xlsx', '.xls')) and not filename.startswith('~$'):
            excel_path = os.path.join(folder_path, filename)
            print(f"Processing {excel_path}...", flush=True)
            xls = pd.ExcelFile(excel_path)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                base = os.path.splitext(filename)[0]
                safe_sheet = "".join([c if c.isalnum() else "_" for c in sheet_name])
                csv_filename = f"{base}_{safe_sheet}.csv"
                csv_path = os.path.join(folder_path, csv_filename)

                # If filename starts with "ZAZ", keep only columns H and I
                if filename.upper().startswith("ZAZ"):
                    try:
                        df = df.iloc[:, [7, 8]]
                    except IndexError:
                        print(f"Warning: '{excel_path}' sheet '{sheet_name}' does not have columns H and I.")
                        continue

                df.to_csv(csv_path, index=False)
                print(f"Converted '{excel_path}' sheet '{sheet_name}' to '{csv_path}'.", flush=True)

if __name__ == "__main__":
    convert_all_excels_to_csv()