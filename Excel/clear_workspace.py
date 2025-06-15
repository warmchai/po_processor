import os
import shutil
import zipfile
from datetime import datetime, timedelta

def get_workspace_dirs():
    excel_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = os.path.abspath(os.path.join(excel_dir, '..'))
    artwork_dir = os.path.join(root_dir, 'Artwork')
    return root_dir, artwork_dir, excel_dir

def get_first_jastintech_date(jastintech_dir):
    folders = [f for f in os.listdir(jastintech_dir) if os.path.isdir(os.path.join(jastintech_dir, f))]
    if not folders:
        raise Exception("No folders found in Artwork/JastinTech.")
    # Assumes folder names start with date in m.d.yy
    first_folder = sorted(folders)[0]
    date_str = first_folder.split(' ')[0]
    return date_str, first_folder

def zip_folder(src_folder, dest_zip):
    with zipfile.ZipFile(dest_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(src_folder):
            for file in files:
                abs_path = os.path.join(root, file)
                rel_path = os.path.relpath(abs_path, os.path.dirname(src_folder))
                zf.write(abs_path, rel_path)

def zip_multiple_folders(folder_list, dest_zip):
    with zipfile.ZipFile(dest_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
        for folder in folder_list:
            for root, dirs, files in os.walk(folder):
                for file in files:
                    abs_path = os.path.join(root, file)
                    rel_path = os.path.relpath(abs_path, os.path.dirname(folder_list[0]))
                    zf.write(abs_path, rel_path)

def main():
    root_dir, artwork_dir, excel_dir = get_workspace_dirs()
    jastintech_dir = os.path.join(artwork_dir, 'JastinTech')

    # 1. Get date from first folder in JastinTech
    date_str, _ = get_first_jastintech_date(jastintech_dir)

    # 2. Zip the JastinTech folder (including the folder itself)
    jastintech_zip = os.path.join(artwork_dir, 'JastinTech.zip')
    zip_folder(jastintech_dir, jastintech_zip)

    # 3. Rename the zip to include the date
    dated_jastintech_zip = os.path.join(artwork_dir, f'JastinTech_{date_str}.zip')
    os.rename(jastintech_zip, dated_jastintech_zip)

    # 4. Zip all other folders in Artwork as Babel (excluding JastinTech)
    babel_zip = os.path.join(artwork_dir, f'Babel_{date_str}.zip')
    other_folders = [
        os.path.join(artwork_dir, f)
        for f in os.listdir(artwork_dir)
        if os.path.isdir(os.path.join(artwork_dir, f)) and f != 'JastinTech'
    ]
    if other_folders:
        zip_multiple_folders(other_folders, babel_zip)
        babel_zip_exists = True
    else:
        babel_zip_exists = False

    # 5. Move both zip files to root
    jastintech_zip_root = os.path.join(root_dir, os.path.basename(dated_jastintech_zip))
    shutil.move(dated_jastintech_zip, jastintech_zip_root)
    if babel_zip_exists:
        babel_zip_root = os.path.join(root_dir, os.path.basename(babel_zip))
        shutil.move(babel_zip, babel_zip_root)
    else:
        babel_zip_root = None

    # 6. Remove all contents from Artwork (folders and files), but not the folder itself
    for entry in os.listdir(artwork_dir):
        entry_path = os.path.join(artwork_dir, entry)
        if os.path.isdir(entry_path):
            shutil.rmtree(entry_path)
        else:
            os.remove(entry_path)

    # 7. Calculate date 2 days in the future
    try:
        dt = datetime.strptime(date_str, "%m.%d.%y")
    except ValueError:
        # fallback: try m.d.yy
        dt = datetime.strptime(date_str, "%m.%d.%y")
    # Windows: use %#m and %#d, others: use %-m and %-d
    if os.name == 'nt':
        future_date = (dt + timedelta(days=2)).strftime("%#m.%#d.%y")
    else:
        future_date = (dt + timedelta(days=2)).strftime("%-m.%-d.%y")

    # 8. Zip both zips and all .txt files in root into PO Archive
    # Archive name: "6.11.25 PO Archive (Delete on 6.13.25).zip"
    archive_name = f"{date_str} PO Archive (Delete on {future_date}).zip"
    archive_path = os.path.join(root_dir, archive_name)
    files_to_zip = [jastintech_zip_root]
    if babel_zip_exists and babel_zip_root:
        files_to_zip.append(babel_zip_root)
    files_to_zip += [
        os.path.join(root_dir, f)
        for f in os.listdir(root_dir)
        if f.lower().endswith('.txt')
    ]
    with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for file in files_to_zip:
            if os.path.exists(file):
                zf.write(file, os.path.basename(file))

    # 9. Delete files in Excel that are NOT .py or ItemCodeLEGEND.*
    for f in os.listdir(excel_dir):
        if f.endswith('.py') or f.startswith('ItemCodeLEGEND'):
            continue
        try:
            os.remove(os.path.join(excel_dir, f))
        except Exception:
            pass

    # 10. Delete .txt files and the two zip files from the root (after archiving)
    for f in os.listdir(root_dir):
        if (
            f.lower().endswith('.txt')
            or (f.startswith('JastinTech_') and f.endswith('.zip'))
            or (f.startswith('Babel_') and f.endswith('.zip'))
        ):
            try:
                os.remove(os.path.join(root_dir, f))
            except Exception:
                pass

    print("Done!")

if __name__ == "__main__":
    main()