import os
import re

# Get the script's directory (Excel), then Artwork/JastinTech
script_dir = os.path.dirname(os.path.abspath(__file__))
jastintech_dir = os.path.join(script_dir, "..", "Artwork", "JastinTech")
jastintech_dir = os.path.abspath(jastintech_dir)

pattern = re.compile(r'^(.*PO \d+)\sA$')

for folder in os.listdir(jastintech_dir):
    match = pattern.match(folder)
    if match:
        base = match.group(1)
        # Check for siblings with same PO number but different suffix
        siblings = [f for f in os.listdir(jastintech_dir) if f.startswith(base) and f != folder]
        if not siblings:
            old_path = os.path.join(jastintech_dir, folder)
            new_path = os.path.join(jastintech_dir, base)
            print(f"Renaming: {old_path} -> {new_path}")
            os.rename(old_path, new_path)