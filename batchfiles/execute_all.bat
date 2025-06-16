import subprocess
import os

# Define the paths to the batch files
batch_files = [
    r"c:\Users\Work\Documents\A3_Resources\PO Processor\001_ProcessJastin.bat",
    r"c:\Users\Work\Documents\A3_Resources\PO Processor\002_ProcessZazzle.bat",
    r"c:\Users\Work\Documents\A3_Resources\PO Processor\003_QualityAssuranceCheck.bat"
]

def run_batch_file(file_path):
    try:
        print(f"Running: {file_path}")
        subprocess.run(file_path, shell=True, check=True)
        print(f"Finished: {file_path}\n")
    except subprocess.CalledProcessError as e:
        print(f"Error running {file_path}: {e}")

def main():
    for batch_file in batch_files:
        if os.path.exists(batch_file):
            run_batch_file(batch_file)
        else:
            print(f"Batch file not found: {batch_file}")

if __name__ == "__main__":
    main()