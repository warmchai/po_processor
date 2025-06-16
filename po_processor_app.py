import streamlit as st
import subprocess
import os

# Define the paths to the batch files
BATCH_FILES = {
    "Process Jastin Orders": r"c:\Users\Usfund\Documents\A3_Resources\PO\001_ProcessJastin.bat",
    "Process Zazzle Orders": r"c:\Users\Usfund\Documents\A3_Resources\PO\002_ProcessZazzle.bat",
    "Quality Assurance Check": r"c:\Users\Usfund\Documents\A3_Resources\PO\003_QualityAssuranceCheck.bat",
    "Zip Babel": r"c:\Users\Usfund\Documents\A3_Resources\PO\004_ZipBabel.bat",
    "Clear Workspace": r"c:\Users\Usfund\Documents\A3_Resources\PO\005_ClearWorkspace.bat",
}

def run_batch_file(file_path):
    """
    Executes a batch file using subprocess.
    """
    try:
        result = subprocess.run(file_path, shell=True, capture_output=True, text=True)
        return result.stdout, result.stderr
    except Exception as e:
        return "", str(e)

# Streamlit UI
st.title("PO Processor App")
st.header("Run Batch Files")

# Display buttons for each batch file
for label, path in BATCH_FILES.items():
    if st.button(f"Run {label}"):
        if os.path.exists(path):
            st.write(f"Running: {label}")
            stdout, stderr = run_batch_file(path)
            st.text_area("Output", stdout, height=200)
            if stderr:
                st.text_area("Errors", stderr, height=200)
        else:
            st.error(f"Batch file not found: {path}")