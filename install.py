# =========================================
# INSTALL (FOR PPTX2TXT AND MULTI PPTX2TXT)
# =========================================
#
# VERSION
# =======
# 1.1.0 | 16 June 2025
#
# COPYRIGHT STATEMENT
# ===================
# This script was created by Oxzeon Limited (oxzeon.com) with assistance from ChatGPT.
# No copyright is asserted - please use/modify/redistribute as you wish.
# CAUTION: Use entirely at your own risk.
# Contact: alan@oxzeon.com
#
# OVERVIEW
# ========
# This script asks the user to select a directory, then creates a new Python venv (virtual environment) in
# that location, installing all required Python libraries and files to run PPTX2TXT and MULTI PPTX2TXT.
#
# REQUIREMENTS
# ============
# - Python 3.13+
#
# SETUP & USAGE
# =============
# Please see the readme.txt file that ships with this Python script.


import os
import shutil
import subprocess
import sys
import tkinter

from tkinter.filedialog import askdirectory


# Ask the user to select the directory in which to create the venv
root = tkinter.Tk()
print("\nSelect parent directory for the venv (see file explorer/finder prompt)")
venv_dir_parent_path = askdirectory()
root.destroy()

# Define venv directory path and Python packages to install
venv_dir_path = os.path.join(venv_dir_parent_path, "venv_ppt2txt")
print(f"venv directory: {venv_dir_path}")
packages = ["python-pptx"]

# Create the venv
subprocess.run([sys.executable, "-m", "venv", venv_dir_path])

# Construct the path to the venv's pip
if os.name == "nt":  # Windows
    pip_path = os.path.join(venv_dir_path, "Scripts", "pip.exe")
else:  # Unix/macOS
    pip_path = os.path.join(venv_dir_path, "bin", "pip")

# Install packages into the venv
subprocess.run([pip_path, "install"] + packages)
print(f"venv created in '{venv_dir_path}' with package(s): {', '.join(packages)}")

# Copy the solution files into the root of the venv
files_to_copy: list = ["pptx2txt.py", "multi_pptx2txt.py", "readme.txt"]
for file_name in files_to_copy:
    source_dir_path: str = os.path.dirname(os.path.abspath(__file__))
    source_file_path: str = os.path.join(source_dir_path, file_name)
    target_file_path: str = os.path.join(venv_dir_path, file_name)
    shutil.copy2(source_file_path, target_file_path)
    print(f"Copied {source_file_path} to {target_file_path}")

# Tell the user installation is complete
print("PPTX2TXT installation complete")
