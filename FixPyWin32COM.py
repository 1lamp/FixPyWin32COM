"""
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ FixPyWin32COM ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
This script searches for the gen_py cache folder inside the Windows temporary directory (%LOCALAPPDATA%\Temp),
and if it exists, deletes it to clean up the win32com cache.

Author: Behnam Rabieyan
Created: 2025
"""

import os
import shutil

print("...Starting win32com cache cleanup script")

try:
    # Find the exact path of the temporary files folder
    temp_path = os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp')
    gen_py_path = os.path.join(temp_path, 'gen_py')

    print(f"Searching for folder in path: {gen_py_path}")

    # Check if the folder exists and delete it
    if os.path.isdir(gen_py_path):
        print("!The 'gen_py' folder was found. Deleting...")
        shutil.rmtree(gen_py_path)
        print("!The folder was successfully deleted")
    else:
        print(".The 'gen_py' folder was not found. No cleanup needed")

except Exception as e:
    print(f"An error occurred during cleanup: {e}")

print("...Operation completed")
input("Press Enter to exit...")
