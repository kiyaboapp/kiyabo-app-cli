PORTABLE PYTHON DISTRIBUTION
============================

This is a fully portable Python 3.8.10 installation.
No installation required - just copy the folders and run.

CONTENTS:
  python32/  - 32-bit Python with all packages
  python64/  - 64-bit Python with all packages

USAGE:
  Navigate to the python folder and run scripts normally:

    cd python32
    python.exe your_script.py

    cd python64
    python.exe your_script.py

  Or use full path:
    python32\python.exe your_script.py
    python64\python.exe your_script.py

INSTALLED PACKAGES:
  - pandas
  - numpy
  - openpyxl
  - pyodbc
  - pyinstaller
  - tqdm
  - rich
  - tabulate

TESTING:
  python32\python.exe python32\test.py
  python64\python.exe python64\test.py

DISTRIBUTION:
  Simply copy the python32 and/or python64 folders to any location.
  No PATH changes or registry modifications needed.
  Each folder is completely self-contained.
  Run python.exe directly like normal Python installations.
