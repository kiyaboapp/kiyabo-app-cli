# KIYABO APP  
**Tanzania School Information System CLI**  
*Professional, Fast, Colorful, Standalone, and Ready for Primary, O-Level, A-Level*

---

## Overview

KIYABO APP is a command-line tool designed to manage student exam results in Tanzanian schools. It supports:

- Uploading results from Excel to Access  
- Exporting ranked results to formatted Excel  
- Processing & Ranking exams with INC handling  
- Standalone 32-bit & 64-bit EXEs (no Python required)  
- Zero dependencies outside your local Python folders  
- Full-color output with ASCII banner  
- Automatic GitHub releases on every tag  
- Custom icon (icon.ico) embedded in EXEs

---

## Project Structure

c:\KiyaboAppPython\
│
├───cli/                          ← All CLI logic
│   ├── __init__.py
│   ├── main.py                   ← Typer CLI entry point
│   ├── colors.py                 ← Rich console instance
│   ├── banner.py                 ← Big colored ASCII banner
│   │
│   └───alevel/                   ← A-Level modules (extend to others)
│       ├── __init__.py
│       ├── exporter.py           ← StudentExamExporter class
│       ├── importer.py           ← ExamDataImporter class
│       └── ranking.py            ← process_exam() function
│
├───python32/                     ← 32-bit Python + dependencies (ignored)
├───python64/                     ← 64-bit Python + dependencies (ignored)
├───downloads/                    ← Local files (ignored)
│
├── run_cli.py                    ← Fixes import path for `cli` package
├── kiyabo32.bat                  ← Launch with 32-bit Python
├── kiyabo64.bat                  ← Launch with 64-bit Python
├── build_32.spec                 ← PyInstaller config (32-bit) - NOW INCLUDES ICON
├── build_64.spec                 ← PyInstaller config (64-bit) - NOW INCLUDES ICON
├── icon.ico                      ← Custom app icon (embedded in EXEs)
│
├── .gitignore
├── README.txt                    ← This file
└── .github/
    └── workflows/
        └── release.yml           ← Auto-build & release EXEs - NOW INCLUDES ICON

---

## Quick Start

### 1. Run Locally

cd c:\KiyaboAppPython

#### 64-bit
kiyabo64.bat --help

#### 32-bit
kiyabo32.bat --help

> You’ll see a colorful ASCII banner and full help.

---

## CLI Commands

kiyabo <command> <level> [options]

### Supported Levels
| Level     | Status       |
|-----------|--------------|
| alevel    | Fully Working     |
| olevel    | Ready for extension |
| primary   | Ready for extension |

---

### upload – Import Excel to Access

kiyabo upload alevel \
  --exam-id MID520250825 \
  --excel "C:\Data\Input_Results.xlsx" \
  --db "C:\Kiyabo App\backend\Kiyabo App Backend v4.0.0.accdb"

- Uses ExamDataImporter.import_exam_data()
- Returns success/failure

---

### export – Export Ranked Results to Excel

kiyabo export alevel \
  --exam-id MID520250825 \
  --db "C:\Kiyabo App\backend\Kiyabo App Backend v4.0.0.accdb" \
  --comb \
  --top 20 \
  --bottom 20

- Instantiates StudentExamExporter(...)
- Output: C:\Kiyabo App\Results\Exam_Results_MID520250825.xlsx
- Includes formatting, rankings, top/bottom students

---

### process – Rank & Process Exam

kiyabo process alevel \
  --exam-id MID520250825 \
  --db "C:\Users\droge\OneDrive\Documents\Kiyabo App Backend v4.0.0.accdb" \
  --no-inc

- Calls process_exam(exam_id, db_path, include_INC=False)
- Handles INC as penalty

---

## Build Standalone EXEs (WITH ICON)

### Generate (once) – includes icon.ico

python32\python.exe -m PyInstaller --name kiyabo32 --onefile --console --icon=icon.ico run_cli.py
python64\python.exe -m PyInstaller --name kiyabo64 --onefile --console --icon=icon.ico run_cli.py

ren kiyabo32.spec build_32.spec
ren kiyabo64.spec build_64.spec

### Build anytime (icon already in .spec)

python32\python.exe -m PyInstaller --noconfirm build_32.spec
python64\python.exe -m PyInstaller --noconfirm build_64.spec

> Output:  
> - dist\kiyabo32.exe (with your icon)  
> - dist\kiyabo64.exe (with your icon)

Console stays open – full color, banner, help.

---

## .spec Files – NOW INCLUDE ICON

### build_32.spec
a = Analysis(
    ['run_cli.py'],
    pathex=[os.path.abspath(".")],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=None)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='kiyabo32',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    icon='icon.ico',   ← ICON ADDED
    disable_windowed_traceback=False,
)

### build_64.spec
a = Analysis(
    ['run_cli.py'],
    pathex=[os.path.abspath(".")],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=None)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='kiyabo64',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    icon='icon.ico',   ← ICON ADDED
    disable_windowed_traceback=False,
)

---

## GitHub Auto-Release – NOW INCLUDES ICON

### .github/workflows/release.yml
name: Build & Release

on:
  push:
    branches: [ main ]
    tags: [ 'v*' ]

jobs:
  build:
    strategy:
      matrix:
        include:
          - python_folder: python32
            exe_name: kiyabo32
            spec_file: build_32.spec
          - python_folder: python64
            exe_name: kiyabo64
            spec_file: build_64.spec
    runs-on: windows-latest
    defaults:
      run:
        shell: cmd
    steps:
      - uses: actions/checkout@v4
      - name: Build ${{ matrix.exe_name }}
        run: |
          ${{ matrix.python_folder }}\python.exe -m PyInstaller --noconfirm ${{ matrix.spec_file }}
      - name: Upload Artifact
        uses: actions/upload-artifact@v4
        with:
          name: ${{ matrix.exe_name }}
          path: dist/${{ matrix.exe_name }}.exe

  release:
    needs: build
    runs-on: ubuntu-latest
    if: startsWith(github.ref, 'refs/tags/')
    permissions:
      contents: write
    steps:
      - uses: actions/download-artifact@v4
      - name: Release
        uses: softprops/action-gh-release@v2
        with:
          files: |
            kiyabo32/kiyabo32.exe
            kiyabo64/kiyabo64.exe

---

## Dependencies (Already Installed in python32 / python64)

typer[all]
rich==14.2.0
pyodbc
openpyxl
pandas
numpy
pywin32
pyinstaller

> No pip install required if already in your Python folders.

---

## Extending to primary or olevel

1. Create folders
   cli/primary/
   cli/olevel/

2. Add files
   cli/primary/exporter.py
   cli/primary/importer.py
   cli/primary/ranking.py

3. Implement
   - process_exam(exam_id, db_path, include_INC)
   - class ExamDataImporter: import_exam_data(...)
   - class StudentExamExporter(...)

4. Import in cli/main.py
   from .primary.ranking import process_exam as primary_process_exam

5. Update commands
   if level == "primary":
       primary_process_exam(...)

---

## .gitignore

# Folders
downloads/
python32/
python64/

# Python
__pycache__/
*.pyc
*.pyo
*.pyd

# PyInstaller
dist/
build/
<!-- *.spec -->

# IDE
.vscode/
.idea/
*.swp

---

## Troubleshooting

| Problem | Solution |
|-------|----------|
| ModuleNotFoundError: cli | Use kiyabo64.bat – it sets sys.path |
| EXE has no console | --console=True in .spec |
| Icon not showing | Ensure icon.ico is in root and .spec has icon='icon.ico' |
| process_exam not found | Must exist in ranking.py |
| ExamDataImporter missing | Must be a class with import_exam_data() |
| Colors not showing | Use Windows Terminal or ConEmu |

---

## Testing Your Modules (Manual)

# Test ranking
python64\python.exe cli\alevel\ranking.py MID520250825 --dbpath "C:\path\to\db.accdb"

# Test importer
python64\python.exe -c "
from cli.alevel.importer import ExamDataImporter
importer = ExamDataImporter()
importer.import_exam_data('TEST', 'input.xlsx', 'db.accdb')
"

---

## Versioning

Use Semantic Versioning:

git tag v1.0.0
git push origin v1.0.0

---

## License

MIT © Kiyabo App Team

---

Made with precision for Tanzanian schools.