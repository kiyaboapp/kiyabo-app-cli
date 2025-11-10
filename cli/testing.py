# -*- coding: utf-8 -*-
"""
PyInstaller-packaged Python-Environment Tester
Focus: pandas • numpy • pyodbc (with fake data)
No external Python required – just run the .exe
"""

import sys
import os
import subprocess
import json
import tempfile
import platform
from pathlib import Path

# --------------------------------------------------------------
# ANSI colour palette (works in Windows 10+ console)
# --------------------------------------------------------------
class C:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    RESET = '\033[0m'


def banner(text: str):
    line = "=" * 80
    print(f"\n{C.OKBLUE}{C.BOLD}{line}{C.RESET}")
    print(f"{C.OKBLUE}{C.BOLD}{text.center(80)}{C.RESET}")
    print(f"{C.OKBLUE}{C.BOLD}{line}{C.RESET}\n")


def result(name: str, ok: bool, details: str = ""):
    mark = f"{C.OKGREEN}PASSED{C.RESET}" if ok else f"{C.FAIL}FAILED{C.RESET}"
    print(f"{C.BOLD}{name:<30}{C.RESET} {mark}")
    if details:
        for line in details.splitlines():
            print(f"   {line}")
    print()


# --------------------------------------------------------------
# 1. Basic Python sanity
# --------------------------------------------------------------
def test_python():
    ver = sys.version_info
    ver_str = f"{ver.major}.{ver.minor}.{ver.micro}"
    ok = ver.major == 3 and ver.minor >= 9
    details = f"Python {ver_str} ({platform.python_implementation()})"
    result("Python version", ok, details)
    return ok


# --------------------------------------------------------------
# 2. NumPy
# --------------------------------------------------------------
def test_numpy():
    try:
        import numpy as np
        a = np.arange(1, 6)
        mean = a.mean()
        details = f"Array [1 2 3 4 5] → mean = {mean}"
        ok = True
    except Exception as e:
        details = f"Import/error: {e}"
        ok = False
    result("NumPy", ok, details)
    return ok


# --------------------------------------------------------------
# 3. Pandas (with fake data)
# --------------------------------------------------------------
def test_pandas():
    try:
        import pandas as pd
        data = {
            "id":   [101, 102, 103, 104, 105],
            "name": ["Alice", "Bob", "Charlie", "Dana", "Eve"],
            "score": [87.5, 92.0, 78.3, 95.1, 88.9],
            "city": ["Dar es Salaam", "Arusha", "Mwanza", "Dodoma", "Zanzibar"]
        }
        df = pd.DataFrame(data)

        # simple ops
        total = df["score"].sum()
        avg   = df["score"].mean()
        top   = df.loc[df["score"].idxmax(), "name"]

        details = (
            f"DataFrame created ({len(df)} rows)\n"
            f"   Sum of scores : {total}\n"
            f"   Avg score     : {avg:.2f}\n"
            f"   Top performer : {top}"
        )
        ok = True
    except Exception as e:
        details = f"Import/error: {e}"
        ok = False
    result("Pandas", ok, details)
    return ok


# --------------------------------------------------------------
# 4. pyodbc – fake ODBC source (CSV driver)
# --------------------------------------------------------------
def test_pyodbc():
    try:
        import pyodbc

        # 1. Create a tiny CSV file that acts as a "table"
        csv_content = """id,name,score,city
101,Alice,87.5,"Dar es Salaam"
102,Bob,92.0,Arusha
103,Charlie,78.3,Mwanza
104,Dana,95.1,Dodoma
105,Eve,88.9,Zanzibar
"""
        tmp_dir = tempfile.mkdtemp(prefix="pyodbc_test_")
        csv_path = Path(tmp_dir) / "sample.csv"
        csv_path.write_text(csv_content, encoding="utf-8")

        # 2. Register a temporary ODBC DSN (Microsoft Text Driver)
        dsn_name = "TMP_CSV_TEST"
        driver = "{Microsoft Access Text Driver (*.txt, *.csv)}"
        conn_str = (
            f"Driver={driver};"
            f"DBQ={tmp_dir};"
            f"Extensions=asc,csv,tab,txt;"
            f"DefaultDir={tmp_dir};"
        )

        # 3. Connect & query
        conn = pyodbc.connect(conn_str, autocommit=True)
        cursor = conn.cursor()
        cursor.execute(f"SELECT * FROM [{csv_path.name}]")
        rows = cursor.fetchall()

        # 4. Clean up
        conn.close()
        # (optional) remove DSN – not needed for temp test

        details = (
            f"Connected via Text ODBC driver\n"
            f"   CSV file : {csv_path}\n"
            f"   Rows read: {len(rows)}\n"
            f"   Sample   : {rows[0] if rows else '—'}"
        )
        ok = True
    except Exception as e:
        details = f"Import/connection error: {e}"
        ok = False
    finally:
        # always clean the temp folder
        try:
            import shutil
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except:
            pass

    result("pyodbc (CSV driver)", ok, details)
    return ok


# --------------------------------------------------------------
# 5. Final summary
# --------------------------------------------------------------
def summary(passed: list, total: int):
    banner("SUMMARY")
    print(f"{C.BOLD}Total tests : {total}{C.RESET}")
    print(f"{C.OKGREEN}{C.BOLD}Passed      : {len(passed)}{C.RESET}")
    print(f"{C.FAIL}{C.BOLD}Failed      : {total - len(passed)}{C.RESET}\n")

    if len(passed) == total:
        print(f"{C.OKGREEN}{C.BOLD}All critical libraries are WORKING!{C.RESET}")
    else:
        print(f"{C.WARNING}{C.BOLD}Some libraries need attention – see details above.{C.RESET}")


# --------------------------------------------------------------
# MAIN
# --------------------------------------------------------------
def main():
    banner("PyInstaller-Packaged Python Tester")
    print(f"{C.OKBLUE}Running on {platform.system()} {platform.release()}{C.RESET}\n")

    tests = [
        ("Python", test_python),
        ("NumPy",  test_numpy),
        ("Pandas", test_pandas),
        ("pyodbc", test_pyodbc),
    ]

    passed = []
    for name, func in tests:
        if func():
            passed.append(name)

    summary(passed, len(tests))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n{C.WARNING}Interrupted by user.{C.RESET}")
        sys.exit(1)
    except Exception as exc:
        print(f"\n{C.FAIL}{C.BOLD}UNHANDLED EXCEPTION:{C.RESET} {exc}")
        import traceback
        traceback.print_exc()
        sys.exit(1)