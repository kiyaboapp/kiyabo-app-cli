#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
KIYABO STUDENT IMPORTER v3.0
Usage:
    py insert.py <CLASS_ID> [EXCEL_PATH] [--db DB_PATH]

Examples:
    py insert.py I
    py insert.py II "C:\admission\Form II Upload.xlsx"
    py insert.py III --db "C:\backup\Kiyabo.accdb"
"""

import pyodbc
import pandas as pd
from pathlib import Path
from datetime import datetime
import re
import openpyxl
from openpyxl.styles import PatternFill
import time
import sys
import argparse
from tkinter import filedialog, Tk
from rich.console import Console
from rich.table import Table
from rich.progress import Progress, SpinnerColumn, BarColumn, TimeElapsedColumn,TextColumn
from rich import box
from rich.panel import Panel

# ===================================================================
# DEFAULTS
# ===================================================================
DEFAULT_DB_PATH = r"C:\Kiyabo App\backend\Kiyabo App Backend v2.0.0.accdb"
SAVE_FOLDER = Path(r"C:\Kiyabo App\admission")
SAVE_FOLDER.mkdir(parents=True, exist_ok=True)

GREEN_FILL = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
RED_FILL   = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

console = Console()

# ===================================================================
# ARGUMENT PARSER
# ===================================================================
def parse_args():
    parser = argparse.ArgumentParser(
        description="Kiyabo Student Importer v3.0",
        epilog="Example: py insert.py I \"C:\\admission\\Form I Upload.xlsx\""
    )
    parser.add_argument("class_id", help="Class ID (I, II, III, IV, V, VI, PC)")
    parser.add_argument("excel_path", nargs="?", default=None, help="Path to Excel file (optional)")
    parser.add_argument("--db", default=DEFAULT_DB_PATH, help=f"Database path (default: {DEFAULT_DB_PATH})")
    return parser.parse_args()

# ===================================================================
# FILE DIALOG
# ===================================================================
def select_excel_file():
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    file_path = filedialog.askopenfilename(
        title="Select Excel File to Upload",
        filetypes=[("Excel Files", "*.xlsx *.xlsm")]
    )
    root.destroy()
    if not file_path:
        console.print("[red]No file selected. Exiting.[/red]")
        sys.exit(1)
    return Path(file_path)

# ===================================================================
# HELPERS
# ===================================================================
def get_class_number(class_id: str) -> int:
    mapping = {"I": 0, "II": 1, "III": 2, "IV": 3, "V": 5, "VI": 6, "PC": 7}
    num = mapping.get(class_id.upper(), -1)
    if num == -1:
        console.print(f"[red]Invalid Class ID: {class_id}[/red]")
        sys.exit(1)
    return num + 1

def get_check_digit(pseudo: str) -> str:
    total = 0
    for i, ch in enumerate(reversed(pseudo)):
        n = int(ch)
        total += n if i % 2 == 0 else (n * 2 if n * 2 <= 9 else n * 2 - 9)
    return str((10 - total % 10) % 10)

def escape_single_quote(s) -> str:
    if pd.isna(s): return ""
    return str(s).strip().replace("'", "''")

def clean_phone(phone) -> str:
    if pd.isna(phone): return ""
    digits = re.sub(r"\D", "", str(phone))
    return digits[-9:] if len(digits) > 9 else digits

def name_exists(cur, first, middle, surname) -> bool:
    if not first or not middle or not surname: return False
    sql = "SELECT COUNT(*) FROM [tbl_student_academic_info] WHERE first_name=? AND middle_name=? AND surname=?"
    cur.execute(sql, first, middle, surname)
    return cur.fetchone()[0] > 0

def get_last_student_id_once(cur, class_number) -> str:
    year_part = datetime.now().year - class_number if class_number == 0 else datetime.now().year + 1 - class_number
    criteria = f"{year_part}{class_number}"
    sql = "SELECT TOP 1 student_id FROM [tbl_student_academic_info] WHERE reg_year_criteria=? ORDER BY student_id DESC"
    cur.execute(sql, criteria)
    row = cur.fetchone()
    return row[0] if row else "0"

def generate_student_ids_in_memory(last_id: str, count: int, class_number: int) -> list:
    start_serial = 1 if last_id == "0" else int(last_id[5:9]) + 1
    ids = []
    for serial in range(start_serial, start_serial + count):
        year_part = datetime.now().year - class_number if class_number == 0 else datetime.now().year + 1 - class_number
        pseudo = f"{year_part}{class_number}{serial:04d}"
        ids.append(pseudo + get_check_digit(pseudo))
    return ids

# ===================================================================
# MAIN
# ===================================================================
def main():
    start_time = time.time()
    args = parse_args()

    CLASS_ID = args.class_id.strip().upper()
    DB_PATH = Path(args.db)
    EXCEL_PATH = Path(args.excel_path) if args.excel_path else None

    # === VALIDATE DB ===
    if not DB_PATH.exists():
        console.print(f"[red]Database not found: {DB_PATH}[/red]")
        sys.exit(1)

    # === GET EXCEL PATH ===
    if not EXCEL_PATH:
        console.print("[bold cyan]Opening file selector...[/bold cyan]")
        EXCEL_PATH = select_excel_file()
    elif not EXCEL_PATH.exists():
        console.print(f"[red]Excel file not found: {EXCEL_PATH}[/red]")
        sys.exit(1)

    CONN_STR = f"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};Dbq={DB_PATH};"

    console.rule(f"[bold magenta]KIYABO IMPORTER – CLASS {CLASS_ID}[/bold magenta]", style="magenta")
    console.print(f"[cyan]DB:[/cyan]  {DB_PATH.name}")
    console.print(f"[cyan]XLS:[/cyan] {EXCEL_PATH.name}\n")

    # === READ EXCEL ===
    try:
        df = pd.read_excel(EXCEL_PATH, header=0, dtype=str)
    except Exception as e:
        console.print(f"[red]Failed to read Excel: {e}[/red]")
        sys.exit(1)

    console.print(Panel("[bold yellow]EXCEL PREVIEW (first 5 rows)[/bold yellow]"))
    if len(df) > 0:
        preview = df[["FIRST NAME", "MIDDLE NAME", "SURNAME", "SEX", "CLASS ID"]].head(5).copy()
        preview.columns = ["FIRST", "MIDDLE", "SURNAME", "SEX", "CLASS"]
        console.print(preview.to_string(index=False), "\n")
    else:
        console.print("[red]Excel is empty.[/red]")
        sys.exit(1)

    # === VALIDATE ROWS ===
    valid_rows = []
    for idx in df.index:
        r = df.loc[idx]
        if pd.isna(r["FIRST NAME"]) or pd.isna(r["MIDDLE NAME"]) or pd.isna(r["SURNAME"]) or pd.isna(r["SEX"]) or pd.isna(r["CLASS ID"]):
            break
        if str(r["CLASS ID"]).strip().upper() != CLASS_ID:
            console.print(f"[yellow]Warning: Row {idx+2} has Class '{r['CLASS ID']}' ≠ '{CLASS_ID}'[/yellow]")
            ans = console.input("   [bold]Continue up to this row? (y/n): [/bold]")
            if ans.lower() != 'y':
                console.print("[red]Aborted.[/red]")
                sys.exit(0)
            break
        valid_rows.append(idx)

    if not valid_rows:
        console.print("[red]No valid students found.[/red]")
        sys.exit(1)

    console.print(f"[bold green]Ready: {len(valid_rows)} students to import.[/bold green]\n")

    # === CONNECT DB ===
    try:
        conn = pyodbc.connect(CONN_STR, autocommit=False)
        cur = conn.cursor()
    except Exception as e:
        console.print(f"[red]Database connection failed: {e}[/red]")
        sys.exit(1)

    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active

    class_number = get_class_number(CLASS_ID)
    last_id = get_last_student_id_once(cur, class_number)
    console.print(f"[bold cyan]Last Student ID in DB: {last_id}[/bold cyan]\n")
    student_ids = generate_student_ids_in_memory(last_id, len(valid_rows), class_number)

    # === STATS ===
    stats = {"total": len(valid_rows), "inserted": 0, "skipped": 0, "skip_reasons": {"duplicate_name": 0}}
    student_records = []
    admission_records = []
    family_records = []

    family_keys = [
        "student_id", "date_of_birth", "father_name", "father_occupation", "father_phone", "father_phone_alternative",
        "mother_name", "mother_occupation", "mother_phone", "mother_phone_alternative",
        "gurdian_name", "gurdian_occupation", "gurdian_relationship", "gurdian_phone", "gurdian_phone_alternative",
        "parent_name", "parent_occupation", "parent_phone", "parent_phone_alternative", "parent_relationship"
    ]

    table = Table(title="Processing Students", box=box.DOUBLE, show_header=True, header_style="bold magenta")
    table.add_column("SN", width=4, style="cyan")
    table.add_column("ID", width=12, style="green")
    table.add_column("NAME", width=28, style="white")
    table.add_column("SEX", width=5, style="yellow")
    table.add_column("CLASS", width=6, style="blue")
    table.add_column("SECTION", width=8, style="dim")
    table.add_column("STATUS", width=10, style="bold")

    serial = 1

    # === PROCESSING WITH PROGRESS BAR ===
    with Progress(
        SpinnerColumn(spinner_name="bouncingBall"),
        BarColumn(bar_width=40),
        "[progress.percentage]{task.percentage:>3.0f}%",
        TimeElapsedColumn(),
        console=console
    ) as progress:
        task = progress.add_task("Processing", total=len(valid_rows))

        for i, excel_row in enumerate(valid_rows):
            r = df.loc[excel_row]
            student_id = student_ids[i]

            first   = escape_single_quote(r["FIRST NAME"]).upper()
            middle  = escape_single_quote(r["MIDDLE NAME"]).upper()
            surname = escape_single_quote(r["SURNAME"]).upper()
            sex     = str(r["SEX"]).upper()
            inactive = -1 if str(r["INACTIVE"]).strip().upper() == "YES" else 0
            section = str(r["SECTION"]) if pd.notna(r["SECTION"]) and str(r["SECTION"]).strip() else ""
            prem_no = str(r["PREM NO"]) if pd.notna(r["PREM NO"]) and str(r["PREM NO"]).strip() else None

            full_name = f"{first} {middle} {surname}"

            # === SKIP CHECK ===
            skip_reason = None
            if name_exists(cur, first.replace("''", "'"), middle.replace("''", "'"), surname.replace("''", "'")):
                skip_reason = "Duplicate Name"
                stats["skipped"] += 1
                stats["skip_reasons"]["duplicate_name"] += 1
                for c in range(2, 10): ws.cell(row=excel_row+2, column=c).fill = RED_FILL
            else:
                stats["inserted"] += 1
                ws.cell(row=excel_row+2, column=21).value = student_id
                for c in range(2, 10): ws.cell(row=excel_row+2, column=c).fill = GREEN_FILL

                section_id = CLASS_ID + section[-1].upper() if section else None
                student_records.append((
                    student_id, first.replace("''", "'"), middle.replace("''", "'"), surname.replace("''", "'"),
                    sex, 1, inactive, section_id, prem_no
                ))

                # ADMISSION
                admn_no = str(r["ADMISSION NO"]) if pd.notna(r["ADMISSION NO"]) and str(r["ADMISSION NO"]).strip() else None
                entrance_mode = escape_single_quote(r["ENTRANCE MODE"]).upper() if pd.notna(r["ENTRANCE MODE"]) and str(r["ENTRANCE MODE"]).strip() else "DIRECT"
                former_school = escape_single_quote(r["FORMER SCHOOL"]).upper().replace("''", "'") if pd.notna(r["FORMER SCHOOL"]) and str(r["FORMER SCHOOL"]).strip() else None
                admission_records.append((student_id, admn_no, entrance_mode, former_school))
                for c in range(10, 13): ws.cell(row=excel_row+2, column=c).fill = GREEN_FILL

                # FAMILY
                dob = r["DATE OF BIRTH"] if pd.notna(r["DATE OF BIRTH"]) else None
                parent_type = str(r["RELATIONSHIP"]).strip().upper() if pd.notna(r["RELATIONSHIP"]) else ""
                parent_name = escape_single_quote(r["PARENT NAME"])
                occupation = escape_single_quote(r["OCCUPATION"])
                phone1 = clean_phone(r["PHONE NUMBER"])
                phone2 = clean_phone(r["ALTERNATIVE PHONE"])
                gurdian_rel = escape_single_quote(r["IF GUARDIAN SPECIFY"])

                fam = {k: None for k in family_keys}
                fam["student_id"] = student_id
                fam["date_of_birth"] = dob

                if parent_type == "FATHER":
                    fam["father_name"] = parent_name or f"{middle} {surname}"
                    fam["father_occupation"] = occupation
                    fam["father_phone"] = phone1
                    fam["father_phone_alternative"] = phone2
                    fam["parent_name"] = fam["father_name"]
                    fam["parent_occupation"] = fam["father_occupation"]
                    fam["parent_phone"] = fam["father_phone"]
                    fam["parent_phone_alternative"] = fam["father_phone_alternative"]
                    fam["parent_relationship"] = "FATHER"
                elif parent_type == "MOTHER":
                    fam["mother_name"] = parent_name
                    fam["mother_occupation"] = occupation
                    fam["mother_phone"] = phone1
                    fam["mother_phone_alternative"] = phone2
                    fam["parent_name"] = fam["mother_name"]
                    fam["parent_occupation"] = fam["mother_occupation"]
                    fam["parent_phone"] = fam["mother_phone"]
                    fam["parent_phone_alternative"] = fam["mother_phone_alternative"]
                    fam["parent_relationship"] = "MOTHER"
                elif parent_type == "GUARDIAN":
                    fam["gurdian_name"] = parent_name
                    fam["gurdian_occupation"] = occupation
                    fam["gurdian_relationship"] = gurdian_rel or "GUARDIAN"
                    fam["gurdian_phone"] = phone1
                    fam["gurdian_phone_alternative"] = phone2
                    fam["parent_name"] = fam["gurdian_name"]
                    fam["parent_occupation"] = fam["gurdian_occupation"]
                    fam["parent_phone"] = fam["gurdian_phone"]
                    fam["parent_phone_alternative"] = fam["gurdian_phone_alternative"]
                    fam["parent_relationship"] = fam["gurdian_relationship"]
                elif phone1 or phone2:
                    fam["gurdian_name"] = parent_name
                    fam["gurdian_occupation"] = occupation
                    fam["gurdian_relationship"] = "GUARDIAN"
                    fam["gurdian_phone"] = phone1
                    fam["gurdian_phone_alternative"] = phone2
                    fam["parent_name"] = fam["gurdian_name"]
                    fam["parent_occupation"] = fam["gurdian_occupation"]
                    fam["parent_phone"] = fam["gurdian_phone"]
                    fam["parent_phone_alternative"] = fam["gurdian_phone_alternative"]
                    fam["parent_relationship"] = "GUARDIAN"

                family_records.append([fam[k] for k in family_keys])
                for c in range(13, 20): ws.cell(row=excel_row+2, column=c).fill = GREEN_FILL

            # === TABLE ROW ===
            status = "[red]SKIPPED[/red]" if skip_reason else "[green]INSERTED[/green]"
            table.add_row(
                str(serial),
                student_id if not skip_reason else "—",
                full_name[:27] + ("…" if len(full_name) > 27 else ""),
                sex,
                CLASS_ID,
                section or "—",
                status
            )
            serial += 1
            progress.advance(task)

    console.print(table)

    # === INSERT INTO DB ===
    with Progress(
        SpinnerColumn(spinner_name="dots12"),
        TextColumn("[bold green]Inserting into database...[/bold green]"),
        console=console
    ) as progress:
        task1 = progress.add_task("Students", total=1)
        task2 = progress.add_task("Admission", total=1)
        task3 = progress.add_task("Family", total=1)

        if student_records:
            sql = """INSERT INTO [tbl_student_academic_info] 
                     (student_id, first_name, middle_name, surname, sex, is_boarding, inactive, section_id, prem_no) 
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)"""
            cur.executemany(sql, student_records)
            conn.commit()
            progress.advance(task1)

        if admission_records:
            sql = """INSERT INTO [tbl_student_admission] 
                     (student_id, admn_no, entrance_mode, former_school) 
                     VALUES (?, ?, ?, ?)"""
            cur.executemany(sql, admission_records)
            conn.commit()
            progress.advance(task2)

        if family_records:
            used = [c for i, c in enumerate(family_keys) if c in ("student_id", "date_of_birth") or any(rec[i] for rec in family_records)]
            sql = f"INSERT INTO [tbl_student_family_info] ({', '.join(f'[{c}]' for c in used)}) VALUES ({', '.join('?' * len(used))})"
            data = [[r[family_keys.index(c)] for c in used] for r in family_records]
            cur.executemany(sql, data)
            conn.commit()
            progress.advance(task3)

    # === SAVE EXCEL ===
    timestamp = datetime.now().strftime("%d%b%Y %H%M%S")
    new_file = SAVE_FOLDER / f"UPLOADED Form {CLASS_ID} {timestamp}.xlsx"
    wb.save(new_file)

    # === FINAL SUMMARY ===
    total_time = time.time() - start_time
    mins, secs = divmod(total_time, 60)

    summary = Table(title="IMPORT SUMMARY", box=box.ROUNDED, show_header=True)
    summary.add_column("Metric", style="bold cyan")
    summary.add_column("Value", style="bold green")
    summary.add_row("Total Students", str(stats["total"]))
    summary.add_row("Inserted", str(stats["inserted"]))
    summary.add_row("Skipped", str(stats["skipped"]))
    summary.add_row("  • Duplicate Name", str(stats["skip_reasons"]["duplicate_name"]))
    summary.add_row("Admission Records", str(len(admission_records)))
    summary.add_row("Family Records", str(len(family_records)))
    summary.add_row("Time Taken", f"{int(mins)} min {secs:.1f} sec")
    summary.add_row("Saved File", str(new_file))

    console.rule("[bold green]IMPORT COMPLETED[/bold green]")
    console.print(summary)

    conn.close()

# ===================================================================
# RUN
# ===================================================================
if __name__ == "__main__":
    main()