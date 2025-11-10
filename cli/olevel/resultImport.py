# resultImport.py
# FINAL — MARKS FROM COLUMN 7 (Python: index 6) — 100% VBA

import sys
import pyodbc
import pandas as pd
from pathlib import Path
from rich.console import Console
from rich.table import Table
from rich.progress import Progress
from rich.prompt import Prompt

console = Console()

# ================================
# CONFIG
# ================================
DB_PATH = r"C:\Kiyabo App\backend\Kiyabo App Backend v2.0.0.accdb"
DRIVER = "{Microsoft Access Driver (*.mdb, *.accdb)}"
RESULTS_TABLE = "tbl_student_exam_results"
SUBJECTS_TABLE = "tbl_school_subjects"

EXCEL_FILE = Path(r"C:\Kiyabo App\exam templates\Form_IV_Exam_Template_20251025 155902.xlsx")
EXAM_ID = "MID420251027"


# ================================
# 100% VBA LOGIC — MARKS FROM COLUMN 7 (Python: 6)
# ================================
def get_connection():
    conn_str = f"DRIVER={DRIVER};DBQ={DB_PATH};"
    try:
        conn = pyodbc.connect(conn_str, autocommit=False)
        console.print("[green]Connected to Access[/green]")
        return conn
    except Exception as e:
        console.print(f"[red]DB Error:[/red] {e}")
        sys.exit(1)


def has_records(conn, exam_id: str) -> bool:
    cursor = conn.cursor()
    sql = f"SELECT COUNT(*) FROM [{RESULTS_TABLE}] WHERE exam_id = ?"
    return cursor.execute(sql, (exam_id,)).fetchone()[0] > 0


def delete_existing_records(conn, exam_id: str):
    cursor = conn.cursor()
    cursor.execute(f"DELETE FROM [{RESULTS_TABLE}] WHERE exam_id = ?", (exam_id,))
    cursor.execute(f"DELETE FROM [{RESULTS_TABLE}_special] WHERE exam_id = ?", (exam_id,))
    conn.commit()
    console.print(f"[green]Deleted existing[/green]")


def get_class_from_exam_id(exam_id: str) -> str:
    digit = exam_id[3]
    return {'1': 'I', '2': 'II', '3': 'III', '4': 'IV', '8': 'PC', '0': 'PRE'}.get(digit, "")


def school_subjects_count(conn, class_id: str) -> int:
    cursor = conn.cursor()
    sql = f"SELECT COUNT(*) FROM [{SUBJECTS_TABLE}] WHERE [is_present_{class_id}] = True"
    try:
        return int(cursor.execute(sql).fetchone()[0])
    except:
        return 0


def load_excel(file_path: Path) -> pd.DataFrame:
    if not file_path.exists():
        console.print(f"[red]FILE NOT FOUND[/red]")
        sys.exit(1)
    console.print(f"[yellow]Loading:[/yellow] {file_path.name}")
    df = pd.read_excel(file_path, sheet_name=0, header=None, engine="openpyxl")
    console.print(f"[green]Loaded:[/green] {len(df)} rows × {len(df.columns)} cols")
    return df


def build_column_to_field_map(subject_count: int) -> dict:
    mapping = {}
    col = 6  # ← EXCEL COLUMN 7 → Python index 6
    base = ["CIV", "HIS", "GEO", "KIS", "ENG", "PHY", "CHE", "BIO", "MAT"]
    extra = ["EDK", "ICS"] + [f"SUB{i}" for i in range(12, 21)]
    for i in range(subject_count):
        field = base[i] if i < 9 else extra[i - 9]
        mapping[col] = field
        col += 2  # SKIP GRADE COLUMN
    return mapping


def preview_data(df: pd.DataFrame, subject_map: dict):
    start_row = 13
    end_row = min(start_row + 20, len(df))
    sample = df.iloc[start_row:end_row]

    table = Table(title=f"PREVIEW: MARKS FROM COL 7 (Python: 6) → Rows {start_row+1}–{end_row}", show_header=True, expand=True)
    table.add_column("Row", style="dim")
    table.add_column("ID", style="bold yellow")
    table.add_column("First", style="cyan")
    table.add_column("Middle", style="cyan")
    table.add_column("Surname", style="cyan")
    table.add_column("Sex", style="magenta")
    for col in subject_map:
        table.add_column(subject_map[col], justify="center", style="bold green")

    for idx, (_, row) in enumerate(sample.iterrows(), start=start_row + 1):
        values = [
            str(idx),
            str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else "",
            str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else "",
            str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else "",
            str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else "",
            str(row.iloc[5]).strip() if pd.notna(row.iloc[5]) else "",
        ]
        for col in subject_map:
            val = row.iloc[col] if col < len(row) else ""
            val_str = str(val).strip() if pd.notna(val) and str(val).strip() else ""
            if val_str and val_str.replace(".", "").replace("-", "").isdigit():
                values.append(val_str)
            else:
                values.append("—")
        table.add_row(*values)

    console.print(table)


# ================================
# MAIN
# ================================
def main():
    console.print("[bold magenta]Kiyabo Importer – MARKS FROM COLUMN 7 (Python: 6)[/bold magenta]\n")
    console.print(f"[cyan]Exam ID:[/cyan] {EXAM_ID}")
    console.print(f"[cyan]File:[/cyan] {EXCEL_FILE}\n")

    conn = get_connection()

    if has_records(conn, EXAM_ID):
        if Prompt.ask("Delete existing?", default="y") == "y":
            delete_existing_records(conn, EXAM_ID)
        else:
            conn.close()
            return

    class_id = get_class_from_exam_id(EXAM_ID)
    subject_count = school_subjects_count(conn, class_id)
    if subject_count == 0 or subject_count > 20:
        console.print(f"[red]Invalid subjects: {subject_count}[/red]")
        conn.close()
        return

    df = load_excel(EXCEL_FILE)
    subject_map = build_column_to_field_map(subject_count)
    console.print(f"[cyan]Reading MARKS from Python indices:[/cyan] {', '.join([f'col {col}' for col in subject_map.keys()])}")
    console.print(f"[cyan]→ Excel columns:[/cyan] {', '.join([f'C{col+1}' for col in subject_map.keys()])}")

    preview_data(df, subject_map)

    if Prompt.ask("\nProceed?", default="y") != "y":
        conn.close()
        return

    records = []
    start_row = 13
    with Progress(console=console) as p:
        task = p.add_task("Importing...", total=len(df) - start_row)
        for idx in range(start_row, len(df)):
            row = df.iloc[idx]
            sid = row.iloc[1]
            if pd.isna(sid) or str(sid).strip() == "": break
            student_id = str(sid).strip()
            record = {"exam_id": EXAM_ID, "student_id": student_id}
            for col, field in subject_map.items():
                if col < len(row):
                    val = row.iloc[col]
                    if pd.notna(val):
                        val_str = str(val).strip()
                        if val_str and val_str.replace(".", "").replace("-", "").isdigit():
                            record[field] = float(val_str) if "." in val_str else int(val_str)
                        else:
                            record[field] = None
            records.append(record)
            p.update(task, advance=1)

    cursor = conn.cursor()
    cols = ["exam_id", "student_id"] + list(subject_map.values())
    col_str = ", ".join(f"[{c}]" for c in cols)
    placeholders = ", ".join("?" for _ in cols)
    sql = f"INSERT INTO [{RESULTS_TABLE}] ({col_str}) VALUES ({placeholders})"
    data = [[r["exam_id"], r["student_id"]] + [r.get(f, None) for f in subject_map.values()] for r in records]
    console.print(f"[yellow]Inserting {len(data)}...[/yellow]")
    cursor.executemany(sql, data)
    conn.commit()
    console.print(f"[bold green]SUCCESS: {len(data)} inserted![/bold green]")

    conn.close()
    console.print("\n[bold green]CASE SOLVED! YOU ARE THE MASTER.[/bold green]")


if __name__ == "__main__":
    main()