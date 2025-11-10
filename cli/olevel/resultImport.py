# resultImport.py
# FINAL — MARKS FROM COLUMN 7 (Python: index 6) — 100% VBA
# BEAUTIFIED • CLASS-BASED • CONFIGURABLE • RICH OUTPUT

import sys
from pathlib import Path
import pyodbc
import pandas as pd
from typing import Union  # ← ADD THIS
from rich.console import Console
from rich.table import Table
from rich.progress import Progress
from rich.prompt import Confirm
from rich.panel import Panel
from rich import box

console = Console()


class OlevelResultImporter:
    """
    O-Level Exam Results Importer
    Imports marks from Excel (column 7+) → Access DB
    100% VBA logic | Rich UI | Fully Configurable
    """

    DRIVER = "{Microsoft Access Driver (*.mdb, *.accdb)}"
    RESULTS_TABLE = "tbl_student_exam_results"
    SUBJECTS_TABLE = "tbl_school_subjects"

    def __init__(
        self,
        exam_id: str,
        excel_file: Union[Path, str],  # ← FIXED: Union instead of |
        db_path: Union[Path, str],     # ← FIXED
        force_import: bool = False,
        show_preview: bool = True,
        start_row: int = 13,
        max_preview_rows: int = 20
    ):
        self.exam_id = exam_id
        self.excel_file = Path(excel_file)  # ← Convert str → Path
        self.db_path = Path(db_path)        # ← Convert str → Path
        self.force_import = force_import
        self.show_preview = show_preview
        self.start_row = start_row
        self.max_preview_rows = max_preview_rows

        self.conn = None
        self.cursor = None
        self.df = None
        self.subject_map = {}

    # ================================
    # CONNECTION
    # ================================
    def _connect(self):
        conn_str = f"DRIVER={self.DRIVER};DBQ={self.db_path};"
        try:
            self.conn = pyodbc.connect(conn_str, autocommit=False)
            self.cursor = self.conn.cursor()
            console.print(Panel(
                f"[bold green]Connected to Database[/bold green]\n{self.db_path.name}",
                style="green", box=box.ROUNDED
            ))
        except Exception as e:
            console.print(Panel(f"[bold red]Connection Failed[/bold red]\n{e}", style="red"))
            sys.exit(1)

    # ================================
    # CHECK & DELETE
    # ================================
    def _has_records(self) -> bool:
        sql = f"SELECT COUNT(*) FROM [{self.RESULTS_TABLE}] WHERE exam_id = ?"
        count = self.cursor.execute(sql, (self.exam_id,)).fetchone()[0]
        return count > 0

    def _delete_existing(self):
        self.cursor.execute(f"DELETE FROM [{self.RESULTS_TABLE}] WHERE exam_id = ?", (self.exam_id,))
        self.cursor.execute(f"DELETE FROM [{self.RESULTS_TABLE}_special] WHERE exam_id = ?", (self.exam_id,))
        deleted = self.cursor.rowcount
        self.conn.commit()
        console.print(Panel(
            f"[bold green]Deleted {deleted} existing records[/bold green]",
            style="green"
        ))

    # ================================
    # CLASS ID FROM EXAM ID
    # ================================
    def _get_class_id(self) -> str:
        digit = self.exam_id[3]
        mapping = {'1': 'I', '2': 'II', '3': 'III', '4': 'IV', '8': 'PC', '0': 'PRE'}
        class_id = mapping.get(digit, "")
        if not class_id:
            console.print(Panel("[bold red]Invalid exam_id format[/bold red]", style="red"))
            sys.exit(1)
        return class_id

    # ================================
    # SUBJECT COUNT
    # ================================
    def _get_subject_count(self, class_id: str) -> int:
        sql = f"SELECT COUNT(*) FROM [{self.SUBJECTS_TABLE}] WHERE [is_present_{class_id}] = True"
        try:
            count = int(self.cursor.execute(sql).fetchone()[0])
            if count == 0 or count > 20:
                console.print(Panel(f"[bold red]Invalid subject count: {count}[/bold red]", style="red"))
                sys.exit(1)
            return count
        except Exception as e:
            console.print(f"[yellow]Warning: {e}[/yellow]")
            return 0

    # ================================
    # LOAD EXCEL
    # ================================
    def _load_excel(self) -> pd.DataFrame:
        if not self.excel_file.exists():
            console.print(Panel(f"[bold red]File Not Found[/bold red]\n{self.excel_file}", style="red"))
            sys.exit(1)

        console.print(f"[bold cyan]Loading Excel...[/bold cyan]")
        df = pd.read_excel(self.excel_file, sheet_name=0, header=None, engine="openpyxl")
        console.print(Panel(
            f"[bold green]Loaded:[/bold green] {len(df):,} rows × {len(df.columns)} cols",
            style="green"
        ))
        return df

    # ================================
    # BUILD COLUMN → FIELD MAP
    # ================================
    def _build_subject_map(self, subject_count: int) -> dict:
        mapping = {}
        col = 6  # ← COLUMN 7 (Python index 6)
        base = ["CIV", "HIS", "GEO", "KIS", "ENG", "PHY", "CHE", "BIO", "MAT"]
        extra = ["EDK", "ICS"] + [f"SUB{i}" for i in range(12, 21)]
        for i in range(subject_count):
            field = base[i] if i < 9 else extra[i - 9]
            mapping[col] = field.lower()  # Access uses lowercase
            col += 2  # Skip grade column
        return mapping

    # ================================
    # PREVIEW DATA
    # ================================
    def _preview_data(self, df: pd.DataFrame):
        if not self.show_preview:
            return

        start = self.start_row
        end = min(start + self.max_preview_rows, len(df))
        sample = df.iloc[start:end]

        table = Table(
            title=f"[bold yellow]PREVIEW: Rows {start+1}–{end} | Marks from Column 7+[/bold yellow]",
            box=box.DOUBLE,
            show_header=True,
            header_style="bold magenta",
            expand=True
        )
        table.add_column("Row", style="dim", width=5)
        table.add_column("ID", style="bold yellow")
        table.add_column("First", style="cyan")
        table.add_column("Middle", style="cyan")
        table.add_column("Surname", style="cyan")
        table.add_column("Sex", style="magenta")

        for field in self.subject_map.values():
            table.add_column(field.upper(), justify="center", style="bold green")

        for idx, (_, row) in enumerate(sample.iterrows(), start=start + 1):
            values = [
                str(idx),
                str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else "",
                str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else "",
                str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else "",
                str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else "",
                str(row.iloc[5]).strip() if pd.notna(row.iloc[5]) else "",
            ]
            for col in self.subject_map:
                val = row.iloc[col] if col < len(row) else ""
                val_str = str(val).strip() if pd.notna(val) and str(val).strip() else "—"
                if val_str.replace(".", "").replace("-", "").isdigit():
                    values.append(f"[green]{val_str}[/green]")
                else:
                    values.append(f"[dim]{val_str}[/dim]")
            table.add_row(*values)

        console.print(table)

    # ================================
    # IMPORT TO DB
    # ================================
    def _import_to_db(self, df: pd.DataFrame):
        records = []
        with Progress(console=console) as progress:
            task = progress.add_task("[cyan]Importing students...", total=len(df) - self.start_row)

            for idx in range(self.start_row, len(df)):
                row = df.iloc[idx]
                sid = row.iloc[1]
                if pd.isna(sid) or str(sid).strip() == "":
                    break

                student_id = str(sid).strip()
                record = {"exam_id": self.exam_id, "student_id": student_id}

                for col, field in self.subject_map.items():
                    if col < len(row):
                        val = row.iloc[col]
                        if pd.notna(val):
                            val_str = str(val).strip()
                            if val_str.replace(".", "").replace("-", "").isdigit():
                                record[field] = float(val_str) if "." in val_str else int(val_str)
                            else:
                                record[field] = None
                        else:
                            record[field] = None
                    else:
                        record[field] = None

                records.append(record)
                progress.update(task, advance=1)

        if not records:
            console.print(Panel("[bold red]No valid records to import[/bold red]", style="red"))
            return

        cols = ["exam_id", "student_id"] + list(self.subject_map.values())
        col_str = ", ".join(f"[{c}]" for c in cols)
        placeholders = ", ".join("?" for _ in cols)
        sql = f"INSERT INTO [{self.RESULTS_TABLE}] ({col_str}) VALUES ({placeholders})"

        data = [[r["exam_id"], r["student_id"]] + [r.get(f, None) for f in self.subject_map.values()] for r in records]

        console.print(f"[bold yellow]Inserting {len(data):,} records...[/bold yellow]")
        self.cursor.executemany(sql, data)
        self.conn.commit()
        console.print(Panel(
            f"[bold green]SUCCESS: {len(data):,} records imported![/bold green]",
            style="green", box=box.ROUNDED
        ))

    # ================================
    # MAIN RUN
    # ================================
    def run(self):
        console.print(Panel(
            f"[bold magenta]O-LEVEL RESULT IMPORTER[/bold magenta]\n"
            f"Exam ID: [bold cyan]{self.exam_id}[/bold cyan] | "
            f"File: [bold cyan]{self.excel_file.name}[/bold cyan]",
            style="magenta", box=box.DOUBLE
        ))

        # Connect
        self._connect()

        # Check existing
        if self._has_records():
            if self.force_import:
                console.print("[yellow]--force-import: Deleting existing...[/yellow]")
                self._delete_existing()
            else:
                if not Confirm.ask("[bold red]Records exist. Delete?[/bold red]", default=False):
                    console.print("[dim]Import cancelled.[/dim]")
                    self.conn.close()
                    return
                self._delete_existing()

        # Setup
        class_id = self._get_class_id()
        subject_count = self._get_subject_count(class_id)
        self.df = self._load_excel()
        self.subject_map = self._build_subject_map(subject_count)

        console.print(f"[cyan]Reading marks from columns:[/cyan] "
                      f"{', '.join([f'C{col+1}' for col in self.subject_map.keys()])}")

        # Preview
        self._preview_data(self.df)

        # Confirm
        if not self.force_import and not Confirm.ask("\n[bold green]Proceed with import?[/bold green]", default=True):
            console.print("[dim]Import cancelled.[/dim]")
            self.conn.close()
            return

        # Import
        self._import_to_db(self.df)

        # Done
        self.conn.close()
        console.print(Panel(
            "[bold green]CASE SOLVED! YOU ARE THE MASTER.[/bold green]",
            style="green", box=box.ROUNDED
        ))


# ========================================
# USAGE EXAMPLE
# ========================================
if __name__ == "__main__":
    importer = OlevelResultImporter(
        exam_id="MID420251027",
        excel_file=r"C:\Kiyabo App\exam templates\Form_IV_Exam_Template_20251025 155902.xlsx",
        db_path=r"C:\Kiyabo App\backend\Kiyabo App Backend v2.0.0.accdb",
        force_import=False,           # Set True to skip delete prompt
        show_preview=True,            # Hide preview if needed
        start_row=13,                 # Change if template changes
        max_preview_rows=20
    )
    importer.run()