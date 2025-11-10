# resultImport.py
# O-LEVEL RESULT IMPORTER + PROCESSOR (Optional)
# 100% VBA LOGIC | RICH UI | PROGRESS BARS | ACCURATE DELETE | PROCESS AFTER
# FULLY COMPATIBLE: Python 3.8+

from __future__ import annotations  # ← Required for tuple[int, int] in Python < 3.9

import sys
from pathlib import Path
import pyodbc
import pandas as pd
from typing import Union, Tuple
from rich.console import Console
from rich.table import Table
from rich.progress import Progress, BarColumn, TextColumn, TimeRemainingColumn
from rich.prompt import Confirm
from rich.panel import Panel
from rich import box

# === IMPORT PROCESSOR ===
try:
    from .processDS import OlevelProcessor
except ImportError as e:
    console = Console()
    console.print(Panel(
        f"[bold red]FATAL: Cannot import OlevelProcessor[/bold red]\n"
        f"Error: {e}\n"
        f"Ensure 'cli/olevel/ProcessDS.py' exists and is importable.",
        style="red", box=box.ROUNDED
    ))
    sys.exit(1)

console = Console()


class OlevelResultImporter:
    """
    O-Level Exam Results Importer + Optional Post-Processing
    Imports Excel → Access DB → (Optional) Runs OlevelProcessor
    100% VBA logic | Rich UI | Full Progress | Accurate Delete Count
    """

    DRIVER = "{Microsoft Access Driver (*.mdb, *.accdb)}"
    RESULTS_TABLE = "tbl_student_exam_results"
    SPECIAL_TABLE = "tbl_student_exam_results_special"
    SUBJECTS_TABLE = "tbl_school_subjects"
    INSERT_BATCH_SIZE = 100

    def __init__(
        self,
        exam_id: str,
        excel_file: Union[Path, str],
        db_path: Union[Path, str],
        force_import: bool = False,
        show_preview: bool = True,
        start_row: int = 13,
        max_preview_rows: int = 20,
        process_after: bool = False,
        base_subjects: int = 7,
        flat_rate: bool = True,
        include_inc: bool = True,
        update_competency: bool = True
    ):
        self.exam_id = exam_id.strip()
        self.excel_file = Path(excel_file)
        self.db_path = Path(db_path)
        self.force_import = force_import
        self.show_preview = show_preview
        self.start_row = start_row
        self.max_preview_rows = max_preview_rows
        self.process_after = process_after

        # Processor defaults (unchanged)
        self.base_subjects = base_subjects
        self.flat_rate = flat_rate
        self.include_inc = include_inc
        self.update_competency = update_competency

        self.conn = None
        self.cursor = None
        self.df = None
        self.subject_map = {}
        self.class_id = ""

    # ================================
    # 1. CONNECT TO DATABASE
    # ================================
    def _connect(self) -> None:
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
    # 2. COUNT EXISTING RECORDS
    # ================================
    def _count_existing_records(self) -> Tuple[int, int]:
        sql_main = f"SELECT COUNT(*) FROM [{self.RESULTS_TABLE}] WHERE exam_id = ?"
        sql_special = f"SELECT COUNT(*) FROM [{self.SPECIAL_TABLE}] WHERE exam_id = ?"
        try:
            main_count = self.cursor.execute(sql_main, (self.exam_id,)).fetchone()[0]
            special_count = self.cursor.execute(sql_special, (self.exam_id,)).fetchone()[0]
            return main_count, special_count
        except Exception as e:
            console.print(f"[yellow]Warning: Count failed: {e}[/yellow]")
            return 0, 0

    # ================================
    # 3. DELETE EXISTING RECORDS
    # ================================
    def _delete_existing_records(self) -> None:
        main_count, special_count = self._count_existing_records()
        total_existing = main_count + special_count

        if total_existing == 0:
            console.print(Panel(
                f"[yellow]No existing records for exam_id: {self.exam_id}[/yellow]",
                style="yellow"
            ))
            return

        if not self.force_import:
            if not Confirm.ask(f"[bold red]Delete {total_existing} record(s)?[/bold red]", default=False):
                console.print("[dim]Import cancelled.[/dim]")
                self.conn.close()
                sys.exit(0)

        self.cursor.execute(f"DELETE FROM [{self.RESULTS_TABLE}] WHERE exam_id = ?", (self.exam_id,))
        deleted_main = self.cursor.rowcount

        self.cursor.execute(f"DELETE FROM [{self.SPECIAL_TABLE}] WHERE exam_id = ?", (self.exam_id,))
        deleted_special = self.cursor.rowcount

        self.conn.commit()

        console.print(Panel(
            f"[bold green]Deleted {deleted_main + deleted_special} record(s)[/bold green]\n"
            f"   • {self.RESULTS_TABLE}: [bold cyan]{deleted_main}[/bold cyan]\n"
            f"   • {self.SPECIAL_TABLE}: [bold cyan]{deleted_special}[/bold cyan]",
            style="green", box=box.ROUNDED
        ))

    # ================================
    # 4. GET CLASS ID FROM EXAM ID
    # ================================
    def _get_class_id(self) -> str:
        digit = self.exam_id[3]
        mapping = {'1': 'I', '2': 'II', '3': 'III', '4': 'IV', '8': 'PC', '0': 'PRE'}
        class_id = mapping.get(digit)
        if not class_id:
            console.print(Panel("[bold red]Invalid exam_id format. Digit 4 must be 0–4, 8[/bold red]", style="red"))
            sys.exit(1)
        self.class_id = class_id
        return class_id

    # ================================
    # 5. GET SUBJECT COUNT FROM DB
    # ================================
    def _get_subject_count(self) -> int:
        sql = f"SELECT COUNT(*) FROM [{self.SUBJECTS_TABLE}] WHERE [is_present_{self.class_id}] = True"
        try:
            count = int(self.cursor.execute(sql).fetchone()[0])
            if count == 0 or count > 20:
                console.print(Panel(f"[bold red]Invalid subject count: {count}[/bold red]", style="red"))
                sys.exit(1)
            return count
        except Exception as e:
            console.print(Panel(f"[bold red]Subject count failed[/bold red]\n{e}", style="red"))
            sys.exit(1)

    # ================================
    # 6. LOAD EXCEL FILE
    # ================================
    def _load_excel(self) -> pd.DataFrame:
        if not self.excel_file.exists():
            console.print(Panel(f"[bold red]File Not Found[/bold red]\n{self.excel_file}", style="red"))
            sys.exit(1)

        console.print(f"[bold cyan]Loading Excel file...[/bold cyan]")
        df = pd.read_excel(self.excel_file, sheet_name=0, header=None, engine="openpyxl")
        console.print(Panel(
            f"[bold green]Loaded:[/bold green] {len(df):,} rows × {len(df.columns)} cols",
            style="green"
        ))
        return df

    # ================================
    # 7. BUILD COLUMN → FIELD MAP (VBA LOGIC)
    # ================================
    def _build_subject_map(self, subject_count: int) -> dict:
        mapping = {}
        col = 6  # ← Excel Column 7 → Python index 6
        base = ["CIV", "HIS", "GEO", "KIS", "ENG", "PHY", "CHE", "BIO", "MAT"]
        extra = ["EDK", "ICS"] + [f"SUB{i}" for i in range(12, 21)]
        for i in range(subject_count):
            field = base[i] if i < 9 else extra[i - 9]
            mapping[col] = field.lower()  # Access uses lowercase
            col += 2  # Skip grade column
        return mapping

    # ================================
    # 8. PREVIEW DATA (RICH TABLE)
    # ================================
    def _preview_data(self, df: pd.DataFrame) -> None:
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
    # 9. IMPORT TO DB WITH PROGRESS
    # ================================
    def _import_to_db(self, df: pd.DataFrame) -> None:
        records = []
        valid_count = 0

        # Step 1: Parse students
        with Progress(
            TextColumn("[bold blue]Parsing students..."),
            BarColumn(),
            TextColumn("{task.completed}/{task.total}"),
            TimeRemainingColumn(),
            console=console
        ) as progress:
            task = progress.add_task("Parsing", total=len(df) - self.start_row)

            for idx in range(self.start_row, len(df)):
                row = df.iloc[idx]
                sid = row.iloc[1]
                if pd.isna(sid) or str(sid).strip() == "":
                    progress.update(task, advance=1)
                    continue

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
                valid_count += 1
                progress.update(task, advance=1)

        if valid_count == 0:
            console.print(Panel("[bold red]No valid student records found[/bold red]", style="red"))
            return

        # Step 2: Insert in batches
        cols = ["exam_id", "student_id"] + list(self.subject_map.values())
        col_str = ", ".join(f"[{c}]" for c in cols)
        placeholders = ", ".join("?" for _ in cols)
        sql = f"INSERT INTO [{self.RESULTS_TABLE}] ({col_str}) VALUES ({placeholders})"

        batch_size = self.INSERT_BATCH_SIZE
        total_batches = (len(records) + batch_size - 1) // batch_size

        with Progress(
            TextColumn("[bold green]Inserting into DB..."),
            BarColumn(),
            TextColumn("{task.completed}/{task.total} batches"),
            TimeRemainingColumn(),
            console=console
        ) as progress:
            task = progress.add_task("Inserting", total=total_batches)

            for i in range(0, len(records), batch_size):
                batch = records[i:i + batch_size]
                data = [[r["exam_id"], r["student_id"]] + [r.get(f, None) for f in self.subject_map.values()] for r in batch]
                self.cursor.executemany(sql, data)
                progress.update(task, advance=1)

        self.conn.commit()
        console.print(Panel(
            f"[bold green]SUCCESS: {len(records):,} records imported![/bold green]",
            style="green", box=box.ROUNDED
        ))

    # ================================
    # 10. RUN PROCESSOR (Optional)
    # ================================
    def _run_processor(self) -> None:
        if not self.process_after:
            return

        console.print(Panel(
            f"[bold magenta]PROCESSING RESULTS[/bold magenta]\n"
            f"Exam ID: [bold cyan]{self.exam_id}[/bold cyan]",
            style="magenta", box=box.ROUNDED
        ))

        try:
            processor = OlevelProcessor(
                exam_id=self.exam_id,
                db_path=str(self.db_path),
                base_subjects=self.base_subjects,
                flat_rate=self.flat_rate,
                include_inc=self.include_inc,
                update_competency=self.update_competency
            )
            processor.run()
            console.print(Panel(
                "[bold green]PROCESSING COMPLETE[/bold green]",
                style="green", box=box.ROUNDED
            ))
        except Exception as e:
            console.print(Panel(f"[bold red]PROCESSING FAILED[/bold red]\n{e}", style="red"))

    # ================================
    # MAIN RUN METHOD
    # ================================
    def run(self) -> None:
        console.print(Panel(
            f"[bold magenta]O-LEVEL RESULT IMPORTER[/bold magenta]\n"
            f"Exam ID: [bold cyan]{self.exam_id}[/bold cyan] | "
            f"File: [bold cyan]{self.excel_file.name}[/bold cyan]",
            style="magenta", box=box.DOUBLE
        ))

        self._connect()
        self._delete_existing_records()
        self._get_class_id()
        subject_count = self._get_subject_count()
        self.df = self._load_excel()
        self.subject_map = self._build_subject_map(subject_count)

        console.print(f"[cyan]Reading marks from:[/cyan] "
                      f"{', '.join([f'C{col+1}' for col in self.subject_map.keys()])}")

        self._preview_data(self.df)

        if not self.force_import and not Confirm.ask("\n[bold green]Proceed with import?[/bold green]", default=True):
            console.print("[dim]Import cancelled.[/dim]")
            self.conn.close()
            return

        self._import_to_db(self.df)
        self.conn.close()

        # === AUTO PROCESS ===
        self._run_processor()

        console.print(Panel(
            "[bold green]CASE SOLVED! YOU ARE THE MASTER.[/bold green]",
            style="green", box=box.ROUNDED
        ))


# ========================================
# USAGE EXAMPLE (CLI)
# ========================================
if __name__ == "__main__":
    importer = OlevelResultImporter(
        exam_id="MID420251027",
        excel_file=r"C:\Kiyabo App\exam templates\Form_IV_Exam_Template_20251025 155902.xlsx",
        db_path=r"C:\Kiyabo App\backend\Kiyabo App Backend v2.0.0.accdb",
        process_after=True
    )
    importer.run()