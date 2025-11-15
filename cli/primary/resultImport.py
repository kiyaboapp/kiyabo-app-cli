# primary_result_importer.py
import pandas as pd
import pyodbc
import os
from rich.console import Console
from rich.table import Table
from rich import box
from tqdm import tqdm

console = Console()


class PrimaryResultImporter:
    def __init__(
        self,
        db_path: str,
        excel_path: str,
        exam_id: str,
        results_table: str = "tbl_pupil_exam_results",
        batch_size: int = 1,
        auto_import: bool = True 
    ):
        self.db_path = db_path
        self.excel_path = excel_path
        self.exam_id = exam_id
        self.results_table = results_table
        self.batch_size = batch_size
        self.auto_import = auto_import

        if not os.path.exists(db_path):
            raise FileNotFoundError(f"Database not found: {db_path}")
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"Excel file not found: {excel_path}")

        self.conn_str = f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_path};"

    def _get_class_from_exam_id(self) -> str:
        n = int(self.exam_id[3])
        return ["Baby", "Middle", "Pre-unit"][n] if n <= 2 else f"Grade {n - 2}"

    def _get_subjects_with_short_names(self):
        class_name = self._get_class_from_exam_id()
        sql = """
        SELECT ss.subject_id, ss.subject_short, ss.subject_name
        FROM tbl_school_subjects ss
        INNER JOIN tbl_class_subjects cs ON ss.subject_number = cs.subject_id
        WHERE ss.is_present = True AND cs.class_id = ?
        ORDER BY cs.ID
        """
        with pyodbc.connect(self.conn_str) as conn:
            cur = conn.cursor()
            cur.execute(sql, class_name)
            return [
                {"id": row.subject_id, "short": (row.subject_short or row.subject_id).strip(), "name": row.subject_name}
                for row in cur.fetchall()
            ]

    def _has_existing_results(self) -> int:
        sql = f"SELECT COUNT(*) FROM {self.results_table} WHERE exam_id = ?"
        with pyodbc.connect(self.conn_str) as conn:
            cur = conn.cursor()
            cur.execute(sql, self.exam_id)
            return cur.fetchone()[0]

    def _delete_existing_results(self):
        with pyodbc.connect(self.conn_str) as conn:
            cur = conn.cursor()
            cur.execute(f"DELETE FROM {self.results_table} WHERE exam_id = ?", self.exam_id)
            cur.execute("DELETE FROM tbl_Competency WHERE exam_id = ?", self.exam_id)
            conn.commit()

    def _show_subject_mapping(self):
        subjects = self._get_subjects_with_short_names()
        table = Table(title="Subject Mapping", box=box.ROUNDED)
        table.add_column("Order", width=6)
        table.add_column("subject_id", style="green")
        table.add_column("Header", style="bold yellow")
        table.add_column("Full Name")
        for i, s in enumerate(subjects, 1):
            table.add_row(str(i), s["id"], s["short"], s["name"])
        console.print(table)
        console.print(f"Total subjects: {len(subjects)}\n")
        return subjects

    def _preview_data(self):
        df = pd.read_excel(self.excel_path, engine="openpyxl", header=None)
        data = df.iloc[13:].copy()
        data = data[data.iloc[:, 1].notna() & (data.iloc[:, 1].astype(str).str.strip() != "")]
        subjects = self._get_subjects_with_short_names()

        table = Table(title=f"Preview — {self.exam_id}", box=box.DOUBLE_EDGE)
        table.add_column("S/N", width=5)
        table.add_column("ID", style="green")
        table.add_column("Name", width=35)
        table.add_column("Sex")
        for s in subjects:
            table.add_column(s["short"], justify="center", style="cyan")

        for idx, (_, row) in enumerate(data.head(12).iterrows(), 1):
            name = " ".join(str(x) for x in row.iloc[2:5] if pd.notna(x))
            sex = str(row.iloc[5]) if len(row) > 5 and pd.notna(row.iloc[5]) else ""
            marks = [
                str(row.iloc[6 + i * 2]).strip() if 6 + i * 2 < len(row) and pd.notna(row.iloc[6 + i * 2]) else "—"
                for i in range(len(subjects))
            ]
            table.add_row(str(idx), str(row.iloc[1]).strip(), name, sex, *marks)

        console.print(table)

    def run(self):
        console.print(f"Starting import → {self.exam_id}\n")


        if self._has_existing_results() > 0:
            console.print(f"[bold red]Results already exist for {self.exam_id}[/bold red]")
            if input("Delete existing and continue? (y/N): ").strip().lower() != "y":
                console.print("Cancelled.")
                return False
            self._delete_existing_results()

        subjects = self._get_subjects_with_short_names()
        if not subjects:
            console.print("[bold red]No subjects found.[/bold red]")
            return False

        self._show_subject_mapping()
        self._preview_data()

        # ← THIS IS WHAT YOU ASKED FOR
        # auto_import=True → NO "Start import?" QUESTION
        if not self.auto_import:
            if input("\nStart import? (y/N): ").strip().lower() != "y":
                console.print("Import cancelled.")
                return False
        else:
            console.print("\n[bold green]auto_import=True → Starting import automatically...[/bold green]")

        # IMPORT STARTS
        df = pd.read_excel(self.excel_path, engine="openpyxl", header=None)
        rows = df.iloc[13:]
        valid = rows[rows.iloc[:, 1].notna() & (rows.iloc[:, 1].astype(str).str.strip() != "")]
        total = len(valid)

        records = []
        with tqdm(total=total, desc="Preparing data", colour="yellow") as pbar:
            for _, row in valid.iterrows():
                student_id = str(row.iloc[1]).strip()
                cols = ["exam_id", "pupil_id"]
                vals = [self.exam_id, student_id]
                for i, subj in enumerate(subjects):
                    col_idx = 6 + i * 2
                    cell = row.iloc[col_idx] if col_idx < len(row) else None
                    if pd.isna(cell) or str(cell).strip() == "":
                        vals.append(None)
                    else:
                        m = str(cell).strip().replace(",", "")
                        if not m.replace(".", "").replace("-", "").isdigit():
                            raise ValueError(f"Invalid mark: {cell}")
                        vals.append(float(m))
                    cols.append(subj["id"])
                records.append((cols, vals))
                pbar.update(1)

        conn = pyodbc.connect(self.conn_str)
        cur = conn.cursor()
        console.print(f"\n[bold green]Inserting {total} records...[/bold green]")
        try:
            with tqdm(total=total, desc=f"Inserting", colour="cyan") as pbar:
                batch = []
                for cols, vals in records:
                    batch.append((cols, vals))
                    if len(batch) >= self.batch_size:
                        for c, v in batch:
                            sql = f"INSERT INTO [{self.results_table}] ({','.join(f'[{x}]' for x in c)}) VALUES ({','.join('?' for _ in v)})"
                            cur.execute(sql, v)
                        pbar.update(len(batch))
                        batch.clear()
                if batch:
                    for c, v in batch:
                        cur.execute(f"INSERT INTO [{self.results_table}] ({','.join(f'[{x}]' for x in c)}) VALUES ({','.join('?' for _ in v)})", v)
                    pbar.update(len(batch))

            console.print("\n[bold green]Committing...[/bold green]")
            with tqdm(total=1, desc="Commit", colour="green") as cbar:
                conn.commit()
                cbar.update(1)

            console.print(f"\n[bold green]SUCCESS:TC {total} records imported![/bold green]")
            return True

        except Exception as e:
            conn.rollback()
            console.print(f"\n[bold red]FAILED: {e}[/bold red]")
            return False
        finally:
            conn.close()



if __name__ == "__main__":
    importer = PrimaryResultImporter(
        db_path=r"C:\Kiyabo App\backend\Kiyabo App Backend v3.0.0.accdb",
        excel_path=r"C:\Kiyabo App\exam templates\Form_Grade 2_Exam_Template_20251115-163121.xlsx",
        exam_id="ANN420251117",
        batch_size=1,
        auto_import=True
    )
    importer.run()