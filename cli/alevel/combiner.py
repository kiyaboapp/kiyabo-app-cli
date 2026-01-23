import pandas as pd
import pyodbc
import numpy as np
from tqdm import tqdm
from rich.console import Console
from rich.table import Table
from rich import box
from rich.text import Text
from datetime import datetime

tqdm.pandas()

console = Console()

class DualAlevelProcessor:
    """Processes A-Level student exam results for TWO exams combined with ranking and grading."""
    
    SUBJECTS = [
        "acc", "agr", "ara", "bio", "bus", "che", "chi", "com", "csc", 
        "eng", "eco", "fsh", "fin", "fhn", "fre", "tex", "geo", "his", 
        "isl", "kis", "lit", "mat", "msc", "spt", "the", "phy", "bam", 
        "gs", "htm", "aco", "cus"
    ]
    
    
    DB_UPDATE_FIELDS = [
        "division", "points", "subject_count", "total_marks", "gpa",
        "position_comb", "position_school", "out_of_comb", "out_of_school",
        "first", "second", "third", 'avg_grade', 'avg_marks',
        "subject_count_all", "necta_results", "necta_results_marks"
    ]
    
    GRADE_POINTS = {'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5, 'S': 6, 'F': 7, None: None}
    DIVISION_VALUES = {'I': 1, 'II': 2, 'III': 3, 'IV': 4, '0': 5, 'ABS': None, 'INC': None}
    POINTS_TO_DIV = {
        range(3, 10): 'I',
        range(10, 13): 'II',
        range(13, 18): 'III',
        range(18, 20): 'IV',
        range(20, 22): '0'
    }
    
    SORT_ASCENDING = {
        'points': True,
        'avg_marks': False,
        'subject_count': False,
    }
    
    def __init__(self, exam_id_1: str, exam_id_2: str, db_path: str, 
                 exam_name_1: str = None, exam_name_2: str = None,
                 sort_columns: list = None,
                 include_inc: bool = True,
                 class_id:str=None,
                 rank_method: str = 'min'):
        self.exam_id_1 = exam_id_1
        self.exam_id_2 = exam_id_2
        self.db_path = db_path
        self.combo_id = f"{exam_id_1}_{exam_id_2}"
        
        # IMPORTANT: keep default ordering in sync with single-exam AlevelProcessor
        # (avg_marks first, then points, then subject_count)
        self.sort_columns = sort_columns or ["avg_marks", "points", "subject_count"]
        self.include_inc = include_inc
        self.rank_method = rank_method
        self.sort_ascending = [self.SORT_ASCENDING[col] for col in self.sort_columns]
        
        self.conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + db_path + ';'
        self.df = None              # combined (averaged) dataframe
        self.df1 = None             # processed exam 1 dataframe (with _1 suffix)
        self.df2 = None             # processed exam 2 dataframe (with _2 suffix)
        self.valid_subjects = []
        self.comb_metadata = None
        self.subject_to_user = {}
        self.db_columns = []
        
        self.exam_name_1 = exam_name_1
        self.exam_name_2 = exam_name_2
        self.class_id = class_id
        
    @staticmethod
    def bracket_field(field: str) -> str:
        return f"[{field}]"
    
    @staticmethod
    def get_grade(marks) -> str:
        if pd.isna(marks) or not isinstance(marks, (int, float)) or marks < 0 or marks > 100:
            return None
        if marks >= 80: return 'A'
        elif marks >= 70: return 'B'
        elif marks >= 60: return 'C'
        elif marks >= 50: return 'D'
        elif marks >= 40: return 'E'
        elif marks >= 35: return 'S'
        else: return 'F'
    
    @classmethod
    def get_div_from_points(cls, points) -> str:
        if pd.isna(points) or not isinstance(points, int):
            return None
        points = int(points)
        for r, div in cls.POINTS_TO_DIV.items():
            if points in r:
                return div
        return None
    
    @staticmethod
    def format_marks_for_necta(marks):
        if pd.isna(marks):
            return ''
        if isinstance(marks, (int, np.integer)) or (isinstance(marks, float) and marks.is_integer()):
            return str(int(marks))
        else:
            return f"{marks:.2f}"
    
    @staticmethod
    def format_grade_for_necta(grade):
        return f"'{grade}'" if grade else ''
    
    @staticmethod
    def format_sig_figs(value, sig_figs):
        """Format a number to specified significant figures"""
        if pd.isna(value):
            return "—"
        if value == 0:
            return "0"
        import math
        try:
            magnitude = math.floor(math.log10(abs(value)))
            decimals = max(0, sig_figs - magnitude - 1)
            return f"{value:.{decimals}f}"
        except (ValueError, OverflowError):
            return f"{value:.4f}"
    
    def _preview(self, title: str, cols: list, highlight=None, rows: int = 12, df=None):
        """Beautiful preview table - matches ranking.py style"""
        if df is None:
            df = self.df
        if df is None or df.empty:
            return
        
        table = Table(
            title=f"[bold gold1]{title}[/]", 
            box=box.ROUNDED, 
            border_style="cyan", 
            expand=True,
            show_lines=False
        )
        
        # Define column widths and formatting
        col_widths = {
            'full_name': 25,
            'division': 8,
            'points': 8,
            'avg_marks': 10,
            'position_school': 12,
            'position_comb': 12,
        }
        
        for col in cols:
            style = "bold bright_yellow" if highlight and col in highlight else "white"
            width = col_widths.get(col, None)
            justify = "left" if col == "full_name" else "center"
            table.add_column(
                col.upper().replace("_", " "), 
                style=style, 
                justify=justify,
                width=width,
                no_wrap=False
            )
        
        preview_df = df[cols].head(rows) if all(c in df.columns for c in cols) else pd.DataFrame()
        
        def format_value(val, col_name):
            if pd.isna(val):
                return "—"
            if 'avg_marks' in col_name.lower():
                # 2 significant figures for averages
                return f"{float(val):.2f}"
            elif 'points' in col_name.lower():
                return f"{int(val)}" if pd.notna(val) else "—"
            elif 'position' in col_name.lower():
                return f"{int(val)}" if pd.notna(val) else "—"
            else:
                return str(val)
        
        for _, row in preview_df.iterrows():
            values = [format_value(row[col], col) for col in cols]
            table.add_row(*values)
        console.print(table)
    
    def connect_db(self):
        return pyodbc.connect(self.conn_str)
    
    def create_dual_exam_table(self):
        console.print("\n[bold cyan]Creating tbl_dual_exams if not exists[/bold cyan]")
        conn = self.connect_db()
        cursor = conn.cursor()
        
        tables = [table.table_name for table in cursor.tables(tableType='TABLE')]
        exists = 'tbl_dual_exams' in tables
        
        if not exists:
            cursor.execute("""
                CREATE TABLE tbl_dual_exams (
                    combo_id TEXT(100) PRIMARY KEY,
                    exam_id_1 TEXT(50),
                    exam_id_2 TEXT(50),
                    exam_name_1 TEXT(100),
                    exam_name_2 TEXT(100),
                    class_id TEXT(50)
                )
            """)
            conn.commit()
            console.print(" [green]- tbl_dual_exams created[/green]")
        else:
            console.print(" [yellow]- tbl_dual_exams already exists[/yellow]")
        
        # Get exam names and class_id if not provided
        if not self.exam_name_1 or not self.exam_name_2 or not self.class_id:
            exam_query = "SELECT exam_id, exam_name, class_id FROM tbl_student_exams WHERE exam_id IN (?, ?)"
            exam_df = pd.read_sql(exam_query, conn, params=[self.exam_id_1, self.exam_id_2])
            
            if not self.exam_name_1:
                row = exam_df[exam_df['exam_id'] == self.exam_id_1]
                self.exam_name_1 = row.iloc[0]['exam_name'] if not row.empty else self.exam_id_1
            
            if not self.exam_name_2:
                row = exam_df[exam_df['exam_id'] == self.exam_id_2]
                self.exam_name_2 = row.iloc[0]['exam_name'] if not row.empty else self.exam_id_2
            
            if not self.class_id:
                row = exam_df[exam_df['exam_id'] == self.exam_id_2]
                self.class_id = row.iloc[0]['class_id'] if not row.empty and pd.notna(row.iloc[0]['class_id']) else None
        
        # Insert/update combo record
        cursor.execute("SELECT COUNT(*) FROM tbl_dual_exams WHERE combo_id = ?", (self.combo_id,))
        if cursor.fetchone()[0] == 0:
            cursor.execute("""
                INSERT INTO tbl_dual_exams (combo_id, exam_id_1, exam_id_2, exam_name_1, exam_name_2, class_id)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (self.combo_id, self.exam_id_1, self.exam_id_2, self.exam_name_1, self.exam_name_2, self.class_id))
        else:
            cursor.execute("""
                UPDATE tbl_dual_exams 
                SET exam_name_1=?, exam_name_2=?, class_id=?
                WHERE combo_id=?
            """, (self.exam_name_1, self.exam_name_2, self.class_id, self.combo_id))
        
        conn.commit()
        conn.close()
        console.print(f" [green]- Combo record saved: {self.combo_id}[/green]")
    
    def create_dual_combined_results_table(self):
        console.print("\n[bold cyan]Creating/Updating tbl_dual_combined_results[/bold cyan]")
        conn = self.connect_db()
        cursor = conn.cursor()
        
        tables = [table.table_name for table in cursor.tables(tableType='TABLE')]
        exists = 'tbl_dual_combined_results' in tables
        
        if not exists:
            cursor.execute("""
                CREATE TABLE tbl_dual_combined_results (
                    record_id AUTOINCREMENT PRIMARY KEY,
                    combo_id TEXT(100),
                    exam_id_1 TEXT(50),
                    exam_id_2 TEXT(50),
                    student_id TEXT(50),
                    full_name TEXT(100),
                    sex TEXT(10),
                    comb_id TEXT(50),
                    division TEXT(10),
                    points INTEGER,
                    subject_count INTEGER,
                    total_marks DOUBLE,
                    gpa DOUBLE,
                    position_comb INTEGER,
                    position_school INTEGER,
                    out_of_comb INTEGER,
                    out_of_school INTEGER,
                    first DOUBLE,
                    second DOUBLE,
                    third DOUBLE,
                    avg_grade TEXT(5),
                    avg_marks DOUBLE,
                    subject_count_all INTEGER,
                    necta_results MEMO,
                    necta_results_marks MEMO,
                    division_1 TEXT(10),
                    points_1 INTEGER,
                    subject_count_1 INTEGER,
                    total_marks_1 DOUBLE,
                    gpa_1 DOUBLE,
                    position_comb_1 INTEGER,
                    position_school_1 INTEGER,
                    out_of_comb_1 INTEGER,
                    out_of_school_1 INTEGER,
                    first_1 DOUBLE,
                    second_1 DOUBLE,
                    third_1 DOUBLE,
                    avg_grade_1 TEXT(5),
                    avg_marks_1 DOUBLE,
                    subject_count_all_1 INTEGER,
                    necta_results_1 MEMO,
                    necta_results_marks_1 MEMO,
                    division_2 TEXT(10),
                    points_2 INTEGER,
                    subject_count_2 INTEGER,
                    total_marks_2 DOUBLE,
                    gpa_2 DOUBLE,
                    position_comb_2 INTEGER,
                    position_school_2 INTEGER,
                    out_of_comb_2 INTEGER,
                    out_of_school_2 INTEGER,
                    first_2 DOUBLE,
                    second_2 DOUBLE,
                    third_2 DOUBLE,
                    avg_grade_2 TEXT(5),
                    avg_marks_2 DOUBLE,
                    subject_count_all_2 INTEGER,
                    necta_results_2 MEMO,
                    necta_results_marks_2 MEMO,
                    class_id TEXT(50),
                    processed_date DATETIME
                )
            """)
            conn.commit()
            console.print(" [green]- tbl_dual_combined_results created[/green]")
        else:
            console.print(" [yellow]- tbl_dual_combined_results exists[/yellow]")

            # Ensure new exam_id_1, exam_id_2, class_id, and processed_date columns exist on older databases
            existing_cols = [c.column_name.lower() for c in cursor.columns(table='tbl_dual_combined_results')]
            for col_name, col_type in [('exam_id_1', 'TEXT(50)'), ('exam_id_2', 'TEXT(50)'), 
                                       ('class_id', 'TEXT(50)'), ('processed_date', 'DATETIME')]:
                if col_name not in existing_cols:
                    try:
                        cursor.execute(
                            f"ALTER TABLE tbl_dual_combined_results "
                            f"ADD COLUMN {self.bracket_field(col_name)} {col_type}"
                        )
                        conn.commit()
                        console.print(f" [green]+ Added column: {col_name}[/green]")
                    except Exception as e:
                        console.print(f" [yellow]Column {col_name} may already exist: {e}[/yellow]")
        
        conn.close()




    def ensure_subject_columns_exist(self, subjects_list, suffix=''):
        """
        Ensure subject columns exist for a specific suffix.
        This is called early for processing, but ensure_all_insert_columns_exist()
        will comprehensively check ALL columns before insert.
        """
        console.print(f"\n[bold cyan]Ensuring subject columns exist (suffix='{suffix}')[/bold cyan]")
        conn = self.connect_db()
        cursor = conn.cursor()
        
        # Check if table exists
        tables = [table.table_name for table in cursor.tables(tableType='TABLE')]
        if 'tbl_dual_combined_results' not in tables:
            console.print(f" [yellow]- Table tbl_dual_combined_results does not exist yet, will be created[/yellow]")
            conn.close()
            return
        
        # Get existing columns
        try:
            cursor.execute("SELECT TOP 1 * FROM tbl_dual_combined_results")
            existing_cols = [desc[0].lower() for desc in cursor.description]
        except Exception as e:
            console.print(f" [yellow]- Could not read table structure: {e}[/yellow]")
            conn.close()
            return
        
        created_count = 0
        for sub in subjects_list:
            cols_to_add = [
                (f"{sub}{suffix}", "DOUBLE"),
                (f"{sub}_grade{suffix}", "TEXT(5)"),
                (f"{sub}_pos{suffix}", "INTEGER"),
                (f"{sub}_out_of{suffix}", "INTEGER")
            ]
            
            for col_name, col_type in cols_to_add:
                if col_name.lower() not in existing_cols:
                    try:
                        cursor.execute(f"ALTER TABLE tbl_dual_combined_results ADD COLUMN {self.bracket_field(col_name)} {col_type}")
                        conn.commit()
                        console.print(f" [green]+ Added column: {col_name}[/green]")
                        existing_cols.append(col_name.lower())
                        created_count += 1
                    except Exception as e:
                        console.print(f" [yellow]Column {col_name} might exist: {e}[/yellow]")
        
        if created_count == 0:
            console.print(f" [yellow]- All subject columns already exist for suffix '{suffix}'[/yellow]")
        
        conn.close()
    
    def fetch_dual_exam_data(self):
        console.print("\n[bold cyan]================================================================================[/bold cyan]")
        console.print("[bold cyan]  FETCHING DATA FOR BOTH EXAMS[/bold cyan]")
        console.print("[bold cyan]================================================================================[/bold cyan]\n")
        conn = self.connect_db()
        
        # Detect columns from exam 1
        dummy_query = "SELECT TOP 1 * FROM tbl_student_exam_results WHERE exam_id = ?"
        df_dummy = pd.read_sql(dummy_query, conn, params=[self.exam_id_1])
        self.db_columns = [col.lower() for col in df_dummy.columns]
        potential_subjects = [sub for sub in self.SUBJECTS if sub in self.db_columns]
        
        console.print(f" [green]- Found {len(potential_subjects)} potential subjects[/green]")
        
        # Fetch exam 1
        query_1 = f"""
        SELECT r.*, i.full_name, i.sex, i.comb_id
        FROM tbl_student_exam_results r
        INNER JOIN tbl_student_academic_info i ON r.student_id = i.student_id
        WHERE r.exam_id = ?
        """
        df1 = pd.read_sql(query_1, conn, params=[self.exam_id_1])
        console.print(f" [green]- Exam 1: {len(df1)} records[/green]")
        
        # Fetch exam 2
        query_2 = f"""
        SELECT r.*, i.full_name, i.sex, i.comb_id
        FROM tbl_student_exam_results r
        INNER JOIN tbl_student_academic_info i ON r.student_id = i.student_id
        WHERE r.exam_id = ?
        """
        df2 = pd.read_sql(query_2, conn, params=[self.exam_id_2])
        console.print(f" [green]- Exam 2: {len(df2)} records[/green]")
        
        conn.close()
        
        # Filter to only students in exam 2
        students_in_exam2 = set(df2['student_id'].values)
        df1 = df1[df1['student_id'].isin(students_in_exam2)]
        console.print(f" [yellow]- Filtered exam 1 to {len(df1)} students (present in exam 2)[/yellow]")
        
        # Identify valid subjects from both exams
        valid_subjects_1 = []
        valid_subjects_2 = []
        
        for sub in potential_subjects:
            if sub in df1.columns:
                df1[sub] = pd.to_numeric(df1[sub], errors='coerce')
                if df1[sub].ge(0).any():
                    valid_subjects_1.append(sub)
            
            if sub in df2.columns:
                df2[sub] = pd.to_numeric(df2[sub], errors='coerce')
                if df2[sub].ge(0).any():
                    valid_subjects_2.append(sub)
        
        self.valid_subjects = sorted(set(valid_subjects_1 + valid_subjects_2))
        console.print(f" [green]- Valid subjects across both exams: {self.valid_subjects}[/green]")
        
        return df1, df2
    
    def process_single_exam(self, df, exam_id, suffix):
        console.print(f"\n[bold cyan]================================================================================[/bold cyan]")
        console.print(f"[bold cyan]  PROCESSING EXAM: {exam_id} (suffix={suffix})[/bold cyan]")
        console.print(f"[bold cyan]================================================================================[/bold cyan]\n")
        
        # Load metadata
        conn = self.connect_db()
        comb_df = pd.read_sql("SELECT serial_id, comb_id, subject_id FROM tbl_student_comb_subjects", conn)
        sub_df = pd.read_sql("SELECT subject_serial, subject_short, subject_user_short, is_core, is_present, subject_name FROM tbl_student_subjects", conn)
        conn.close()
        
        comb_metadata_all = comb_df.merge(sub_df, left_on='subject_id', right_on='subject_serial')
        self.comb_metadata = comb_metadata_all[comb_metadata_all['is_present'] == True]
        self.subject_to_user = dict(zip(sub_df['subject_short'].str.lower(), sub_df['subject_user_short']))
        
        df['comb_id'] = df['comb_id'].astype(str)
        
        # Compute grades
        for sub in self.valid_subjects:
            if sub in df.columns:
                df[sub] = df[sub].where(df[sub].between(0, 100), np.nan)
                df[f'{sub}_grade{suffix}'] = df[sub].apply(self.get_grade)
        
        df[f'total_marks{suffix}'] = df[self.valid_subjects].sum(axis=1, skipna=True)
        
        # Process each row
        df = df.apply(lambda row: self.process_student_row(row, suffix), axis=1)
        
        # Rankings
        self.compute_ranking(df, suffix)
        self.compute_subject_rankings(df, suffix)
        
        # Beautiful preview after processing
        preview_cols = [f'full_name', f'division{suffix}', f'points{suffix}', 
                       f'avg_marks{suffix}', 
                       f'position_school{suffix}', f'position_comb{suffix}']
        available_cols = [c for c in preview_cols if c in df.columns]
        if available_cols:
            # Sort by position_school for preview (handle NaN/object dtype)
            pos_col = f'position_school{suffix}'
            if pos_col in df.columns:
                df_preview = df.copy()
                df_preview[pos_col] = pd.to_numeric(df_preview[pos_col], errors='coerce')
                df_preview = df_preview.nsmallest(12, pos_col) if df_preview[pos_col].notna().any() else df.head(12)
            else:
                df_preview = df.head(12)
            
            self._preview(
                f"EXAM {exam_id} RESULTS — Top 12 Students (suffix={suffix})",
                available_cols,
                highlight=[f'division{suffix}', f'points{suffix}', f'position_school{suffix}'],
                rows=12,
                df=df_preview
            )
        
        console.print(f" [bold green][OK] Exam {exam_id} processed successfully[/bold green]")
        return df
    
    def process_student_row(self, row, suffix):
        student_comb = self.comb_metadata[self.comb_metadata['comb_id'] == row['comb_id']]
        comb_shorts = set(student_comb['subject_short'].str.lower())
        core_shorts = set(student_comb[student_comb['is_core'] == True]['subject_short'].str.lower())
        
        effective_cores = core_shorts.union([sub for sub in self.valid_subjects if pd.notna(row.get(sub)) and sub not in comb_shorts])
        attempted_effective = [sub for sub in effective_cores if pd.notna(row.get(sub))]
        missing_count = len(effective_cores) - len(attempted_effective)
        
        attempted_all = [sub for sub in self.valid_subjects if pd.notna(row.get(sub))]
        row[f'subject_count_all{suffix}'] = len(attempted_all) if attempted_all else 0
        row[f'subject_count{suffix}'] = len(attempted_effective) if attempted_effective else 0
        
        if row[f'subject_count_all{suffix}'] > 0:
            row[f'avg_marks{suffix}'] = row[f'total_marks{suffix}'] / row[f'subject_count_all{suffix}']
            row[f'avg_grade{suffix}'] = self.get_grade(row[f'avg_marks{suffix}'])
        else:
            row[f'avg_marks{suffix}'] = None
            row[f'avg_grade{suffix}'] = None
        
        core_marks = sorted([row[sub] for sub in attempted_effective if pd.notna(row.get(sub))], reverse=True)
        row[f'first{suffix}'] = core_marks[0] if len(core_marks) >= 1 else None
        row[f'second{suffix}'] = core_marks[1] if len(core_marks) >= 2 else None
        row[f'third{suffix}'] = core_marks[2] if len(core_marks) >= 3 else None
        
        grade_pts = [self.GRADE_POINTS.get(row.get(f'{sub}_grade{suffix}')) for sub in attempted_effective]
        valid_pts = [p for p in grade_pts if p is not None]
        computed_points = sum(valid_pts) if valid_pts else None
        
        has_invalid = len(grade_pts) != len(valid_pts)
        is_complete = (len(effective_cores) >= 3 and missing_count == 0 and not has_invalid)
        is_abs = len(attempted_effective) == 0
        is_inc = not is_complete and not is_abs
        
        if is_abs:
            save_div = 'ABS' if self.include_inc else ('0' if self.get_div_from_points(7 * missing_count) == '0' else 'IV')
            save_points = None
        elif is_inc:
            save_div = 'INC' if self.include_inc else ('0' if self.get_div_from_points(sum(valid_pts) + 7 * missing_count) == '0' else 'IV')
            save_points = None
        else:
            save_div = self.get_div_from_points(computed_points)
            save_points = computed_points if save_div is not None else None
        
        row[f'division{suffix}'] = save_div
        row[f'points{suffix}'] = save_points
        
        div_val = self.DIVISION_VALUES.get(save_div)
        if div_val is not None and row[f'subject_count{suffix}'] > 0:
            row[f'gpa{suffix}'] = div_val / row[f'subject_count{suffix}']
        else:
            row[f'gpa{suffix}'] = None
        
        row = self.build_necta(row, suffix)
        return row
    
    def build_necta(self, row, suffix):
        parts = []
        parts_marks = []
        student_comb = self.comb_metadata[self.comb_metadata['comb_id'] == row['comb_id']]
        comb_shorts = set(student_comb['subject_short'].str.lower())
        
        def get_user(short):
            meta_row = student_comb[student_comb['subject_short'].str.lower() == short]
            if not meta_row.empty:
                return meta_row.iloc[0]['subject_user_short']
            return self.subject_to_user.get(short, short.upper())
        
        for _, sub_row in student_comb.iterrows():
            short = sub_row['subject_short'].lower()
            user = sub_row['subject_user_short']
            marks = row.get(short)
            grade = row.get(f'{short}_grade{suffix}')
            if pd.notna(marks):
                parts.append(f"{user}-{self.format_grade_for_necta(grade)}")
                marks_str = self.format_marks_for_necta(marks)
                parts_marks.append(f"{user}-{marks_str} {self.format_grade_for_necta(grade)}")
            else:
                parts.append(f"{user}-X")
                parts_marks.append(f"{user}-X")
        
        for sub in self.valid_subjects:
            if pd.notna(row.get(sub)) and sub not in comb_shorts:
                user = get_user(sub)
                grade = row.get(f'{sub}_grade{suffix}')
                parts.append(f"{user}-{self.format_grade_for_necta(grade)}")
                marks_str = self.format_marks_for_necta(row[sub])
                parts_marks.append(f"{user}-{marks_str} {self.format_grade_for_necta(grade)}")
        
        avg_grade = row.get(f'avg_grade{suffix}')
        avg_grade_fmt = self.format_grade_for_necta(avg_grade)
        parts.append(f"AVG-{avg_grade_fmt}")
        
        avg_marks_val = row.get(f'avg_marks{suffix}')
        avg_marks_str = self.format_marks_for_necta(avg_marks_val) if pd.notna(avg_marks_val) else 'X'
        parts_marks.append(f"AVG-{avg_marks_str}{avg_grade_fmt}")
        
        row[f'necta_results{suffix}'] = ', '.join(parts)
        row[f'necta_results_marks{suffix}'] = ', '.join(parts_marks)
        return row

    def compute_ranking(self, df, suffix):
        """
        Compute school-wide and combination-specific rankings with PROPER TIE HANDLING,
        mirroring the single-exam AlevelProcessor logic but supporting a suffix.
        """
        console.print(f"\n[bold cyan]Stage: Computing Rankings (suffix='{suffix}')[/bold cyan]")
        
        # ============================================================
        # 1. Create masks for validity (must match single-exam logic)
        # ============================================================
        div_col = f'division{suffix}'
        points_col = f'points{suffix}'
        avg_col = f'avg_marks{suffix}'
        
        if div_col not in df.columns or points_col not in df.columns or avg_col not in df.columns:
            console.print(f" [yellow]- Missing ranking columns for suffix='{suffix}', skipping ranking.[/yellow]")
            return df
        
        invalid_mask = (
            (df[div_col] == 'ABS') |
            (df[points_col].isna()) |
            (df[avg_col].isna())
        )
        
        valid_students = df[~invalid_mask].copy()
        invalid_students = df[invalid_mask]
        
        if valid_students.empty:
            # Clear ranks for all students
            rank_cols = [f'position_school{suffix}', f'out_of_school{suffix}', 
                        f'position_comb{suffix}', f'out_of_comb{suffix}']
            df[rank_cols] = pd.NA
            console.print(f" [yellow]- No valid students to rank (suffix='{suffix}')[/yellow]")
            return df
        
        console.print(f" [green][OK] Found {len(valid_students):,} valid students to rank (suffix='{suffix}')[/green]")
        console.print(f" [yellow][OK] Excluded {len(invalid_students):,} invalid students (suffix='{suffix}')[/yellow]")
        
        # ============================================================
        # 2. Prepare sort columns & NaN handling (synchronized with AlevelProcessor)
        # ============================================================
        sort_cols_suffixed = [f"{col}{suffix}" for col in self.sort_columns]
        console.print(f"\n[bold yellow]Preparing sort columns for suffix='{suffix}': {sort_cols_suffixed}[/bold yellow]")
        
        for base_col, col in zip(self.sort_columns, sort_cols_suffixed):
            if col not in valid_students.columns:
                continue
            
            before_fill = valid_students[col].isna().sum()
            
            if base_col == 'points':
                valid_students[col] = valid_students[col].fillna(999999)
            elif base_col == 'avg_marks':
                valid_students[col] = valid_students[col].fillna(-1)
            elif base_col in ['subject_count', 'subject_count_all']:
                valid_students[col] = valid_students[col].fillna(0)
            
            after_fill = valid_students[col].isna().sum()
            console.print(f" [cyan]- {col}: filled {before_fill} NaN values -> {after_fill} remaining[/cyan]")
        
        # Round numeric sort columns to 4 decimals to avoid float comparison issues
        for base_col, col in zip(self.sort_columns, sort_cols_suffixed):
            if base_col in ['points', 'avg_marks'] and col in valid_students.columns:
                valid_students[col] = pd.to_numeric(valid_students[col], errors='coerce').round(4)
                console.print(f" [cyan]- {col}: rounded to 4 decimal places[/cyan]")
        
        # ============================================================
        # 3. SCHOOL-WIDE RANKING (IDENTICAL STRATEGY AS SINGLE-EXAM)
        # ============================================================
        console.print(f"\n[bold yellow]School-Wide Ranking (suffix='{suffix}')[/bold yellow]")
        console.print(
            f" [cyan]- Sort criteria: "
            f"{list(zip(sort_cols_suffixed, ['ASC' if asc else 'DESC' for asc in self.sort_ascending]))}[/cyan]"
        )
        
        valid_students_sorted = valid_students.sort_values(
            by=sort_cols_suffixed,
            ascending=self.sort_ascending
        ).copy()
        
        # Tuple of sort values for stable tie detection
        valid_students_sorted['_sort_tuple'] = valid_students_sorted[sort_cols_suffixed].apply(
            lambda row: tuple(row), axis=1
        )
        
        unique_tuples = valid_students_sorted['_sort_tuple'].unique()
        rank_map = {tuple_val: idx + 1 for idx, tuple_val in enumerate(unique_tuples)}
        
        pos_school_col = f'position_school{suffix}'
        out_school_col = f'out_of_school{suffix}'
        
        valid_students_sorted[pos_school_col] = valid_students_sorted['_sort_tuple'].map(rank_map)
        valid_students_sorted[out_school_col] = len(valid_students_sorted)
        
        console.print(f" [green][OK] Ranked {len(valid_students_sorted):,} students school-wide (suffix='{suffix}')[/green]")
        console.print(f" [green][OK] Unique performance levels: {len(unique_tuples):,}[/green]")
        
        # ============================================================
        # 4. COMBINATION-SPECIFIC RANKING (IDENTICAL STRATEGY)
        # ============================================================
        console.print(f"\n[bold yellow]Combination-Specific Ranking (suffix='{suffix}')[/bold yellow]")
        
        pos_comb_col = f'position_comb{suffix}'
        out_comb_col = f'out_of_comb{suffix}'
        
        def rank_within_combination(group):
            group['_comb_sort_tuple'] = group[sort_cols_suffixed].apply(
                lambda row: tuple(row), axis=1
            )
            unique_tuples_local = group['_comb_sort_tuple'].unique()
            rank_map_local = {tuple_val: idx + 1 for idx, tuple_val in enumerate(unique_tuples_local)}
            group[pos_comb_col] = group['_comb_sort_tuple'].map(rank_map_local)
            group[out_comb_col] = len(group)
            return group.drop(columns=['_comb_sort_tuple'])
        
        ranked_df = valid_students_sorted.groupby('comb_id', group_keys=False).apply(
            rank_within_combination
        )
        
        ranked_df = ranked_df.drop(columns=['_sort_tuple'], errors='ignore')
        
        num_combs = ranked_df['comb_id'].nunique()
        console.print(f" [green][OK] Ranked students within {num_combs} combinations (suffix='{suffix}')[/green]")
        
        # ============================================================
        # 5. WRITE BACK TO DATAFRAME & CLEAR INVALIDS
        # ============================================================
        rank_cols = [pos_school_col, out_school_col, pos_comb_col, out_comb_col]
        df.loc[ranked_df.index, rank_cols] = ranked_df[rank_cols]
        
        if not invalid_students.empty:
            df.loc[invalid_students.index, rank_cols] = pd.NA
            console.print(
                f" [yellow]- Cleared ranks for {len(invalid_students)} invalid students (suffix='{suffix}')[/yellow]"
            )
        
        console.print(f"[bold green][OK] Ranking complete for suffix='{suffix}'[/bold green]")
        return df


    def compute_subject_rankings(self, df, suffix):
        for sub in self.valid_subjects:
            if sub not in df.columns:
                continue
            
            pos_col = f"{sub}_pos{suffix}"
            out_col = f"{sub}_out_of{suffix}"
            
            sub_df = df[df[sub].notna()].copy()
            if sub_df.empty:
                continue
            
            sub_df = sub_df.sort_values(sub, ascending=False)
            sub_df['rank_temp'] = sub_df[sub].rank(method='min', ascending=False)
            sub_df[pos_col] = sub_df['rank_temp'].astype('float64').round().astype('Int64')
            sub_df[out_col] = len(sub_df)
            
            df.loc[sub_df.index, [pos_col, out_col]] = sub_df[[pos_col, out_col]]
        
        return df
    
    def process_combined_exam(self, df_combined):
        console.print("\n[bold cyan]================================================================================[/bold cyan]")
        console.print("[bold cyan]  PROCESSING COMBINED (AVERAGED) EXAM[/bold cyan]")
        console.print("[bold cyan]================================================================================[/bold cyan]\n")
        
        self.df = df_combined
        
        # Load metadata
        conn = self.connect_db()
        comb_df = pd.read_sql("SELECT serial_id, comb_id, subject_id FROM tbl_student_comb_subjects", conn)
        sub_df = pd.read_sql("SELECT subject_serial, subject_short, subject_user_short, is_core, is_present, subject_name FROM tbl_student_subjects", conn)
        conn.close()
        
        comb_metadata_all = comb_df.merge(sub_df, left_on='subject_id', right_on='subject_serial')
        self.comb_metadata = comb_metadata_all[comb_metadata_all['is_present'] == True]
        self.subject_to_user = dict(zip(sub_df['subject_short'].str.lower(), sub_df['subject_user_short']))
        
        self.df['comb_id'] = self.df['comb_id'].astype(str)
        
        # Compute grades for averaged marks
        for sub in self.valid_subjects:
            if sub in self.df.columns:
                self.df[sub] = self.df[sub].where(self.df[sub].between(0, 100), np.nan)
                self.df[f'{sub}_grade'] = self.df[sub].apply(self.get_grade)
        
        self.df['total_marks'] = self.df[self.valid_subjects].sum(axis=1, skipna=True)
        
        # Process each row
        self.df = self.df.apply(lambda row: self.process_student_row(row, ''), axis=1)
        
        # Rankings
        self.compute_ranking(self.df, '')
        self.compute_subject_rankings(self.df, '')
        
        # Beautiful preview after processing combined exam
        preview_cols = ['full_name', 'division', 'points', 'avg_marks', 
                       'position_school', 'position_comb']
        available_cols = [c for c in preview_cols if c in self.df.columns]
        if available_cols:
            # Sort by position_school for preview (handle NaN/object dtype)
            if 'position_school' in self.df.columns:
                df_preview = self.df.copy()
                df_preview['position_school'] = pd.to_numeric(df_preview['position_school'], errors='coerce')
                df_preview = df_preview.nsmallest(12, 'position_school') if df_preview['position_school'].notna().any() else self.df.head(12)
            else:
                df_preview = self.df.head(12)
            
            self._preview(
                "COMBINED (AVERAGED) EXAM RESULTS — Top 12 Students",
                available_cols,
                highlight=['division', 'points', 'position_school'],
                rows=12,
                df=df_preview
            )
        
        console.print(" [bold green][OK] Combined exam processed successfully[/bold green]")
    
    def create_dual_competency_table(self):
        console.print("\n[bold cyan]Creating tbl_dual_competency if not exists[/bold cyan]")
        conn = self.connect_db()
        cursor = conn.cursor()
        
        tables = [table.table_name for table in cursor.tables(tableType='TABLE')]
        exists = 'tbl_dual_competency' in tables
        
        if not exists:
            cursor.execute("""
                CREATE TABLE tbl_dual_competency (
                    record_id AUTOINCREMENT PRIMARY KEY,
                    combo_id TEXT(100),
                    subject_serial INTEGER,
                    A_s INTEGER,
                    B_s INTEGER,
                    C_s INTEGER,
                    D_s INTEGER,
                    E_s INTEGER,
                    S_s INTEGER,
                    F_s INTEGER,
                    total INTEGER,
                    pass INTEGER,
                    fail INTEGER,
                    gpa DOUBLE,
                    competency_level TEXT(50)
                )
            """)
            conn.commit()
            console.print(" [green]- tbl_dual_competency created[/green]")
        else:
            console.print(" [yellow]- tbl_dual_competency already exists[/yellow]")
        
        conn.close()


    @staticmethod
    def get_competency_level(gpa):
        if gpa is None or not isinstance(gpa, (int, float)):
            return None
        gpa = float(gpa)
        if 1.0 <= gpa <= 1.4999:   return "Grade A (Excellent)"
        if 1.5 <= gpa <= 2.4999:   return "Grade B (Very Good)"
        if 2.5 <= gpa <= 3.4999:   return "Grade C (Good)"
        if 3.5 <= gpa <= 4.4999:   return "Grade D (Average)"
        if 4.5 <= gpa <= 5.4999:   return "Grade E (Satisfactory)"
        if 5.5 <= gpa <= 6.4999:   return "Grade S (Poor)"
        if 6.5 <= gpa <= 7.0:      return "Grade F (Fail)"
        return None


    def merge_and_average_exams(self, df1, df2):
        console.print("\n[bold cyan]================================================================================[/bold cyan]")
        console.print("[bold cyan]  MERGING EXAMS AND COMPUTING AVERAGES[/bold cyan]")
        console.print("[bold cyan]================================================================================[/bold cyan]\n")
        
        # Start with exam 2 base data
        merged = df2[['student_id', 'full_name', 'sex', 'comb_id']].copy()
        merged = merged.set_index('student_id')
        
        # Create lookup dataframes
        df1_indexed = df1.set_index('student_id')
        df2_indexed = df2.set_index('student_id')
        
        # Add all _1 and _2 fields for subjects
        for sub in self.valid_subjects:
            # Exam 1 - marks and metadata
            if sub in df1.columns:
                merged[f'{sub}_1'] = df1_indexed[sub]
                if f'{sub}_grade_1' in df1_indexed.columns:
                    merged[f'{sub}_grade_1'] = df1_indexed[f'{sub}_grade_1']
                if f'{sub}_pos_1' in df1_indexed.columns:
                    merged[f'{sub}_pos_1'] = df1_indexed[f'{sub}_pos_1']
                if f'{sub}_out_of_1' in df1_indexed.columns:
                    merged[f'{sub}_out_of_1'] = df1_indexed[f'{sub}_out_of_1']
            
            # Exam 2 - marks and metadata
            if sub in df2.columns:
                merged[f'{sub}_2'] = df2_indexed[sub]
                if f'{sub}_grade_2' in df2_indexed.columns:
                    merged[f'{sub}_grade_2'] = df2_indexed[f'{sub}_grade_2']
                if f'{sub}_pos_2' in df2_indexed.columns:
                    merged[f'{sub}_pos_2'] = df2_indexed[f'{sub}_pos_2']
                if f'{sub}_out_of_2' in df2_indexed.columns:
                    merged[f'{sub}_out_of_2'] = df2_indexed[f'{sub}_out_of_2']
            
            # Average marks
            merged[sub] = merged[[f'{sub}_1', f'{sub}_2']].mean(axis=1, skipna=True)
        
        # Add historic fields with suffixes
        for field in self.DB_UPDATE_FIELDS:
            field_1 = f'{field}_1'
            field_2 = f'{field}_2'
            
            if field_1 in df1_indexed.columns:
                merged[field_1] = df1_indexed[field_1]
            
            if field_2 in df2_indexed.columns:
                merged[field_2] = df2_indexed[field_2]
        
        # Reset index to get student_id back as a column
        merged = merged.reset_index()
        
        # Add combo metadata (exam_id is not needed in this table; combo_id identifies the combined exam)
        merged['combo_id'] = self.combo_id
        
        console.print(f" [green]- Merged data: {len(merged)} students[/green]")
        console.print(f" [green]- Columns in merged: {len(merged.columns)}[/green]")
        
        return merged


    def ensure_all_insert_columns_exist(self):
        """
        Ensure ALL required columns exist in tbl_dual_combined_results before insert.
        This handles dynamic subjects - if new subjects are found, all their columns are created.
        For example, if 'math' is a new valid subject, it creates:
        - math, math_1, math_2 (marks)
        - math_grade, math_grade_1, math_grade_2 (grades)
        - math_pos, math_pos_1, math_pos_2 (positions)
        - math_out_of, math_out_of_1, math_out_of_2 (out_of counts)
        """
        console.print("\n[bold cyan]Ensuring all required columns exist in tbl_dual_combined_results[/bold cyan]")
        conn = self.connect_db()
        cursor = conn.cursor()
        
        # Verify table exists
        tables = [table.table_name for table in cursor.tables(tableType='TABLE')]
        if 'tbl_dual_combined_results' not in tables:
            console.print(" [red]ERROR: tbl_dual_combined_results table does not exist![/red]")
            conn.close()
            raise ValueError("tbl_dual_combined_results table must exist before insert")
        
        # Get existing columns
        try:
            cursor.execute("SELECT TOP 1 * FROM tbl_dual_combined_results")
            existing_cols = [desc[0].lower() for desc in cursor.description]
        except Exception as e:
            console.print(f" [yellow]Could not read table structure: {e}[/yellow]")
            existing_cols = []
        
        # Base fields that should always exist
        base_fields = [
            ('combo_id', 'TEXT(100)'),
            ('exam_id_1', 'TEXT(50)'),
            ('exam_id_2', 'TEXT(50)'),
            ('student_id', 'TEXT(50)'),
            ('full_name', 'TEXT(100)'),
            ('sex', 'TEXT(10)'),
            ('comb_id', 'TEXT(50)'),
            ('division', 'TEXT(10)'),
            ('points', 'INTEGER'),
            ('subject_count', 'INTEGER'),
            ('total_marks', 'DOUBLE'),
            ('gpa', 'DOUBLE'),
            ('position_comb', 'INTEGER'),
            ('position_school', 'INTEGER'),
            ('out_of_comb', 'INTEGER'),
            ('out_of_school', 'INTEGER'),
            ('first', 'DOUBLE'),
            ('second', 'DOUBLE'),
            ('third', 'DOUBLE'),
            ('avg_grade', 'TEXT(5)'),
            ('avg_marks', 'DOUBLE'),
            ('subject_count_all', 'INTEGER'),
            ('necta_results', 'MEMO'),
            ('necta_results_marks', 'MEMO'),
            ('class_id', 'TEXT(50)'),
            ('processed_date', 'DATETIME'),
        ]
        
        # Historic fields with _1 and _2 suffixes
        historic_fields = []
        for field in self.DB_UPDATE_FIELDS:
            if field == 'division':
                historic_fields.extend([(f'{field}_1', 'TEXT(10)'), (f'{field}_2', 'TEXT(10)')])
            elif field in ['points', 'subject_count', 'subject_count_all', 'position_comb', 'position_school', 
                          'out_of_comb', 'out_of_school']:
                historic_fields.extend([(f'{field}_1', 'INTEGER'), (f'{field}_2', 'INTEGER')])
            elif field in ['total_marks', 'gpa', 'first', 'second', 'third', 'avg_marks']:
                historic_fields.extend([(f'{field}_1', 'DOUBLE'), (f'{field}_2', 'DOUBLE')])
            elif field == 'avg_grade':
                historic_fields.extend([(f'{field}_1', 'TEXT(5)'), (f'{field}_2', 'TEXT(5)')])
            elif field in ['necta_results', 'necta_results_marks']:
                historic_fields.extend([(f'{field}_1', 'MEMO'), (f'{field}_2', 'MEMO')])
        
        # Subject fields - dynamically created for each valid subject
        subject_fields = []
        for sub in self.valid_subjects:
            # Combined (averaged) - 4 columns per subject
            subject_fields.extend([
                (sub, 'DOUBLE'),
                (f'{sub}_grade', 'TEXT(5)'),
                (f'{sub}_pos', 'INTEGER'),
                (f'{sub}_out_of', 'INTEGER')
            ])
            # Exam 1 - 4 columns per subject
            subject_fields.extend([
                (f'{sub}_1', 'DOUBLE'),
                (f'{sub}_grade_1', 'TEXT(5)'),
                (f'{sub}_pos_1', 'INTEGER'),
                (f'{sub}_out_of_1', 'INTEGER')
            ])
            # Exam 2 - 4 columns per subject
            subject_fields.extend([
                (f'{sub}_2', 'DOUBLE'),
                (f'{sub}_grade_2', 'TEXT(5)'),
                (f'{sub}_pos_2', 'INTEGER'),
                (f'{sub}_out_of_2', 'INTEGER')
            ])
        
        # Combine all fields
        all_fields = base_fields + historic_fields + subject_fields
        
        # Check and create missing columns
        created_count = 0
        for col_name, col_type in all_fields:
            if col_name.lower() not in existing_cols:
                try:
                    cursor.execute(
                        f"ALTER TABLE tbl_dual_combined_results "
                        f"ADD COLUMN {self.bracket_field(col_name)} {col_type}"
                    )
                    conn.commit()
                    existing_cols.append(col_name.lower())
                    created_count += 1
                    console.print(f" [green]+ Added column: {col_name}[/green]")
                except Exception as e:
                    console.print(f" [yellow]Column {col_name} might already exist: {e}[/yellow]")
        
        if created_count > 0:
            console.print(f" [green]- Created {created_count} new columns[/green]")
        else:
            console.print(f" [yellow]- All required columns already exist[/yellow]")
        
        conn.close()
        return all_fields

    def save_to_dual_combined_results(self):
        console.print("\n[bold cyan]================================================================================[/bold cyan]")
        console.print("[bold cyan]  SAVING TO DATABASE: tbl_dual_combined_results[/bold cyan]")
        console.print("[bold cyan]================================================================================[/bold cyan]\n")
        
        # CRITICAL: Ensure all columns exist before attempting insert
        all_required_fields = self.ensure_all_insert_columns_exist()
        
        conn = self.connect_db()
        cursor = conn.cursor()
        
        # Delete existing records for this combo
        cursor.execute("DELETE FROM tbl_dual_combined_results WHERE combo_id = ?", (self.combo_id,))
        conn.commit()
        
        # Prepare insert - base fields for combined results
        base_fields = ['combo_id', 'exam_id_1', 'exam_id_2', 'student_id', 'full_name', 'sex', 'comb_id',
                    'division', 'points', 'subject_count', 'total_marks', 'gpa',
                    'position_comb', 'position_school', 'out_of_comb', 'out_of_school',
                    'first', 'second', 'third', 'avg_grade', 'avg_marks', 'subject_count_all',
                    'necta_results', 'necta_results_marks', 'class_id', 'processed_date']
        
        # Subject fields - combined, exam 1, and exam 2
        subject_fields = []
        for sub in self.valid_subjects:
            # Combined (averaged)
            subject_fields.extend([sub, f'{sub}_grade', f'{sub}_pos', f'{sub}_out_of'])
            # Exam 1
            subject_fields.extend([f'{sub}_1', f'{sub}_grade_1', f'{sub}_pos_1', f'{sub}_out_of_1'])
            # Exam 2
            subject_fields.extend([f'{sub}_2', f'{sub}_grade_2', f'{sub}_pos_2', f'{sub}_out_of_2'])
        
        # Historic fields from both exams
        historic_fields = []
        for field in self.DB_UPDATE_FIELDS:
            historic_fields.extend([f'{field}_1', f'{field}_2'])
        
        # Combine all fields
        all_fields = base_fields + subject_fields + historic_fields
        
        # Filter to columns that actually exist in df OR are special fields we provide
        insert_fields = []
        for f in all_fields:
            if f in self.df.columns:
                insert_fields.append(f)
            elif f in ('combo_id', 'exam_id_1', 'exam_id_2', 'class_id', 'processed_date'):
                insert_fields.append(f)
        
        console.print(f" [yellow]- Inserting {len(insert_fields)} fields[/yellow]")
        
        # Create SQL query
        placeholders = ', '.join(['?' for _ in insert_fields])
        field_str = ', '.join([self.bracket_field(f) for f in insert_fields])
        
        insert_query = f"INSERT INTO tbl_dual_combined_results ({field_str}) VALUES ({placeholders})"
        
        # Get current date/time for processed_date
        current_datetime = datetime.now()
        
        # Insert records
        for _, row in tqdm(self.df.iterrows(), total=len(self.df), desc="Inserting records"):
            values = []
            for field in insert_fields:
                if field == 'combo_id':
                    values.append(self.combo_id)
                elif field == 'exam_id_1':
                    values.append(self.exam_id_1)
                elif field == 'exam_id_2':
                    values.append(self.exam_id_2)
                elif field == 'class_id':
                    values.append(self.class_id)
                elif field == 'processed_date':
                    values.append(current_datetime)
                else:
                    val = row.get(field)
                    # Convert numpy/pandas types to native Python types
                    if pd.isna(val):
                        values.append(None)
                    elif isinstance(val, (np.integer, pd.Int64Dtype)):
                        values.append(int(val))
                    elif isinstance(val, (np.floating, float)):
                        values.append(float(val))
                    else:
                        values.append(val)
            
            cursor.execute(insert_query, tuple(values))
        
        conn.commit()
        conn.close()
        
        console.print(f" [bold green][OK] Saved {len(self.df)} records to tbl_dual_combined_results[/bold green]")


    def debug_compare_results(self, sample_size: int = 15):
        """
        Enhanced diagnostic comparison:
        - Compares ALL key metrics between exam 1 (_1), exam 2 (_2), and combined.
        - Shows division, points, GPA, positions (school & comb), avg_marks, etc.
        - Prints beautiful Rich tables with statistics.
        """
        if self.df is None or self.df1 is None or self.df2 is None:
            console.print("[yellow][WARN] debug_compare_results: missing dataframes; skipping comparison.[/yellow]")
            return

        console.print("\n[bold magenta]================================================================================[/bold magenta]")
        console.print("[bold magenta]  COMPREHENSIVE COMPARISON: EXAM 1 vs EXAM 2 vs COMBINED[/bold magenta]")
        console.print("[bold magenta]================================================================================[/bold magenta]\n")

        # Extended column lists
        cols_1 = ['student_id', 'full_name', 'comb_id', 'division_1', 'points_1', 'avg_marks_1', 
                  'gpa_1', 'position_school_1', 'position_comb_1', 'subject_count_1', 'subject_count_all_1']
        cols_2 = ['student_id', 'division_2', 'points_2', 'avg_marks_2', 
                  'gpa_2', 'position_school_2', 'position_comb_2', 'subject_count_2', 'subject_count_all_2']
        cols_c = ['student_id', 'division', 'points', 'avg_marks', 
                  'gpa', 'position_school', 'position_comb', 'subject_count', 'subject_count_all']

        # Check columns exist
        missing_cols = []
        for col in cols_1 + cols_2 + cols_c:
            if col not in (self.df1.columns.union(self.df2.columns).union(self.df.columns)):
                missing_cols.append(col)
        
        if missing_cols:
            console.print(f"[yellow][WARN] Missing columns: {missing_cols[:5]}... (some comparisons may be limited)[/yellow]")
            cols_1 = [c for c in cols_1 if c in self.df1.columns]
            cols_2 = [c for c in cols_2 if c in self.df2.columns]
            cols_c = [c for c in cols_c if c in self.df.columns]

        # Build comparison dataframe
        d1 = self.df1[cols_1].copy()
        d2 = self.df2[cols_2].copy()
        dc = self.df[cols_c].copy()

        merged = d1.merge(d2, on='student_id', how='inner').merge(dc, on='student_id', how='inner')

        if merged.empty:
            console.print("[red][ERROR] No overlapping students to compare.[/red]")
            return

        # Compute diagnostics
        merged['avg_marks_mean_1_2'] = merged[['avg_marks_1', 'avg_marks_2']].mean(axis=1)
        merged['avg_marks_diff'] = merged['avg_marks'] - merged['avg_marks_mean_1_2']
        
        # Points difference
        if 'points_1' in merged.columns and 'points_2' in merged.columns and 'points' in merged.columns:
            merged['points_mean_1_2'] = merged[['points_1', 'points_2']].mean(axis=1)
            merged['points_diff'] = merged['points'] - merged['points_mean_1_2']

        # ============================================================
        # SUMMARY STATISTICS TABLE
        # ============================================================
        console.print("[bold cyan]SUMMARY STATISTICS[/bold cyan]")
        stats_table = Table(box=box.ROUNDED, show_header=True, header_style="bold cyan")
        stats_table.add_column("METRIC", style="yellow", justify="left")
        stats_table.add_column("EXAM 1", justify="right")
        stats_table.add_column("EXAM 2", justify="right")
        stats_table.add_column("COMBINED", justify="right")
        stats_table.add_column("DIFFERENCE", justify="right", style="magenta")

        # Avg marks stats - 2 significant figures
        if 'avg_marks_1' in merged.columns and 'avg_marks_2' in merged.columns and 'avg_marks' in merged.columns:
            avg1_mean = merged['avg_marks_1'].mean()
            avg2_mean = merged['avg_marks_2'].mean()
            avgc_mean = merged['avg_marks'].mean()
            stats_table.add_row(
                "Avg Marks (Mean)",
                self.format_sig_figs(avg1_mean, 2),
                self.format_sig_figs(avg2_mean, 2),
                self.format_sig_figs(avgc_mean, 2),
                f"Diff: {merged['avg_marks_diff'].mean():.2f}"
            )
            std1 = merged['avg_marks_1'].std()
            std2 = merged['avg_marks_2'].std()
            stdc = merged['avg_marks'].std()
            stats_table.add_row(
                "Avg Marks (Std)",
                self.format_sig_figs(std1, 2),
                self.format_sig_figs(std2, 2),
                self.format_sig_figs(stdc, 2),
                f"Range: [{merged['avg_marks_diff'].min():.2f}, {merged['avg_marks_diff'].max():.2f}]"
            )

        # Points stats
        if 'points_1' in merged.columns and 'points_2' in merged.columns and 'points' in merged.columns:
            valid_points_1 = merged['points_1'].dropna()
            valid_points_2 = merged['points_2'].dropna()
            valid_points_c = merged['points'].dropna()
            if len(valid_points_1) > 0 and len(valid_points_2) > 0 and len(valid_points_c) > 0:
                stats_table.add_row(
                    "Points (Mean)",
                    f"{valid_points_1.mean():.2f}",
                    f"{valid_points_2.mean():.2f}",
                    f"{valid_points_c.mean():.2f}",
                    f"Diff: {merged['points_diff'].mean():.4f}" if 'points_diff' in merged.columns else "—"
                )

        # GPA stats - 4 significant figures
        if 'gpa_1' in merged.columns and 'gpa_2' in merged.columns and 'gpa' in merged.columns:
            valid_gpa_1 = merged['gpa_1'].dropna()
            valid_gpa_2 = merged['gpa_2'].dropna()
            valid_gpa_c = merged['gpa'].dropna()
            if len(valid_gpa_1) > 0 and len(valid_gpa_2) > 0 and len(valid_gpa_c) > 0:
                gpa1_mean = valid_gpa_1.mean()
                gpa2_mean = valid_gpa_2.mean()
                gpac_mean = valid_gpa_c.mean()
                stats_table.add_row(
                    "GPA (Mean)",
                    self.format_sig_figs(gpa1_mean, 4),
                    self.format_sig_figs(gpa2_mean, 4),
                    self.format_sig_figs(gpac_mean, 4),
                    "—"
                )

        stats_table.add_row(
            "Students Compared",
            str(len(merged)),
            str(len(merged)),
            str(len(merged)),
            "[OK]"
        )

        console.print(stats_table)
        console.print(f"\n[cyan]Students compared: {len(merged):,}[/cyan]")
        if 'avg_marks_diff' in merged.columns:
            console.print(f"[cyan]Avg marks difference range: [{merged['avg_marks_diff'].min():.4f}, {merged['avg_marks_diff'].max():.4f}][/cyan]")

        # ============================================================
        # DETAILED COMPARISON TABLE (SAMPLE)
        # ============================================================
        console.print(f"\n[bold cyan]DETAILED SAMPLE COMPARISON (First {sample_size} Students)[/bold cyan]")
        
        sample = merged.head(sample_size).copy()
        
        detail_table = Table(
            title=f"[bold white] EXAM 1 vs EXAM 2 vs COMBINED — Sample View [/bold white]",
            box=box.ROUNDED,
            expand=True,
            show_lines=False
        )
        
        detail_table.add_column("NAME", justify="left", style="cyan", width=22, no_wrap=False)
        
        # Division columns
        detail_table.add_column("DIV_1", justify="center", style="yellow", width=6)
        detail_table.add_column("DIV_2", justify="center", style="yellow", width=6)
        detail_table.add_column("DIV_C", justify="center", style="bold yellow", width=6)
        
        # Points columns
        detail_table.add_column("PTS_1", justify="right", style="green", width=6)
        detail_table.add_column("PTS_2", justify="right", style="green", width=6)
        detail_table.add_column("PTS_C", justify="right", style="bold green", width=6)
        
        # Avg marks columns - 2 sig figs
        detail_table.add_column("AVG_1", justify="right", style="blue", width=7)
        detail_table.add_column("AVG_2", justify="right", style="blue", width=7)
        detail_table.add_column("AVG_C", justify="right", style="bold blue", width=7)
        
        # Position columns
        detail_table.add_column("POS_SCH_1", justify="right", style="red", width=9)
        detail_table.add_column("POS_SCH_2", justify="right", style="red", width=9)
        detail_table.add_column("POS_SCH_C", justify="right", style="bold red", width=9)
        detail_table.add_column("POS_COMB_1", justify="right", style="yellow", width=10)
        detail_table.add_column("POS_COMB_2", justify="right", style="yellow", width=10)
        detail_table.add_column("POS_COMB_C", justify="right", style="bold yellow", width=10)

        for _, row in sample.iterrows():
            name = str(row.get('full_name', ''))[:20] if 'full_name' in row else "—"
            
            # Format values with proper handling
            def fmt(val, sig_figs=None):
                if pd.isna(val):
                    return "—"
                if isinstance(val, (int, np.integer)):
                    return str(int(val))
                if sig_figs:
                    return self.format_sig_figs(float(val), sig_figs)
                return f"{float(val):.2f}"
            
            detail_table.add_row(
                name,
                str(row.get('division_1', '—')) if pd.notna(row.get('division_1')) else "—",
                str(row.get('division_2', '—')) if pd.notna(row.get('division_2')) else "—",
                str(row.get('division', '—')) if pd.notna(row.get('division')) else "—",
                fmt(row.get('points_1'), 0),
                fmt(row.get('points_2'), 0),
                fmt(row.get('points'), 0),
                fmt(row.get('avg_marks_1'), 2),
                fmt(row.get('avg_marks_2'), 2),
                fmt(row.get('avg_marks'), 2),
                fmt(row.get('position_school_1'), 0),
                fmt(row.get('position_school_2'), 0),
                fmt(row.get('position_school'), 0),
                fmt(row.get('position_comb_1'), 0),
                fmt(row.get('position_comb_2'), 0),
                fmt(row.get('position_comb'), 0),
            )

        console.print(detail_table)
        
        # ============================================================
        # DIVISION DISTRIBUTION COMPARISON
        # ============================================================
        console.print(f"\n[bold cyan]DIVISION DISTRIBUTION COMPARISON[/bold cyan]")
        div_table = Table(box=box.ROUNDED, show_header=True, header_style="bold cyan")
        div_table.add_column("DIVISION", justify="center", style="yellow")
        div_table.add_column("EXAM 1", justify="right", style="green")
        div_table.add_column("EXAM 2", justify="right", style="green")
        div_table.add_column("COMBINED", justify="right", style="bold green")
        
        if 'division_1' in merged.columns and 'division_2' in merged.columns and 'division' in merged.columns:
            all_divs = set(merged['division_1'].dropna().unique()) | \
                      set(merged['division_2'].dropna().unique()) | \
                      set(merged['division'].dropna().unique())
            for div in sorted(all_divs, key=lambda x: {'I': 1, 'II': 2, 'III': 3, 'IV': 4, '0': 5, 'ABS': 6, 'INC': 7}.get(x, 99)):
                count_1 = (merged['division_1'] == div).sum() if 'division_1' in merged.columns else 0
                count_2 = (merged['division_2'] == div).sum() if 'division_2' in merged.columns else 0
                count_c = (merged['division'] == div).sum() if 'division' in merged.columns else 0
                div_table.add_row(str(div), str(count_1), str(count_2), str(count_c))
        
        console.print(div_table)
        
        console.print("\n[bold green][OK] COMPREHENSIVE COMPARISON COMPLETE [OK][/bold green]\n")


    def update_dual_competency(self):
        console.print("\n[bold cyan]Processing Dual Subject Competency Report[/bold cyan]")
        
        conn = self.connect_db()
        cur = conn.cursor()
        cur.execute("DELETE FROM tbl_dual_competency WHERE combo_id = ?", (self.combo_id,))
        conn.commit()
        
        GPA_COLORS = {
            (1.0, 1.4999): "#00A82A",
            (1.5, 2.4999): "#1FEE0B",
            (2.5, 3.4999): "#1FEE0B",
            (3.5, 4.4999): "#DEF043",
            (4.5, 5.4999): "#DEF043",
            (5.5, 6.4999): "#FF772F",
            (6.5, 7.0):    "#FF272F",
        }
        
        records = []
        rows = []
        
        total_A = total_B = total_C = total_D = total_E = total_S = total_F = 0
        total_students = 0
        total_gpa = 0
        
        for sub in self.valid_subjects:
            grade_col = f"{sub}_grade"
            
            # Check if grade column exists and has data
            if grade_col not in self.df.columns:
                console.print(f" [yellow][SKIP] Skipping {sub}: grade column not found[/yellow]")
                continue
                
            # Filter students who attempted this subject (have valid marks)
            sat = self.df[self.df[sub].notna()].copy()
            
            if sat.empty:
                console.print(f" [yellow][SKIP] Skipping {sub}: no students attempted[/yellow]")
                continue
            
            # Count grades
            c = sat[grade_col].value_counts()
            A, B, C_grade, D, E, S, F = [int(c.get(g, 0)) for g in "ABCDEFS"]
            total = len(sat)
            
            # Calculate GPA
            gpa = (A*1 + B*2 + C_grade*3 + D*4 + E*5 + S*6 + F*7) / total if total > 0 else None
            
            if gpa is None:
                console.print(f" [yellow][SKIP] Skipping {sub}: could not calculate GPA[/yellow]")
                continue
                
            level = self.get_competency_level(gpa)
            
            # Get subject metadata
            meta = self.comb_metadata[self.comb_metadata['subject_short'].str.lower() == sub]
            if meta.empty:
                console.print(f" [yellow][SKIP] Skipping {sub}: metadata not found[/yellow]")
                continue
                
            name = meta.iloc[0]['subject_name'].upper()
            serial = int(meta.iloc[0]['subject_serial'])
            
            # Store record for database
            records.append((
                str(self.combo_id), serial, 
                A, B, C_grade, D, E, S, F, 
                total, A+B+C_grade+D+E, S+F, 
                round(gpa, 4), level
            ))
            
            # Determine background color
            bg = None
            for (lo, hi), color in GPA_COLORS.items():
                if lo <= gpa <= hi:
                    bg = color
                    break
            
            # Store row for table display
            rows.append({
                'name': name, 
                'A': A, 'B': B, 'C': C_grade, 'D': D, 'E': E, 'S': S, 'F': F, 
                'total': total, 
                'gpa': gpa, 
                'level': level, 
                'bg_color': bg
            })
            
            # Update totals
            total_A += A
            total_B += B
            total_C += C_grade
            total_D += D
            total_E += E
            total_S += S
            total_F += F
            total_students += total
            total_gpa += gpa
        
        # Insert records into database
        if records:
            cur.executemany("""
                INSERT INTO tbl_dual_competency 
                (combo_id, subject_serial, A_s, B_s, C_s, D_s, E_s, S_s, F_s, 
                total, pass, fail, gpa, competency_level) 
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, records)
            conn.commit()
            console.print(f" [green][OK] Inserted {len(records)} subject records into tbl_dual_competency[/green]")
        else:
            console.print("[red]- No competency data to save[/red]")
        
        conn.close()
        
        if not rows:
            console.print("[red]No competency data to display[/red]")
            return
        
        # Sort rows by GPA (best first)
        rows.sort(key=lambda x: x['gpa'])
        
        # Create and display table
        table = Table(
            title=f"[bold white on #1e1b4b] DUAL COMPETENCY REPORT – {self.combo_id} [/]", 
            box=box.DOUBLE_EDGE, 
            expand=True
        )
        
        table.add_column("SUBJECT", justify="left", style="cyan")
        table.add_column("A", justify="center")
        table.add_column("B", justify="center")
        table.add_column("C", justify="center")
        table.add_column("D", justify="center")
        table.add_column("E", justify="center")
        table.add_column("S", justify="center")
        table.add_column("F", justify="center")
        table.add_column("TOTAL", justify="center", style="bold")
        table.add_column("PASS", justify="center", style="green")
        table.add_column("FAIL", justify="center", style="red")
        table.add_column("GPA", justify="center", style="bold yellow")
        table.add_column("COMPETENCY LEVEL", justify="center")
        
        for r in rows:
            if r['bg_color']:
                competency_cell = Text(r['level'] or "N/A", style=f"bold white on {r['bg_color']}")
            else:
                competency_cell = r['level'] or "N/A"
            
            table.add_row(
                r['name'], 
                str(r['A']), 
                str(r['B']), 
                str(r['C']), 
                str(r['D']),
                str(r['E']), 
                str(r['S']), 
                str(r['F']), 
                str(r['total']),
                str(r['A'] + r['B'] + r['C'] + r['D'] + r['E']),
                str(r['S'] + r['F']),
                self.format_sig_figs(r['gpa'], 4),
                competency_cell
            )
        
        # Add totals row
        total_pass = total_A + total_B + total_C + total_D + total_E
        total_fail = total_S + total_F
        avg_gpa = total_gpa / len(rows) if rows else 0
        
        table.add_row(
            "TOTAL/AVERAGE",
            str(total_A),
            str(total_B),
            str(total_C),
            str(total_D),
            str(total_E),
            str(total_S),
            str(total_F),
            str(total_students),
            str(total_pass),
            str(total_fail),
            self.format_sig_figs(avg_gpa, 4),
            "-",
            style="bold cyan"
        )
        
        console.print(table)
        console.print(f" [bold green][OK] Competency report complete — {len(records)} subjects processed[/bold green]")




    def run(self):
        """
        Main execution method - processes dual exams exactly like single AlevelProcessor.
        Ensures all tables and columns exist before operations.
        """
        try:
            # STEP 1: Create all required tables if missing
            self.create_dual_exam_table()
            self.create_dual_combined_results_table()
            self.create_dual_competency_table()  # Create early to ensure it exists
            
            # STEP 2: Fetch data for both exams
            df1, df2 = self.fetch_dual_exam_data()
            
            # STEP 3: Ensure subject columns exist (for early processing)
            # Note: ensure_all_insert_columns_exist() will be called before insert to handle ALL columns
            self.ensure_subject_columns_exist(self.valid_subjects, '_1')
            self.ensure_subject_columns_exist(self.valid_subjects, '_2')
            self.ensure_subject_columns_exist(self.valid_subjects, '')
            
            # STEP 4: Process both exams separately (identical logic to AlevelProcessor with suffix support)
            df1 = self.process_single_exam(df1, self.exam_id_1, '_1')
            df2 = self.process_single_exam(df2, self.exam_id_2, '_2')

            # Keep processed per-exam dataframes for debugging/comparison
            self.df1 = df1
            self.df2 = df2
            
            # STEP 5: Merge and average exams
            df_combined = self.merge_and_average_exams(df1, df2)
            
            # STEP 6: Process combined exam (identical logic to AlevelProcessor, no suffix)
            self.process_combined_exam(df_combined)

            # STEP 7: Optional comparison diagnostics
            self.debug_compare_results()
            
            # STEP 8: Save to database (ensure_all_insert_columns_exist() called here)
            self.save_to_dual_combined_results()
            
            # STEP 9: Update competency report
            self.update_dual_competency()
            
            console.print("\n[bold green]================================================================================[/bold green]")
            console.print("[bold green]  [OK]  DUAL EXAM PROCESSING COMPLETE  [OK][/bold green]")
            console.print("[bold green]================================================================================[/bold green]\n")
            
        except Exception as e:
            console.print(f"[bold red]Error: {e}[/bold red]")
            raise


if __name__ == "__main__":
    processor = DualAlevelProcessor(
    exam_id_1="MID520250403",
    exam_id_2="ANN520250526",
    db_path=r"C:\Kiyabo App\backend\Kiyabo App Backend v4.0.0.accdb"
    )
    processor.run()