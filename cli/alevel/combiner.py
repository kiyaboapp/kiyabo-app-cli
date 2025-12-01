import pandas as pd
import pyodbc
import numpy as np
from tqdm import tqdm
from rich.console import Console
from rich.table import Table
from rich import box
from rich.text import Text

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
        "first", "second", "third", 'avg_grade',
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
                 rank_method: str = 'min'):
        self.exam_id_1 = exam_id_1
        self.exam_id_2 = exam_id_2
        self.db_path = db_path
        self.combo_id = f"{exam_id_1}_{exam_id_2}"
        
        self.sort_columns = sort_columns or ["points", "avg_marks", "subject_count"]
        self.include_inc = include_inc
        self.rank_method = rank_method
        self.sort_ascending = [self.SORT_ASCENDING[col] for col in self.sort_columns]
        
        self.conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + db_path + ';'
        self.df = None
        self.valid_subjects = []
        self.comb_metadata = None
        self.subject_to_user = {}
        self.db_columns = []
        
        self.exam_name_1 = exam_name_1
        self.exam_name_2 = exam_name_2
        
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
    
    def connect_db(self):
        return pyodbc.connect(self.conn_str)
    
    def create_dual_exam_table(self):
        console.print("\n[bold cyan]Creating tbl_dual_exam if not exists[/bold cyan]")
        conn = self.connect_db()
        cursor = conn.cursor()
        
        tables = [table.table_name for table in cursor.tables(tableType='TABLE')]
        exists = 'tbl_dual_exam' in tables
        
        if not exists:
            cursor.execute("""
                CREATE TABLE tbl_dual_exam (
                    combo_id TEXT(100) PRIMARY KEY,
                    exam_id_1 TEXT(50),
                    exam_id_2 TEXT(50),
                    exam_name_1 TEXT(100),
                    exam_name_2 TEXT(100),
                    class_id TEXT(50)
                )
            """)
            conn.commit()
            console.print(" [green]- tbl_dual_exam created[/green]")
        else:
            console.print(" [yellow]- tbl_dual_exam already exists[/yellow]")
        
        # Get exam names if not provided
        if not self.exam_name_1 or not self.exam_name_2:
            exam_query = "SELECT exam_id, exam_name FROM tbl_student_exams WHERE exam_id IN (?, ?)"
            exam_df = pd.read_sql(exam_query, conn, params=[self.exam_id_1, self.exam_id_2])
            
            if not self.exam_name_1:
                row = exam_df[exam_df['exam_id'] == self.exam_id_1]
                self.exam_name_1 = row.iloc[0]['exam_name'] if not row.empty else self.exam_id_1
            
            if not self.exam_name_2:
                row = exam_df[exam_df['exam_id'] == self.exam_id_2]
                self.exam_name_2 = row.iloc[0]['exam_name'] if not row.empty else self.exam_id_2
        
        # Insert/update combo record
        cursor.execute("SELECT COUNT(*) FROM tbl_dual_exam WHERE combo_id = ?", (self.combo_id,))
        if cursor.fetchone()[0] == 0:
            cursor.execute("""
                INSERT INTO tbl_dual_exam (combo_id, exam_id_1, exam_id_2, exam_name_1, exam_name_2, class_id)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (self.combo_id, self.exam_id_1, self.exam_id_2, self.exam_name_1, self.exam_name_2, None))
        else:
            cursor.execute("""
                UPDATE tbl_dual_exam 
                SET exam_name_1=?, exam_name_2=?
                WHERE combo_id=?
            """, (self.exam_name_1, self.exam_name_2, self.combo_id))
        
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
                    student_id TEXT(50),
                    exam_id TEXT(50),
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
                    necta_results_marks_2 MEMO
                )
            """)
            conn.commit()
            console.print(" [green]- tbl_dual_combined_results created[/green]")
        else:
            console.print(" [yellow]- tbl_dual_combined_results exists[/yellow]")
        
        conn.close()




    def ensure_subject_columns_exist(self, subjects_list, suffix=''):
        console.print(f"\n[bold cyan]Ensuring subject columns exist (suffix='{suffix}')[/bold cyan]")
        conn = self.connect_db()
        cursor = conn.cursor()
        
        # Get existing columns
        cursor.execute("SELECT TOP 1 * FROM tbl_dual_combined_results")
        existing_cols = [desc[0].lower() for desc in cursor.description]
        
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
                    except Exception as e:
                        console.print(f" [yellow]Column {col_name} might exist: {e}[/yellow]")
        
        conn.close()
    
    def fetch_dual_exam_data(self):
        console.print("\n[bold cyan]Fetching data for both exams[/bold cyan]")
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
        console.print(f"\n[bold cyan]Processing exam {exam_id} (suffix={suffix})[/bold cyan]")
        
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
        
        console.print(f" [green]- Exam {exam_id} processed[/green]")
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
        valid_students = df[
            (df[f'division{suffix}'] != 'ABS') &
            (df[f'points{suffix}'].notna()) &
            (df[f'avg_marks{suffix}'].notna())
        ].copy()
        
        if valid_students.empty:
            return df
        
        sort_cols_suffixed = [f"{col}{suffix}" for col in self.sort_columns]
        sorted_df = valid_students.sort_values(by=sort_cols_suffixed, ascending=self.sort_ascending).copy()
        
        primary_col = sort_cols_suffixed[0]
        ascending_primary = self.sort_ascending[0]
        
        sorted_df[f'position_school{suffix}'] = sorted_df[primary_col].rank(method=self.rank_method, ascending=ascending_primary).astype('Int64')
        sorted_df[f'out_of_school{suffix}'] = len(sorted_df)
        
        def rank_comb(group):
            group = group.sort_values(by=sort_cols_suffixed, ascending=self.sort_ascending)
            group[f'position_comb{suffix}'] = group[primary_col].rank(method=self.rank_method, ascending=ascending_primary).astype('Int64')
            group[f'out_of_comb{suffix}'] = len(group)
            return group
        
        ranked_df = sorted_df.groupby('comb_id', group_keys=False).apply(rank_comb)
        
        df.loc[ranked_df.index, [f'position_school{suffix}', f'out_of_school{suffix}', f'position_comb{suffix}', f'out_of_comb{suffix}']] = \
            ranked_df[[f'position_school{suffix}', f'out_of_school{suffix}', f'position_comb{suffix}', f'out_of_comb{suffix}']]
        
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
            sub_df[pos_col] = sub_df['rank_temp'].astype('Int64')
            sub_df[out_col] = len(sub_df)
            
            df.loc[sub_df.index, [pos_col, out_col]] = sub_df[[pos_col, out_col]]
        
        return df
    
    def merge_and_average_exams(self, df1, df2):
        console.print("\n[bold cyan]Merging exams and computing averages[/bold cyan]")
        
        # Merge on student_id
        merged = df2[['student_id', 'full_name', 'sex', 'comb_id']].copy()
        
        # Add all _1 and _2 fields
        for sub in self.valid_subjects:
            # Exam 1
            if sub in df1.columns:
                merged[f'{sub}_1'] = df1.set_index('student_id')[sub]
                merged[f'{sub}_grade_1'] = df1.set_index('student_id').get(f'{sub}_grade_1', None)
                merged[f'{sub}_pos_1'] = df1.set_index('student_id').get(f'{sub}_pos_1', None)
                merged[f'{sub}_out_of_1'] = df1.set_index('student_id').get(f'{sub}_out_of_1', None)
            
            # Exam 2
            if sub in df2.columns:
                merged[f'{sub}_2'] = df2.set_index('student_id')[sub]
                merged[f'{sub}_grade_2'] = df2.set_index('student_id').get(f'{sub}_grade_2', None)
                merged[f'{sub}_pos_2'] = df2.set_index('student_id').get(f'{sub}_pos_2', None)
                merged[f'{sub}_out_of_2'] = df2.set_index('student_id').get(f'{sub}_out_of_2', None)
            
            # Average
            merged[sub] = merged[[f'{sub}_1', f'{sub}_2']].mean(axis=1, skipna=True)
        
        # Add historic fields with suffixes
        for field in self.DB_UPDATE_FIELDS:
            if f'{field}_1' in df1.columns:
                merged[f'{field}_1'] = df1.set_index('student_id')[f'{field}_1']
            if f'{field}_2' in df2.columns:
                merged[f'{field}_2'] = df2.set_index('student_id')[f'{field}_2']
        
        merged['combo_id'] = self.combo_id
        merged['exam_id'] = self.combo_id
        
        console.print(f" [green]- Merged data: {len(merged)} students[/green]")
        return merged
    
    def process_combined_exam(self, df_combined):
        console.print("\n[bold cyan]Processing combined (averaged) exam[/bold cyan]")
        
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
        
        console.print(" [green]- Combined exam processed[/green]")
    
    def save_to_dual_combined_results(self):
        console.print("\n[bold cyan]Saving to tbl_dual_combined_results[/bold cyan]")
        
        conn = self.connect_db()
        cursor = conn.cursor()
        
        # Delete existing records for this combo
        cursor.execute("DELETE FROM tbl_dual_combined_results WHERE combo_id = ?", (self.combo_id,))
        conn.commit()
        
        # Prepare insert
        base_fields = ['combo_id', 'student_id', 'exam_id', 'full_name', 'sex', 'comb_id',
                       'division', 'points', 'subject_count', 'total_marks', 'gpa',
                       'position_comb', 'position_school', 'out_of_comb', 'out_of_school',
                       'first', 'second', 'third', 'avg_grade', 'avg_marks', 'subject_count_all',
                       'necta_results', 'necta_results_marks']
        
        
        subject_fields = []
        for sub in self.valid_subjects:
            subject_fields.extend([
            sub, f'{sub}_grade', f'{sub}_pos', f'{sub}_out_of',
            f'{sub}_1', f'{sub}_grade_1', f'{sub}_pos_1', f'{sub}_out_of_1',
            f'{sub}_2', f'{sub}_grade_2', f'{sub}_pos_2', f'{sub}_out_of_2'
            ])

        historic_fields = []
        for field in self.DB_UPDATE_FIELDS:
            historic_fields.extend([f'{field}_1', f'{field}_2'])
        
        all_fields = base_fields + subject_fields + historic_fields
        
        # Filter to columns that exist in df
        insert_fields = [f for f in all_fields if f in self.df.columns or f == 'combo_id' or f == 'exam_id']
        
        placeholders = ', '.join(['?' for _ in insert_fields])
        field_str = ', '.join([self.bracket_field(f) for f in insert_fields])
        
        insert_query = f"INSERT INTO tbl_dual_combined_results ({field_str}) VALUES ({placeholders})"
        
        for _, row in tqdm(self.df.iterrows(), total=len(self.df), desc="Inserting records"):
            values = []
            for field in insert_fields:
                if field == 'combo_id':
                    values.append(self.combo_id)
                elif field == 'exam_id':
                    values.append(self.combo_id)
                else:
                    val = row.get(field)
                    values.append(None if pd.isna(val) else val)
            
            cursor.execute(insert_query, tuple(values))
        
        conn.commit()
        conn.close()
        
        console.print(f" [bold green]✓ Saved {len(self.df)} records to tbl_dual_combined_results[/bold green]")

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
            if f"{sub}_grade" not in self.df.columns: continue
            sat = self.df[self.df[sub].notna()]
            if sat.empty: continue
            
            c = sat[f"{sub}_grade"].value_counts()
            A,B,C,D,E,S,F = [int(c.get(g,0)) for g in "ABCDEFS"]
            total = len(sat)
            gpa = (A*1+B*2+C*3+D*4+E*5+S*6+F*7) / total
            level = self.get_competency_level(gpa)
            
            meta = self.comb_metadata[self.comb_metadata['subject_short'].str.lower() == sub]
            if meta.empty: continue
            name = meta.iloc[0]['subject_name'].upper()
            serial = int(meta.iloc[0]['subject_serial'])
            
            records.append((str(self.combo_id), serial, A,B,C,D,E,S,F, total, A+B+C+D+E, S+F, round(gpa,4), level))
            
            bg = None
            for (lo, hi), color in GPA_COLORS.items():
                if lo <= gpa <= hi:
                    bg = color
                    break
            
            rows.append({
                'name': name, 'A': A, 'B': B, 'C': C, 'D': D, 'E': E, 'S': S, 'F': F, 
                'total': total, 'gpa': gpa, 'level': level, 'bg_color': bg
            })
            
            total_A += A
            total_B += B
            total_C += C
            total_D += D
            total_E += E
            total_S += S
            total_F += F
            total_students += total
            total_gpa += gpa
        
        if records:
            cur.executemany("INSERT INTO tbl_dual_competency (combo_id,subject_serial,A_s,B_s,C_s,D_s,E_s,S_s,F_s,total,pass,fail,gpa,competency_level) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", records)
            conn.commit()
        conn.close()
        
        if not rows:
            console.print("[red]No data[/red]")
            return
        
        rows.sort(key=lambda x: x['gpa'])
        
        table = Table(title=f"[bold white on #1e1b4b] DUAL COMPETENCY REPORT – {self.combo_id} [/]", box=box.DOUBLE_EDGE, expand=True)
        
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
                f"{r['gpa']:.4f}",
                competency_cell
            )
        
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
            f"{avg_gpa:.4f}",
            "-"
        )
        
        console.print(table)
        console.print(f" [bold green]Done — {len(records)} subjects[/bold green]")

    def run(self):
        try:
            self.create_dual_exam_table()
            self.create_dual_combined_results_table()
            
            df1, df2 = self.fetch_dual_exam_data()
            
            # Ensure columns exist for both exams
            self.ensure_subject_columns_exist(self.valid_subjects, '_1')
            self.ensure_subject_columns_exist(self.valid_subjects, '_2')
            self.ensure_subject_columns_exist(self.valid_subjects, '')
            
            # Process both exams separately
            df1 = self.process_single_exam(df1, self.exam_id_1, '_1')
            df2 = self.process_single_exam(df2, self.exam_id_2, '_2')
            
            # Merge and average
            df_combined = self.merge_and_average_exams(df1, df2)
            
            # Process combined exam
            self.process_combined_exam(df_combined)
            
            # Save to database
            self.save_to_dual_combined_results()
            
            # Competency
            self.create_dual_competency_table()
            self.update_dual_competency()
            
            console.print("\n[bold green]✓✓✓ DUAL EXAM PROCESSING COMPLETE ✓✓✓[/bold green]")
            
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