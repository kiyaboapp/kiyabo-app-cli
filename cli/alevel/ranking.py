import pandas as pd
import pyodbc
import numpy as np
from tqdm import tqdm
from rich.console import Console
from rich.table import Table
from rich import box
from rich.text import Text

# ONLY THIS LINE ADDED – MAKES .progress_apply() WORK WITHOUT CHANGING ANY OF YOUR CODE
tqdm.pandas()

console = Console()

class AlevelProcessor:
    """Processes A-Level student exam results with ranking and grading."""
    
    # Class-level constants
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
    
    # Fixed sorting rules
    SORT_ASCENDING = {
        'points_for_rank': True,   # Lower is better
        'avg_marks': False,         # Higher is better
        'subject_count': False,     # Higher is better
    }
    
    def __init__(self, exam_id: str, db_path: str, 
                 sort_columns: list = None,
                 include_inc: bool = True):
        """
        Initialize the A-Level processor.
        
        Args:
            exam_id: The exam identifier (e.g., 'ANN520250526')
            db_path: Path to the Access database file
            sort_columns: List of column names to sort by (default: ["points_for_rank", "avg_marks", "subject_count"])
            include_inc: Whether to include INC status or convert to penalty
        """
        self.exam_id = exam_id
        self.db_path = db_path
        self.sort_columns = sort_columns or ["points_for_rank", "avg_marks", "subject_count"]
        self.include_inc = include_inc
        
        # Build ascending list based on fixed rules
        self.sort_ascending = [self.SORT_ASCENDING[col] for col in self.sort_columns]
        
        # Runtime attributes
        self.conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + db_path + ';'
        self.df = None
        self.valid_subjects = []
        self.comb_metadata = None
        self.subject_to_user = {}
        self.db_columns = []
        
    @staticmethod
    def bracket_field(field: str) -> str:
        """Wrap field name in brackets for SQL queries."""
        return f"[{field}]"
    
    @staticmethod
    def get_grade(marks) -> str:
        """Convert numeric marks to letter grade."""
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
        """Convert points to division."""
        if pd.isna(points) or not isinstance(points, int):
            return None
        points = int(points)
        for r, div in cls.POINTS_TO_DIV.items():
            if points in r:
                return div
        return None
    
    @staticmethod
    def format_marks_for_necta(marks):
        """Format marks for NECTA results string."""
        if pd.isna(marks):
            return ''
        if isinstance(marks, (int, np.integer)) or (isinstance(marks, float) and marks.is_integer()):
            return str(int(marks))
        else:
            return f"{marks:.2f}"
    
    @staticmethod
    def format_grade_for_necta(grade):
        """Format grade for NECTA results string."""
        return f"'{grade}'" if grade else ''
    
    def connect_db(self):
        """Establish database connection."""
        return pyodbc.connect(self.conn_str)
    
    def _display_sample_data(self, potential_subjects):
        """Display sample data in tabular form."""
        console.print("\n[bold magenta]Sample Data Preview (First 5 Students):[/bold magenta]")
        
        # Select columns to display: student info + first few subjects (max 10 columns total)
        display_cols = ['student_id', 'full_name', 'sex', 'comb_id']
        subject_cols = potential_subjects[:10]  # Show first 10 subjects
        display_cols.extend(subject_cols)
        
        sample_df = self.df[display_cols].head(10)
        
        # Create rich table
        table = Table(show_header=True, header_style="bold cyan", box=box.ROUNDED)
        
        # Add columns
        for col in display_cols:
            table.add_column(col.upper(), style="white")
        
        # Add rows
        for _, row in sample_df.iterrows():
            table.add_row(*[str(row[col]) if pd.notna(row[col]) else "-" for col in display_cols])
        
        console.print(table)

    # PURE DISPLAY – NEVER TOUCHES OR CHANGES ANY DATA
    def _preview(self, title: str, cols: list, highlight=None, rows: int = 12):
        table = Table(title=f"[bold gold1 on #1e1b4b]{title}[/]", box=box.DOUBLE_EDGE, border_style="#7209b7", expand=True)
        for col in cols:
            style = "bold bright_yellow" if highlight and col in highlight else "white"
            table.add_column(col.upper().replace("_", " "), style=style, justify="center")
        for _, row in self.df[cols].head(rows).iterrows():
            values = [str(val) if pd.notna(val) else "—" for val in row]
            table.add_row(*values)
        console.print(table)
    
    def detect_columns(self):
        """Detect available columns in the database."""
        console.print("\n[bold cyan]Stage 1: Connecting to Database and Detecting Columns[/bold cyan]")
        conn = self.connect_db()
        
        dummy_query = "SELECT TOP 1 * FROM tbl_student_exam_results WHERE exam_id = ?"
        df_dummy = pd.read_sql(dummy_query, conn, params=[self.exam_id])
        if df_dummy.empty:
            console.print(" [yellow]- No records for exam_id, falling back to general structure.[/yellow]")
            df_dummy = pd.read_sql("SELECT TOP 1 * FROM tbl_student_exam_results", conn)
        
        self.db_columns = [col.lower() for col in df_dummy.columns]
        potential_subjects = [sub for sub in self.SUBJECTS if sub in self.db_columns]
        console.print(f" [green]- Detected {len(potential_subjects)} potential subject mark fields: {potential_subjects}[/green]")
        
        if 'student_id' not in self.db_columns or 'exam_id' not in self.db_columns:
            conn.close()
            raise ValueError("Missing essential fields.")
        
        conn.close()
        return potential_subjects
    
    def fetch_data(self, potential_subjects):
        """Fetch exam data from database."""
        console.print("\n[bold cyan]Stage 2: Building SQL Query and Fetching Data[/bold cyan]")
        conn = self.connect_db()
        
        update_fields = [field for field in self.DB_UPDATE_FIELDS if field in self.db_columns]
        
        select_parts = [f"r.{self.bracket_field('student_id')}", f"r.{self.bracket_field('exam_id')}"]
        fixed_select = ', '.join([f"r.{self.bracket_field(f)}" for f in update_fields])
        if fixed_select: select_parts.append(fixed_select)
        subject_select = ', '.join([f"r.{self.bracket_field(sub)}" for sub in potential_subjects])
        if subject_select: select_parts.append(subject_select)
        select_clause = ', '.join(select_parts)
        
        query = f"""
        SELECT {select_clause},
               i.{self.bracket_field('full_name')}, i.{self.bracket_field('sex')}, i.{self.bracket_field('comb_id')}
        FROM tbl_student_exam_results r
        INNER JOIN tbl_student_academic_info i ON r.{self.bracket_field('student_id')} = i.{self.bracket_field('student_id')}
        WHERE r.{self.bracket_field('exam_id')} = ?
        """
        self.df = pd.read_sql(query, conn, params=[self.exam_id])
        console.print(f" [green]- Fetched {len(self.df)} student records.[/green]")
        
        conn.close()
        
        if self.df.empty:
            raise ValueError("No records found for this exam_id.")
        
        # ORIGINAL DISPLAY KEPT
        self._display_sample_data(potential_subjects)
        # EXTRA 12-ROW PREVIEW
        self._preview("RAW MARKS – First 12 Students", ['full_name','sex','comb_id'] + potential_subjects[:9], highlight=potential_subjects[:9])
    
    def filter_valid_subjects(self, potential_subjects):
        """Filter subjects that have at least one valid mark."""
        console.print("\n[bold cyan]Stage 2.5: Filtering Subjects with At Least One Valid Mark (>=0)[/bold cyan]")
        self.valid_subjects = []
        
        for sub in potential_subjects:
            self.df[sub] = pd.to_numeric(self.df[sub], errors='coerce')
            if self.df[sub].ge(0).any():
                self.valid_subjects.append(sub)
            else:
                console.print(f" [yellow]- Dropping subject {sub}: no valid marks >=0 in any record.[/yellow]")
                if sub in self.df.columns:
                    self.df = self.df.drop(columns=[sub])
        
        console.print(f" [green]- Retained {len(self.valid_subjects)} subjects with valid marks: {self.valid_subjects}[/green]")
        self._preview("VALID SUBJECT MARKS – First 12 Students", ['full_name'] + self.valid_subjects[:10], highlight=self.valid_subjects[:10])
    
    def load_metadata(self):
        """Load combination and subject metadata."""
        console.print("\n[bold cyan]Stage 3: Fetching Metadata for Combinations and Subjects[/bold cyan]")
        conn = self.connect_db()
        
        comb_df = pd.read_sql("SELECT serial_id, comb_id, subject_id FROM tbl_student_comb_subjects", conn)
        sub_df = pd.read_sql("SELECT subject_serial, subject_short, subject_user_short, is_core, is_present,subject_name FROM tbl_student_subjects", conn)
        
        conn.close()
        console.print(f" [green]- Loaded {len(comb_df)} comb-subject links and {len(sub_df)} subjects.[/green]")
        
        comb_metadata_all = comb_df.merge(sub_df, left_on='subject_id', right_on='subject_serial')
        self.comb_metadata = comb_metadata_all[comb_metadata_all['is_present'] == True]
        self.subject_to_user = dict(zip(sub_df['subject_short'].str.lower(), sub_df['subject_user_short']))
        
        self.df['comb_id'] = self.df['comb_id'].astype(str)
    
    def reset_computed_fields(self):
        """Reset all computed fields in the database."""
        console.print("\n[bold cyan]Stage 4: Resetting Computed Fields in DB[/bold cyan]")
        
        grade_fields = [f"{sub}_grade" for sub in self.valid_subjects if f"{sub}_grade" in self.db_columns]
        pos_out_fields = [f"{sub}_pos" for sub in self.valid_subjects if f"{sub}_pos" in self.db_columns] + \
                         [f"{sub}_out_of" for sub in self.valid_subjects if f"{sub}_out_of" in self.db_columns]
        all_reset_fields = [f for f in self.DB_UPDATE_FIELDS if f in self.db_columns] + grade_fields + pos_out_fields
        
        conn = self.connect_db()
        cursor = conn.cursor()
        
        for field in all_reset_fields:
            bracketed = self.bracket_field(field)
            cursor.execute(f"UPDATE tbl_student_exam_results SET {bracketed} = NULL WHERE {self.bracket_field('exam_id')} = ?", (self.exam_id,))
        
        conn.commit()
        conn.close()
        console.print(f" [green]- Reset {len(all_reset_fields)} fields in DB for exam.[/green]")
    
    def compute_grades(self):
        """Compute letter grades from numeric marks."""
        console.print("\n[bold cyan]Stage 5: Computing Grades from Marks[/bold cyan]")
        
        for sub in self.valid_subjects:
            self.df[sub] = self.df[sub].where(self.df[sub].between(0, 100), np.nan)
            self.df[f'{sub}_grade'] = self.df[sub].apply(self.get_grade)
        
        self.df['total_marks'] = self.df[self.valid_subjects].sum(axis=1, skipna=True)
        console.print(" [green]- Grades computed successfully.[/green]")
        self._preview("LETTER GRADES – First 12 Students", ['full_name'] + [f"{s}_grade" for s in self.valid_subjects[:10]], highlight=[f"{s}_grade" for s in self.valid_subjects[:10]])
    
    def build_necta(self, row):
        """Build NECTA results strings for a student row."""
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
            grade = row.get(f'{short}_grade')
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
                grade = row.get(f'{sub}_grade')
                parts.append(f"{user}-{self.format_grade_for_necta(grade)}")
                marks_str = self.format_marks_for_necta(row[sub])
                parts_marks.append(f"{user}-{marks_str} {self.format_grade_for_necta(grade)}")
        
        avg_grade = row.get('avg_grade')
        avg_grade_fmt = self.format_grade_for_necta(avg_grade)
        parts.append(f"AVG-{avg_grade_fmt}")
        
        avg_marks_val = row.get('avg_marks')
        avg_marks_str = self.format_marks_for_necta(avg_marks_val) if pd.notna(avg_marks_val) else 'X'
        parts_marks.append(f"AVG-{avg_marks_str}{avg_grade_fmt}")
        
        row['necta_results'] = ', '.join(parts)
        row['necta_results_marks'] = ', '.join(parts_marks)
        return row
    
    def process_student_row(self, row):
        """Process a single student row to compute all metrics."""
        student_comb = self.comb_metadata[self.comb_metadata['comb_id'] == row['comb_id']]
        comb_shorts = set(student_comb['subject_short'].str.lower())
        core_shorts = set(student_comb[student_comb['is_core'] == True]['subject_short'].str.lower())
        
        effective_cores = core_shorts.union([sub for sub in self.valid_subjects if pd.notna(row.get(sub)) and sub not in comb_shorts])
        attempted_effective = [sub for sub in effective_cores if pd.notna(row.get(sub))]
        missing_count = len(effective_cores) - len(attempted_effective)
        
        attempted_all = [sub for sub in self.valid_subjects if pd.notna(row.get(sub))]
        row['subject_count_all'] = len(attempted_all) if attempted_all else 0
        row['subject_count'] = len(attempted_effective) if attempted_effective else 0
        
        if row['subject_count_all'] > 0:
            row['avg_marks'] = row['total_marks'] / row['subject_count_all']
            row['avg_grade'] = self.get_grade(row['avg_marks'])
        else:
            row['avg_marks'] = None
            row['avg_grade'] = None
        
        core_marks = sorted([row[sub] for sub in attempted_effective if pd.notna(row.get(sub))], reverse=True)
        row['first'] = core_marks[0] if len(core_marks) >= 1 else None
        row['second'] = core_marks[1] if len(core_marks) >= 2 else None
        row['third'] = core_marks[2] if len(core_marks) >= 3 else None
        
        grade_pts = [self.GRADE_POINTS.get(row.get(f'{sub}_grade')) for sub in attempted_effective]
        valid_pts = [p for p in grade_pts if p is not None]
        computed_points = sum(valid_pts) if valid_pts else None
        
        row['computed_points'] = computed_points
        
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
        
        row['division'] = save_div
        row['points'] = save_points
        
        div_val = self.DIVISION_VALUES.get(save_div)
        if div_val is not None and row['subject_count'] > 0:
            row['gpa'] = div_val / row['subject_count']
        else:
            row['gpa'] = None
        
        row = self.build_necta(row)
        return row
    
    def process_all_students(self):
        """Process all student rows."""
        console.print("\n[bold cyan]Stage 6: Processing Each Student Row...[/bold cyan]")
        # YOUR ORIGINAL LINE – ONLY .progress_apply() ADDED (tqdm.pandas() makes it work)
        self.df = self.df.progress_apply(self.process_student_row, axis=1)
        console.print(" [green]- All student rows processed.[/green]")
        self._preview("FINAL RESULTS – First 12 Students", ['full_name','division','points','gpa','first','second','third'], highlight=['division','points','gpa'])
    
    def compute_ranking(self):
        """Compute school-wide and combination-specific rankings."""
        exam_df = self.df[self.df['exam_id'] == self.exam_id].copy()
        if exam_df.empty:
            return
        
        # Valid students: NOT ABS + HAS AT LEAST ONE MARK
        valid_students = exam_df[
            (exam_df['division'] != 'ABS') &
            (exam_df[self.valid_subjects].notna().any(axis=1))
        ].copy()
        
        invalid_students = exam_df[
            (exam_df['division'] == 'ABS') |
            (~exam_df[self.valid_subjects].notna().any(axis=1))
        ].copy()
        
        if valid_students.empty:
            self.df.loc[exam_df.index, ['position_school', 'out_of_school', 'position_comb', 'out_of_comb']] = pd.NA
            return
        
        # Prepare points_for_rank to handle None/NaN in points column
        valid_students['points_for_rank'] = valid_students['points'].replace({None: np.inf, np.nan: np.inf})
        
        # Sort using parameterized columns and directions
        valid_students = valid_students.sort_values(self.sort_columns, ascending=self.sort_ascending)
        
        n_valid = len(valid_students)
        valid_students['position_school'] = np.arange(1, n_valid + 1)
        valid_students['out_of_school'] = n_valid
        
        def assign_comb_rank(group):
            group['points_for_rank'] = group['points'].replace({None: np.inf, np.nan: np.inf})
            group = group.sort_values(self.sort_columns, ascending=self.sort_ascending)
            group['position_comb'] = np.arange(1, len(group) + 1)
            group['out_of_comb'] = len(group)
            return group
        
        valid_students = valid_students.groupby('comb_id', group_keys=False).apply(assign_comb_rank)
        
        self.df.loc[valid_students.index, ['position_school', 'out_of_school', 'position_comb', 'out_of_comb']] = \
            valid_students[['position_school', 'out_of_school', 'position_comb', 'out_of_comb']]
        
        if not invalid_students.empty:
            self.df.loc[invalid_students.index, ['position_school', 'out_of_school', 'position_comb', 'out_of_comb']] = pd.NA
        
        if 'points_for_rank' in self.df.columns:
            self.df = self.df.drop(columns=['points_for_rank'], errors='ignore')
    
    def compute_subject_rankings(self):
        """Compute per-subject rankings."""
        for sub in tqdm(self.valid_subjects, desc="Ranking subjects"):
            pos_col = f"{sub}_pos"
            out_col = f"{sub}_out_of"
            
            sub_df = self.df[self.df[sub].notna()].copy()
            if sub_df.empty:
                continue
            
            sub_df = sub_df.sort_values(sub, ascending=False)
            sub_df['rank_temp'] = sub_df[sub].rank(method='min', ascending=False)
            sub_df[pos_col] = sub_df['rank_temp'].astype('Int64')
            sub_df[out_col] = len(sub_df)
            
            self.df.loc[sub_df.index, [pos_col, out_col]] = sub_df[[pos_col, out_col]]
    
    def finalize_data_types(self):
        """Finalize data types for all fields."""
        console.print("\n[bold cyan]Stage 9: Finalizing Data Types and Updating DB[/bold cyan]")
        
        int_fields = ['position_school', 'position_comb', 'out_of_school', 'out_of_comb',
                      'first', 'second', 'third', 'points', 'subject_count', 'subject_count_all'] + \
                     [f"{sub}_pos" for sub in self.valid_subjects] + [f"{sub}_out_of" for sub in self.valid_subjects]
        
        for field in int_fields:
            if field in self.df.columns:
                self.df[field] = pd.to_numeric(self.df[field], errors='coerce').astype('Int64')
        
        if 'gpa' in self.df.columns:
            self.df['gpa'] = pd.to_numeric(self.df['gpa'], errors='coerce').round(4)
        
        if 'computed_points' in self.df.columns:
            self.df = self.df.drop(columns=['computed_points'])
        
        console.print(" [green]- Data types finalized.[/green]")
    
    def update_database(self):
        """Update database with computed results."""
        grade_fields = [f"{sub}_grade" for sub in self.valid_subjects if f"{sub}_grade" in self.db_columns]
        pos_out_fields = [f"{sub}_pos" for sub in self.valid_subjects if f"{sub}_pos" in self.db_columns] + \
                         [f"{sub}_out_of" for sub in self.valid_subjects if f"{sub}_out_of" in self.db_columns]
        
        all_update_fields = [f for f in [
            'points', 'division', 'subject_count', 'total_marks', 'gpa',
            'subject_count_all', 'necta_results', 'necta_results_marks',
            'first', 'second', 'third', 'avg_grade',
            'position_comb', 'position_school', 'out_of_comb', 'out_of_school'
        ] if f in self.db_columns] + grade_fields + pos_out_fields
        
        console.print(f" [blue]- Updating {len(all_update_fields)} fields back to DB.[/blue]")
        
        conn = self.connect_db()
        cursor = conn.cursor()
        
        for idx, row in tqdm(self.df.iterrows(), total=len(self.df), desc="Updating records", colour="cyan"):
            set_clause = ', '.join([f"{self.bracket_field(field)} = ?" for field in all_update_fields])
            params = tuple(None if pd.isna(row.get(field)) else row.get(field) for field in all_update_fields) + (row['student_id'], self.exam_id)
            update_query = f"UPDATE tbl_student_exam_results SET {set_clause} WHERE {self.bracket_field('student_id')} = ? AND {self.bracket_field('exam_id')} = ?"
            cursor.execute(update_query, params)
        
        conn.commit()
        
        cursor.execute("UPDATE tbl_student_exam_results SET subject_count_all=NULL,subject_count=NULL WHERE division=? AND exam_id=?", ('ABS', self.exam_id))
        conn.commit()
        
        conn.close()
        console.print("\n[bold green]✓ All updates committed to DB. Process complete.[/bold green]")
        
        # ORIGINAL FINAL SUMMARY KEPT
        self._display_final_summary()
    
    def _display_final_summary(self):
        """Display final results summary."""
        console.print("\n[bold magenta]Final Results Summary (Top 5 Students):[/bold magenta]")
        
        # Select top 5 students and relevant columns
        summary_cols = ['student_id', 'full_name', 'division', 'points', 'avg_marks', 
                        'position_school', 'subject_count']
        
        # Get valid students only
        valid_df = self.df[self.df['division'] != 'ABS'].copy()
        if valid_df.empty:
            console.print("[yellow]No valid students to display.[/yellow]")
            return
        
        top_students = valid_df.nsmallest(5, 'position_school')[summary_cols]
        
        # Create rich table
        table = Table(show_header=True, header_style="bold green", box=box.DOUBLE)
        
        for col in summary_cols:
            table.add_column(col.upper().replace('_', ' '), style="cyan")
        
        for _, row in top_students.iterrows():
            table.add_row(*[
                str(row[col]) if pd.notna(row[col]) else "-" 
                for col in summary_cols
            ])
        
        console.print(table)
        
        # Division distribution
        console.print("\n[bold yellow]Division Distribution:[/bold yellow]")
        div_counts = self.df['division'].value_counts().sort_index()
        for div, count in div_counts.items():
            console.print(f"  {div}: {count} students")



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


    def update_and_display_competency(self):
        console.print("\n[bold cyan]Processing Subject Competency Report[/bold cyan]")

        conn = self.connect_db()
        cur = conn.cursor()
        cur.execute("DELETE FROM tbl_competency WHERE exam_id = ?", (self.exam_id,))
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

        # Initialize totals
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

            records.append((str(self.exam_id), serial, A,B,C,D,E,S,F, total, A+B+C+D+E, S+F, round(gpa,4), level))
            
            # Find background color for this row's competency level
            bg = None
            for (lo, hi), color in GPA_COLORS.items():
                if lo <= gpa <= hi:
                    bg = color
                    break
            
            rows.append({
                'name': name, 'A': A, 'B': B, 'C': C, 'D': D, 'E': E, 'S': S, 'F': F, 
                'total': total, 'gpa': gpa, 'level': level, 'bg_color': bg
            })

            # Accumulate totals
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
            cur.executemany("INSERT INTO tbl_competency (exam_id,subject_serial,A_s,B_s,C_s,D_s,E_s,S_s,F_s,total,pass,fail,gpa,competency_level) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", records)
            conn.commit()
        conn.close()

        if not rows:
            console.print("[red]No data[/red]")
            return

        rows.sort(key=lambda x: x['gpa'])

        table = Table(title=f"[bold white on #1e1b4b] SUBJECT COMPETENCY REPORT – {self.exam_id} [/]", box=box.DOUBLE_EDGE, expand=True)

        # Normal columns
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
            # Create styled text ONLY for competency level cell if bg_color exists
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

        # Add totals/average row
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
        """Execute the complete processing pipeline."""
        try:
            potential_subjects = self.detect_columns()
            self.fetch_data(potential_subjects)
            self.filter_valid_subjects(potential_subjects)
            self.load_metadata()
            self.reset_computed_fields()
            self.compute_grades()
            self.process_all_students()
            self.compute_ranking()
            self.compute_subject_rankings()
            self.finalize_data_types()
            self.update_database()
            self.update_and_display_competency()
            
        except Exception as e:
            print(f"Error during processing: {e}")
            raise


# Usage example:
if __name__ == "__main__":
    # Default sorting: points_for_rank, avg_marks, subject_count
    processor = AlevelProcessor(
        exam_id='ANN520250526',
        db_path=r"C:\Users\droge\OneDrive\Documents\Kiyabo App Backend v4.0.0.accdb",
        include_inc=True
    )
    processor.run()
    
    # Custom sorting example: prioritize avg_marks first
    # processor = AlevelProcessor(
    #     exam_id='ANN520250526',
    #     db_path=r"C:\Users\droge\OneDrive\Documents\Kiyabo App Backend v4.0.0.accdb",
    #     sort_columns=['avg_marks', 'points_for_rank', 'subject_count'],
    #     include_inc=True
    # )
    # processor.run()