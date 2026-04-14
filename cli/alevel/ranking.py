import pandas as pd
import pyodbc
import numpy as np
from datetime import datetime
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
        'points': True,   # Lower is better
        'avg_marks': False,         # Higher is better
        'subject_count': False,     # Higher is better
    }
    
    def __init__(self, exam_id: str, db_path: str, 
                 sort_columns: list = None,
                 include_inc: bool = True,
                 rank_method: str = 'min',
                 rank_incs: bool = False):
        """
        Initialize the A-Level processor.
        
        Args:
            exam_id: The exam identifier (e.g., 'ANN520250526')
            db_path: Path to the Access database file
            sort_columns: List of column names to sort by (default: ["points", "avg_marks", "subject_count"])
            include_inc: Whether to include INC status or convert to penalty
        """
        self.exam_id = exam_id
        self.db_path = db_path
        self.sort_columns = sort_columns or ["avg_marks","points", "subject_count"]
        self.include_inc = include_inc
        self.rank_method = rank_method
        self.rank_incs = rank_incs
        
        # Build ascending list based on fixed rules
        self.sort_ascending = [self.SORT_ASCENDING[col] for col in self.sort_columns]
        
        # Runtime attributes
        self.conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + db_path + ';'
        self.df = None
        self.valid_subjects = []
        self.comb_metadata = None
        self.subject_to_user = {}
        self.db_columns = []
        self.grade_boundaries = None
        self.class_id = ""  # NEW: Store class_id for curriculum handling
        
    @staticmethod
    def bracket_field(field: str) -> str:
        """Wrap field name in brackets for SQL queries."""
        return f"[{field}]"
    
    def get_class_from_exam_id(self):
        """Extract class ID from exam record in database."""
        try:
            conn = self.connect_db()
            cursor = conn.cursor()
            
            sql = "SELECT class_id FROM tbl_student_exams WHERE exam_id = ?"
            cursor.execute(sql, (self.exam_id,))
            result = cursor.fetchone()
            
            cursor.close()
            conn.close()
            
            if result:
                return result[0]
            return ""
            
        except Exception as e:
            # Try to extract class from exam_id format
            if self.exam_id and len(self.exam_id) >= 4:
                if 'VI' in self.exam_id.upper():
                    return 'VI'
                elif self.exam_id[4:4] == '6':  # Form VI pattern
                    return 'VI'
                elif self.exam_id[4:4] == '5':  # Form V pattern
                    return 'V'
            return ""
    
    def get_valid_subjects_for_curriculum(self):
        """Get list of valid subjects based on curriculum (Form VI old vs new)."""
        conn = self.connect_db()
        
        # Build SQL based on curriculum
        if self.class_id == "VI" and datetime.now() < datetime(2026, 7, 1):
            # Form VI old curriculum: include subject 31 (gs), exclude 30 & 34
            sql = """
                SELECT subject_short 
                FROM tbl_student_subjects 
                WHERE (is_present=True OR subject_serial=31) 
                AND subject_serial NOT IN (30,34)
            """
            console.print("[bold magenta]✓ Using OLD CURRICULUM (Form VI before July 2026)[/bold magenta]")
            console.print("[yellow]  - Including subject 31 (GS) regardless of is_present[/yellow]")
            console.print("[yellow]  - Excluding subjects 30 & 34[/yellow]")
        else:
            # New curriculum or other classes
            sql = "SELECT subject_short FROM tbl_student_subjects WHERE is_present=True"
            console.print("[bold cyan]✓ Using NEW CURRICULUM / Standard[/bold cyan]")
        
        df = pd.read_sql(sql, conn)
        conn.close()
        
        curriculum_subjects = [s.lower() for s in df['subject_short'].tolist()]
        console.print(f"[green]✓ Curriculum defines {len(curriculum_subjects)} subjects: {curriculum_subjects}[/green]")
        
        return curriculum_subjects
    
    def get_grade(self, marks) -> str:
        """Convert numeric marks to letter grade using database boundaries."""
        if pd.isna(marks) or not isinstance(marks, (int, float)) or marks < 0 or marks > 100:
            return None
        
        for _, row in self.grade_boundaries.iterrows():
            if marks >= row['lower']:
                return row['grade']
        
        return None
    
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
        
        display_cols = ['student_id', 'full_name', 'sex', 'comb_id']
        subject_cols = potential_subjects[:10]
        display_cols.extend(subject_cols)
        
        sample_df = self.df[display_cols].head(10)
        
        table = Table(show_header=True, header_style="bold cyan", box=box.ROUNDED)
        
        for col in display_cols:
            table.add_column(col.upper(), style="white")
        
        for _, row in sample_df.iterrows():
            table.add_row(*[str(row[col]) if pd.notna(row[col]) else "-" for col in display_cols])
        
        console.print(table)

    def _preview(self, title: str, cols: list, highlight=None, rows: int = 12):
        table = Table(title=f"[bold gold1 on #1e1b4b]{title}[/]", box=box.DOUBLE_EDGE, border_style="#7209b7", expand=True)
        for col in cols:
            style = "bold bright_yellow" if highlight and col in highlight else "white"
            table.add_column(col.upper().replace("_", " "), style=style, justify="center")
        for _, row in self.df[cols].head(rows).iterrows():
            values = [str(val) if pd.notna(val) else "—" for val in row]
            table.add_row(*values)
        console.print(table)
    
    def _display_combination_structure(self):
        """Display debug table showing what subjects belong to each combination."""
        console.print("\n[bold magenta]═══════════════════════════════════════════════════════════════════════[/bold magenta]")
        console.print("[bold magenta]              COMBINATION STRUCTURE DEBUG TABLE                        [/bold magenta]")
        console.print("[bold magenta]═══════════════════════════════════════════════════════════════════════[/bold magenta]")
        
        student_combs = self.df['comb_id'].unique()
        
        table = Table(
            box=box.DOUBLE_EDGE,
            border_style="cyan",
            expand=True,
            title="[bold white on #1e1b4b] COMBINATION SUBJECTS BREAKDOWN [/]"
        )
        
        table.add_column("COMB ID", style="bold cyan", justify="center")
        table.add_column("TOTAL\nSUBJECTS", style="bold yellow", justify="center")
        table.add_column("CORE\nSUBJECTS", style="bold green", justify="center")
        table.add_column("ELECTIVE\nSUBJECTS", style="bold magenta", justify="center")
        table.add_column("ALL SUBJECTS IN COMBINATION", style="white", justify="left")
        table.add_column("STUDENTS\nIN COMB", style="bold blue", justify="center")
        
        for comb_id in sorted(student_combs):
            comb_subjects = self.comb_metadata[self.comb_metadata['comb_id'] == str(comb_id)]
            
            if comb_subjects.empty:
                console.print(f" [red]⚠ Combination {comb_id} has NO subjects defined in metadata![/red]")
                continue
            
            core_subjects = comb_subjects[comb_subjects['is_core'] == True]
            elective_subjects = comb_subjects[comb_subjects['is_core'] == False]
            
            all_subs = []
            for _, row in comb_subjects.iterrows():
                user = row['subject_user_short']
                marker = " ★ " if row['is_core'] else " ○ "
                all_subs.append(f"{marker}{user}")
            
            student_count = len(self.df[self.df['comb_id'] == str(comb_id)])
            
            table.add_row(
                str(comb_id),
                str(len(comb_subjects)),
                str(len(core_subjects)),
                str(len(elective_subjects)),
                ", ".join(all_subs),
                str(student_count)
            )
        
        console.print(table)
        console.print("\n[cyan]Legend: ★ = Core Subject  |  ○ = Elective Subject[/cyan]")
        console.print("[yellow]Note: subject_count_all = ALL subjects in combination (with/without marks) + Extra subjects with marks[/yellow]")
        console.print("[yellow]      subject_count = ONLY subjects the student attempted (has marks for)[/yellow]\n")
    
    def _display_grade_boundaries(self):
        """Display the grade boundaries table loaded from database."""
        console.print("\n[bold magenta]═══════════════════════════════════════════════════════════════════════[/bold magenta]")
        console.print("[bold magenta]                    GRADE BOUNDARIES TABLE                             [/bold magenta]")
        console.print("[bold magenta]═══════════════════════════════════════════════════════════════════════[/bold magenta]")
        
        table = Table(
            box=box.DOUBLE_EDGE,
            border_style="green",
            expand=True,
            title="[bold white on #1e1b4b] GRADING SYSTEM CONFIGURATION [/]"
        )
        
        table.add_column("GRADE", style="bold yellow", justify="center", width=10)
        table.add_column("LOWER\nBOUND", style="bold cyan", justify="center", width=10)
        table.add_column("UPPER\nBOUND", style="bold cyan", justify="center", width=10)
        table.add_column("MARK RANGE", style="bold white", justify="center", width=15)
        table.add_column("DESCRIPTION", style="white", justify="left")
        
        for _, row in self.grade_boundaries.iterrows():
            lower = int(row['lower']) if pd.notna(row['lower']) else 0
            higher = int(row['higher']) if pd.notna(row['higher']) else 100
            grade = str(row['grade'])
            description = str(row['description']) if pd.notna(row['description']) else ""
            
            if grade == 'A':
                grade_style = "bold green"
            elif grade in ['B', 'C']:
                grade_style = "bold cyan"
            elif grade in ['D', 'E']:
                grade_style = "bold yellow"
            else:
                grade_style = "bold red"
            
            table.add_row(
                Text(grade, style=grade_style),
                str(lower),
                str(higher),
                f"{lower} - {higher}",
                description
            )
        
        console.print(table)
        console.print(f"\n[green]✓ Loaded {len(self.grade_boundaries)} grade boundaries from database[/green]")
        console.print("[cyan]Note: Grading uses LOWER bound only (marks >= lower)[/cyan]\n")
    
    def _display_subject_count_discrepancy(self):
        """Display students where subject_count_all differs from subject_count."""
        console.print("\n[bold yellow]Subject Count Discrepancy Analysis[/bold yellow]")
        
        discrepancy_df = self.df[
            (self.df['subject_count_all'] != self.df['subject_count']) & 
            (self.df['division'] != 'ABS')
        ].copy()
        
        if discrepancy_df.empty:
            console.print(" [green]✓ No discrepancies found - all students have matching counts[/green]")
            return
        
        console.print(f" [yellow]⚠ Found {len(discrepancy_df)} students with count discrepancies[/yellow]")
        
        table = Table(
            title="[bold white on red] SUBJECT COUNT DISCREPANCIES [/]",
            box=box.DOUBLE_EDGE,
            border_style="red",
            expand=True
        )
        
        table.add_column("STUDENT ID", style="cyan", justify="left")
        table.add_column("NAME", style="white", justify="left")
        table.add_column("COMB", style="yellow", justify="center")
        table.add_column("COUNT", style="bold green", justify="center")
        table.add_column("COUNT ALL", style="bold magenta", justify="center")
        table.add_column("DIFF", style="bold red", justify="center")
        table.add_column("DIVISION", style="white", justify="center")
        
        for _, row in discrepancy_df.head(20).iterrows():
            count = int(row['subject_count']) if pd.notna(row['subject_count']) else 0
            count_all = int(row['subject_count_all']) if pd.notna(row['subject_count_all']) else 0
            diff = count_all - count
            
            table.add_row(
                str(row['student_id']),
                str(row['full_name'])[:30],
                str(row['comb_id']),
                str(count),
                str(count_all),
                f"+{diff}" if diff > 0 else str(diff),
                str(row['division'])
            )
        
        console.print(table)
        
        console.print(f"\n[bold cyan]Summary:[/bold cyan]")
        console.print(f" - Total students with discrepancy: {len(discrepancy_df)}")
        console.print(f" - Average count: {discrepancy_df['subject_count'].mean():.2f}")
        console.print(f" - Average count_all: {discrepancy_df['subject_count_all'].mean():.2f}")
        console.print(f" - Max difference: {(discrepancy_df['subject_count_all'] - discrepancy_df['subject_count']).max()}")
    
    def detect_columns(self):
        """Detect available columns in the database."""
        console.print("\n[bold cyan]Stage 1: Connecting to Database and Detecting Columns[/bold cyan]")
        
        # Get class ID first
        self.class_id = self.get_class_from_exam_id()
        if self.class_id:
            console.print(f"[green]✓ Detected class: {self.class_id}[/green]")
        else:
            console.print("[yellow]⚠ Class not detected - using standard curriculum[/yellow]")
        
        conn = self.connect_db()
        
        dummy_query = "SELECT TOP 1 * FROM tbl_student_exam_results WHERE exam_id = ?"
        df_dummy = pd.read_sql(dummy_query, conn, params=[self.exam_id])
        if df_dummy.empty:
            console.print(" [yellow]- No records for exam_id, falling back to general structure.[/yellow]")
            df_dummy = pd.read_sql("SELECT TOP 1 * FROM tbl_student_exam_results", conn)
        
        self.db_columns = [col.lower() for col in df_dummy.columns]
        
        # Get curriculum-aware subjects
        curriculum_subjects = self.get_valid_subjects_for_curriculum()
        
        # Only keep subjects that exist in BOTH curriculum AND database columns
        potential_subjects = [sub for sub in curriculum_subjects if sub in self.db_columns]
        console.print(f" [green]✓ Found {len(potential_subjects)} subjects in database: {potential_subjects}[/green]")
        
        if 'student_id' not in self.db_columns or 'exam_id' not in self.db_columns:
            conn.close()
            raise ValueError("Missing essential fields.")
        
        # Load grade boundaries
        console.print("\n[bold cyan]- Loading grade boundaries from tbl_student_grades...[/bold cyan]")
        grade_query = "SELECT grade, lower, higher, description FROM tbl_student_grades ORDER BY lower DESC"
        self.grade_boundaries = pd.read_sql(grade_query, conn)
        
        self._display_grade_boundaries()
        
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
        
        self._display_sample_data(potential_subjects)
        self._preview("RAW MARKS – First 12 Students", ['full_name','sex','comb_id'] + potential_subjects[:9], highlight=potential_subjects[:9])
    
    def filter_valid_subjects(self, potential_subjects):
        """Filter subjects that have at least one valid mark and REMOVE those with no marks."""
        console.print("\n[bold cyan]Stage 2.5: Filtering Subjects with At Least One Valid Mark (>=0)[/bold cyan]")
        self.valid_subjects = []
        removed_subjects = []
        
        for sub in potential_subjects:
            self.df[sub] = pd.to_numeric(self.df[sub], errors='coerce')
            if self.df[sub].ge(0).any():
                self.valid_subjects.append(sub)
            else:
                removed_subjects.append(sub)
                self.df = self.df.drop(columns=[sub])
                console.print(f" [red]✗ Subject {sub} has NO valid marks - REMOVED from dataset[/red]")
        
        if removed_subjects:
            console.print(f" [yellow]⚠ Removed {len(removed_subjects)} subjects with no marks: {removed_subjects}[/yellow]")
        
        console.print(f" [green]✓ Retained {len(self.valid_subjects)} subjects with valid marks: {self.valid_subjects}[/green]")
        self._preview("VALID SUBJECT MARKS – First 12 Students", ['full_name'] + self.valid_subjects[:10], highlight=self.valid_subjects[:10])
    

    def load_metadata(self):
        """Load combination and subject metadata with curriculum handling."""
        console.print("\n[bold cyan]Stage 3: Fetching Metadata for Combinations and Subjects[/bold cyan]")
        conn = self.connect_db()
        
        # Build query based on curriculum
        if self.class_id == "VI" and datetime.now() < datetime(2026, 7, 1):
            comb_sql = """
                SELECT serial_id, comb_id, subject_id 
                FROM tbl_student_comb_subjects 
                WHERE subject_id NOT IN (30, 34)
            """
            sub_sql = """
                SELECT subject_serial, subject_short, subject_user_short, is_core, is_present, subject_name 
                FROM tbl_student_subjects 
                WHERE (is_present=True OR subject_serial=31) 
                AND subject_serial NOT IN (30, 34)
            """
            console.print("[magenta]Using OLD CURRICULUM metadata (excluding 30,34; including 31)[/magenta]")
        else:
            comb_sql = "SELECT serial_id, comb_id, subject_id FROM tbl_student_comb_subjects"
            sub_sql = "SELECT subject_serial, subject_short, subject_user_short, is_core, is_present, subject_name FROM tbl_student_subjects"
        
        comb_df = pd.read_sql(comb_sql, conn)
        sub_df = pd.read_sql(sub_sql, conn)
        conn.close()
        
        console.print(f" [green]- Loaded {len(comb_df)} comb-subject links and {len(sub_df)} subjects.[/green]")
        
        comb_metadata_all = comb_df.merge(sub_df, left_on='subject_id', right_on='subject_serial')
        self.comb_metadata = comb_metadata_all[comb_metadata_all['is_present'] == True]
        
        # For Form VI old curriculum: Add gs(31) to EVERY combination
        if self.class_id == "VI" and datetime.now() < datetime(2026, 7, 1):
            sub31_rows = comb_metadata_all[comb_metadata_all['subject_serial'] == 31]
            if not sub31_rows.empty:
                self.comb_metadata = pd.concat([self.comb_metadata, sub31_rows]).drop_duplicates()
            
            # Get gs(31) info from sub_df
            gs_info = sub_df[sub_df['subject_serial'] == 31]
            if not gs_info.empty and 'gs' in self.valid_subjects:
                unique_combs = comb_df['comb_id'].unique()
                gs_rows_to_add = []
                
                for comb_id in unique_combs:
                    existing = self.comb_metadata[
                        (self.comb_metadata['comb_id'] == comb_id) & 
                        (self.comb_metadata['subject_serial'] == 31)
                    ]
                    
                    if existing.empty:
                        new_row = {
                            'serial_id': None,
                            'comb_id': comb_id,
                            'subject_id': 31,
                            'subject_serial': 31,
                            'subject_short': gs_info.iloc[0]['subject_short'],
                            'subject_user_short': gs_info.iloc[0]['subject_user_short'],
                            'is_core': False,
                            'is_present': gs_info.iloc[0]['is_present'],
                            'subject_name': gs_info.iloc[0]['subject_name']
                        }
                        gs_rows_to_add.append(new_row)
                
                if gs_rows_to_add:
                    gs_df = pd.DataFrame(gs_rows_to_add)
                    self.comb_metadata = pd.concat([self.comb_metadata, gs_df], ignore_index=True)
                    console.print(f"[green]✓ Added GS(31) to {len(gs_rows_to_add)} combinations for Form VI old curriculum[/green]")
        
        # Store ALL subjects metadata for competency report
        self.all_subjects_metadata = sub_df
        self.subject_to_user = dict(zip(sub_df['subject_short'].str.lower(), sub_df['subject_user_short']))
        self.df['comb_id'] = self.df['comb_id'].astype(str)
        
        self._display_combination_structure()



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
        """
        Process a single student row to compute all metrics including division, points, and rankings.
        
        [... keep all the existing docstring ...]
        """
        
        student_comb = self.comb_metadata[self.comb_metadata['comb_id'] == row['comb_id']]
        comb_shorts = set(student_comb['subject_short'].str.lower())
        core_shorts = set(student_comb[student_comb['is_core'] == True]['subject_short'].str.lower())
        
        # ═══════════════════════════════════════════════════════════════
        # CRITICAL: Effective cores = ONLY declared core subjects in combination
        # Non-core subjects (GS, BAM) are NEVER used for division/points
        # ═══════════════════════════════════════════════════════════════
        effective_cores = core_shorts  # CBG: [bio, che, geo] only - NO GS, NO BAM
        
        attempted_effective = [sub for sub in effective_cores if pd.notna(row.get(sub))]
        missing_count = len(effective_cores) - len(attempted_effective)
        
        # ═══════════════════════════════════════════════════════════════
        # Subject counts: ALL subjects in combination (core + non-core)
        # ═══════════════════════════════════════════════════════════════
        all_in_comb = len(comb_shorts)  # CBG: 5 (bio, che, geo, bam, gs)
        
        # Extra subjects NOT in combination but student has marks for
        attempted_extra = [sub for sub in self.valid_subjects 
                        if pd.notna(row.get(sub)) and row.get(sub) >= 0 and sub not in comb_shorts]
        
        row['subject_count_all'] = all_in_comb + len(attempted_extra)
        
        # subject_count = ONLY subjects student actually attempted
        attempted_all = [sub for sub in self.valid_subjects if pd.notna(row.get(sub)) and row.get(sub) >= 0]
        row['subject_count'] = len(attempted_all)
        
        # ═══════════════════════════════════════════════════════════════
        # Average marks: uses ALL subjects (core + non-core + extras)
        # ═══════════════════════════════════════════════════════════════
        if row['subject_count_all'] > 0:
            row['avg_marks'] = row['total_marks'] / row['subject_count_all']
            row['avg_grade'] = self.get_grade(row['avg_marks'])
        else:
            row['avg_marks'] = None
            row['avg_grade'] = None
        
        # ═══════════════════════════════════════════════════════════════
        # First/Second/Third: ONLY from core subjects
        # ═══════════════════════════════════════════════════════════════
        core_marks = sorted([row[sub] for sub in attempted_effective if pd.notna(row.get(sub))], reverse=True)
        row['first'] = core_marks[0] if len(core_marks) >= 1 else None
        row['second'] = core_marks[1] if len(core_marks) >= 2 else None
        row['third'] = core_marks[2] if len(core_marks) >= 3 else None
        
        # ═══════════════════════════════════════════════════════════════
        # Points: ONLY from core subjects
        # ═══════════════════════════════════════════════════════════════
        grade_pts = [self.GRADE_POINTS.get(row.get(f'{sub}_grade')) for sub in attempted_effective]
        valid_pts = [p for p in grade_pts if p is not None]
        computed_points = sum(valid_pts) if valid_pts else None
        
        # Store raw computed points for potential future use (debugging, reports, etc.)
        row['computed_points_raw'] = computed_points
        
        # ═══════════════════════════════════════════════════════════════
        # Completeness: Based on CORE subjects only
        # ═══════════════════════════════════════════════════════════════
        has_invalid = len(grade_pts) != len(valid_pts)
        is_complete = (len(effective_cores) >= 3 and missing_count == 0 and not has_invalid)
        is_abs = len(attempted_effective) == 0
        is_inc = not is_complete and not is_abs
        
        # ═══════════════════════════════════════════════════════════════
        # Division and Points Assignment
        # ═══════════════════════════════════════════════════════════════
        if is_abs:
            save_div = 'ABS' if self.include_inc else ('0' if self.get_div_from_points(7 * len(effective_cores)) == '0' else 'IV')
            save_points = None
            
        elif is_inc:
            # Calculate penalty points for determining converted division
            penalty_points = sum(valid_pts) + 7 * missing_count if valid_pts else 7 * missing_count
            
            # Determine division based on include_inc setting
            if self.include_inc:
                save_div = 'INC'
            else:
                # Convert INC to '0' or 'IV' based on penalty
                save_div = '0' if self.get_div_from_points(penalty_points) == '0' else 'IV'
            
            # CRITICAL: ALWAYS save points=None for INC students
            # Even when converted to '0' or 'IV', they remain INC students with no points
            save_points = None
                
        else:
            # Complete student - normal processing
            save_div = self.get_div_from_points(computed_points)
            save_points = computed_points if save_div is not None else None
        
        row['division'] = save_div
        row['points'] = save_points
        
        # ═══════════════════════════════════════════════════════════════
        # GPA: uses division value / subject_count
        # ═══════════════════════════════════════════════════════════════
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
        self.df = self.df.progress_apply(self.process_student_row, axis=1)
        console.print(" [green]- All student rows processed.[/green]")
        self._preview("FINAL RESULTS – First 12 Students", ['full_name','division','points','gpa','first','second','third'], highlight=['division','points','gpa'])
        
        self._display_subject_count_discrepancy()


    def compute_ranking(self):
        """Compute school-wide and combination-specific rankings with PROPER TIE HANDLING."""
        console.print("\n[bold cyan]Stage 7: Computing Rankings[/bold cyan]")
        
        exam_df = self.df[self.df['exam_id'] == self.exam_id].copy()
        if exam_df.empty:
            console.print(" [yellow]- No exam data found[/yellow]")
            return

        print(f" [cyan]🔍 DEBUG: Total students before filtering: {len(exam_df):,}[/cyan]")
        
        # Build invalid conditions
        invalid_conditions = [(exam_df['division'] == 'ABS'), (exam_df['avg_marks'].isna())]
        
        # Only exclude students with NaN points if NOT ranking INCs
        if not self.rank_incs:
            invalid_conditions.append(exam_df['points'].isna())
        
        invalid_mask = invalid_conditions[0]
        for condition in invalid_conditions[1:]:
            invalid_mask = invalid_mask | condition
        
        valid_students = exam_df[~invalid_mask].copy()
        invalid_students = exam_df[invalid_mask]

        print(f" [cyan]🔍 DEBUG: Students with division='ABS': {(exam_df['division'] == 'ABS').sum()}[/cyan]")
        print(f" [cyan]🔍 DEBUG: Students with NaN points: {exam_df['points'].isna().sum()}[/cyan]")
        print(f" [cyan]🔍 DEBUG: Students with NaN avg_marks: {exam_df['avg_marks'].isna().sum()}[/cyan]")
        
        # Show INC student stats
        inc_mask = exam_df['points'].isna() & (exam_df['division'] != 'ABS')
        inc_count = inc_mask.sum()
        if inc_count > 0:
            inc_divisions = exam_df[inc_mask]['division'].value_counts()
            print(f" [cyan]🔍 DEBUG: INC students (points=None): {inc_count}[/cyan]")
            print(f" [cyan]   Divisions: {dict(inc_divisions)}[/cyan]")
        
        print(f" [cyan]🔍 DEBUG: Total invalid students: {len(invalid_students):,}[/cyan]")
        print(f" [cyan]🔍 DEBUG: Total valid students: {len(valid_students):,}[/cyan]")

        if valid_students.empty:
            self.df.loc[exam_df.index, ['position_school', 'out_of_school', 'position_comb', 'out_of_comb']] = pd.NA
            console.print(" [yellow]- No valid students to rank[/yellow]")
            return

        console.print(f" [green]✓ Found {len(valid_students):,} valid students to rank[/green]")
        console.print(f" [yellow]✓ Excluded {len(invalid_students):,} invalid students[/yellow]")

        console.print(f"\n[bold yellow]🔧 Preparing sort columns: {self.sort_columns}[/bold yellow]")
        
        for col in self.sort_columns:
            before_fill = valid_students[col].isna().sum()
            
            if col == 'points':
                # Fill NaN points with 999999 (worst possible) - this handles INC students when rank_incs=True
                fill_value = 999999
                valid_students[col] = valid_students[col].fillna(fill_value)
                if before_fill > 0:
                    console.print(f" [yellow]- {col}: filled {before_fill} NaN values (INC students) with {fill_value} → they rank last[/yellow]")
            elif col == 'avg_marks':
                valid_students[col] = valid_students[col].fillna(-1)
            elif col in ['subject_count', 'subject_count_all']:
                valid_students[col] = valid_students[col].fillna(0)
            
            after_fill = valid_students[col].isna().sum()
            console.print(f" [cyan]- {col}: {before_fill} NaN → {after_fill} remaining[/cyan]")
        
        for col in ['points', 'avg_marks']:
            if col in valid_students.columns:
                valid_students[col] = valid_students[col].round(4)
                console.print(f" [cyan]- {col}: rounded to 4 decimal places[/cyan]")

        console.print(f"\n[bold yellow]🏆 School-Wide Ranking[/bold yellow]")
        console.print(f" [cyan]- Sort criteria: {list(zip(self.sort_columns, ['↑ ASC' if asc else '↓ DESC' for asc in self.sort_ascending]))}[/cyan]")
        
        valid_students_sorted = valid_students.sort_values(
            by=self.sort_columns,
            ascending=self.sort_ascending
        ).copy()
        
        valid_students_sorted['_sort_tuple'] = valid_students_sorted[self.sort_columns].apply(
            lambda row: tuple(row), axis=1
        )
        
        unique_tuples = valid_students_sorted['_sort_tuple'].unique()
        console.print(f" [cyan]🔍 DEBUG: Found {len(unique_tuples):,} unique performance levels[/cyan]")
        
        rank_map = {tuple_val: idx + 1 for idx, tuple_val in enumerate(unique_tuples)}
        
        valid_students_sorted['position_school'] = valid_students_sorted['_sort_tuple'].map(rank_map)
        valid_students_sorted['out_of_school'] = len(valid_students_sorted)
        
        max_rank = valid_students_sorted['position_school'].max()
        num_unique = len(unique_tuples)
        console.print(f" [green]✓ Ranked {len(valid_students_sorted):,} students[/green]")
        console.print(f" [green]✓ Unique performance levels: {num_unique:,}[/green]")
        console.print(f" [green]✓ Highest rank assigned: {max_rank}[/green]")
        
        rank_counts = valid_students_sorted['position_school'].value_counts()
        ties = rank_counts[rank_counts > 1]
        if len(ties) > 0:
            total_tied = ties.sum()
            console.print(f" [yellow]⚠ Ties detected: {len(ties)} rank positions have ties ({total_tied} students total)[/yellow]")
            for rank, count in ties.head(3).items():
                console.print(f"   • Rank {rank}: {count} students tied")
        else:
            console.print(f" [green]✓ No ties detected (all students have unique ranks)[/green]")
        
        debug_df = valid_students_sorted.head(20)[['student_id', 'full_name', 'points', 'avg_marks', 
                                                    'subject_count_all', '_sort_tuple', 'position_school']]
        console.print(f"\n[bold magenta]🔍 DEBUG: First 20 ranked students (for verification)[/bold magenta]")
        for idx, row in debug_df.iterrows():
            console.print(f" {row['position_school']:3d}. {row['student_id']:15s} | "
                        f"pts:{row['points']:6.2f} avg:{row['avg_marks']:6.2f} cnt:{row['subject_count_all']:2.0f}")
        
        console.print(f"\n[bold yellow]📚 Combination-Specific Ranking[/bold yellow]")
        
        def rank_within_combination(group):
            """Rank students within combination - IDENTICAL to exporter logic"""
            group['_comb_sort_tuple'] = group[self.sort_columns].apply(
                lambda row: tuple(row), axis=1
            )
            unique_tuples = group['_comb_sort_tuple'].unique()
            rank_map = {tuple_val: idx + 1 for idx, tuple_val in enumerate(unique_tuples)}
            group['position_comb'] = group['_comb_sort_tuple'].map(rank_map)
            group['out_of_comb'] = len(group)
            return group.drop(columns=['_comb_sort_tuple'])
        
        ranked_df = valid_students_sorted.groupby('comb_id', group_keys=False).apply(
            rank_within_combination
        )
        
        ranked_df = ranked_df.drop(columns=['_sort_tuple'], errors='ignore')
        
        num_combs = ranked_df['comb_id'].nunique()
        console.print(f" [green]✓ Ranked students within {num_combs} combinations[/green]")
        
        rank_cols = ['position_school', 'out_of_school', 'position_comb', 'out_of_comb']
        self.df.loc[ranked_df.index, rank_cols] = ranked_df[rank_cols]
        
        if not invalid_students.empty:
            self.df.loc[invalid_students.index, rank_cols] = pd.NA
        
        console.print(f"\n[bold magenta]🔍 Validation Checks[/bold magenta]")
        
        if ranked_df['position_school'].max() > len(ranked_df):
            console.print(" [red]✗ ERROR: position_school exceeds total students![/red]")
        else:
            console.print(" [green]✓ School ranks within valid range[/green]")
        
        first_student = ranked_df.iloc[0]
        if first_student['position_school'] != 1:
            console.print(f" [red]✗ ERROR: First student has rank {first_student['position_school']} (should be 1)[/red]")
        else:
            console.print(" [green]✓ First student has rank 1[/green]")
        
        all_ranks = sorted(ranked_df['position_school'].unique())
        expected_max = len(all_ranks)
        actual_max = all_ranks[-1]
        console.print(f" [cyan]- Rank sequence: 1 to {actual_max} ({expected_max} unique values)[/cyan]")
        
        preview_cols = ['full_name', 'comb_id', 'division', 'points', 'avg_marks', 'subject_count_all',
                        'position_school', 'out_of_school', 'position_comb', 'out_of_comb']
        
        top_students = ranked_df.nsmallest(12, 'position_school')
        console.print(f"\n[bold gold1]🏆 TOP 12 STUDENTS[/bold gold1]")
        temp_df = self.df.copy()
        self.df = self.df.loc[top_students.index]
        self._preview("Top Performers", preview_cols, 
                    highlight=['position_school', 'position_comb', 'division', 'points'],
                    rows=len(top_students))
        self.df = temp_df
        
        if len(ranked_df) > 12:
            bottom_students = ranked_df.nlargest(12, 'position_school')
            console.print(f"\n[bold red]⚠️ BOTTOM 12 STUDENTS[/bold red]")
            self.df = self.df.loc[bottom_students.index]
            self._preview("Bottom Performers", preview_cols,
                        highlight=['position_school', 'position_comb', 'division', 'points'],
                        rows=len(bottom_students))
            self.df = temp_df
        
        console.print("\n[bold green]✅ Ranking complete with synchronized logic![/bold green]")

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
            sub_df[pos_col] = sub_df['rank_temp'].round().astype('Int64')
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
                numeric_series = pd.to_numeric(self.df[field], errors='coerce')
                self.df[field] = numeric_series.astype('float64').round().astype('Int64')
        
        if 'gpa' in self.df.columns:
            self.df['gpa'] = pd.to_numeric(self.df['gpa'], errors='coerce').round(4)
        
        if 'avg_marks' in self.df.columns:
            self.df['avg_marks'] = pd.to_numeric(self.df['avg_marks'], errors='coerce').round(4)
        
        # Keep computed_points_raw for reference but don't save to database
        if 'computed_points_raw' in self.df.columns:
            self.df['computed_points_raw'] = pd.to_numeric(self.df['computed_points_raw'], errors='coerce').round().astype('Int64')
        
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
        
        self._display_final_summary()
    
    def _display_final_summary(self):
        """Display final results summary."""
        console.print("\n[bold magenta]Final Results Summary (Top 5 Students):[/bold magenta]")
        
        summary_cols = ['student_id', 'full_name', 'division', 'points', 'avg_marks', 
                        'position_school', 'subject_count']
        
        valid_df = self.df[self.df['division'] != 'ABS'].copy()
        if valid_df.empty:
            console.print("[yellow]No valid students to display.[/yellow]")
            return
        
        top_students = valid_df.nsmallest(5, 'position_school')[summary_cols]
        
        table = Table(show_header=True, header_style="bold green", box=box.DOUBLE)
        
        for col in summary_cols:
            table.add_column(col.upper().replace('_', ' '), style="cyan")
        
        for _, row in top_students.iterrows():
            table.add_row(*[
                str(row[col]) if pd.notna(row[col]) else "-" 
                for col in summary_cols
            ])
        
        console.print(table)
        
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

        total_A = total_B = total_C = total_D = total_E = total_S = total_F = 0
        total_students = 0
        total_gpa = 0

        for sub in self.valid_subjects:
            if f"{sub}_grade" not in self.df.columns: continue
            sat = self.df[self.df[sub].notna()]
            if sat.empty: continue

            c = sat[f"{sub}_grade"].value_counts()
            A,B,C,D,E,S,F = [int(c.get(g,0)) for g in "ABCDESF"]
            total = len(sat)
            gpa = (A*1+B*2+C*3+D*4+E*5+S*6+F*7) / total
            level = self.get_competency_level(gpa)

            # Use all_subjects_metadata instead of comb_metadata
            meta = self.all_subjects_metadata[self.all_subjects_metadata['subject_short'].str.lower() == sub]
            if meta.empty: 
                console.print(f"[red]⚠ Skipping {sub} - not found in metadata[/red]")
                continue
            
            name = meta.iloc[0]['subject_name'].upper()
            serial = int(meta.iloc[0]['subject_serial'])

            records.append((str(self.exam_id), serial, A,B,C,D,E,S,F, total, A+B+C+D+E, S+F, round(gpa,4), level))
            
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
            cur.executemany("INSERT INTO tbl_competency (exam_id,subject_serial,A_s,B_s,C_s,D_s,E_s,S_s,F_s,total,pass,fail,gpa,competency_level) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", records)
            conn.commit()
        conn.close()

        if not rows:
            console.print("[red]No data[/red]")
            return

        rows.sort(key=lambda x: x['gpa'])

        table = Table(title=f"[bold white on #1e1b4b] SUBJECT COMPETENCY REPORT – {self.exam_id} [/]", box=box.DOUBLE_EDGE, expand=True)

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
    processor = AlevelProcessor(
        exam_id='ANN520250526',
        db_path=r"C:\Users\droge\OneDrive\Documents\Kiyabo App Backend v4.0.0.accdb",
        include_inc=True,
    )
    processor.run()