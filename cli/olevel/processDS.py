# ========================================
# OLEVEL PROCESSOR — FLEXIBLE RANKING SYSTEM
# ========================================
import pyodbc
import pandas as pd
import numpy as np
from tqdm import tqdm
import warnings
from colorama import init, Fore, Back, Style
from tabulate import tabulate
import datetime

# Initialize colorama for colored output
init(autoreset=True)

warnings.filterwarnings("ignore")


class OlevelProcessor:
    """
    O-Level Results Processing Engine with Flexible Ranking System
    Processes student exam results from Access DB with configurable ranking.
    """

    # ANSI Color Codes
    GREEN = Fore.GREEN + Style.BRIGHT
    CYAN = Fore.CYAN + Style.BRIGHT
    YELLOW = Fore.YELLOW + Style.BRIGHT
    MAGENTA = Fore.MAGENTA + Style.BRIGHT
    RED = Fore.RED + Style.BRIGHT
    BLUE = Fore.BLUE + Style.BRIGHT
    WHITE = Fore.WHITE + Style.BRIGHT
    RESET = Style.RESET_ALL

    def __init__(
        self,
        exam_id: str,
        db_path: str,
        base_subjects: int = 7,
        flat_rate: bool = True,
        include_inc: bool = True,
        update_competency: bool = True,
        sort_columns: list = None,
        ranking_method: str = "min"
    ):
        """
        Initialize O-Level Processor with flexible ranking configuration.
        
        Args:
            exam_id: Exam identifier
            db_path: Path to Access database
            base_subjects: Number of subjects for base calculation (default: 7)
            flat_rate: Use flat rate calculation (default: True)
            include_inc: Include incomplete students (default: True)
            update_competency: Update competency analysis (default: True)
            sort_columns: List of columns to sort by (default: ['ranking_points', 'avg_marks', 'subject_count_real'])
            ranking_method: Ranking method - 'min', 'max', 'average', 'dense', 'first' (default: 'min')
        """
        self.EXAM_ID = exam_id
        self.DB_PATH = db_path
        self.BASE_SUBJECTS = base_subjects
        self.FLAT_RATE = flat_rate
        self.INCLUDE_INC = include_inc
        self.UPDATE_COMPETENCY = update_competency
        self.RANKING_METHOD = ranking_method.lower()

        # Validate ranking method
        valid_methods = ['min', 'max', 'average', 'dense', 'first']
        if self.RANKING_METHOD not in valid_methods:
            raise ValueError(f"Invalid ranking_method '{ranking_method}'. Must be one of: {valid_methods}")

        # Default sort columns if none provided
        if sort_columns is None:
            sort_columns = ['avg_marks', 'ranking_points', 'subject_count_real']
        self.sort_columns = sort_columns

        # Build sort directions dynamically based on column semantics
        self.ascending = self._build_sort_directions(sort_columns)

        # Database connections and dataframes
        self.conn = None
        self.cursor = None
        self.df = None
        self.students_df = None
        self.exam_df = None
        self.subjects_df = None
        self.class_id = None
        self.valid_subject_cols = []
        self.subject_column_map = {}
        self.mark_columns = ['civ','his','geo','kis','eng','phy','che','bio','mat','edk','ics'] + \
                           [f'sub{i}' for i in range(12,21)]

    def _build_sort_directions(self, columns: list) -> list:
        """
        Build ascending/descending list based on column semantics.
        
        Rules:
        - ranking_points, points → ASC (lower is better)
        - avg_marks, total_marks, subject_count_real, subject_count → DESC (higher is better)
        
        Args:
            columns: List of column names to sort by
            
        Returns:
            list: Boolean list for pandas sort (True = ascending, False = descending)
        """
        ascending = []
        for col in columns:
            if col in ['ranking_points', 'points']:
                # Lower points = better ranking
                ascending.append(True)
            elif col in ['avg_marks', 'total_marks', 'subject_count_real', 'subject_count']:
                # Higher values = better ranking
                ascending.append(False)
            else:
                # Default to descending for unknown columns
                print(f"{self.YELLOW}⚠️  Warning: Unknown column '{col}' - defaulting to DESC sort")
                ascending.append(False)
        
        return ascending

    # -------------------------------
    # CONNECT
    # -------------------------------
    def _connect_to_database(self):
        """Establish connection to Access database."""
        print(f"{self.CYAN}CONNECTING TO ACCESS DATABASE...")
        self.conn = pyodbc.connect(
            f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};"
            f"DBQ={self.DB_PATH};"
        )
        self.cursor = self.conn.cursor()

    # -------------------------------
    # 1. LOAD DATA
    # -------------------------------
    def load_data(self):
        """Load exam results, student info, and exam configuration from database."""
        print(f"\n{self.GREEN}1. LOADING EXAM RESULTS DATA")
        print("=" * 60)

        results_sql = """
            SELECT result_id, student_id, 
                   civ, his, geo, kis, eng, phy, che, bio, mat, edk, ics,
                   sub12, sub13, sub14, sub15, sub16, sub17, sub18, sub19, sub20
            FROM tbl_student_exam_results 
            WHERE exam_id = ?
        """
        self.df = pd.read_sql(results_sql, self.conn, params=[self.EXAM_ID])

        students_sql = "SELECT student_id, full_name, sex FROM tbl_student_academic_info"
        self.students_df = pd.read_sql(students_sql, self.conn)

        exam_sql = "SELECT exam_id, class_id FROM tbl_student_exams WHERE exam_id = ?"
        self.exam_df = pd.read_sql(exam_sql, self.conn, params=[self.EXAM_ID])

        self.df = self.df.merge(self.students_df, on='student_id', how='left')
        self.class_id = self.exam_df['class_id'].iloc[0] if not self.exam_df.empty else None

        if self.class_id is None:
            raise ValueError(f"{self.RED}Class ID could not be determined from tbl_student_exams for exam_id: {self.EXAM_ID}")

        print(f"{self.YELLOW}Class ID: {self.WHITE}{self.class_id}")
        print(f"{self.YELLOW}Total Students: {self.WHITE}{len(self.df):,}")

    # -------------------------------
    # 2. SUBJECT MAPPING
    # -------------------------------
    def configure_subjects(self):
        """Configure subject mapping and identify active subject columns."""
        print(f"\n{self.GREEN}2. SUBJECT CONFIGURATION")
        print("=" * 60)

        subjects_sql = f"""
            SELECT subject_id, subject_name, subject_code, subject_short
            FROM tbl_school_subjects 
            WHERE is_present_{self.class_id} = True
        """
        self.subjects_df = pd.read_sql(subjects_sql, self.conn)

        def sort_subjects_vba_logic(df_sub):
            """Sort subjects using VBA-like logic."""
            has_41_42 = (41 in df_sub['subject_id'].values) or (42 in df_sub['subject_id'].values)
            if has_41_42:
                def priority(row):
                    if row['subject_id'] == 41: return (0, row['subject_code'])
                    if row['subject_id'] == 42: return (1, row['subject_code'])
                    return (2, row['subject_code'])
                df_sub['priority'] = df_sub.apply(priority, axis=1)
                df_sub = df_sub.sort_values('priority').drop('priority', axis=1)
            else:
                df_sub = df_sub.sort_values('subject_id')
            return df_sub.reset_index(drop=True)

        self.subjects_df = sort_subjects_vba_logic(self.subjects_df)

        print(f"{self.CYAN}SUBJECTS CONFIGURED FOR THIS CLASS:")
        subjects_display = self.subjects_df[['subject_id', 'subject_short', 'subject_name']].copy()
        subjects_display['No'] = range(1, len(subjects_display) + 1)
        table = subjects_display[['No', 'subject_id', 'subject_short', 'subject_name']].head(10)
        print(tabulate(table, headers='keys', tablefmt='fancy_grid', showindex=False))

        # Build subject column mapping
        self.subject_column_map = {}
        for i, subject_row in self.subjects_df.iterrows():
            if i < len(self.mark_columns):
                col_name = self.mark_columns[i]
                self.subject_column_map[col_name] = {
                    'subject_short': subject_row['subject_short'],
                    'subject_id': int(subject_row['subject_id']),
                    'subject_name': subject_row['subject_name']
                }

        # Identify valid subject columns with data
        self.valid_subject_cols = []
        for col in self.mark_columns:
            if col in self.df.columns:
                self.df[col] = pd.to_numeric(self.df[col], errors='coerce')
                if self.df[col].notna().any():
                    self.valid_subject_cols.append(col)

        print(f"\n{self.GREEN}ACTIVE SUBJECT COLUMNS: {len(self.valid_subject_cols)}")
        valid_cols_df = pd.DataFrame({
            'No': range(1, len(self.valid_subject_cols) + 1),
            'Column': self.valid_subject_cols,
            'Subject': [self.subject_column_map[col]['subject_name'] for col in self.valid_subject_cols]
        })
        print(tabulate(valid_cols_df.head(10), headers='keys', tablefmt='fancy_grid', showindex=False))

        # Keep only relevant columns
        keep_cols = ['result_id', 'student_id', 'full_name', 'sex'] + self.valid_subject_cols
        self.df = self.df[keep_cols].copy()

    # -------------------------------
    # 3. GRADE CALCULATION
    # -------------------------------
    def calculate_grades(self):
        """Calculate letter grades for all subjects."""
        print(f"\n{self.GREEN}3. GRADE CALCULATION PROCESS")
        print("=" * 60)

        def calculate_grade(mark):
            """Convert numeric mark to letter grade."""
            if pd.isna(mark): return None
            try:
                mark_float = float(mark)
                if mark_float >= 75: return 'A'
                if mark_float >= 65: return 'B'
                if mark_float >= 45: return 'C'
                if mark_float >= 30: return 'D'
                if mark_float >= 0: return 'F'
                return None
            except:
                return None

        # Apply grading to all subject columns
        for col in self.valid_subject_cols:
            self.df[f"{col}_grade"] = self.df[col].apply(calculate_grade)

        print(f"{self.CYAN}STUDENT MARKS SAMPLE (First 8 students):")
        sample_data = []
        for i in range(min(8, len(self.df))):
            row = self.df.iloc[i]
            sample_row = {'No': i+1, 'Student': row['full_name']}
            for j, col in enumerate(self.valid_subject_cols[:4]):
                short = self.subject_column_map[col]["subject_short"]
                mark = row[col]
                sample_row[f'{short}'] = f"{mark:.0f}" if pd.notna(mark) else "—"
            sample_data.append(sample_row)

        sample_df = pd.DataFrame(sample_data)
        print(tabulate(sample_df, headers='keys', tablefmt='fancy_grid', showindex=False))

    # -------------------------------
    # 4. SUBJECT COUNT & TOTAL MARKS
    # -------------------------------
    def aggregate_performance(self):
        """Calculate subject counts, total marks, and average marks."""
        print(f"\n{self.GREEN}4. ACADEMIC AGGREGATION")
        print("=" * 60)

        # Count actual subjects attempted
        self.df['subject_count_real'] = self.df[self.valid_subject_cols].notna().sum(axis=1)

        if self.FLAT_RATE:
            # Flat rate: take top N subjects
            self.df['subject_count'] = self.BASE_SUBJECTS
            self.df['total_marks'] = self.df[self.valid_subject_cols].apply(
                lambda row: sum(sorted([m for m in row if not pd.isna(m)], reverse=True)[:self.BASE_SUBJECTS]), 
                axis=1
            )
        else:
            # Dynamic: count all or minimum base subjects
            self.df['subject_count'] = self.df['subject_count_real'].apply(
                lambda x: self.BASE_SUBJECTS if x <= self.BASE_SUBJECTS else x
            )
            self.df['total_marks'] = self.df[self.valid_subject_cols].sum(axis=1)

        # Calculate average marks
        self.df['avg_marks'] = np.where(
            self.df['subject_count'] > 0, 
            self.df['total_marks'] / self.df['subject_count'], 
            np.nan
        )

        print(f"{self.CYAN}STUDENT PERFORMANCE SUMMARY:")
        summary_data = []
        for i in range(min(8, len(self.df))):
            row = self.df.iloc[i]
            summary_data.append({
                'No': i+1,
                'Student': row['full_name'],
                'Real Subjects': int(row['subject_count_real']),
                'Counted Subjects': int(row['subject_count']),
                'Total Marks': f"{row['total_marks']:.0f}",
                'Average Marks': f"{row['avg_marks']:.2f}"
            })

        summary_df = pd.DataFrame(summary_data)
        print(tabulate(summary_df, headers='keys', tablefmt='fancy_grid', showindex=False))

    # -------------------------------
    # 5. POINTS & DIVISION
    # -------------------------------
    def calculate_points_and_division(self):
        """Calculate ranking points and division for each student."""
        print(f"\n{self.GREEN}5. POINTS & DIVISION CALCULATION")
        print("=" * 60)

        grade_points = {'A': 1, 'B': 2, 'C': 3, 'D': 4, 'F': 5}

        def calculate_points_and_division(row):
            """Calculate points and division for a student."""
            subject_count_real = row['subject_count_real']
            
            # Handle students with no subjects
            if subject_count_real == 0:
                return None, 'ABS'
            
            # Collect valid grades
            valid_grades = [row[f"{col}_grade"] for col in self.valid_subject_cols 
                           if row[f"{col}_grade"] in grade_points]
            
            if not valid_grades:
                return None, 'ABS'
            
            # Take best 7 subjects
            points_list = sorted([grade_points[g] for g in valid_grades])[:7]
            points_list = points_list + [5] * (7 - len(points_list))
            points = sum(points_list)
            
            # Handle incomplete students
            if self.INCLUDE_INC:
                if subject_count_real < self.BASE_SUBJECTS:
                    return None, 'INC'
            else:
                if subject_count_real < self.BASE_SUBJECTS:
                    potential_points = points + (5 * (self.BASE_SUBJECTS - subject_count_real))
                    return None, '0' if potential_points >= 34 else 'IV'
            
            # Assign division based on points
            if points <= 17: return points, 'I'
            if points <= 22: return points, 'II'
            if points <= 25: return points, 'III'
            if points <= 33: return points, 'IV'
            return points, '0'

        # Apply calculation to all students
        points_division = self.df.apply(calculate_points_and_division, axis=1, result_type='expand')
        self.df['points'] = points_division[0]
        self.df['division'] = points_division[1]

        print(f"{self.CYAN}STUDENT POINTS & DIVISIONS:")
        points_data = []
        for i in range(min(8, len(self.df))):
            row = self.df.iloc[i]
            points_data.append({
                'No': i+1,
                'Student': row['full_name'],
                'Subjects Attempted': int(row['subject_count_real']),
                'Points': int(row['points']) if pd.notna(row['points']) else 'N/A',
                'Division': row['division']
            })

        points_df = pd.DataFrame(points_data)
        print(tabulate(points_df, headers='keys', tablefmt='fancy_grid', showindex=False))

        print(f"\n{self.MAGENTA}DIVISION DISTRIBUTION:")
        div_counts = self.df['division'].value_counts().reset_index()
        div_counts.columns = ['Division', 'Students']
        div_counts['Percentage'] = (div_counts['Students'] / len(self.df) * 100).round(1)
        div_counts = div_counts.sort_values('Students', ascending=False)
        print(tabulate(div_counts, headers='keys', tablefmt='fancy_grid', showindex=False))

    # -------------------------------
    # 6. FLEXIBLE RANKING SYSTEM
    # -------------------------------
    def rank_students(self):
        """
        Flexible ranking system supporting:
        - Any combination of sort columns in any order
        - Automatic sort direction inference
        - Configurable ranking method (min, max, average, first, dense)
        - Proper NULL handling
        """
        print(f"\n{self.GREEN}6. ACADEMIC RANKING ({self.RANKING_METHOD.upper()} METHOD)")
        print("=" * 60)

        # Separate ABS students (they don't get ranked)
        abs_students = self.df[self.df['division'] == 'ABS'].copy()
        valid_students = self.df[self.df['division'] != 'ABS'].copy()

        print(f"{self.YELLOW}RANKING BREAKDOWN:")
        print(f"   • ABS Students: {self.WHITE}{len(abs_students):,}")
        print(f"   • Valid Students: {self.WHITE}{len(valid_students):,}")
        print(f"   • Sort Columns: {self.WHITE}{', '.join(self.sort_columns)}")
        print(f"   • Sort Directions: {self.WHITE}{', '.join(['ASC' if asc else 'DESC' for asc in self.ascending])}")
        print(f"   • Ranking Method: {self.WHITE}{self.RANKING_METHOD}")

        if len(valid_students) == 0:
            print(f"{self.RED}No valid students to rank!")
            return

        # Handle ranking_points column for students with NULL points
        if 'ranking_points' in self.sort_columns:
            max_points = valid_students['points'].max() if not valid_students['points'].isna().all() else 35
            valid_students['ranking_points'] = valid_students['points'].fillna(max_points + 1)

        # Verify all sort columns exist
        missing_cols = [col for col in self.sort_columns if col not in valid_students.columns]
        if missing_cols:
            print(f"{self.RED}Error: Missing columns for sorting: {missing_cols}")
            print(f"{self.YELLOW}Available columns: {list(valid_students.columns)}")
            return

        # Sort students by specified columns with appropriate directions
        valid_students = valid_students.sort_values(
            self.sort_columns,
            ascending=self.ascending,
            na_position='last'
        )

        # Apply ranking based on selected method
        if self.RANKING_METHOD == 'first':
            # Simple sequential ranking based on sort order
            valid_students['position'] = range(1, len(valid_students) + 1)
            
        elif self.RANKING_METHOD == 'dense':
            # Dense ranking: no gaps (1, 2, 2, 3, 4)
            valid_students['position'] = (
                valid_students.groupby(self.sort_columns, dropna=False)
                .ngroup() + 1
            )
            
        elif self.RANKING_METHOD in ['min', 'max', 'average']:
            # Create groups based on tie-breaking columns
            valid_students['temp_position'] = range(1, len(valid_students) + 1)
            
            # Group by all sort columns to identify ties
            for cols_tuple, group in valid_students.groupby(self.sort_columns, dropna=False):
                indices = group.index
                positions = group['temp_position'].values
                
                if self.RANKING_METHOD == 'min':
                    # All tied students get the minimum position
                    rank_value = positions.min()
                elif self.RANKING_METHOD == 'max':
                    # All tied students get the maximum position
                    rank_value = positions.max()
                else:  # average
                    # All tied students get the average position
                    rank_value = positions.mean()
                
                valid_students.loc[indices, 'position'] = rank_value
            
            valid_students.drop('temp_position', axis=1, inplace=True)
            
            # Round average positions if using average method
            if self.RANKING_METHOD == 'average':
                valid_students['position'] = valid_students['position'].round(1)
            else:
                valid_students['position'] = valid_students['position'].astype(int)
        
        # Set out_of value
        valid_students['out_of'] = len(valid_students)

        # Merge rankings back to main dataframe
        self.df = self.df.merge(
            valid_students[['result_id', 'position', 'out_of']], 
            on='result_id', 
            how='left'
        )

        # Display top 10 students
        print(f"\n{self.GREEN}TOP 10 RANKED STUDENTS:")
        display_cols_base = ['full_name']
        
        # Add sort columns to display (use 'points' instead of 'ranking_points' for display)
        for col in self.sort_columns:
            if col == 'ranking_points':
                if 'points' not in display_cols_base:
                    display_cols_base.append('points')
            elif col not in display_cols_base:
                display_cols_base.append(col)
        
        display_cols_base.extend(['position', 'division'])
        
        # Filter to only existing columns
        display_cols = [c for c in display_cols_base if c in valid_students.columns]
        
        top_10 = valid_students[display_cols].head(10).copy()
        
        # Format numeric columns
        if 'avg_marks' in top_10.columns:
            top_10['avg_marks'] = top_10['avg_marks'].round(2)
        if 'points' in top_10.columns:
            top_10['points'] = top_10['points'].fillna('N/A')
        
        top_10.insert(0, 'No', range(1, len(top_10) + 1))
        print(tabulate(top_10, headers='keys', tablefmt='fancy_grid', showindex=False))

        # Display bottom 10 students
        print(f"\n{self.RED}BOTTOM 10 RANKED STUDENTS:")
        bottom_10 = valid_students[display_cols].tail(10).copy()
        
        if 'avg_marks' in bottom_10.columns:
            bottom_10['avg_marks'] = bottom_10['avg_marks'].round(2)
        if 'points' in bottom_10.columns:
            bottom_10['points'] = bottom_10['points'].fillna('N/A')
        
        bottom_10.insert(0, 'No', range(1, len(bottom_10) + 1))
        print(tabulate(bottom_10, headers='keys', tablefmt='fancy_grid', showindex=False))

        # Show tie statistics for methods that handle ties
        if self.RANKING_METHOD in ['min', 'max', 'average', 'dense']:
            # Count positions with ties
            position_counts = valid_students.groupby('position').size()
            tied_positions = position_counts[position_counts > 1]
            
            if len(tied_positions) > 0:
                print(f"\n{self.YELLOW}TIED RANKINGS ({self.RANKING_METHOD.upper()} method):")
                print(f"   • Total positions with ties: {self.WHITE}{len(tied_positions)}")
                print(f"   • Largest tie group: {self.WHITE}{tied_positions.max()} students")
                print(f"   • Total students in ties: {self.WHITE}{tied_positions.sum()}")
                
                # Show example of tied ranking
                example_pos = tied_positions.index[0]
                tied_students = valid_students[valid_students['position'] == example_pos][
                    ['full_name'] + [c for c in self.sort_columns if c != 'ranking_points' and c in valid_students.columns] + ['position']
                ].head(5).copy()
                
                tied_students.insert(0, 'No', range(1, len(tied_students) + 1))
                print(f"\n{self.CYAN}Example tied group (Position {example_pos}):")
                print(tabulate(tied_students, headers='keys', tablefmt='fancy_grid', showindex=False))

    # -------------------------------
    # 7. SUBJECT-WISE RANKING
    # -------------------------------
    def rank_subjects(self):
        """Rank students within each subject based on marks."""
        print(f"\n{self.GREEN}7. SUBJECT-WISE RANKING (MARKS DESC, NOT NULL)")
        print("=" * 60)

        print(f"{self.CYAN}CALCULATING SUBJECT POSITIONS...")
        
        for col in self.valid_subject_cols:
            # Only rank students who have marks for this subject
            subject_rank_df = self.df[self.df[col].notna()].copy()
            
            if len(subject_rank_df) > 0:
                # Sort by marks descending
                subject_rank_df = subject_rank_df.sort_values(col, ascending=False)
                
                # Apply min ranking (students with same marks get same position)
                subject_rank_df[f'{col}_pos'] = subject_rank_df[col].rank(
                    method='min', 
                    ascending=False
                ).astype(int)
                
                subject_rank_df[f'{col}_out_of'] = len(subject_rank_df)
                
                # Merge back to main dataframe
                self.df = self.df.merge(
                    subject_rank_df[['result_id', f'{col}_pos', f'{col}_out_of']],
                    on='result_id', 
                    how='left'
                )

        # Display sample subject ranking
        print(f"{self.CYAN}SUBJECT RANKING SAMPLE (First Subject - Top 10):")
        if len(self.valid_subject_cols) > 0:
            first_subject = self.valid_subject_cols[0]
            subject_name = self.subject_column_map[first_subject]['subject_name']
            print(f"{self.YELLOW}Subject: {self.WHITE}{subject_name}")
            
            subject_top_10 = self.df[self.df[first_subject].notna()].nlargest(10, first_subject)[
                ['full_name', first_subject, f'{first_subject}_pos', f'{first_subject}_out_of']
            ].copy()
            
            subject_top_10.insert(0, 'No', range(1, len(subject_top_10) + 1))
            subject_top_10 = subject_top_10.rename(columns={
                first_subject: 'Marks',
                f'{first_subject}_pos': 'Position',
                f'{first_subject}_out_of': 'Out Of'
            })
            print(tabulate(subject_top_10, headers='keys', tablefmt='fancy_grid', showindex=False))

    # -------------------------------
    # 8. NECTA RESULTS STRING
    # -------------------------------
    def generate_necta_strings(self):
        """Generate NECTA-format result strings for each student."""
        print(f"\n{self.GREEN}8. NECTA RESULTS STRING GENERATION")
        print("=" * 60)

        # Check if new curriculum (subjects 41 or 42 present)
        is_new_curriculum = (41 in self.subjects_df['subject_id'].values) or \
                           (42 in self.subjects_df['subject_id'].values)
        max_compulsory = 8 if is_new_curriculum else 7

        def build_necta_string(row):
            """Build NECTA result string for a student."""
            parts = []
            for i, col in enumerate(self.valid_subject_cols):
                if i >= len(self.subjects_df): 
                    break
                    
                mark = row[col]
                grade = row[f"{col}_grade"]
                short = self.subject_column_map[col]['subject_short']
                
                # Compulsory subjects (first N subjects)
                if i < max_compulsory:
                    if pd.isna(grade) or grade is None:
                        parts.append(f" {short}-'X'")
                    else:
                        parts.append(f" {short} {int(mark)} -'{grade}'")
                # Optional subjects (only if attempted)
                elif not (pd.isna(grade) or grade is None):
                    parts.append(f" {short} {int(mark)} -'{grade}'")
            
            result = "".join(parts).strip()
            if result.endswith("-"):
                result = result[:-1].strip()
            return result

        tqdm.pandas(desc="Generating NECTA strings")
        self.df['necta_results'] = self.df.progress_apply(build_necta_string, axis=1)

        print(f"{self.CYAN}NECTA RESULTS SAMPLE:")
        necta_sample = self.df[['full_name', 'necta_results']].head(6).copy()
        necta_sample.insert(0, 'No', range(1, len(necta_sample) + 1))
        
        # Truncate long result strings for display
        for i, row in necta_sample.iterrows():
            txt = row['necta_results']
            necta_sample.at[i, 'necta_results'] = txt[:100] + '...' if len(txt) > 100 else txt
        
        print(tabulate(necta_sample, headers='keys', tablefmt='fancy_grid', showindex=False))

    # -------------------------------
    # 9. UPDATE DATABASE (OPTIMIZED)
    # -------------------------------
    def update_database(self):
        """Update database with calculated results using batch processing."""
        print(f"\n{self.GREEN}9. DATABASE UPDATE PROCESS (OPTIMIZED)")
        print("=" * 60)

        # Build list of columns to update
        update_columns = [
            'necta_results', 'subject_count', 'total_marks', 
            'points', 'division', 'position', 'out_of'
        ]
        
        # Add grade columns
        for col in self.valid_subject_cols:
            update_columns.append(f"{col}_grade")
            if f"{col}_pos" in self.df.columns:
                update_columns.extend([f"{col}_pos", f"{col}_out_of"])

        print(f"{self.YELLOW}UPDATING {len(update_columns)} COLUMNS ACROSS {len(self.df):,} RECORDS")

        # Prepare SQL update statement
        set_clause = ", ".join([f"{col} = ?" for col in update_columns])
        update_sql = f"UPDATE tbl_student_exam_results SET {set_clause} WHERE result_id = ?"

        # Prepare all data at once (vectorized operation)
        update_data = []
        for _, row in self.df.iterrows():
            values = []
            for col in update_columns:
                value = row[col]
                if pd.isna(value) or value is None:
                    values.append(None)
                else:
                    # Convert numpy types to Python types
                    if hasattr(value, 'item'):
                        value = value.item()
                    values.append(value)
            values.append(int(row['result_id']))
            update_data.append(tuple(values))

        # Execute in batches with transaction optimization
        batch_size = 20  # Optimal batch size for Access DB
        success_count = 0
        
        try:
            # Disable autocommit for better performance
            self.cursor.connection.autocommit = False
            
            with tqdm(total=len(update_data), desc="Updating Database") as pbar:
                for i in range(0, len(update_data), batch_size):
                    batch = update_data[i:i + batch_size]
                    
                    # Use executemany for batch execution
                    self.cursor.executemany(update_sql, batch)
                    success_count += len(batch)
                    
                    # Commit each batch
                    self.conn.commit()
                    
                    # Update progress bar
                    pbar.update(len(batch))
            
            # Re-enable autocommit
            self.cursor.connection.autocommit = True
            
        except Exception as e:
            print(f"{self.RED}Error during batch update: {e}")
            self.conn.rollback()
            self.cursor.connection.autocommit = True
            raise

        status = f"{self.GREEN}COMPLETE" if success_count == len(self.df) else f"{self.RED}ISSUE"
        print(f"{self.CYAN}DATABASE UPDATE {status}: {success_count:,}/{len(self.df):,} records")
        return success_count

    # -------------------------------
    # 9.5 FINALIZE NECTA WITH AVG FROM DB
    # -------------------------------
    def finalize_necta_with_avg(self):
        """Append average marks and grade to NECTA results string."""
        print(f"\n{self.GREEN}9.5 FINALIZE NECTA WITH AVG FROM DB")
        print("=" * 60)

        # Query avg_marks from DB after update
        avg_sql = """
            SELECT result_id, avg_marks
            FROM tbl_student_exam_results
            WHERE exam_id = ? AND subject_count > 0
        """
        avg_df = pd.read_sql(avg_sql, self.conn, params=[self.EXAM_ID])

        # Merge into main df
        self.df = self.df.merge(
            avg_df[['result_id', 'avg_marks']],
            on='result_id', 
            how='left', 
            suffixes=('', '_db')
        )

        # Grade calculation function
        def calculate_grade(mark):
            if pd.isna(mark): 
                return None
            try:
                mark_float = float(mark)
                if mark_float >= 75: return 'A'
                if mark_float >= 65: return 'B'
                if mark_float >= 45: return 'C'
                if mark_float >= 30: return 'D'
                if mark_float >= 0: return 'F'
                return None
            except:
                return None

        # Calculate avg_grade for each row
        self.df['avg_grade'] = self.df['avg_marks'].apply(calculate_grade)

        # Append AVG part to existing necta_results
        def append_avg(row):
            base = row['necta_results']
            if pd.isna(row['avg_marks']) or pd.isna(row['avg_grade']):
                return base
            
            avg_marks_formatted = f"{row['avg_marks']:.2f}"
            return f"{base} AVG {avg_marks_formatted} -'{row['avg_grade']}'".strip()

        self.df['necta_results'] = self.df.apply(append_avg, axis=1)

        # Update BOTH necta_results AND avg_grade in DB
        final_sql = """
            UPDATE tbl_student_exam_results
            SET necta_results = ?, avg_grade = ?
            WHERE result_id = ?
        """
        
        # Prepare all update data
        update_data = [
            (str(row['necta_results']), 
             str(row['avg_grade']) if not pd.isna(row['avg_grade']) else None, 
             int(row['result_id']))
            for _, row in self.df.iterrows()
        ]
        
        # Execute updates in batches with progress bar
        batch_size = 20  # Optimal batch size for Access DB
        with tqdm(total=len(update_data), desc="Final NECTA Update") as pbar:
            for i in range(0, len(update_data), batch_size):
                batch = update_data[i:i + batch_size]
                self.cursor.executemany(final_sql, batch)
                pbar.update(len(batch))
        
        self.conn.commit()

        print(f"{self.CYAN}FINAL NECTA SAMPLE (WITH AVG):")
        final_sample = self.df[['full_name', 'avg_marks', 'avg_grade', 'necta_results']].head(22).copy()
        final_sample.insert(0, 'No', range(1, len(final_sample) + 1))
        
        # Format avg_marks and truncate NECTA strings
        for i, row in final_sample.iterrows():
            txt = row['necta_results']
            if len(txt) > 80:
                final_sample.at[i, 'necta_results'] = '...' + txt[-80:]
            else:
                final_sample.at[i, 'necta_results'] = txt
            
            final_sample.at[i, 'avg_marks'] = f"{row['avg_marks']:.2f}" if not pd.isna(row['avg_marks']) else "N/A"
        
        print(tabulate(final_sample, headers='keys', tablefmt='fancy_grid', showindex=False))

    # -------------------------------
    # 10. COMPETENCY ANALYSIS
    # -------------------------------
    def update_competency_analysis(self):
        """Update competency analysis table with grade distributions."""
        if not self.UPDATE_COMPETENCY:
            return

        print(f"\n{self.GREEN}10. COMPETENCY ANALYSIS UPDATE")
        print("=" * 60)
        
        # Clear existing competency data for this exam
        self.cursor.execute("DELETE FROM tbl_competency WHERE exam_id = ?", (self.EXAM_ID,))
        
        def calculate_gpa(A, B, C, D, F):
            total = A + B + C + D + F
            return (A*1 + B*2 + C*3 + D*4 + F*5) / total if total > 0 else None
        
        def get_competency_level(gpa):
            if gpa is None: 
                return "No Data"
            if gpa >= 4.6: return "Grade F (Fail)"
            if gpa >= 3.6: return "Grade D (Satisfactory)"
            if gpa >= 2.6: return "Grade C (Good)"
            if gpa >= 1.6: return "Grade B (Very Good)"
            if gpa >= 1.0: return "Grade A (Excellent)"
            return "No Data"
        
        competency_data = []
        
        for _, subject in self.subjects_df.iterrows():
            subject_id = int(subject['subject_id'])
            
            # Find corresponding column
            col_name = next(
                (col for col, info in self.subject_column_map.items() 
                if info['subject_id'] == subject_id), 
                None
            )
            
            if not col_name: 
                print(f"{self.YELLOW}Warning: No column mapping found for subject_id {subject_id}")
                continue
            
            grade_col = f"{col_name}_grade"
            
            # Check if grade column exists and has valid data
            if grade_col not in self.df.columns:
                print(f"{self.YELLOW}Warning: Grade column '{grade_col}' not found for '{subject['subject_name']}'")
                # Insert NULL record for this subject
                self.cursor.execute("""
                    INSERT INTO tbl_competency 
                    (exam_id, subject_id, A_s, B_s, C_s, D_s, F_s, gpa, competency_level)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (self.EXAM_ID, subject_id, 0, 0, 0, 0, 0, None, "No Data"))
                continue
                
            # Get grade counts - handle null/empty cases
            grade_series = self.df[grade_col]
            
            # Check if all grades are null/empty
            if grade_series.isna().all() or grade_series.empty:
                print(f"{self.YELLOW}No valid grade data for '{subject['subject_name']}' - all null/empty")
                A, B, C, D, F = 0, 0, 0, 0, 0
                gpa = None
                level = "No Data"
            else:
                # Count non-null grades only
                valid_grades = grade_series.dropna()
                if valid_grades.empty:
                    print(f"{self.YELLOW}No valid grade data for '{subject['subject_name']}' - all dropped as null")
                    A, B, C, D, F = 0, 0, 0, 0, 0
                    gpa = None
                    level = "No Data"
                else:
                    counts = valid_grades.value_counts()
                    A = int(counts.get('A', 0))
                    B = int(counts.get('B', 0))
                    C = int(counts.get('C', 0))
                    D = int(counts.get('D', 0))
                    F = int(counts.get('F', 0))
                    total = A + B + C + D + F

                    # Only calculate GPA if we have valid grades
                    if total > 0:
                        gpa = float(calculate_gpa(A, B, C, D, F))
                        level = get_competency_level(gpa)
                        print(f"{self.GREEN}Processed '{subject['subject_name']}': A={A}, B={B}, C={C}, D={D}, F={F}, GPA={gpa:.4f}")
                    else:
                        gpa = None
                        level = "No Data"
                        print(f"{self.YELLOW}No valid grades for '{subject['subject_name']}' - all counts are zero")
            
            competency_data.append({
                'No': len(competency_data) + 1,
                'Subject': subject['subject_name'],
                'A': A, 'B': B, 'C': C, 'D': D, 'F': F,
                'GPA': round(gpa, 4) if gpa is not None else "N/A",
                'Level': level
            })
            
            # Insert into database
            self.cursor.execute("""
                INSERT INTO tbl_competency 
                (exam_id, subject_id, A_s, B_s, C_s, D_s, F_s, gpa, competency_level)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (self.EXAM_ID, subject_id, A, B, C, D, F, gpa, level))
        
        self.conn.commit()
        
        if competency_data:
            print(f"{self.MAGENTA}SUBJECT COMPETENCY ANALYSIS:")
            comp_df = pd.DataFrame(competency_data)
            print(tabulate(
                comp_df[['No', 'Subject', 'A', 'B', 'C', 'D', 'F', 'GPA', 'Level']].head(15),
                headers='keys', 
                tablefmt='fancy_grid', 
                showindex=False
            ))
            
            # Show summary
            valid_subjects = len([d for d in competency_data if d['GPA'] != "N/A"])
            total_subjects = len(competency_data)
            print(f"{self.CYAN}Summary: {valid_subjects}/{total_subjects} subjects have valid GPA data")
        else:
            print(f"{self.RED}No competency data was processed!")

    # -------------------------------
    # 11. FINAL SUMMARY REPORT
    # -------------------------------
    def generate_summary_report(self, success_count):
        """Generate comprehensive summary of processing results."""
        print(f"\n{self.GREEN}11. COMPREHENSIVE SUMMARY REPORT")
        print("=" * 60)

        valid_students = self.df[self.df['division'] != 'ABS']
        status_db = "UPDATED" if success_count == len(self.df) else "ISSUE"

        summary_data = {
            'Category': [
                'STUDENT DATA',
                'SUBJECT DATA', 
                'PERFORMANCE',
                'DIVISIONS',
                'RANKING',
                'DATABASE'
            ],
            'Metric': [
                'Total Students Processed',
                'Active Subjects Count',
                'Average Total Marks',
                'Division I Students',
                'Unique Ranking Positions',
                'Successfully Updated Records'
            ],
            'Value': [
                f"{len(self.df):,}",
                f"{len(self.valid_subject_cols)}",
                f"{self.df['total_marks'].mean():.1f}",
                f"{len(self.df[self.df['division'] == 'I']):,}",
                f"{valid_students['position'].nunique():,}" if 'position' in valid_students.columns else "0",
                f"{success_count:,}/{len(self.df):,}"
            ],
            'Status': [
                'COMPLETE',
                'CONFIGURED',
                'CALCULATED',
                'CLASSIFIED',
                'RANKED',
                status_db
            ]
        }

        summary_df = pd.DataFrame(summary_data)
        print(tabulate(summary_df, headers='keys', tablefmt='fancy_grid', showindex=False))

        print(f"\n{self.YELLOW}RANKING METHODOLOGY:")
        sort_desc = ', '.join([
            f"{col} {'ASC' if asc else 'DESC'}" 
            for col, asc in zip(self.sort_columns, self.ascending)
        ])
        print(f"   • Sort order: {self.WHITE}{sort_desc}")
        print(f"   • Ranking method: {self.WHITE}{self.RANKING_METHOD}")
        print(f"   • Students with NULL points: {self.WHITE}Assigned penalty value")
        print(f"   • ABS students: {self.WHITE}Not ranked")
        print(f"   • Subject ranking: {self.WHITE}Marks DESC (not null students only)")

        now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(f"\n{self.CYAN}PROCESS COMPLETED AT: {self.WHITE}{now}")

    # -------------------------------
    # MAIN EXECUTION
    # -------------------------------
    def run(self):
        """Execute the complete O-Level processing pipeline."""
        try:
            self._connect_to_database()
            self.load_data()
            self.configure_subjects()
            self.calculate_grades()
            self.aggregate_performance()
            self.calculate_points_and_division()
            self.rank_students()
            self.rank_subjects()
            self.generate_necta_strings()
            success_count = self.update_database()
            self.finalize_necta_with_avg()
            self.update_competency_analysis()
            self.generate_summary_report(success_count)

        except Exception as e:
            print(f"{self.RED}ERROR: {e}")
            import traceback
            traceback.print_exc()
            raise
        finally:
            if self.conn:
                self.conn.close()

        print("\n" + "=" * 60)
        print(f"{self.GREEN}PROCESS COMPLETED SUCCESSFULLY!")
        print("=" * 60)


# ========================================
# USAGE EXAMPLES:
# ========================================

if __name__ == "__main__":
    # Example 1: Default ranking (points, avg_marks, subject_count)
    processor = OlevelProcessor(
        exam_id="MID420251027",
        db_path=r"C:\Kiyabo App\backend\Kiyabo App Backend v2.0.0.accdb",
        base_subjects=7,
        flat_rate=True,
        include_inc=True,
        update_competency=True,
        sort_columns=['ranking_points', 'avg_marks', 'subject_count_real'],
        ranking_method='min'
    )
    processor.run()

    # Example 2: Rank by points only (ascending)
    processor2 = OlevelProcessor(
        exam_id="MID420251027",
        db_path=r"C:\Kiyabo App\backend\Kiyabo App Backend v2.0.0.accdb",
        sort_columns=['ranking_points'],
        ranking_method='min'
    )
    # processor2.run()

    # Example 3: Rank by average marks only (descending)
    processor3 = OlevelProcessor(
        exam_id="MID420251027",
        db_path=r"C:\Kiyabo App\backend\Kiyabo App Backend v2.0.0.accdb",
        sort_columns=['avg_marks'],
        ranking_method='dense'
    )
    # processor3.run()

    # Example 4: Custom order - prioritize subject count
    processor4 = OlevelProcessor(
        exam_id="MID420251027",
        db_path=r"C:\Kiyabo App\backend\Kiyabo App Backend v2.0.0.accdb",
        sort_columns=['subject_count_real', 'avg_marks', 'ranking_points'],
        ranking_method='min'
    )
    # processor4.run()

    # Example 5: Using 'average' method for ties
    processor5 = OlevelProcessor(
        exam_id="MID420251027",
        db_path=r"C:\Kiyabo App\backend\Kiyabo App Backend v2.0.0.accdb",
        sort_columns=['ranking_points', 'avg_marks'],
        ranking_method='average'
    )
    # processor5.run()

    # Example 6: Rank by total marks, then subject count
    processor6 = OlevelProcessor(
        exam_id="MID420251027",
        db_path=r"C:\Kiyabo App\backend\Kiyabo App Backend v2.0.0.accdb",
        sort_columns=['total_marks', 'subject_count_real'],
        ranking_method='first'
    )
    # processor6.run()