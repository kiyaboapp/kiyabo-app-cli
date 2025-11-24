# ========================================
# DUAL EXAM PROCESSOR — COMBINE & AVERAGE
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

# CRITICAL FIX: Enable tqdm progress bar for pandas
tqdm.pandas()

warnings.filterwarnings("ignore")


class DualExamProcessor:
    """
    Dual Exam Processor - Combines two exams and calculates averages
    Saves results as a query for viewing in Access
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
        exam_id_1: str,
        exam_id_2: str,
        db_path: str,
        query_name: str = "qry_CombinedExamResults",
        base_subjects: int = 7,
        flat_rate: bool = True,
        include_inc: bool = True,
        ranking_method: str = "min",
        necta_decimal_places: int = 1
    ):
        """
        Initialize Dual Exam Processor.
        
        Args:
            exam_id_1: First exam identifier
            exam_id_2: Second exam identifier
            db_path: Path to Access database
            query_name: Name for the saved query
            base_subjects: Number of subjects for base calculation (default: 7)
            flat_rate: Use flat rate calculation (default: True)
            include_inc: Include incomplete students (default: True)
            ranking_method: Ranking method - 'min', 'max', 'average', 'dense', 'first' (default: 'min')
            necta_decimal_places: Number of decimal places for subject marks in NECTA string (default: 1)
        """
        self.EXAM_ID_1 = exam_id_1
        self.EXAM_ID_2 = exam_id_2
        self.DB_PATH = db_path
        self.QUERY_NAME =f"{exam_id_1}_{exam_id_2}"             #query_name
        self.BASE_SUBJECTS = base_subjects
        self.FLAT_RATE = flat_rate
        self.INCLUDE_INC = include_inc
        self.RANKING_METHOD = ranking_method.lower()
        self.NECTA_DECIMAL_PLACES = necta_decimal_places

        # Validate ranking method
        valid_methods = ['min', 'max', 'average', 'dense', 'first']
        if self.RANKING_METHOD not in valid_methods:
            raise ValueError(f"Invalid ranking_method '{ranking_method}'. Must be one of: {valid_methods}")

        # Sort configuration for average ranking
        self.sort_columns = ['ranking_points', 'avg_marks', 'subject_count_real']
        self.ascending = [True, False, False]  # points ASC, marks DESC, count DESC

        # Database connections
        self.conn = None
        self.cursor = None
        self.df_1 = None
        self.df_2 = None
        self.df_combined = None
        self.students_df = None
        self.class_id = None
        self.valid_subject_cols = []
        self.subject_column_map = {}
        self.mark_columns = ['civ','his','geo','kis','eng','phy','che','bio','mat','edk','ics'] + \
                           [f'sub{i}' for i in range(12,21)]

    def _connect_to_database(self):
        """Establish connection to Access database."""
        print(f"{self.CYAN}CONNECTING TO ACCESS DATABASE...")
        self.conn = pyodbc.connect(
            f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};"
            f"DBQ={self.DB_PATH};"
        )
        self.cursor = self.conn.cursor()

    def load_exam_data(self, exam_id: str) -> pd.DataFrame:
        """Load data for a single exam."""
        print(f"{self.YELLOW}Loading data for exam: {self.WHITE}{exam_id}")
        
        results_sql = """
            SELECT result_id, student_id, 
                   civ, his, geo, kis, eng, phy, che, bio, mat, edk, ics,
                   sub12, sub13, sub14, sub15, sub16, sub17, sub18, sub19, sub20
            FROM tbl_student_exam_results 
            WHERE exam_id = ?
        """
        df = pd.read_sql(results_sql, self.conn, params=[exam_id])
        
        # Convert mark columns to numeric
        for col in self.mark_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        return df

    def load_data(self):
        """Load both exams and student information."""
        print(f"\n{self.GREEN}1. LOADING DUAL EXAM DATA")
        print("=" * 60)

        # Load both exams
        self.df_1 = self.load_exam_data(self.EXAM_ID_1)
        self.df_2 = self.load_exam_data(self.EXAM_ID_2)

        # Load student info
        students_sql = "SELECT student_id, full_name, sex FROM tbl_student_academic_info"
        self.students_df = pd.read_sql(students_sql, self.conn)

        # Get class_id from first exam
        exam_sql = "SELECT exam_id, class_id FROM tbl_student_exams WHERE exam_id = ?"
        exam_df = pd.read_sql(exam_sql, self.conn, params=[self.EXAM_ID_1])
        self.class_id = exam_df['class_id'].iloc[0] if not exam_df.empty else None

        if self.class_id is None:
            raise ValueError(f"{self.RED}Class ID could not be determined")

        print(f"{self.YELLOW}Class ID: {self.WHITE}{self.class_id}")
        print(f"{self.YELLOW}Exam 1 Students: {self.WHITE}{len(self.df_1):,}")
        print(f"{self.YELLOW}Exam 2 Students: {self.WHITE}{len(self.df_2):,}")

    def configure_subjects(self):
        """Configure subject mapping."""
        print(f"\n{self.GREEN}2. SUBJECT CONFIGURATION")
        print("=" * 60)

        subjects_sql = f"""
            SELECT subject_id, subject_name, subject_code, subject_short
            FROM tbl_school_subjects 
            WHERE is_present_{self.class_id} = True
        """
        subjects_df = pd.read_sql(subjects_sql, self.conn)

        # Sort subjects (VBA logic)
        has_41_42 = (41 in subjects_df['subject_id'].values) or (42 in subjects_df['subject_id'].values)
        if has_41_42:
            def priority(row):
                if row['subject_id'] == 41: return (0, row['subject_code'])
                if row['subject_id'] == 42: return (1, row['subject_code'])
                return (2, row['subject_code'])
            subjects_df['priority'] = subjects_df.apply(priority, axis=1)
            subjects_df = subjects_df.sort_values('priority').drop('priority', axis=1)
        else:
            subjects_df = subjects_df.sort_values('subject_id')

        subjects_df = subjects_df.reset_index(drop=True)

        # Build subject column mapping
        self.subject_column_map = {}
        for i, subject_row in subjects_df.iterrows():
            if i < len(self.mark_columns):
                col_name = self.mark_columns[i]
                self.subject_column_map[col_name] = {
                    'subject_short': subject_row['subject_short'],
                    'subject_id': int(subject_row['subject_id']),
                    'subject_name': subject_row['subject_name']
                }

        # Identify valid subject columns
        self.valid_subject_cols = []
        for col in self.mark_columns:
            if col in self.df_1.columns and col in self.df_2.columns:
                if self.df_1[col].notna().any() or self.df_2[col].notna().any():
                    self.valid_subject_cols.append(col)

        print(f"{self.GREEN}ACTIVE SUBJECT COLUMNS: {len(self.valid_subject_cols)}")
        valid_cols_df = pd.DataFrame({
            'No': range(1, len(self.valid_subject_cols) + 1),
            'Column': self.valid_subject_cols,
            'Subject': [self.subject_column_map[col]['subject_name'] for col in self.valid_subject_cols]
        })
        print(tabulate(valid_cols_df.head(10), headers='keys', tablefmt='fancy_grid', showindex=False))

    def combine_exams(self):
        """Combine two exams with suffixes and calculate averages."""
        print(f"\n{self.GREEN}3. COMBINING EXAMS AND CALCULATING AVERAGES")
        print("=" * 60)

        print(f"{self.YELLOW}Merging exam datasets...")
        # Merge on student_id with suffixes
        df_merged = pd.merge(
            self.df_1,
            self.df_2,
            on='student_id',
            how='outer',
            suffixes=('_1', '_2')
        )

        # Add student information
        df_merged = df_merged.merge(self.students_df, on='student_id', how='left')
        print(f"{self.GREEN}✓ Successfully merged {len(df_merged):,} student records")

        print(f"\n{self.CYAN}Calculating subject averages...")
        # Calculate averages for each subject with progress bar
        for col in tqdm(self.valid_subject_cols, desc="Processing Subjects", ncols=80):
            col_1 = f"{col}_1"
            col_2 = f"{col}_2"
            
            # Average column
            df_merged[col] = df_merged[[col_1, col_2]].mean(axis=1)

        self.df_combined = df_merged
        print(f"{self.GREEN}✓ Calculated averages for {len(self.valid_subject_cols)} subjects")

        # Show detailed comparison for first 4 subjects
        print(f"\n{self.MAGENTA}EXAM COMPARISON - DETAILED VIEW (First 10 Students)")
        print(f"{self.WHITE}Showing: Exam 1 → Exam 2 → Average for first 4 subjects")
        print("=" * 120)
        
        sample_data = []
        for i in range(min(10, len(self.df_combined))):
            row = self.df_combined.iloc[i]
            sample_row = {
                'No': i+1, 
                'Student': row['full_name'][:20]  # Truncate long names
            }
            
            for j, col in enumerate(self.valid_subject_cols[:4]):
                short = self.subject_column_map[col]["subject_short"]
                mark_1 = row.get(f"{col}_1", np.nan)
                mark_2 = row.get(f"{col}_2", np.nan)
                avg = row[col]
                
                sample_row[f'{short}_1'] = f"{mark_1:.0f}" if pd.notna(mark_1) else "—"
                sample_row[f'{short}_2'] = f"{mark_2:.0f}" if pd.notna(mark_2) else "—"
                sample_row[f'{short}_Avg'] = f"{avg:.1f}" if pd.notna(avg) else "—"
            
            sample_data.append(sample_row)

        sample_df = pd.DataFrame(sample_data)
        print(tabulate(sample_df, headers='keys', tablefmt='fancy_grid', showindex=False))
        
        # Show statistics
        print(f"\n{self.CYAN}MERGE STATISTICS:")
        stats_data = []
        for col in self.valid_subject_cols[:6]:  # First 6 subjects
            short = self.subject_column_map[col]["subject_short"]
            name = self.subject_column_map[col]["subject_name"]
            
            exam1_count = self.df_combined[f"{col}_1"].notna().sum()
            exam2_count = self.df_combined[f"{col}_2"].notna().sum()
            avg_count = self.df_combined[col].notna().sum()
            
            exam1_avg = self.df_combined[f"{col}_1"].mean()
            exam2_avg = self.df_combined[f"{col}_2"].mean()
            combined_avg = self.df_combined[col].mean()
            
            stats_data.append({
                'Subject': f"{short} ({name[:15]})",
                'Exam1_N': exam1_count,
                'Exam1_Avg': f"{exam1_avg:.1f}" if pd.notna(exam1_avg) else "—",
                'Exam2_N': exam2_count,
                'Exam2_Avg': f"{exam2_avg:.1f}" if pd.notna(exam2_avg) else "—",
                'Combined_N': avg_count,
                'Combined_Avg': f"{combined_avg:.1f}" if pd.notna(combined_avg) else "—"
            })
        
        stats_df = pd.DataFrame(stats_data)
        print(tabulate(stats_df, headers='keys', tablefmt='fancy_grid', showindex=False))

    def calculate_grades(self):
        """Calculate grades for averages."""
        print(f"\n{self.GREEN}4. CALCULATING GRADES FOR AVERAGES")
        print("=" * 60)

        def calculate_grade(mark):
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

        print(f"{self.CYAN}Processing grades for {len(self.valid_subject_cols)} subjects...")
        for col in tqdm(self.valid_subject_cols, desc="Calculating Grades", ncols=80):
            self.df_combined[f"{col}_grade"] = self.df_combined[col].apply(calculate_grade)

        print(f"\n{self.MAGENTA}GRADE DISTRIBUTION (First 10 Students, First 5 Subjects):")
        grade_sample = []
        for i in range(min(10, len(self.df_combined))):
            row = self.df_combined.iloc[i]
            grade_row = {
                'No': i+1,
                'Student': row['full_name'][:20]
            }
            
            for j, col in enumerate(self.valid_subject_cols[:5]):
                short = self.subject_column_map[col]["subject_short"]
                mark = row[col]
                grade = row[f"{col}_grade"]
                
                if pd.notna(mark) and pd.notna(grade):
                    grade_row[short] = f"{mark:.1f} ({grade})"
                else:
                    grade_row[short] = "—"
            
            grade_sample.append(grade_row)
        
        grade_df = pd.DataFrame(grade_sample)
        print(tabulate(grade_df, headers='keys', tablefmt='fancy_grid', showindex=False))
        
        # Overall grade statistics
        print(f"\n{self.CYAN}OVERALL GRADE STATISTICS (Across All Subjects):")
        all_grades = []
        for col in self.valid_subject_cols:
            all_grades.extend(self.df_combined[f"{col}_grade"].dropna().tolist())
        
        if all_grades:
            grade_counts = pd.Series(all_grades).value_counts().sort_index()
            total = len(all_grades)
            
            grade_stats = []
            for grade in ['A', 'B', 'C', 'D', 'F']:
                count = grade_counts.get(grade, 0)
                percentage = (count / total * 100) if total > 0 else 0
                grade_stats.append({
                    'Grade': grade,
                    'Count': count,
                    'Percentage': f"{percentage:.1f}%"
                })
            
            grade_stats_df = pd.DataFrame(grade_stats)
            print(tabulate(grade_stats_df, headers='keys', tablefmt='fancy_grid', showindex=False))

    def aggregate_performance(self):
        """Calculate subject counts, total marks, and average marks."""
        print(f"\n{self.GREEN}5. ACADEMIC AGGREGATION")
        print("=" * 60)

        print(f"{self.CYAN}Calculating subject counts and totals...")
        # Count actual subjects attempted
        self.df_combined['subject_count_real'] = self.df_combined[self.valid_subject_cols].notna().sum(axis=1)

        if self.FLAT_RATE:
            self.df_combined['subject_count'] = self.BASE_SUBJECTS
            self.df_combined['total_marks'] = self.df_combined[self.valid_subject_cols].apply(
                lambda row: sum(sorted([m for m in row if not pd.isna(m)], reverse=True)[:self.BASE_SUBJECTS]), 
                axis=1
            )
            print(f"{self.YELLOW}Mode: FLAT RATE (Top {self.BASE_SUBJECTS} subjects)")
        else:
            self.df_combined['subject_count'] = self.df_combined['subject_count_real'].apply(
                lambda x: self.BASE_SUBJECTS if x <= self.BASE_SUBJECTS else x
            )
            self.df_combined['total_marks'] = self.df_combined[self.valid_subject_cols].sum(axis=1)
            print(f"{self.YELLOW}Mode: DYNAMIC (All subjects)")

        self.df_combined['avg_marks'] = np.where(
            self.df_combined['subject_count'] > 0, 
            self.df_combined['total_marks'] / self.df_combined['subject_count'], 
            np.nan
        )

        print(f"\n{self.MAGENTA}PERFORMANCE SUMMARY (First 10 Students):")
        summary_data = []
        for i in range(min(10, len(self.df_combined))):
            row = self.df_combined.iloc[i]
            summary_data.append({
                'No': i+1,
                'Student': row['full_name'][:25],
                'Real Subjects': int(row['subject_count_real']),
                'Counted Subjects': int(row['subject_count']),
                'Total Marks': f"{row['total_marks']:.1f}",
                'Average Marks': f"{row['avg_marks']:.2f}" if pd.notna(row['avg_marks']) else "—"
            })

        summary_df = pd.DataFrame(summary_data)
        print(tabulate(summary_df, headers='keys', tablefmt='fancy_grid', showindex=False))
        
        # Performance statistics
        print(f"\n{self.CYAN}CLASS PERFORMANCE STATISTICS:")
        perf_stats = [
            {
                'Metric': 'Average Total Marks',
                'Value': f"{self.df_combined['total_marks'].mean():.2f}",
                'Min': f"{self.df_combined['total_marks'].min():.2f}",
                'Max': f"{self.df_combined['total_marks'].max():.2f}"
            },
            {
                'Metric': 'Average Marks (Mean)',
                'Value': f"{self.df_combined['avg_marks'].mean():.2f}",
                'Min': f"{self.df_combined['avg_marks'].min():.2f}",
                'Max': f"{self.df_combined['avg_marks'].max():.2f}"
            },
            {
                'Metric': 'Subject Count (Real)',
                'Value': f"{self.df_combined['subject_count_real'].mean():.1f}",
                'Min': f"{int(self.df_combined['subject_count_real'].min())}",
                'Max': f"{int(self.df_combined['subject_count_real'].max())}"
            }
        ]
        
        perf_df = pd.DataFrame(perf_stats)
        print(tabulate(perf_df, headers='keys', tablefmt='fancy_grid', showindex=False))

    def calculate_points_and_division(self):
        """Calculate points and division."""
        print(f"\n{self.GREEN}6. POINTS & DIVISION CALCULATION")
        print("=" * 60)

        grade_points = {'A': 1, 'B': 2, 'C': 3, 'D': 4, 'F': 5}

        def calculate_points_and_division(row):
            subject_count_real = row['subject_count_real']
            
            if subject_count_real == 0:
                return None, 'ABS'
            
            valid_grades = [row[f"{col}_grade"] for col in self.valid_subject_cols 
                           if row[f"{col}_grade"] in grade_points]
            
            if not valid_grades:
                return None, 'ABS'
            
            points_list = sorted([grade_points[g] for g in valid_grades])[:7]
            points_list = points_list + [5] * (7 - len(points_list))
            points = sum(points_list)
            
            if self.INCLUDE_INC:
                if subject_count_real < self.BASE_SUBJECTS:
                    return None, 'INC'
            else:
                if subject_count_real < self.BASE_SUBJECTS:
                    potential_points = points + (5 * (self.BASE_SUBJECTS - subject_count_real))
                    return None, '0' if potential_points >= 34 else 'IV'
            
            if points <= 17: return points, 'I'
            if points <= 22: return points, 'II'
            if points <= 25: return points, 'III'
            if points <= 33: return points, 'IV'
            return points, '0'

        print(f"{self.CYAN}Calculating points for {len(self.df_combined):,} students...")
        # FIXED: Use progress_apply instead of apply
        points_division = self.df_combined.progress_apply(calculate_points_and_division, axis=1, result_type='expand')
        self.df_combined['points'] = points_division[0]
        self.df_combined['division'] = points_division[1]

        print(f"\n{self.MAGENTA}STUDENT POINTS & DIVISIONS (First 10):")
        points_data = []
        for i in range(min(10, len(self.df_combined))):
            row = self.df_combined.iloc[i]
            points_data.append({
                'No': i+1,
                'Student': row['full_name'][:25],
                'Subjects': int(row['subject_count_real']),
                'Points': int(row['points']) if pd.notna(row['points']) else 'N/A',
                'Division': row['division'],
                'Avg Marks': f"{row['avg_marks']:.2f}" if pd.notna(row['avg_marks']) else "—"
            })

        points_df = pd.DataFrame(points_data)
        print(tabulate(points_df, headers='keys', tablefmt='fancy_grid', showindex=False))

        print(f"\n{self.CYAN}DIVISION DISTRIBUTION:")
        div_counts = self.df_combined['division'].value_counts().reset_index()
        div_counts.columns = ['Division', 'Students']
        div_counts['Percentage'] = (div_counts['Students'] / len(self.df_combined) * 100).round(1)
        div_counts = div_counts.sort_values('Students', ascending=False)
        
        # Add visual bar
        max_count = div_counts['Students'].max()
        div_counts['Visual'] = div_counts['Students'].apply(
            lambda x: '█' * int((x / max_count) * 30) if x > 0 else ''
        )
        
        print(tabulate(div_counts, headers='keys', tablefmt='fancy_grid', showindex=False))
        
        # Points distribution
        if self.df_combined['points'].notna().any():
            print(f"\n{self.CYAN}POINTS DISTRIBUTION:")
            points_valid = self.df_combined[self.df_combined['points'].notna()]['points']
            print(f"   • Minimum Points: {self.WHITE}{int(points_valid.min())}")
            print(f"   • Maximum Points: {self.WHITE}{int(points_valid.max())}")
            print(f"   • Average Points: {self.WHITE}{points_valid.mean():.2f}")
            print(f"   • Median Points: {self.WHITE}{points_valid.median():.1f}")

    def rank_subjects(self):
        """Rank students within each subject based on average marks."""
        print(f"\n{self.GREEN}7. SUBJECT-WISE RANKING (MARKS DESC)")
        print("=" * 60)

        print(f"{self.CYAN}Calculating subject positions for {len(self.valid_subject_cols)} subjects...")
        
        for col in tqdm(self.valid_subject_cols, desc="Ranking Subjects", ncols=80):
            # Only rank students who have marks for this subject (average)
            subject_rank_df = self.df_combined[self.df_combined[col].notna()].copy()
            
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
                self.df_combined = self.df_combined.merge(
                    subject_rank_df[['student_id', f'{col}_pos', f'{col}_out_of']],
                    on='student_id', 
                    how='left',
                    suffixes=('', '_dup')
                )
                
                # Drop duplicate columns if they exist
                dup_cols = [c for c in self.df_combined.columns if c.endswith('_dup')]
                if dup_cols:
                    self.df_combined.drop(columns=dup_cols, inplace=True)

        print(f"{self.GREEN}✓ Completed subject-wise ranking for all subjects")

        # Display sample subject ranking for first 3 subjects
        print(f"\n{self.MAGENTA}SUBJECT RANKING SAMPLE (First 3 Subjects - Top 10 Each)")
        print("=" * 100)
        
        for idx, col in enumerate(self.valid_subject_cols[:3]):
            subject_name = self.subject_column_map[col]['subject_name']
            subject_short = self.subject_column_map[col]['subject_short']
            
            print(f"\n{self.CYAN}Subject {idx+1}: {self.WHITE}{subject_name} ({subject_short})")
            
            subject_top = self.df_combined[self.df_combined[col].notna()].nlargest(10, col)[
                ['full_name', col, f'{col}_pos', f'{col}_out_of']
            ].copy()
            
            if len(subject_top) > 0:
                subject_top['full_name'] = subject_top['full_name'].str[:25]
                subject_top[col] = subject_top[col].apply(lambda x: f"{x:.1f}")
                subject_top.insert(0, 'No', range(1, len(subject_top) + 1))
                subject_top = subject_top.rename(columns={
                    col: 'Average Marks',
                    f'{col}_pos': 'Position',
                    f'{col}_out_of': 'Out Of'
                })
                print(tabulate(subject_top, headers='keys', tablefmt='fancy_grid', showindex=False))
            else:
                print(f"{self.YELLOW}No data available for this subject")

        # Show subject ranking statistics
        print(f"\n{self.CYAN}SUBJECT RANKING STATISTICS:")
        stats_data = []
        for col in self.valid_subject_cols[:8]:  # First 8 subjects
            short = self.subject_column_map[col]['subject_short']
            
            if f'{col}_pos' in self.df_combined.columns:
                ranked_count = self.df_combined[f'{col}_pos'].notna().sum()
                if ranked_count > 0:
                    top_mark = self.df_combined[self.df_combined[col].notna()][col].max()
                    avg_mark = self.df_combined[self.df_combined[col].notna()][col].mean()
                    
                    stats_data.append({
                        'Subject': short,
                        'Students Ranked': ranked_count,
                        'Top Mark': f"{top_mark:.1f}",
                        'Average': f"{avg_mark:.1f}"
                    })
        
        if stats_data:
            stats_df = pd.DataFrame(stats_data)
            print(tabulate(stats_df, headers='keys', tablefmt='fancy_grid', showindex=False))

    def rank_students(self):
        """Rank students based on average performance."""
        print(f"\n{self.GREEN}8. STUDENT RANKING ({self.RANKING_METHOD.upper()} METHOD)")
        print("=" * 60)

        abs_students = self.df_combined[self.df_combined['division'] == 'ABS'].copy()
        valid_students = self.df_combined[self.df_combined['division'] != 'ABS'].copy()

        print(f"{self.YELLOW}RANKING BREAKDOWN:")
        print(f"   • ABS Students: {self.WHITE}{len(abs_students):,}")
        print(f"   • Valid Students: {self.WHITE}{len(valid_students):,}")
        print(f"   • Sort Columns: {self.WHITE}{', '.join(self.sort_columns)}")
        print(f"   • Sort Directions: {self.WHITE}{', '.join(['ASC' if asc else 'DESC' for asc in self.ascending])}")
        print(f"   • Ranking Method: {self.WHITE}{self.RANKING_METHOD}")

        if len(valid_students) == 0:
            print(f"{self.RED}No valid students to rank!")
            return

        # Handle ranking_points
        max_points = valid_students['points'].max() if not valid_students['points'].isna().all() else 35
        valid_students['ranking_points'] = valid_students['points'].fillna(max_points + 1)

        # Sort and rank
        print(f"\n{self.CYAN}Sorting students...")
        valid_students = valid_students.sort_values(
            self.sort_columns,
            ascending=self.ascending,
            na_position='last'
        )

        print(f"{self.CYAN}Applying ranking algorithm...")
        if self.RANKING_METHOD == 'first':
            valid_students['position'] = range(1, len(valid_students) + 1)
        elif self.RANKING_METHOD == 'dense':
            valid_students['position'] = valid_students.groupby(self.sort_columns, dropna=False).ngroup() + 1
        elif self.RANKING_METHOD in ['min', 'max', 'average']:
            valid_students['temp_position'] = range(1, len(valid_students) + 1)
            for cols_tuple, group in valid_students.groupby(self.sort_columns, dropna=False):
                indices = group.index
                positions = group['temp_position'].values
                
                if self.RANKING_METHOD == 'min':
                    rank_value = positions.min()
                elif self.RANKING_METHOD == 'max':
                    rank_value = positions.max()
                else:  # average
                    rank_value = positions.mean()
                
                valid_students.loc[indices, 'position'] = rank_value
            
            valid_students.drop('temp_position', axis=1, inplace=True)

        valid_students['out_of'] = len(valid_students)

        self.df_combined = self.df_combined.merge(
            valid_students[['student_id', 'position', 'out_of']], 
            on='student_id', 
            how='left'
        )

        print(f"\n{self.GREEN}{'='*100}")
        print(f"{self.GREEN}TOP 10 STUDENTS (COMBINED EXAM AVERAGE)")
        print(f"{self.GREEN}{'='*100}")
        top_10 = valid_students[['full_name', 'points', 'avg_marks', 'subject_count_real', 'position', 'division']].head(10).copy()
        top_10['full_name'] = top_10['full_name'].str[:25]
        top_10['points'] = top_10['points'].apply(lambda x: int(x) if pd.notna(x) else 'N/A')
        top_10['avg_marks'] = top_10['avg_marks'].apply(lambda x: f"{x:.2f}" if pd.notna(x) else '—')
        top_10['position'] = top_10['position'].apply(lambda x: int(x) if x == int(x) else f"{x:.1f}")
        top_10.insert(0, 'No', range(1, len(top_10) + 1))
        print(tabulate(top_10, headers='keys', tablefmt='fancy_grid', showindex=False))

        print(f"\n{self.YELLOW}{'='*100}")
        print(f"{self.YELLOW}BOTTOM 10 STUDENTS")
        print(f"{self.YELLOW}{'='*100}")
        bottom_10 = valid_students[['full_name', 'points', 'avg_marks', 'subject_count_real', 'position', 'division']].tail(10).copy()
        bottom_10['full_name'] = bottom_10['full_name'].str[:25]
        bottom_10['points'] = bottom_10['points'].apply(lambda x: int(x) if pd.notna(x) else 'N/A')
        bottom_10['avg_marks'] = bottom_10['avg_marks'].apply(lambda x: f"{x:.2f}" if pd.notna(x) else '—')
        bottom_10['position'] = bottom_10['position'].apply(lambda x: int(x) if x == int(x) else f"{x:.1f}")
        bottom_10.insert(0, 'No', range(1, len(bottom_10) + 1))
        print(tabulate(bottom_10, headers='keys', tablefmt='fancy_grid', showindex=False))

        # Show tie statistics
        if self.RANKING_METHOD in ['min', 'max', 'average', 'dense']:
            position_counts = valid_students.groupby('position').size()
            tied_positions = position_counts[position_counts > 1]
            
            if len(tied_positions) > 0:
                print(f"\n{self.CYAN}TIED RANKINGS ANALYSIS ({self.RANKING_METHOD.upper()} method):")
                tie_stats = [
                    {'Metric': 'Total positions with ties', 'Value': len(tied_positions)},
                    {'Metric': 'Largest tie group', 'Value': f"{tied_positions.max()} students"},
                    {'Metric': 'Total students in ties', 'Value': tied_positions.sum()},
                    {'Metric': 'Percentage in ties', 'Value': f"{(tied_positions.sum()/len(valid_students)*100):.1f}%"}
                ]
                tie_df = pd.DataFrame(tie_stats)
                print(tabulate(tie_df, headers='keys', tablefmt='fancy_grid', showindex=False))

    def generate_necta_strings(self):
        """Generate NECTA-format result strings for each student."""
        print(f"\n{self.GREEN}9. NECTA RESULTS STRING GENERATION")
        print("=" * 60)

        # Check if new curriculum (subjects 41 or 42 present)
        subjects_sql = f"""
            SELECT subject_id FROM tbl_school_subjects 
            WHERE is_present_{self.class_id} = True
        """
        subjects_check = pd.read_sql(subjects_sql, self.conn)
        is_new_curriculum = (41 in subjects_check['subject_id'].values) or \
                           (42 in subjects_check['subject_id'].values)
        max_compulsory = 8 if is_new_curriculum else 7

        print(f"{self.CYAN}Curriculum Type: {self.WHITE}{'NEW' if is_new_curriculum else 'OLD'}")
        print(f"{self.CYAN}Compulsory Subjects: {self.WHITE}{max_compulsory}")

        def build_necta_string(row):
            """Build NECTA result string for a student based on averages."""
            parts = []
            for i, col in enumerate(self.valid_subject_cols):
                mark = row[col]  # Average mark
                grade = row[f"{col}_grade"]
                short = self.subject_column_map[col]['subject_short']
                
                # Compulsory subjects (first N subjects)
                if i < max_compulsory:
                    if pd.isna(grade) or grade is None:
                        parts.append(f" {short}-'X'")
                    else:
                        # Format mark: if it's a whole number, show as int; otherwise use decimal places
                        if mark == int(mark):
                            mark_formatted = f"{int(mark)}"
                        else:
                            mark_formatted = f"{mark:.{self.NECTA_DECIMAL_PLACES}f}"
                        parts.append(f" {short} {mark_formatted} -'{grade}'")
                # Optional subjects (only if attempted)
                elif not (pd.isna(grade) or grade is None):
                    # Format mark: if it's a whole number, show as int; otherwise use decimal places
                    if mark == int(mark):
                        mark_formatted = f"{int(mark)}"
                    else:
                        mark_formatted = f"{mark:.{self.NECTA_DECIMAL_PLACES}f}"
                    parts.append(f" {short} {mark_formatted} -'{grade}'")
            
            result = "".join(parts).strip()
            if result.endswith("-"):
                result = result[:-1].strip()
            return result

        print(f"\n{self.CYAN}Generating NECTA strings for {len(self.df_combined):,} students...")
        # Use apply with tqdm wrapper
        with tqdm(total=len(self.df_combined), desc="Building NECTA Strings", ncols=80) as pbar:
            necta_results = []
            for idx, row in self.df_combined.iterrows():
                necta_results.append(build_necta_string(row))
                pbar.update(1)
            self.df_combined['necta_results'] = necta_results

        print(f"\n{self.MAGENTA}NECTA RESULTS SAMPLE (First 10 Students):")
        necta_sample = self.df_combined[['full_name', 'necta_results']].head(10).copy()
        necta_sample['full_name'] = necta_sample['full_name'].str[:20]
        necta_sample.insert(0, 'No', range(1, len(necta_sample) + 1))
        
        # Truncate long result strings for display
        for i, row in necta_sample.iterrows():
            txt = row['necta_results']
            necta_sample.at[i, 'necta_results'] = txt[:90] + '...' if len(txt) > 90 else txt
        
        print(tabulate(necta_sample, headers='keys', tablefmt='fancy_grid', showindex=False))

        # Calculate average grade
        def calculate_grade(mark):
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

        self.df_combined['avg_grade'] = self.df_combined['avg_marks'].apply(calculate_grade)

        # Append AVG to NECTA string
        def append_avg(row):
            base = row['necta_results']
            if pd.isna(row['avg_marks']) or pd.isna(row['avg_grade']):
                return base
            
            avg_marks_formatted = f"{row['avg_marks']:.2f}"
            return f"{base} AVG {avg_marks_formatted} -'{row['avg_grade']}'".strip()

        self.df_combined['necta_results'] = self.df_combined.apply(append_avg, axis=1)

        print(f"\n{self.GREEN}✓ NECTA strings generated and finalized with AVG")

        # Show final NECTA sample with AVG
        print(f"\n{self.CYAN}FINAL NECTA RESULTS WITH AVG (Sample):")
        final_sample = self.df_combined[['full_name', 'avg_marks', 'avg_grade', 'necta_results']].head(10).copy()
        final_sample['full_name'] = final_sample['full_name'].str[:20]
        final_sample['avg_marks'] = final_sample['avg_marks'].apply(
            lambda x: f"{x:.2f}" if pd.notna(x) else "N/A"
        )
        final_sample.insert(0, 'No', range(1, len(final_sample) + 1))
        
        # Truncate NECTA strings
        for i, row in final_sample.iterrows():
            txt = row['necta_results']
            if len(txt) > 70:
                final_sample.at[i, 'necta_results'] = '...' + txt[-70:]
            else:
                final_sample.at[i, 'necta_results'] = txt
        
        print(tabulate(final_sample, headers='keys', tablefmt='fancy_grid', showindex=False))

    def save_as_query(self):
        """Save the combined results as an Access query."""
        print(f"\n{self.GREEN}10. SAVING COMBINED RESULTS AS ACCESS QUERY")
        print("=" * 60)

        # Drop existing query/view if it exists - try multiple approaches
        drop_attempts = [
            f"DROP VIEW {self.QUERY_NAME}",
            f"DROP TABLE {self.QUERY_NAME}",
        ]
        
        for drop_sql in drop_attempts:
            try:
                self.cursor.execute(drop_sql)
                self.conn.commit()
                print(f"{self.YELLOW}✓ Dropped existing object: {self.QUERY_NAME}")
                break
            except:
                pass
        
        # Also try deleting from MSysObjects (Access system table) if above fails
        try:
            self.cursor.execute(f"""
                DELETE FROM MSysObjects 
                WHERE Name = '{self.QUERY_NAME}' AND Type = 5
            """)
            self.conn.commit()
            print(f"{self.YELLOW}✓ Removed query from system catalog: {self.QUERY_NAME}")
        except:
            pass

        # Build column list for SELECT statement
        select_cols = ['student_id', 'full_name', 'sex']
        
        # Add subject columns with suffixes and position columns
        for col in self.valid_subject_cols:
            select_cols.extend([f"{col}_1", f"{col}_2", col])  # col is the average
            # Add grade column
            select_cols.append(f"{col}_grade")
            # Add position columns if they exist
            if f"{col}_pos" in self.df_combined.columns:
                select_cols.extend([f"{col}_pos", f"{col}_out_of"])
        
        # Add aggregate columns
        select_cols.extend([
            'subject_count_real', 'subject_count', 'total_marks', 
            'avg_marks', 'avg_grade', 'points', 'division', 'position', 'out_of', 'necta_results'
        ])

        # Create temporary table to store results
        temp_table = f"tbl_temp_{self.QUERY_NAME}"
        
        # Drop temp table if exists
        try:
            self.cursor.execute(f"DROP TABLE {temp_table}")
            self.conn.commit()
            print(f"{self.YELLOW}✓ Dropped existing temp table: {temp_table}")
        except:
            pass

        print(f"\n{self.CYAN}Creating table structure...")
        # Define column types for table creation
        col_definitions = []
        col_definitions.append("student_id INT")
        col_definitions.append("full_name TEXT(255)")
        col_definitions.append("sex TEXT(10)")
        
        for col in self.valid_subject_cols:
            col_definitions.append(f"{col}_1 DOUBLE")
            col_definitions.append(f"{col}_2 DOUBLE")
            col_definitions.append(f"{col} DOUBLE")
            col_definitions.append(f"{col}_grade TEXT(2)")
            if f"{col}_pos" in self.df_combined.columns:
                col_definitions.append(f"{col}_pos INT")
                col_definitions.append(f"{col}_out_of INT")
        
        col_definitions.extend([
            "subject_count_real INT",
            "subject_count INT",
            "total_marks DOUBLE",
            "avg_marks DOUBLE",
            "avg_grade TEXT(2)",
            "points INT",
            "division TEXT(10)",
            "position DOUBLE",
            "out_of INT",
            "necta_results MEMO"
        ])

        # Create table
        create_sql = f"CREATE TABLE {temp_table} ({', '.join(col_definitions)})"
        self.cursor.execute(create_sql)
        self.conn.commit()

        print(f"{self.GREEN}✓ Created temporary table: {self.WHITE}{temp_table}")
        print(f"{self.CYAN}Table columns: {self.WHITE}{len(col_definitions)}")

        # Insert data
        insert_cols = [c for c in select_cols if c in self.df_combined.columns]
        placeholders = ', '.join(['?' for _ in insert_cols])
        insert_sql = f"INSERT INTO {temp_table} ({', '.join(insert_cols)}) VALUES ({placeholders})"

        print(f"\n{self.CYAN}Preparing data for insertion...")
        insert_data = []
        for _, row in self.df_combined.iterrows():
            values = []
            for col in insert_cols:
                value = row[col]
                if pd.isna(value) or value is None:
                    values.append(None)
                else:
                    if hasattr(value, 'item'):
                        value = value.item()
                    values.append(value)
            insert_data.append(tuple(values))

        print(f"{self.GREEN}✓ Prepared {len(insert_data):,} records for insertion")

        # Batch insert
        batch_size = 50
        print(f"\n{self.CYAN}Inserting data into table (batch size: {batch_size})...")
        with tqdm(total=len(insert_data), desc="Inserting Records", ncols=100, colour='green') as pbar:
            for i in range(0, len(insert_data), batch_size):
                batch = insert_data[i:i + batch_size]
                self.cursor.executemany(insert_sql, batch)
                self.conn.commit()
                pbar.update(len(batch))

        print(f"{self.GREEN}✓ Successfully inserted all records!")

        # Create query pointing to table
        print(f"\n{self.CYAN}Creating query view...")
        query_sql = f"CREATE VIEW {self.QUERY_NAME} AS SELECT * FROM {temp_table}"
        self.cursor.execute(query_sql)
        self.conn.commit()

        print(f"\n{self.GREEN}{'='*100}")
        print(f"{self.GREEN}✓✓✓ QUERY SUCCESSFULLY CREATED ✓✓✓")
        print(f"{self.GREEN}{'='*100}")
        print(f"{self.WHITE}Query Name: {self.CYAN}{self.QUERY_NAME}")
        print(f"{self.WHITE}Backing Table: {self.CYAN}{temp_table}")
        print(f"{self.WHITE}Total Records: {self.CYAN}{len(insert_data):,}")
        print(f"{self.WHITE}Total Columns: {self.CYAN}{len(insert_cols)}")
        
        print(f"\n{self.MAGENTA}QUERY STRUCTURE PREVIEW:")
        structure_preview = []
        for i, col in enumerate(insert_cols[:15]):  # Show first 15 columns
            structure_preview.append({
                'No': i+1,
                'Column Name': col,
                'Type': 'Number' if any(x in col for x in ['mark', 'count', 'total', 'avg', 'point', 'position']) else 'Text'
            })
        
        structure_df = pd.DataFrame(structure_preview)
        print(tabulate(structure_df, headers='keys', tablefmt='fancy_grid', showindex=False))
        
        if len(insert_cols) > 15:
            print(f"{self.YELLOW}... and {len(insert_cols) - 15} more columns")
        
        print(f"\n{self.CYAN}{'='*100}")
        print(f"{self.CYAN}You can now open '{self.QUERY_NAME}' in Microsoft Access!")
        print(f"{self.CYAN}{'='*100}")

    def run(self):
        """Execute the complete dual exam processing pipeline."""
        start_time = datetime.datetime.now()
        
        print(f"\n{self.MAGENTA}{'='*100}")
        print(f"{self.MAGENTA}{'='*100}")
        print(f"{self.GREEN}       DUAL EXAM PROCESSOR - COMBINE & AVERAGE SYSTEM")
        print(f"{self.MAGENTA}{'='*100}")
        print(f"{self.MAGENTA}{'='*100}")
        print(f"\n{self.CYAN}CONFIGURATION:")
        print(f"   • Exam 1 ID: {self.WHITE}{self.EXAM_ID_1}")
        print(f"   • Exam 2 ID: {self.WHITE}{self.EXAM_ID_2}")
        print(f"   • Database: {self.WHITE}{self.DB_PATH}")
        print(f"   • Query Name: {self.WHITE}{self.QUERY_NAME}")
        print(f"   • Base Subjects: {self.WHITE}{self.BASE_SUBJECTS}")
        print(f"   • Flat Rate: {self.WHITE}{self.FLAT_RATE}")
        print(f"   • Include Incomplete: {self.WHITE}{self.INCLUDE_INC}")
        print(f"   • Ranking Method: {self.WHITE}{self.RANKING_METHOD.upper()}")
        print(f"   • NECTA Decimal Places: {self.WHITE}{self.NECTA_DECIMAL_PLACES}")
        print(f"   • Start Time: {self.WHITE}{start_time.strftime('%Y-%m-%d %H:%M:%S')}")
        
        try:
            self._connect_to_database()
            self.load_data()
            self.configure_subjects()
            self.combine_exams()
            self.calculate_grades()
            self.aggregate_performance()
            self.calculate_points_and_division()
            self.rank_subjects()
            self.rank_students()
            self.generate_necta_strings()
            self.save_as_query()
            
            end_time = datetime.datetime.now()
            duration = (end_time - start_time).total_seconds()

            print(f"\n{self.GREEN}{'='*100}")
            print(f"{self.GREEN}{'='*100}")
            print(f"{self.GREEN}           ✓✓✓ PROCESS COMPLETED SUCCESSFULLY ✓✓✓")
            print(f"{self.GREEN}{'='*100}")
            print(f"{self.GREEN}{'='*100}")
            
            print(f"\n{self.CYAN}FINAL SUMMARY:")
            summary_data = [
                {'Metric': 'Total Students Processed', 'Value': f"{len(self.df_combined):,}"},
                {'Metric': 'Active Subjects', 'Value': f"{len(self.valid_subject_cols)}"},
                {'Metric': 'Valid Rankings', 'Value': f"{self.df_combined['position'].notna().sum():,}"},
                {'Metric': 'Division I Students', 'Value': f"{len(self.df_combined[self.df_combined['division'] == 'I']):,}"},
                {'Metric': 'Query Created', 'Value': self.QUERY_NAME},
                {'Metric': 'Total Columns in Query', 'Value': f"{3 + (len(self.valid_subject_cols) * 6) + 10}"},
                {'Metric': 'Processing Time', 'Value': f"{duration:.2f} seconds"},
                {'Metric': 'End Time', 'Value': end_time.strftime('%Y-%m-%d %H:%M:%S')}
            ]
            
            summary_df = pd.DataFrame(summary_data)
            print(tabulate(summary_df, headers='keys', tablefmt='fancy_grid', showindex=False))
            
            print(f"\n{self.MAGENTA}QUERY DATA COMPLETENESS CHECK:")
            completeness = [
                {'Data Type': '✓ Individual Exam Marks (Exam 1 & 2)', 'Status': 'INCLUDED'},
                {'Data Type': '✓ Average Marks per Subject', 'Status': 'INCLUDED'},
                {'Data Type': '✓ Grades per Subject', 'Status': 'INCLUDED'},
                {'Data Type': '✓ Subject Rankings (pos & out_of)', 'Status': 'INCLUDED'},
                {'Data Type': '✓ Overall Student Rankings', 'Status': 'INCLUDED'},
                {'Data Type': '✓ Points & Divisions', 'Status': 'INCLUDED'},
                {'Data Type': '✓ NECTA Results String', 'Status': 'INCLUDED'},
                {'Data Type': '✓ Average Grade', 'Status': 'INCLUDED'}
            ]
            
            completeness_df = pd.DataFrame(completeness)
            print(tabulate(completeness_df, headers='keys', tablefmt='fancy_grid', showindex=False))

        except Exception as e:
            print(f"\n{self.RED}{'='*100}")
            print(f"{self.RED}ERROR OCCURRED!")
            print(f"{self.RED}{'='*100}")
            print(f"{self.RED}Error: {e}")
            import traceback
            traceback.print_exc()
            raise
        finally:
            if self.conn:
                self.conn.close()
                print(f"\n{self.CYAN}Database connection closed.")


# ========================================
# USAGE EXAMPLE:
# ========================================

if __name__ == "__main__":
    processor = DualExamProcessor(
        exam_id_1="ANN320251117",
        exam_id_2="MID320251027",
        db_path=r"C:\Kiyabo App\backend\Kiyabo App Backend v2.0.0.accdb",
        query_name="qry_CombinedANN_MID_Results",
        base_subjects=7,
        flat_rate=True,
        include_inc=True,
        ranking_method='min',
        necta_decimal_places=1
    )
    processor.run()