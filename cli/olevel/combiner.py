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
        ranking_method: str = "min"
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
        """
        self.EXAM_ID_1 = exam_id_1
        self.EXAM_ID_2 = exam_id_2
        self.DB_PATH = db_path
        self.QUERY_NAME = query_name
        self.BASE_SUBJECTS = base_subjects
        self.FLAT_RATE = flat_rate
        self.INCLUDE_INC = include_inc
        self.RANKING_METHOD = ranking_method.lower()

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

        # Calculate averages for each subject
        for col in self.valid_subject_cols:
            col_1 = f"{col}_1"
            col_2 = f"{col}_2"
            
            # Average column
            df_merged[col] = df_merged[[col_1, col_2]].mean(axis=1)

        self.df_combined = df_merged

        print(f"{self.CYAN}COMBINED DATA SAMPLE (First 5 students, First 3 subjects):")
        sample_data = []
        for i in range(min(5, len(self.df_combined))):
            row = self.df_combined.iloc[i]
            sample_row = {'No': i+1, 'Student': row['full_name']}
            
            for j, col in enumerate(self.valid_subject_cols[:3]):
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

        for col in self.valid_subject_cols:
            self.df_combined[f"{col}_grade"] = self.df_combined[col].apply(calculate_grade)

    def aggregate_performance(self):
        """Calculate subject counts, total marks, and average marks."""
        print(f"\n{self.GREEN}5. ACADEMIC AGGREGATION")
        print("=" * 60)

        # Count actual subjects attempted
        self.df_combined['subject_count_real'] = self.df_combined[self.valid_subject_cols].notna().sum(axis=1)

        if self.FLAT_RATE:
            self.df_combined['subject_count'] = self.BASE_SUBJECTS
            self.df_combined['total_marks'] = self.df_combined[self.valid_subject_cols].apply(
                lambda row: sum(sorted([m for m in row if not pd.isna(m)], reverse=True)[:self.BASE_SUBJECTS]), 
                axis=1
            )
        else:
            self.df_combined['subject_count'] = self.df_combined['subject_count_real'].apply(
                lambda x: self.BASE_SUBJECTS if x <= self.BASE_SUBJECTS else x
            )
            self.df_combined['total_marks'] = self.df_combined[self.valid_subject_cols].sum(axis=1)

        self.df_combined['avg_marks'] = np.where(
            self.df_combined['subject_count'] > 0, 
            self.df_combined['total_marks'] / self.df_combined['subject_count'], 
            np.nan
        )

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

        points_division = self.df_combined.apply(calculate_points_and_division, axis=1, result_type='expand')
        self.df_combined['points'] = points_division[0]
        self.df_combined['division'] = points_division[1]

        print(f"{self.MAGENTA}DIVISION DISTRIBUTION:")
        div_counts = self.df_combined['division'].value_counts().reset_index()
        div_counts.columns = ['Division', 'Students']
        div_counts['Percentage'] = (div_counts['Students'] / len(self.df_combined) * 100).round(1)
        print(tabulate(div_counts, headers='keys', tablefmt='fancy_grid', showindex=False))

    def rank_students(self):
        """Rank students based on average performance."""
        print(f"\n{self.GREEN}7. STUDENT RANKING")
        print("=" * 60)

        abs_students = self.df_combined[self.df_combined['division'] == 'ABS'].copy()
        valid_students = self.df_combined[self.df_combined['division'] != 'ABS'].copy()

        if len(valid_students) == 0:
            print(f"{self.RED}No valid students to rank!")
            return

        # Handle ranking_points
        max_points = valid_students['points'].max() if not valid_students['points'].isna().all() else 35
        valid_students['ranking_points'] = valid_students['points'].fillna(max_points + 1)

        # Sort and rank
        valid_students = valid_students.sort_values(
            self.sort_columns,
            ascending=self.ascending,
            na_position='last'
        )

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

        print(f"{self.GREEN}TOP 10 STUDENTS:")
        top_10 = valid_students[['full_name', 'points', 'avg_marks', 'subject_count_real', 'position', 'division']].head(10).copy()
        top_10.insert(0, 'No', range(1, len(top_10) + 1))
        print(tabulate(top_10, headers='keys', tablefmt='fancy_grid', showindex=False))

    def save_as_query(self):
        """Save the combined results as an Access query."""
        print(f"\n{self.GREEN}8. SAVING AS ACCESS QUERY")
        print("=" * 60)

        # Drop existing query if it exists
        try:
            self.cursor.execute(f"DROP VIEW {self.QUERY_NAME}")
            self.conn.commit()
            print(f"{self.YELLOW}Dropped existing query: {self.QUERY_NAME}")
        except:
            pass

        # Build column list for SELECT statement
        select_cols = ['student_id', 'full_name', 'sex']
        
        # Add subject columns with suffixes
        for col in self.valid_subject_cols:
            select_cols.extend([f"{col}_1", f"{col}_2", col])  # col is the average
        
        # Add aggregate columns
        select_cols.extend([
            'subject_count_real', 'subject_count', 'total_marks', 
            'avg_marks', 'points', 'division', 'position', 'out_of'
        ])

        # Create temporary table to store results
        temp_table = f"tbl_temp_{self.QUERY_NAME}"
        
        # Drop temp table if exists
        try:
            self.cursor.execute(f"DROP TABLE {temp_table}")
            self.conn.commit()
        except:
            pass

        # Define column types for table creation
        col_definitions = []
        col_definitions.append("student_id INT")
        col_definitions.append("full_name TEXT(255)")
        col_definitions.append("sex TEXT(10)")
        
        for col in self.valid_subject_cols:
            col_definitions.append(f"{col}_1 DOUBLE")
            col_definitions.append(f"{col}_2 DOUBLE")
            col_definitions.append(f"{col} DOUBLE")
        
        col_definitions.extend([
            "subject_count_real INT",
            "subject_count INT",
            "total_marks DOUBLE",
            "avg_marks DOUBLE",
            "points INT",
            "division TEXT(10)",
            "position DOUBLE",
            "out_of INT"
        ])

        # Create table
        create_sql = f"CREATE TABLE {temp_table} ({', '.join(col_definitions)})"
        self.cursor.execute(create_sql)
        self.conn.commit()

        print(f"{self.CYAN}Created temporary table: {temp_table}")

        # Insert data
        insert_cols = [c for c in select_cols if c in self.df_combined.columns]
        placeholders = ', '.join(['?' for _ in insert_cols])
        insert_sql = f"INSERT INTO {temp_table} ({', '.join(insert_cols)}) VALUES ({placeholders})"

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

        # Batch insert
        batch_size = 50
        with tqdm(total=len(insert_data), desc="Inserting Data") as pbar:
            for i in range(0, len(insert_data), batch_size):
                batch = insert_data[i:i + batch_size]
                self.cursor.executemany(insert_sql, batch)
                self.conn.commit()
                pbar.update(len(batch))

        # Create query pointing to table
        query_sql = f"CREATE VIEW {self.QUERY_NAME} AS SELECT * FROM {temp_table}"
        self.cursor.execute(query_sql)
        self.conn.commit()

        print(f"{self.GREEN}✓ Query saved: {self.WHITE}{self.QUERY_NAME}")
        print(f"{self.CYAN}You can now open this query in Access to view the combined results!")

    def run(self):
        """Execute the complete dual exam processing pipeline."""
        try:
            self._connect_to_database()
            self.load_data()
            self.configure_subjects()
            self.combine_exams()
            self.calculate_grades()
            self.aggregate_performance()
            self.calculate_points_and_division()
            self.rank_students()
            self.save_as_query()

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
        ranking_method='min'
    )
    processor.run()