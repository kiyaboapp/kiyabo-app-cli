# ========================================
# DUAL TABLE PROCESSOR — CUSTOMIZED FOR DUAL RESULTS
# ========================================
import pyodbc
import pandas as pd
import numpy as np
from tqdm import tqdm
import warnings
from colorama import init, Fore, Back, Style
from tabulate import tabulate
import datetime

init(autoreset=True)
warnings.filterwarnings("ignore")


class DualTableProcessor:
    """
    Processes tbl_student_exam_results_dual with exam results.
    No competency analysis, works with combo_id instead of exam_id.
    """

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
        combo_id: str,
        db_path: str,
        base_subjects: int = 7,
        flat_rate: bool = True,
        include_inc: bool = True,
        sort_columns: list = None,
        ranking_method: str = "min"
    ):
        """
        Initialize Dual Table Processor.
        
        Args:
            combo_id: Combo identifier (exam_id1_exam_id2)
            db_path: Path to Access database
            base_subjects: Number of subjects for base calculation (default: 7)
            flat_rate: Use flat rate calculation (default: True)
            include_inc: Include incomplete students (default: True)
            sort_columns: List of columns to sort by
            ranking_method: Ranking method - 'min', 'max', 'average', 'dense', 'first'
        """
        self.COMBO_ID = combo_id
        self.DB_PATH = db_path
        self.BASE_SUBJECTS = base_subjects
        self.FLAT_RATE = flat_rate
        self.INCLUDE_INC = include_inc
        self.RANKING_METHOD = ranking_method.lower()

        valid_methods = ['min', 'max', 'average', 'dense', 'first']
        if self.RANKING_METHOD not in valid_methods:
            raise ValueError(f"Invalid ranking_method. Must be one of: {valid_methods}")

        if sort_columns is None:
            sort_columns = ['ranking_points', 'avg_marks', 'subject_count_real']
        self.sort_columns = sort_columns
        self.ascending = self._build_sort_directions(sort_columns)

        self.conn = None
        self.cursor = None
        self.df = None
        self.valid_subject_cols = []
        self.subject_column_map = {}
        self.mark_columns = ['civ','his','geo','kis','eng','phy','che','bio','mat','edk','ics'] + \
                           [f'sub{i}' for i in range(12,21)]

    def _build_sort_directions(self, columns: list) -> list:
        """Build ascending/descending list based on column semantics."""
        ascending = []
        for col in columns:
            if col in ['ranking_points', 'points']:
                ascending.append(True)
            elif col in ['avg_marks', 'total_marks', 'subject_count_real', 'subject_count']:
                ascending.append(False)
            else:
                print(f"{self.YELLOW}⚠️  Warning: Unknown column '{col}' - defaulting to DESC")
                ascending.append(False)
        return ascending

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
        """Load exam results from dual table + STRICTLY keep only students present in exam2."""
        print(f"\n{self.GREEN}1. LOADING DUAL TABLE DATA")
        print("=" * 60)

        results_sql = """
            SELECT result_id, student_id, exam_id, combo_id,
                first_name, middle_name, surname, sex,
                civ, his, geo, kis, eng, phy, che, bio, mat, edk, ics,
                sub12, sub13, sub14, sub15, sub16, sub17, sub18, sub19, sub20
            FROM tbl_student_exam_results_dual 
            WHERE combo_id = ?
        """
        self.df = pd.read_sql(results_sql, self.conn, params=[self.COMBO_ID])

        # Create full_name
        self.df['full_name'] = (
            self.df['first_name'].fillna('') + ' ' + 
            self.df['middle_name'].fillna('') + ' ' + 
            self.df['surname'].fillna('')
        ).str.strip()

        total_before = len(self.df)

        # ==================== CRITICAL FILTER: ONLY KEEP STUDENTS IN EXAM2 ====================
        # Assuming your combo_id format is: exam1_id_exam2_id  → e.g., "MID120250818_ANN120251119"
        exam1_id, exam2_id = self.COMBO_ID.split("_")

        # Keep ONLY rows where exam_id matches the SECOND (exam2) part
        self.df = self.df[self.df['exam_id'] == exam2_id].copy()

        total_after = len(self.df)

        print(f"{self.YELLOW}Combo ID: {self.WHITE}{self.COMBO_ID}")
        print(f"{self.YELLOW}Total Records (Before Filter): {self.WHITE}{total_before:,}")
        print(f"{self.RED}→ Filtered to ONLY exam2 sitters: {self.WHITE}{total_after:,} records")
        print(f"{self.RED}→ Excluded {total_before - total_after:,} students missing in exam2")

        if total_after == 0:
            raise ValueError(f"No students found in exam2 ({exam2_id}). Check combo_id or data.")
        # -------------------------------
    # 2. SUBJECT MAPPING
    # -------------------------------
    def configure_subjects(self):
        """Configure subject mapping and identify active subject columns."""
        print(f"\n{self.GREEN}2. SUBJECT CONFIGURATION")
        print("=" * 60)

        # Identify valid subject columns with data
        self.valid_subject_cols = []
        for col in self.mark_columns:
            if col in self.df.columns:
                self.df[col] = pd.to_numeric(self.df[col], errors='coerce')
                if self.df[col].notna().any():
                    self.valid_subject_cols.append(col)

        print(f"{self.GREEN}ACTIVE SUBJECT COLUMNS: {len(self.valid_subject_cols)}")
        print(f"{self.CYAN}Columns: {', '.join(self.valid_subject_cols[:10])}")

        # Keep only relevant columns
        keep_cols = ['result_id', 'student_id', 'exam_id', 'combo_id', 
                    'full_name', 'sex'] + self.valid_subject_cols
        self.df = self.df[keep_cols].copy()

    # -------------------------------
    # 3. GRADE CALCULATION
    # -------------------------------
    def calculate_grades(self):
        """Calculate letter grades for all subjects."""
        print(f"\n{self.GREEN}3. GRADE CALCULATION PROCESS")
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
            self.df[f"{col}_grade"] = self.df[col].apply(calculate_grade)

        print(f"{self.CYAN}Grades calculated for {len(self.valid_subject_cols)} subjects")

    # -------------------------------
    # 4. SUBJECT COUNT & TOTAL MARKS
    # -------------------------------
    def aggregate_performance(self):
        """Calculate subject counts, total marks, and average marks."""
        print(f"\n{self.GREEN}4. ACADEMIC AGGREGATION")
        print("=" * 60)

        self.df['subject_count_real'] = self.df[self.valid_subject_cols].notna().sum(axis=1)

        if self.FLAT_RATE:
            self.df['subject_count'] = self.BASE_SUBJECTS
            self.df['total_marks'] = self.df[self.valid_subject_cols].apply(
                lambda row: sum(sorted([m for m in row if not pd.isna(m)], reverse=True)[:self.BASE_SUBJECTS]), 
                axis=1
            )
        else:
            self.df['subject_count'] = self.df['subject_count_real'].apply(
                lambda x: self.BASE_SUBJECTS if x <= self.BASE_SUBJECTS else x
            )
            self.df['total_marks'] = self.df[self.valid_subject_cols].sum(axis=1)

        self.df['avg_marks'] = np.where(
            self.df['subject_count'] > 0, 
            self.df['total_marks'] / self.df['subject_count'], 
            np.nan
        )

        print(f"{self.CYAN}Aggregation completed")

    # -------------------------------
    # 5. POINTS & DIVISION
    # -------------------------------
    def calculate_points_and_division(self):
        """Calculate ranking points and division for each student."""
        print(f"\n{self.GREEN}5. POINTS & DIVISION CALCULATION")
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

        points_division = self.df.apply(calculate_points_and_division, axis=1, result_type='expand')
        self.df['points'] = points_division[0]
        self.df['division'] = points_division[1]

        print(f"{self.MAGENTA}DIVISION DISTRIBUTION:")
        div_counts = self.df['division'].value_counts().reset_index()
        div_counts.columns = ['Division', 'Students']
        print(tabulate(div_counts, headers='keys', tablefmt='fancy_grid', showindex=False))

    # -------------------------------
    # 6. FLEXIBLE RANKING SYSTEM
    # -------------------------------
    def rank_students(self):
        """Flexible ranking system."""
        print(f"\n{self.GREEN}6. ACADEMIC RANKING ({self.RANKING_METHOD.upper()} METHOD)")
        print("=" * 60)

        abs_students = self.df[self.df['division'] == 'ABS'].copy()
        valid_students = self.df[self.df['division'] != 'ABS'].copy()

        print(f"{self.YELLOW}RANKING BREAKDOWN:")
        print(f"   • ABS Students: {self.WHITE}{len(abs_students):,}")
        print(f"   • Valid Students: {self.WHITE}{len(valid_students):,}")

        if len(valid_students) == 0:
            print(f"{self.RED}No valid students to rank!")
            return

        if 'ranking_points' in self.sort_columns:
            max_points = valid_students['points'].max() if not valid_students['points'].isna().all() else 35
            valid_students['ranking_points'] = valid_students['points'].fillna(max_points + 1)

        valid_students = valid_students.sort_values(
            self.sort_columns,
            ascending=self.ascending,
            na_position='last'
        )

        if self.RANKING_METHOD == 'first':
            valid_students['position'] = range(1, len(valid_students) + 1)
        elif self.RANKING_METHOD == 'dense':
            valid_students['position'] = (
                valid_students.groupby(self.sort_columns, dropna=False).ngroup() + 1
            )
        elif self.RANKING_METHOD in ['min', 'max', 'average']:
            valid_students['temp_position'] = range(1, len(valid_students) + 1)
            for cols_tuple, group in valid_students.groupby(self.sort_columns, dropna=False):
                indices = group.index
                positions = group['temp_position'].values
                
                if self.RANKING_METHOD == 'min':
                    rank_value = positions.min()
                elif self.RANKING_METHOD == 'max':
                    rank_value = positions.max()
                else:
                    rank_value = positions.mean()
                
                valid_students.loc[indices, 'position'] = rank_value
            
            valid_students.drop('temp_position', axis=1, inplace=True)
            
            if self.RANKING_METHOD == 'average':
                valid_students['position'] = valid_students['position'].round(1)
            else:
                valid_students['position'] = valid_students['position'].astype(int)
        
        valid_students['out_of'] = len(valid_students)

        self.df = self.df.merge(
            valid_students[['result_id', 'position', 'out_of']], 
            on='result_id', 
            how='left'
        )

        print(f"{self.GREEN}TOP 10 RANKED STUDENTS:")
        top_10 = valid_students[['full_name', 'points', 'avg_marks', 'position', 'division']].head(10).copy()
        top_10.insert(0, 'No', range(1, len(top_10) + 1))
        print(tabulate(top_10, headers='keys', tablefmt='fancy_grid', showindex=False))

    # -------------------------------
    # 7. SUBJECT-WISE RANKING
    # -------------------------------
    def rank_subjects(self):
        """Rank students within each subject based on marks."""
        print(f"\n{self.GREEN}7. SUBJECT-WISE RANKING")
        print("=" * 60)

        for col in self.valid_subject_cols:
            subject_rank_df = self.df[self.df[col].notna()].copy()
            
            if len(subject_rank_df) > 0:
                subject_rank_df = subject_rank_df.sort_values(col, ascending=False)
                subject_rank_df[f'{col}_pos'] = subject_rank_df[col].rank(
                    method='min', ascending=False
                ).astype(int)
                subject_rank_df[f'{col}_out_of'] = len(subject_rank_df)
                
                self.df = self.df.merge(
                    subject_rank_df[['result_id', f'{col}_pos', f'{col}_out_of']],
                    on='result_id', how='left'
                )

        print(f"{self.CYAN}Subject rankings calculated")

    # -------------------------------
    # 8. NECTA RESULTS STRING
    # -------------------------------
    def generate_necta_strings(self):
        """Generate NECTA-format result strings."""
        print(f"\n{self.GREEN}8. NECTA RESULTS STRING GENERATION")
        print("=" * 60)

        def build_necta_string(row):
            parts = []
            for col in self.valid_subject_cols[:11]:  # First 11 subjects
                mark = row[col]
                grade = row[f"{col}_grade"]
                short = col.upper()
                
                if pd.isna(grade) or grade is None:
                    parts.append(f" {short}-'X'")
                else:
                    parts.append(f" {short} {int(mark)} -'{grade}'")
            
            return "".join(parts).strip()

        tqdm.pandas(desc="Generating NECTA strings")
        self.df['necta_results'] = self.df.progress_apply(build_necta_string, axis=1)

        print(f"{self.CYAN}NECTA strings generated")

    # -------------------------------
    # 9. UPDATE DATABASE
    # -------------------------------
    def update_database(self):
        """Update database with calculated results."""
        print(f"\n{self.GREEN}9. DATABASE UPDATE PROCESS")
        print("=" * 60)

        update_columns = [
            'necta_results', 'subject_count', 'total_marks', 
            'points', 'division', 'position', 'out_of'
        ]
        
        for col in self.valid_subject_cols:
            update_columns.append(f"{col}_grade")
            if f"{col}_pos" in self.df.columns:
                update_columns.extend([f"{col}_pos", f"{col}_out_of"])

        print(f"{self.YELLOW}UPDATING {len(update_columns)} COLUMNS")

        set_clause = ", ".join([f"{col} = ?" for col in update_columns])
        update_sql = f"UPDATE tbl_student_exam_results_dual SET {set_clause} WHERE result_id = ?"

        update_data = []
        for _, row in self.df.iterrows():
            values = []
            for col in update_columns:
                value = row[col]
                if pd.isna(value) or value is None:
                    values.append(None)
                else:
                    if hasattr(value, 'item'):
                        value = value.item()
                    values.append(value)
            values.append(int(row['result_id']))
            update_data.append(tuple(values))

        batch_size = 20
        success_count = 0
        
        try:
            self.cursor.connection.autocommit = False
            
            with tqdm(total=len(update_data), desc="Updating Database") as pbar:
                for i in range(0, len(update_data), batch_size):
                    batch = update_data[i:i + batch_size]
                    self.cursor.executemany(update_sql, batch)
                    success_count += len(batch)
                    self.conn.commit()
                    pbar.update(len(batch))
            
            self.cursor.connection.autocommit = True
            
        except Exception as e:
            print(f"{self.RED}Error during batch update: {e}")
            self.conn.rollback()
            self.cursor.connection.autocommit = True
            raise

        print(f"{self.GREEN}DATABASE UPDATE COMPLETE: {success_count:,}/{len(self.df):,} records")
        return success_count

    # -------------------------------
    # 9.5 FINALIZE NECTA WITH AVG
    # -------------------------------
    def finalize_necta_with_avg(self):
        """Append average marks to NECTA results."""
        print(f"\n{self.GREEN}9.5 FINALIZE NECTA WITH AVG")
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

        self.df['avg_grade'] = self.df['avg_marks'].apply(calculate_grade)

        def append_avg(row):
            base = row['necta_results']
            if pd.isna(row['avg_marks']) or pd.isna(row['avg_grade']):
                return base
            
            avg_marks_formatted = f"{row['avg_marks']:.2f}"
            return f"{base} AVG {avg_marks_formatted} -'{row['avg_grade']}'".strip()

        self.df['necta_results'] = self.df.apply(append_avg, axis=1)

        final_sql = """
            UPDATE tbl_student_exam_results_dual
            SET necta_results = ?, avg_grade = ?
            WHERE result_id = ?
        """
        
        update_data = [
            (str(row['necta_results']), 
             str(row['avg_grade']) if not pd.isna(row['avg_grade']) else None, 
             int(row['result_id']))
            for _, row in self.df.iterrows()
        ]
        
        batch_size = 20
        with tqdm(total=len(update_data), desc="Final NECTA Update") as pbar:
            for i in range(0, len(update_data), batch_size):
                batch = update_data[i:i + batch_size]
                self.cursor.executemany(final_sql, batch)
                pbar.update(len(batch))
        
        self.conn.commit()
        print(f"{self.CYAN}NECTA finalization complete")

    # -------------------------------
    # 10. SUMMARY REPORT
    # -------------------------------
    def generate_summary_report(self, success_count):
        """Generate comprehensive summary."""
        print(f"\n{self.GREEN}10. COMPREHENSIVE SUMMARY REPORT")
        print("=" * 60)

        summary_data = {
            'Category': ['TOTAL RECORDS', 'ACTIVE SUBJECTS', 'AVG MARKS', 'DIVISION I', 'DATABASE'],
            'Value': [
                f"{len(self.df):,}",
                f"{len(self.valid_subject_cols)}",
                f"{self.df['total_marks'].mean():.1f}",
                f"{len(self.df[self.df['division'] == 'I']):,}",
                f"{success_count:,}/{len(self.df):,}"
            ]
        }

        summary_df = pd.DataFrame(summary_data)
        print(tabulate(summary_df, headers='keys', tablefmt='fancy_grid', showindex=False))

        now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(f"\n{self.CYAN}COMPLETED AT: {self.WHITE}{now}")

    # -------------------------------
    # MAIN EXECUTION
    # -------------------------------
    def run(self):
        """Execute the complete processing pipeline."""
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
# USAGE:
# ========================================
if __name__ == "__main__":
    processor = DualTableProcessor(
        combo_id="MID120250818_ANN120251119",
        db_path=r"C:\Kiyabo App\backend\famgi\Kiyabo App Backend Famgi.accdb",
        base_subjects=7,
        flat_rate=True,
        include_inc=True,
        sort_columns=['avg_marks', 'ranking_points', 'subject_count_real'],
        ranking_method='min'
    )
    processor.run()