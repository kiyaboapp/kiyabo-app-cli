# ========================================
# OLEVEL PROCESSOR — FIXED RANKING
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
    O-Level Results Processing Engine — Fixed Ranking System
    Processes student exam results from Access DB with colorful, tabular output.
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
        update_competency: bool = True
    ):
        self.EXAM_ID = exam_id
        self.DB_PATH = db_path
        self.BASE_SUBJECTS = base_subjects
        self.FLAT_RATE = flat_rate
        self.INCLUDE_INC = include_inc
        self.UPDATE_COMPETENCY = update_competency

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

    # -------------------------------
    # CONNECT
    # -------------------------------
    def _connect_to_database(self):
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
        print(f"\n{self.GREEN}2. SUBJECT CONFIGURATION")
        print("=" * 60)

        subjects_sql = f"""
            SELECT subject_id, subject_name, subject_code, subject_short
            FROM tbl_school_subjects 
            WHERE is_present_{self.class_id} = True
        """
        self.subjects_df = pd.read_sql(subjects_sql, self.conn)

        def sort_subjects_vba_logic(df_sub):
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

        self.subject_column_map = {}
        for i, subject_row in self.subjects_df.iterrows():
            if i < len(self.mark_columns):
                col_name = self.mark_columns[i]
                self.subject_column_map[col_name] = {
                    'subject_short': subject_row['subject_short'],
                    'subject_id': int(subject_row['subject_id']),
                    'subject_name': subject_row['subject_name']
                }

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

        keep_cols = ['result_id', 'student_id', 'full_name', 'sex'] + self.valid_subject_cols
        self.df = self.df[keep_cols].copy()

    # -------------------------------
    # 3. GRADE CALCULATION
    # -------------------------------
    def calculate_grades(self):
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

        self.df['avg_marks'] = np.where(self.df['subject_count'] > 0, self.df['total_marks'] / self.df['subject_count'], np.nan)

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
    # 6. RANKING - FIXED MIN METHOD
    # -------------------------------
    def rank_students(self):
        print(f"\n{self.GREEN}6. ACADEMIC RANKING (FIXED MIN METHOD)")
        print("=" * 60)

        abs_students = self.df[self.df['division'] == 'ABS'].copy()
        valid_students = self.df[self.df['division'] != 'ABS'].copy()

        print(f"{self.YELLOW}RANKING BREAKDOWN:")
        print(f"   • ABS Students: {self.WHITE}{len(abs_students):,}")
        print(f"   • Valid Students: {self.WHITE}{len(valid_students):,}")

        max_points = valid_students['points'].max() if not valid_students['points'].isna().all() else 35
        valid_students['ranking_points'] = valid_students['points'].fillna(max_points + 1)

        if len(valid_students) > 0:
            valid_students = valid_students.sort_values(
                ['ranking_points', 'avg_marks', 'subject_count_real'], 
                ascending=[True, False, False]
            )
            valid_students['position'] = valid_students.groupby(
                ['ranking_points', 'avg_marks', 'subject_count_real']
            ).ngroup() + 1
            valid_students['out_of'] = len(valid_students)

        self.df = self.df.merge(valid_students[['result_id', 'position', 'out_of']], on='result_id', how='left')

        print(f"\n{self.GREEN}TOP 10 RANKED STUDENTS:")
        top_10 = valid_students[['full_name', 'points', 'avg_marks', 'position', 'division']].head(10).copy()
        top_10['avg_marks'] = top_10['avg_marks'].round(2)
        top_10['No'] = range(1, 11)
        top_10 = top_10[['No', 'full_name', 'points', 'avg_marks', 'position', 'division']]
        print(tabulate(top_10, headers='keys', tablefmt='fancy_grid', showindex=False))

        print(f"\n{self.RED}BOTTOM 10 RANKED STUDENTS:")
        bottom_10 = valid_students[['full_name', 'points', 'avg_marks', 'position', 'division']].tail(10).copy()
        bottom_10['avg_marks'] = bottom_10['avg_marks'].round(2)
        bottom_10['No'] = range(1, 11)
        bottom_10 = bottom_10[['No', 'full_name', 'points', 'avg_marks', 'position', 'division']]
        print(tabulate(bottom_10, headers='keys', tablefmt='fancy_grid', showindex=False))

        null_points_students = valid_students[valid_students['points'].isna()]
        if len(null_points_students) > 0:
            print(f"\n{self.YELLOW}STUDENTS WITHOUT POINTS (INC/0/IV) - RANKED AT BOTTOM:")
            inc_students = null_points_students[['full_name', 'points', 'avg_marks', 'position', 'division']].copy()
            inc_students['avg_marks'] = inc_students['avg_marks'].round(2)
            inc_students['No'] = range(1, len(inc_students) + 1)
            inc_students = inc_students[['No', 'full_name', 'points', 'avg_marks', 'position', 'division']]
            print(tabulate(inc_students, headers='keys', tablefmt='fancy_grid', showindex=False))

    # -------------------------------
    # 7. SUBJECT-WISE RANKING
    # -------------------------------
    def rank_subjects(self):
        print(f"\n{self.GREEN}7. SUBJECT-WISE RANKING (MARKS DESC, NOT NULL)")
        print("=" * 60)

        print(f"{self.CYAN}CALCULATING SUBJECT POSITIONS...")
        for col in self.valid_subject_cols:
            subject_rank_df = self.df[self.df[col].notna()].copy()
            if len(subject_rank_df) > 0:
                subject_rank_df = subject_rank_df.sort_values(col, ascending=False)
                subject_rank_df[f'{col}_pos'] = subject_rank_df.groupby(col).ngroup() + 1
                subject_rank_df[f'{col}_out_of'] = len(subject_rank_df)
                self.df = self.df.merge(
                    subject_rank_df[['result_id', f'{col}_pos', f'{col}_out_of']],
                    on='result_id', how='left'
                )

        print(f"{self.CYAN}SUBJECT RANKING SAMPLE (First Subject - Top 10):")
        if len(self.valid_subject_cols) > 0:
            first_subject = self.valid_subject_cols[0]
            subject_name = self.subject_column_map[first_subject]['subject_name']
            print(f"{self.YELLOW}Subject: {self.WHITE}{subject_name}")
            subject_top_10 = self.df[self.df[first_subject].notna()].nlargest(10, first_subject)[
                ['full_name', first_subject, f'{first_subject}_pos', f'{first_subject}_out_of']
            ]
            subject_top_10['No'] = range(1, len(subject_top_10) + 1)
            subject_top_10 = subject_top_10.rename(columns={
                first_subject: 'Marks',
                f'{first_subject}_pos': 'Position',
                f'{first_subject}_out_of': 'Out Of'
            })
            subject_top_10 = subject_top_10[['No', 'full_name', 'Marks', 'Position', 'Out Of']]
            print(tabulate(subject_top_10, headers='keys', tablefmt='fancy_grid', showindex=False))

    # -------------------------------
    # 8. NECTA RESULTS STRING
    # -------------------------------
    def generate_necta_strings(self):
        print(f"\n{self.GREEN}8. NECTA RESULTS STRING GENERATION")
        print("=" * 60)

        is_new_curriculum = (41 in self.subjects_df['subject_id'].values) or (42 in self.subjects_df['subject_id'].values)
        max_compulsory = 8 if is_new_curriculum else 7

        def build_necta_string(row):
            parts = []
            for i, col in enumerate(self.valid_subject_cols):
                if i >= len(self.subjects_df): break
                mark = row[col]
                grade = row[f"{col}_grade"]
                short = self.subject_column_map[col]['subject_short']
                if i < max_compulsory:
                    if pd.isna(grade) or grade is None:
                        parts.append(f" {short}-'X'")
                    else:
                        parts.append(f" {short} {int(mark)} -'{grade}'")
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
        necta_sample['No'] = range(1, 7)
        for i, row in necta_sample.iterrows():
            txt = row['necta_results']
            necta_sample.at[i, 'necta_results'] = txt[:100] + '...' if len(txt) > 100 else txt
        necta_sample = necta_sample[['No', 'full_name', 'necta_results']]
        print(tabulate(necta_sample, headers='keys', tablefmt='fancy_grid', showindex=False))

    # -------------------------------
    # 9. UPDATE DATABASE
    # -------------------------------
    def update_database(self):
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

        print(f"{self.YELLOW}UPDATING {len(update_columns)} COLUMNS ACROSS {len(self.df):,} RECORDS")

        set_clause = ", ".join([f"{col} = ?" for col in update_columns])
        update_sql = f"UPDATE tbl_student_exam_results SET {set_clause} WHERE result_id = ?"

        success_count = 0
        for _, row in tqdm(self.df.iterrows(), total=len(self.df), desc="Updating Database"):
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
            try:
                self.cursor.execute(update_sql, values)
                success_count += 1
            except Exception as e:
                print(f"{self.RED}Error updating {row['full_name']}: {e}")
                break

        self.conn.commit()
        status = f"{self.GREEN}COMPLETE" if success_count == len(self.df) else f"{self.RED}ISSUE"
        print(f"{self.CYAN}DATABASE UPDATE {status}: {success_count:,}/{len(self.df):,} records")
        return success_count

    # -------------------------------
    # 10. COMPETENCY TABLE
    # -------------------------------
    def update_competency_analysis(self):
        if not self.UPDATE_COMPETENCY:
            return

        print(f"\n{self.GREEN}10. COMPETENCY ANALYSIS UPDATE")
        print("=" * 60)
        
        self.cursor.execute("DELETE FROM tbl_competency WHERE exam_id = ?", self.EXAM_ID)
        
        def calculate_gpa(A, B, C, D, F):
            total = A + B + C + D + F
            return (A*1 + B*2 + C*3 + D*4 + F*5) / total if total > 0 else 0.0
        
        def get_competency_level(gpa):
            if gpa >= 4.6: return "Grade F (Fail)"
            if gpa >= 3.6: return "Grade D (Satisfactory)"
            if gpa >= 2.6: return "Grade C (Good)"
            if gpa >= 1.6: return "Grade B (Very Good)"
            if gpa >= 1.0: return "Grade A (Excellent)"
            return ""
        
        competency_data = []
        for _, subject in self.subjects_df.iterrows():
            subject_id = int(subject['subject_id'])
            col_name = next((col for col, info in self.subject_column_map.items() if info['subject_id'] == subject_id), None)
            if not col_name: continue
            grade_col = f"{col_name}_grade"
            counts = self.df[grade_col].value_counts()
            A = int(counts.get('A', 0))
            B = int(counts.get('B', 0))
            C = int(counts.get('C', 0))
            D = int(counts.get('D', 0))
            F = int(counts.get('F', 0))
            total = A + B + C + D + F
            gpa = float(calculate_gpa(A, B, C, D, F))
            level = get_competency_level(gpa)
            competency_data.append({
                'No': len(competency_data) + 1,
                'Subject': subject['subject_name'],
                'A': A, 'B': B, 'C': C, 'D': D, 'F': F,
                'Total': total,
                'GPA': round(gpa, 4),
                'Level': level
            })
            self.cursor.execute("""
                INSERT INTO tbl_competency (exam_id, subject_id, A_s, B_s, C_s, D_s, F_s, gpa, competency_level)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, self.EXAM_ID, subject_id, A, B, C, D, F, gpa, level)
        
        self.conn.commit()
        
        print(f"{self.MAGENTA}SUBJECT COMPETENCY ANALYSIS:")
        comp_df = pd.DataFrame(competency_data)
        print(tabulate(comp_df[['No', 'Subject', 'A', 'B', 'C','D','F', 'GPA', 'Level']].head(10),
                      headers='keys', tablefmt='fancy_grid', showindex=False))

    # -------------------------------
    # 11. FINAL SUMMARY REPORT
    # -------------------------------
    def generate_summary_report(self, success_count):
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
        print(f"   • All valid students ranked by: Points ASC, Avg Marks DESC, Subject Count DESC")
        print(f"   • Students with NULL points assigned worst ranking position")
        print(f"   • ABS students: Not ranked")
        print(f"   • Subject ranking: Marks DESC (not null students only)")

        now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(f"\n{self.CYAN}PROCESS COMPLETED AT: {self.WHITE}{now}")

    # -------------------------------
    # MAIN EXECUTION
    # -------------------------------
    def run(self):
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
            self.update_competency_analysis()
            self.generate_summary_report(success_count)

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
    processor = OlevelProcessor(
        exam_id="MID420251027",
        db_path=r"C:\Kiyabo App\backend\Kiyabo App Backend v2.0.0.accdb",
        base_subjects=7,
        flat_rate=True,
        include_inc=True,
        update_competency=True
    )
    processor.run()