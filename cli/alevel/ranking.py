import pandas as pd
import pyodbc
import numpy as np
from tqdm import tqdm
import argparse
import sys

# SUBJECTS: Base mark columns
SUBJECTS = ["acc", "agr", "ara", "bio", "bus", "che", "chi", "com", "csc", "eng", "eco", "fsh", "fin", "fhn", "fre", "tex", "geo", "his", "isl", "kis", "lit", "mat", "msc", "spt", "the", "phy", "bam", "gs", "htm", "aco", "cus"]

# UPDATE FIELDS: avg_grade SAVED, avg_marks NOT SAVED
DB_UPDATE_FIELDS = [
    "division", "points", "subject_count", "total_marks", "gpa",
    "position_comb", "position_school", "out_of_comb", "out_of_school",
    "first", "second", "third", 'avg_grade',
    "subject_count_all", "necta_results", "necta_results_marks"
]

def bracket_field(field):
    return f"[{field}]"

def get_grade(marks):
    if pd.isna(marks) or not isinstance(marks, (int, float)) or marks < 0 or marks > 100:
        return None
    if marks >= 80: return 'A'
    elif marks >= 70: return 'B'
    elif marks >= 60: return 'C'
    elif marks >= 50: return 'D'
    elif marks >= 40: return 'E'
    elif marks >= 35: return 'S'
    else: return 'F'

GRADE_POINTS = {'A':1, 'B':2, 'C':3, 'D':4, 'E':5, 'S':6, 'F':7, None: None}
DIVISION_VALUES = {'I': 1, 'II': 2, 'III': 3, 'IV': 4, '0': 5, 'ABS': None, 'INC': None}
POINTS_TO_DIV = {
    range(3, 10): 'I',
    range(10, 13): 'II',
    range(13, 18): 'III',
    range(18, 20): 'IV',
    range(20, 22): '0'
}

def get_div_from_points(points):
    if pd.isna(points) or not isinstance(points, int):
        return None
    points = int(points)
    for r, div in POINTS_TO_DIV.items():
        if points in r:
            return div
    return None

def format_marks_for_necta(marks):
    if pd.isna(marks):
        return ''
    if isinstance(marks, (int, np.integer)) or (isinstance(marks, float) and marks.is_integer()):
        return str(int(marks))
    else:
        return f"{marks:.2f}"

def format_grade_for_necta(grade):
    return f"'{grade}'" if grade else ''

def build_necta(row, comb_metadata, subject_to_user, valid_subjects):
    parts = []
    parts_marks = []
    student_comb = comb_metadata[comb_metadata['comb_id'] == row['comb_id']]
    comb_shorts = set(student_comb['subject_short'].str.lower())
   
    def get_user(short):
        meta_row = student_comb[student_comb['subject_short'].str.lower() == short]
        if not meta_row.empty:
            return meta_row.iloc[0]['subject_user_short']
        return subject_to_user.get(short, short.upper())
   
    for _, sub_row in student_comb.iterrows():
        short = sub_row['subject_short'].lower()
        user = sub_row['subject_user_short']
        marks = row.get(short)
        grade = row.get(f'{short}_grade')
        if pd.notna(marks):
            parts.append(f"{user}-{format_grade_for_necta(grade)}")
            marks_str = format_marks_for_necta(marks)
            parts_marks.append(f"{user}-{marks_str} {format_grade_for_necta(grade)}")
        else:
            parts.append(f"{user}-X")
            parts_marks.append(f"{user}-X")
   
    for sub in valid_subjects:
        if pd.notna(row.get(sub)) and sub not in comb_shorts:
            user = get_user(sub)
            grade = row.get(f'{sub}_grade')
            parts.append(f"{user}-{format_grade_for_necta(grade)}")
            marks_str = format_marks_for_necta(row[sub])
            parts_marks.append(f"{user}-{marks_str} {format_grade_for_necta(grade)}")
   
    avg_grade = row.get('avg_grade')
    avg_grade_fmt = format_grade_for_necta(avg_grade)
    parts.append(f"AVG-{avg_grade_fmt}")
   
    avg_marks_val = row.get('avg_marks')
    avg_marks_str = format_marks_for_necta(avg_marks_val) if pd.notna(avg_marks_val) else 'X'
    parts_marks.append(f"AVG-{avg_marks_str}{avg_grade_fmt}")
   
    row['necta_results'] = ', '.join(parts)
    row['necta_results_marks'] = ', '.join(parts_marks)
    return row

# === FINAL compute_ranking: ABS + NO MARKS → NULL ===
def compute_ranking(df, exam_id, valid_subjects, ranking_order=["points", "avg_marks", "subject_count"], ascending=[True, False, False]):
    exam_df = df[df['exam_id'] == exam_id].copy()
    if exam_df.empty:
        return df

    # VALID: NOT ABS + HAS AT LEAST ONE MARK (0 is valid)
    valid_students = exam_df[
        (exam_df['division'] != 'ABS') &
        (exam_df[valid_subjects].notna().any(axis=1))
    ].copy()

    invalid_students = exam_df[
        (exam_df['division'] == 'ABS') |
        (~exam_df[valid_subjects].notna().any(axis=1))
    ].copy()

    if valid_students.empty:
        df.loc[exam_df.index, ['position_school', 'out_of_school', 'position_comb', 'out_of_comb']] = pd.NA
        return df

    valid_students['points_for_rank'] = valid_students['points'].replace({None: np.inf, np.nan: np.inf})
    valid_students = valid_students.sort_values(
        ['points_for_rank', 'avg_marks', 'subject_count'],
        ascending=[True, False, False]
    )

    n_valid = len(valid_students)
    valid_students['position_school'] = np.arange(1, n_valid + 1)
    valid_students['out_of_school'] = n_valid

    def assign_comb_rank(group):
        group['points_for_rank'] = group['points'].replace({None: np.inf, np.nan: np.inf})
        group = group.sort_values(['points_for_rank', 'avg_marks', 'subject_count'], ascending=[True, False, False])
        group['position_comb'] = np.arange(1, len(group) + 1)
        group['out_of_comb'] = len(group)
        return group

    valid_students = valid_students.groupby('comb_id', group_keys=False).apply(assign_comb_rank)

    df.loc[valid_students.index, ['position_school', 'out_of_school', 'position_comb', 'out_of_comb']] = \
        valid_students[['position_school', 'out_of_school', 'position_comb', 'out_of_comb']]

    if not invalid_students.empty:
        df.loc[invalid_students.index, ['position_school', 'out_of_school', 'position_comb', 'out_of_comb']] = pd.NA

    if 'points_for_rank' in df.columns:
        df = df.drop(columns=['points_for_rank'], errors='ignore')

    return df

# === FINAL compute_subject_rankings: 0 is valid ===
def compute_subject_rankings(df, valid_subjects):
    for sub in tqdm(valid_subjects, desc="Ranking subjects"):
        pos_col = f"{sub}_pos"
        out_col = f"{sub}_out_of"

        sub_df = df[df[sub].notna()].copy()
        if sub_df.empty:
            continue

        sub_df = sub_df.sort_values(sub, ascending=False)
        sub_df['rank_temp'] = sub_df[sub].rank(method='min', ascending=False)
        sub_df[pos_col] = sub_df['rank_temp'].astype('Int64')
        sub_df[out_col] = len(sub_df)

        df.loc[sub_df.index, [pos_col, out_col]] = sub_df[[pos_col, out_col]]

    return df

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

def colored_print(text, color=bcolors.OKBLUE):
    print(f"{color}{text}{bcolors.ENDC}")

def display_all_fields_in_results(exam_id='ANN520250526', dbpath=r"C:\Kiyabo App\backend\Kiyabo App Backend v4.0.0.accdb"):
    conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + dbpath + ';'
    conn = pyodbc.connect(conn_str)
    query = "SELECT TOP 5 * FROM tbl_student_exam_results WHERE exam_id = ?"
    df_sample = pd.read_sql(query, conn, params=[exam_id])
    conn.close()
    colored_print("All fields in tbl_student_exam_results:", bcolors.HEADER)
    print(df_sample.columns.tolist())
    colored_print("\nSample data (TOP 5 rows, all fields):", bcolors.OKCYAN)
    print(df_sample.head())

def process_exam(exam_id='ANN520250526', dbpath=r"C:\Kiyabo App\backend\Kiyabo App Backend v4.0.0.accdb",
                 ranking_order=["points", "avg_marks", "subject_count"], ranking_asc=[True, False, False], include_INC=True):
    colored_print("Stage 1: Connecting to Database and Detecting Columns", bcolors.HEADER)
    conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + dbpath + ';'
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
   
    dummy_query = "SELECT TOP 1 * FROM tbl_student_exam_results WHERE exam_id = ?"
    df_dummy = pd.read_sql(dummy_query, conn, params=[exam_id])
    if df_dummy.empty:
        colored_print(" - No records for exam_id, falling back to general structure.", bcolors.WARNING)
        df_dummy = pd.read_sql("SELECT TOP 1 * FROM tbl_student_exam_results", conn)
    db_columns = [col.lower() for col in df_dummy.columns]
   
    potential_subjects = [sub for sub in SUBJECTS if sub in db_columns]
    colored_print(f" - Detected {len(potential_subjects)} potential subject mark fields: {potential_subjects}", bcolors.OKGREEN)
   
    if 'student_id' not in db_columns or 'exam_id' not in db_columns:
        raise ValueError("Missing essential fields.")
   
    colored_print("Stage 2: Building SQL Query and Fetching Data", bcolors.HEADER)
    update_fields = [field for field in DB_UPDATE_FIELDS if field in db_columns]
   
    select_parts = [f"r.{bracket_field('student_id')}", f"r.{bracket_field('exam_id')}"]
    fixed_select = ', '.join([f"r.{bracket_field(f)}" for f in update_fields])
    if fixed_select: select_parts.append(fixed_select)
    subject_select = ', '.join([f"r.{bracket_field(sub)}" for sub in potential_subjects])
    if subject_select: select_parts.append(subject_select)
    select_clause = ', '.join(select_parts)
   
    query = f"""
    SELECT {select_clause},
           i.{bracket_field('full_name')}, i.{bracket_field('sex')}, i.{bracket_field('comb_id')}
    FROM tbl_student_exam_results r
    INNER JOIN tbl_student_academic_info i ON r.{bracket_field('student_id')} = i.{bracket_field('student_id')}
    WHERE r.{bracket_field('exam_id')} = ?
    """
    df = pd.read_sql(query, conn, params=[exam_id])
    colored_print(f" - Fetched {len(df)} student records.", bcolors.OKGREEN)
   
    if df.empty:
        colored_print("No records. Skipping further processing.", bcolors.WARNING)
        conn.close()
        return
   
    colored_print("Stage 2.5: Filtering Subjects with At Least One Valid Mark (>=0)", bcolors.HEADER)
    valid_subjects = []
    for sub in potential_subjects:
        df[sub] = pd.to_numeric(df[sub], errors='coerce')
        if df[sub].ge(0).any():
            valid_subjects.append(sub)
        else:
            colored_print(f" - Dropping subject {sub}: no valid marks >=0 in any record.", bcolors.WARNING)
            if sub in df.columns:
                df = df.drop(columns=[sub])
    colored_print(f" - Retained {len(valid_subjects)} subjects with valid marks: {valid_subjects}", bcolors.OKGREEN)
   
    sample_subjects_dynamic = valid_subjects[:5] if len(valid_subjects) >= 5 else valid_subjects
    sample_grades_dynamic = [f'{sub}_grade' for sub in sample_subjects_dynamic]
   
    colored_print("Stage 3: Fetching Metadata for Combinations and Subjects", bcolors.HEADER)
    comb_df = pd.read_sql("SELECT serial_id, comb_id, subject_id FROM tbl_student_comb_subjects", conn)
    sub_df = pd.read_sql("SELECT subject_serial, subject_short, subject_user_short, is_core, is_present FROM tbl_student_subjects", conn)
   
    conn.close()
    colored_print(f" - Loaded {len(comb_df)} comb-subject links and {len(sub_df)} subjects.", bcolors.OKGREEN)
   
    comb_metadata_all = comb_df.merge(sub_df, left_on='subject_id', right_on='subject_serial')
    comb_metadata = comb_metadata_all[comb_metadata_all['is_present'] == True]
    subject_to_user = dict(zip(sub_df['subject_short'].str.lower(), sub_df['subject_user_short']))
   
    df['comb_id'] = df['comb_id'].astype(str)
   
    colored_print("Stage 4: Resetting Computed Fields in DB", bcolors.HEADER)
    grade_fields = [f"{sub}_grade" for sub in valid_subjects if f"{sub}_grade" in db_columns]
    pos_out_fields = [f"{sub}_pos" for sub in valid_subjects if f"{sub}_pos" in db_columns] + \
                     [f"{sub}_out_of" for sub in valid_subjects if f"{sub}_out_of" in db_columns]
    all_reset_fields = update_fields + grade_fields + pos_out_fields
   
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    for field in all_reset_fields:
        bracketed = bracket_field(field)
        cursor.execute(f"UPDATE tbl_student_exam_results SET {bracketed} = NULL WHERE {bracket_field('exam_id')} = ?", (exam_id,))
    conn.commit()
    colored_print(f" - Reset {len(all_reset_fields)} fields in DB for exam.", bcolors.OKGREEN)
   
    colored_print("Stage 5: Computing Grades from Marks", bcolors.HEADER)
    for sub in valid_subjects:
        df[sub] = df[sub].where(df[sub].between(0, 100), np.nan)
        df[f'{sub}_grade'] = df[sub].apply(get_grade)
   
    # === total_marks = FLOAT (NO ROUNDING) ===
    df['total_marks'] = df[valid_subjects].sum(axis=1, skipna=True)

    colored_print("Stage 6: Processing Each Student Row...", bcolors.HEADER)
    def process_row(row):
        student_comb = comb_metadata[comb_metadata['comb_id'] == row['comb_id']]
        comb_shorts = set(student_comb['subject_short'].str.lower())
        core_shorts = set(student_comb[student_comb['is_core'] == True]['subject_short'].str.lower())
       
        effective_cores = core_shorts.union([sub for sub in valid_subjects if pd.notna(row.get(sub)) and sub not in comb_shorts])
        attempted_effective = [sub for sub in effective_cores if pd.notna(row.get(sub))]
        missing_count = len(effective_cores) - len(attempted_effective)
       
        attempted_all = [sub for sub in valid_subjects if pd.notna(row.get(sub))]
        row['subject_count_all'] = len(attempted_all) if attempted_all else 0
        row['subject_count'] = len(attempted_effective) if attempted_effective else 0
       
        if row['subject_count_all'] > 0:
            row['avg_marks'] = row['total_marks'] / row['subject_count_all']
            row['avg_grade'] = get_grade(row['avg_marks'])
        else:
            row['avg_marks'] = None
            row['avg_grade'] = None
       
        core_marks = sorted([row[sub] for sub in attempted_effective if pd.notna(row.get(sub))], reverse=True)
        row['first'] = core_marks[0] if len(core_marks) >= 1 else None
        row['second'] = core_marks[1] if len(core_marks) >= 2 else None
        row['third'] = core_marks[2] if len(core_marks) >= 3 else None
       
        grade_pts = [GRADE_POINTS.get(row.get(f'{sub}_grade')) for sub in attempted_effective]
        valid_pts = [p for p in grade_pts if p is not None]
        computed_points = sum(valid_pts) if valid_pts else None
       
        row['computed_points'] = computed_points
       
        has_invalid = len(grade_pts) != len(valid_pts)
        is_complete = (len(effective_cores) >= 3 and missing_count == 0 and not has_invalid)
        is_abs = len(attempted_effective) == 0
        is_inc = not is_complete and not is_abs
       
        if is_abs:
            save_div = 'ABS' if include_INC else ('0' if get_div_from_points(7 * missing_count) == '0' else 'IV')
            save_points = None
        elif is_inc:
            save_div = 'INC' if include_INC else ('0' if get_div_from_points(sum(valid_pts) + 7 * missing_count) == '0' else 'IV')
            save_points = None
        else:
            save_div = get_div_from_points(computed_points)
            save_points = computed_points if save_div is not None else None
       
        row['division'] = save_div
        row['points'] = save_points
       
        div_val = DIVISION_VALUES.get(save_div)
        if div_val is not None and row['subject_count'] > 0:
            row['gpa'] = div_val / row['subject_count']
        else:
            row['gpa'] = None
       
        row = build_necta(row, comb_metadata, subject_to_user, valid_subjects)
        return row
   
    df = df.apply(process_row, axis=1)
   
    # === PASS valid_subjects TO compute_ranking ===
    df = compute_ranking(df, exam_id, valid_subjects, ranking_order, ranking_asc)
    df = compute_subject_rankings(df, valid_subjects)
   
    colored_print("Stage 9: Finalizing Data Types and Updating DB", bcolors.HEADER)
    
    # === int_fields: total_marks REMOVED (float) ===
    int_fields = ['position_school', 'position_comb', 'out_of_school', 'out_of_comb',
                  'first', 'second', 'third', 'points', 'subject_count', 'subject_count_all'] + \
                 [f"{sub}_pos" for sub in valid_subjects] + [f"{sub}_out_of" for sub in valid_subjects]
    
    for field in int_fields:
        if field in df.columns:
            df[field] = pd.to_numeric(df[field], errors='coerce').astype('Int64')
   
    if 'gpa' in df.columns:
        df['gpa'] = pd.to_numeric(df['gpa'], errors='coerce').round(4)
   
    if 'computed_points' in df.columns:
        df = df.drop(columns=['computed_points'])
   
    # === all_update_fields: avg_marks REMOVED ===
    all_update_fields = [f for f in [
        'points', 'division', 'subject_count', 'total_marks', 'gpa',
        'subject_count_all', 'necta_results', 'necta_results_marks',
        'first', 'second', 'third', 'avg_grade',
        'position_comb', 'position_school', 'out_of_comb', 'out_of_school'
    ] if f in db_columns] + grade_fields + pos_out_fields

    colored_print(f" - Updating {len(all_update_fields)} fields back to DB.", bcolors.OKBLUE)


    for idx, row in tqdm(df.iterrows(), total=len(df), desc="Updating records"):
        set_clause = ', '.join([f"{bracket_field(field)} = ?" for field in all_update_fields])
        params = tuple(None if pd.isna(row.get(field)) else row.get(field) for field in all_update_fields) + (row['student_id'], exam_id)
        update_query = f"UPDATE tbl_student_exam_results SET {set_clause} WHERE {bracket_field('student_id')} = ? AND {bracket_field('exam_id')} = ?"
        cursor.execute(update_query, params)

    
    conn.commit()
    
    cursor.execute ("UPDATE tbl_student_exam_results SET subject_count_all=NULL,subject_count=NULL where division=? AND exam_id=?", ('ABS', exam_id))
    conn.commit()
    
    conn.close()
    colored_print("\nAll updates committed to DB. Process complete.", bcolors.OKGREEN)

import warnings
warnings.filterwarnings("ignore", message="pandas only supports SQLAlchemy")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Process student exam results.")
    parser.add_argument('exam_id', type=str, help="The exam ID to process (e.g., MID520250825)")
    parser.add_argument('--dbpath', type=str, default=r"C:\Users\droge\OneDrive\Documents\Kiyabo App Backend v4.0.0.accdb", help="Path to the Access database file")
    parser.add_argument('--no-include-inc', action='store_false', dest='include_INC', help="Do not include INC (treat as penalty)")
    parser.add_argument('--display-fields', action='store_true', help="Display all fields and sample data before processing")
    
    args = parser.parse_args()
    
    if args.display_fields:
        display_all_fields_in_results(args.exam_id, args.dbpath)
    
    process_exam(args.exam_id, args.dbpath, include_INC=args.include_INC)