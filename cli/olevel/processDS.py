# ========================================
# NECTA RESULTS ENGINE — FIXED RANKING
# ========================================
import pyodbc
import pandas as pd
import numpy as np
from tqdm import tqdm
import warnings
warnings.filterwarnings("ignore")

# -------------------------------
# CONFIG
# -------------------------------
EXAM_ID = "MID420251027"
DB_PATH = r"C:\Kiyabo App\backend\Kiyabo App Backend v2.0.0.accdb"

BASE_SUBJECTS = 7
FLAT_RATE = True
INCLUDE_INC = True
UPDATE_COMPETENCY = True

# -------------------------------
# CONNECT
# -------------------------------
print("🔌 CONNECTING TO ACCESS DATABASE...")
conn = pyodbc.connect(
    f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};"
    f"DBQ={DB_PATH};"
)
cursor = conn.cursor()

# -------------------------------
# 1. LOAD DATA
# -------------------------------
print("\n📊 1. LOADING EXAM RESULTS DATA")
print("=" * 60)

results_sql = """
    SELECT result_id, student_id, 
           civ, his, geo, kis, eng, phy, che, bio, mat, edk, ics,
           sub12, sub13, sub14, sub15, sub16, sub17, sub18, sub19, sub20
    FROM tbl_student_exam_results 
    WHERE exam_id = ?
"""
df = pd.read_sql(results_sql, conn, params=[EXAM_ID])

students_sql = "SELECT student_id, full_name, sex FROM tbl_student_academic_info"
students_df = pd.read_sql(students_sql, conn)

exam_sql = "SELECT exam_id, class_id FROM tbl_student_exams WHERE exam_id = ?"
exam_df = pd.read_sql(exam_sql, conn, params=[EXAM_ID])

df = df.merge(students_df, on='student_id', how='left')
class_id = exam_df['class_id'].iloc[0] if not exam_df.empty else 'IV'

print(f"🏫 Class ID: {class_id}")
print(f"👥 Total Students: {len(df):,}")

# -------------------------------
# 2. SUBJECT MAPPING
# -------------------------------
print("\n📚 2. SUBJECT CONFIGURATION")
print("=" * 60)

subjects_sql = f"""
    SELECT subject_id, subject_name, subject_code, subject_short
    FROM tbl_school_subjects 
    WHERE is_present_{class_id} = True
"""
subjects_df = pd.read_sql(subjects_sql, conn)

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

subjects_df = sort_subjects_vba_logic(subjects_df)

print("📖 SUBJECTS CONFIGURED FOR THIS CLASS:")
subjects_display = subjects_df[['subject_id', 'subject_short', 'subject_name']].copy()
subjects_display['No'] = range(1, len(subjects_display) + 1)
print(subjects_display[['No', 'subject_id', 'subject_short', 'subject_name']].head(10).to_string(index=False))

mark_columns = ['civ','his','geo','kis','eng','phy','che','bio','mat','edk','ics'] + \
              [f'sub{i}' for i in range(12,21)]

subject_column_map = {}
for i, subject_row in subjects_df.iterrows():
    if i < len(mark_columns):
        col_name = mark_columns[i]
        subject_column_map[col_name] = {
            'subject_short': subject_row['subject_short'],
            'subject_id': int(subject_row['subject_id']),
            'subject_name': subject_row['subject_name']
        }

valid_subject_cols = []
for col in mark_columns:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')
        if df[col].notna().any():
            valid_subject_cols.append(col)

print(f"\n✅ ACTIVE SUBJECT COLUMNS: {len(valid_subject_cols)}")
valid_cols_df = pd.DataFrame({
    'No': range(1, len(valid_subject_cols) + 1),
    'Column': valid_subject_cols,
    'Subject': [subject_column_map[col]['subject_name'] for col in valid_subject_cols]
})
print(valid_cols_df.head(10).to_string(index=False))

keep_cols = ['result_id', 'student_id', 'full_name', 'sex'] + valid_subject_cols
df = df[keep_cols].copy()

# -------------------------------
# 3. GRADE CALCULATION
# -------------------------------
print("\n🎓 3. GRADE CALCULATION PROCESS")
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

for col in valid_subject_cols:
    df[f"{col}_grade"] = df[col].apply(calculate_grade)

print("📝 STUDENT MARKS SAMPLE (First 8 students):")
sample_data = []
for i in range(min(8, len(df))):
    row = df.iloc[i]
    sample_row = {'No': i+1, 'Student': row['full_name']}
    for j, col in enumerate(valid_subject_cols[:4]):
        sample_row[f'{subject_column_map[col]["subject_short"]}'] = f"{row[col]:.0f}"
    sample_data.append(sample_row)

sample_df = pd.DataFrame(sample_data)
print(sample_df.to_string(index=False))

# -------------------------------
# 4. SUBJECT COUNT & TOTAL MARKS
# -------------------------------
print("\n🧮 4. ACADEMIC AGGREGATION")
print("=" * 60)

df['subject_count_real'] = df[valid_subject_cols].notna().sum(axis=1)

if FLAT_RATE:
    df['subject_count'] = BASE_SUBJECTS
    df['total_marks'] = df[valid_subject_cols].apply(
        lambda row: sum(sorted([m for m in row if not pd.isna(m)], reverse=True)[:BASE_SUBJECTS]), 
        axis=1
    )
else:
    df['subject_count'] = df['subject_count_real'].apply(
        lambda x: BASE_SUBJECTS if x <= BASE_SUBJECTS else x
    )
    df['total_marks'] = df[valid_subject_cols].sum(axis=1)

df['avg_marks'] = np.where(df['subject_count'] > 0, df['total_marks'] / df['subject_count'], np.nan)

print("📈 STUDENT PERFORMANCE SUMMARY:")
summary_data = []
for i in range(min(8, len(df))):
    row = df.iloc[i]
    summary_data.append({
        'No': i+1,
        'Student': row['full_name'],
        'Real Subjects': int(row['subject_count_real']),
        'Counted Subjects': int(row['subject_count']),
        'Total Marks': f"{row['total_marks']:.0f}",
        'Average Marks': f"{row['avg_marks']:.2f}"
    })

summary_df = pd.DataFrame(summary_data)
print(summary_df.to_string(index=False))

# -------------------------------
# 5. POINTS & DIVISION
# -------------------------------
print("\n⭐ 5. POINTS & DIVISION CALCULATION")
print("=" * 60)

grade_points = {'A': 1, 'B': 2, 'C': 3, 'D': 4, 'F': 5}

def calculate_points_and_division(row):
    subject_count_real = row['subject_count_real']
    
    if subject_count_real == 0:
        return None, 'ABS'
    
    valid_grades = [row[f"{col}_grade"] for col in valid_subject_cols 
                   if row[f"{col}_grade"] in grade_points]
    
    if not valid_grades:
        return None, 'ABS'
    
    points_list = sorted([grade_points[g] for g in valid_grades])[:7]
    points_list = points_list + [5] * (7 - len(points_list))
    points = sum(points_list)
    
    if INCLUDE_INC:
        if subject_count_real < BASE_SUBJECTS:
            return None, 'INC'
    else:
        if subject_count_real < BASE_SUBJECTS:
            potential_points = points + (5 * (BASE_SUBJECTS - subject_count_real))
            return None, '0' if potential_points >= 34 else 'IV'
    
    if points <= 17: return points, 'I'
    if points <= 22: return points, 'II'
    if points <= 25: return points, 'III'
    if points <= 33: return points, 'IV'
    return points, '0'

points_division = df.apply(calculate_points_and_division, axis=1, result_type='expand')
df['points'] = points_division[0]
df['division'] = points_division[1]

print("🎯 STUDENT POINTS & DIVISIONS:")
points_data = []
for i in range(min(8, len(df))):
    row = df.iloc[i]
    points_data.append({
        'No': i+1,
        'Student': row['full_name'],
        'Subjects Attempted': int(row['subject_count_real']),
        'Points': int(row['points']) if pd.notna(row['points']) else 'N/A',
        'Division': row['division']
    })

points_df = pd.DataFrame(points_data)
print(points_df.to_string(index=False))

print("\n📊 DIVISION DISTRIBUTION:")
div_counts = df['division'].value_counts().reset_index()
div_counts.columns = ['Division', 'Students']
div_counts['Percentage'] = (div_counts['Students'] / len(df) * 100).round(1)
div_counts = div_counts.sort_values('Students', ascending=False)
print(div_counts.to_string(index=False))

# -------------------------------
# 6. RANKING - FIXED MIN METHOD
# -------------------------------
print("\n🏆 6. ACADEMIC RANKING (FIXED MIN METHOD)")
print("=" * 60)

# Separate students into groups for ranking
abs_students = df[df['division'] == 'ABS'].copy()
valid_students = df[df['division'] != 'ABS'].copy()

print(f"📊 RANKING BREAKDOWN:")
print(f"   • ABS Students: {len(abs_students):,}")
print(f"   • Valid Students: {len(valid_students):,}")

# For students with NULL points (INC/0/IV), assign them the worst possible points
# This ensures they are ranked at the bottom
max_points = valid_students['points'].max() if not valid_students['points'].isna().all() else 35
valid_students['ranking_points'] = valid_students['points'].fillna(max_points + 1)

# Rank ALL valid students by: ranking_points ASC, avg_marks DESC, subject_count_real DESC
if len(valid_students) > 0:
    valid_students = valid_students.sort_values(
        ['ranking_points', 'avg_marks', 'subject_count_real'], 
        ascending=[True, False, False]
    )
    valid_students['position'] = valid_students.groupby(
        ['ranking_points', 'avg_marks', 'subject_count_real']
    ).ngroup() + 1
    valid_students['out_of'] = len(valid_students)

# Merge back to main dataframe
df = df.merge(valid_students[['result_id', 'position', 'out_of']], on='result_id', how='left')

print("\n🥇 TOP 10 RANKED STUDENTS:")
top_10 = valid_students[['full_name', 'points', 'avg_marks', 'position', 'division']].head(10).copy()
top_10['avg_marks'] = top_10['avg_marks'].round(2)
top_10['No'] = range(1, 11)
print(top_10[['No', 'full_name', 'points', 'avg_marks', 'position', 'division']].to_string(index=False))

print("\n📉 BOTTOM 10 RANKED STUDENTS:")
bottom_10 = valid_students[['full_name', 'points', 'avg_marks', 'position', 'division']].tail(10).copy()
bottom_10['avg_marks'] = bottom_10['avg_marks'].round(2)
bottom_10['No'] = range(1, 11)
print(bottom_10[['No', 'full_name', 'points', 'avg_marks', 'position', 'division']].to_string(index=False))

# Show students with NULL points if any
null_points_students = valid_students[valid_students['points'].isna()]
if len(null_points_students) > 0:
    print(f"\n🔄 STUDENTS WITHOUT POINTS (INC/0/IV) - RANKED AT BOTTOM:")
    inc_students = null_points_students[['full_name', 'points', 'avg_marks', 'position', 'division']].copy()
    inc_students['avg_marks'] = inc_students['avg_marks'].round(2)
    inc_students['No'] = range(1, len(inc_students) + 1)
    print(inc_students[['No', 'full_name', 'points', 'avg_marks', 'position', 'division']].to_string(index=False))

# -------------------------------
# 7. SUBJECT-WISE RANKING
# -------------------------------
print("\n📚 7. SUBJECT-WISE RANKING (MARKS DESC, NOT NULL)")
print("=" * 60)

print("🔢 CALCULATING SUBJECT POSITIONS...")
for col in valid_subject_cols:
    # Only rank students with valid marks in this subject (NOT NULL)
    subject_rank_df = df[df[col].notna()].copy()
    if len(subject_rank_df) > 0:
        # Sort by marks DESCENDING
        subject_rank_df = subject_rank_df.sort_values(col, ascending=False)
        # MIN method for subject positions
        subject_rank_df[f'{col}_pos'] = subject_rank_df.groupby(col).ngroup() + 1
        subject_rank_df[f'{col}_out_of'] = len(subject_rank_df)
        
        # Merge back to main dataframe
        df = df.merge(
            subject_rank_df[['result_id', f'{col}_pos', f'{col}_out_of']],
            on='result_id', how='left'
        )

print("📊 SUBJECT RANKING SAMPLE (First Subject - Top 10):")
if len(valid_subject_cols) > 0:
    first_subject = valid_subject_cols[0]
    subject_name = subject_column_map[first_subject]['subject_name']
    print(f"Subject: {subject_name}")
    
    # Get top 10 students in this subject (marks DESC)
    subject_top_10 = df[df[first_subject].notna()].nlargest(10, first_subject)[
        ['full_name', first_subject, f'{first_subject}_pos', f'{first_subject}_out_of']
    ]
    subject_top_10['No'] = range(1, len(subject_top_10) + 1)
    subject_top_10 = subject_top_10.rename(columns={
        first_subject: 'Marks',
        f'{first_subject}_pos': 'Position',
        f'{first_subject}_out_of': 'Out Of'
    })
    print(subject_top_10[['No', 'full_name', 'Marks', 'Position', 'Out Of']].to_string(index=False))

# -------------------------------
# 8. NECTA RESULTS STRING
# -------------------------------
print("\n📄 8. NECTA RESULTS STRING GENERATION")
print("=" * 60)

is_new_curriculum = (41 in subjects_df['subject_id'].values) or (42 in subjects_df['subject_id'].values)
max_compulsory = 8 if is_new_curriculum else 7

def build_necta_string(row):
    parts = []
    for i, col in enumerate(valid_subject_cols):
        if i >= len(subjects_df): break
            
        mark = row[col]
        grade = row[f"{col}_grade"]
        short = subject_column_map[col]['subject_short']
        
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

tqdm.pandas(desc="🔄 Generating NECTA strings")
df['necta_results'] = df.progress_apply(build_necta_string, axis=1)

print("📋 NECTA RESULTS SAMPLE:")
necta_sample = df[['full_name', 'necta_results']].head(6).copy()
necta_sample['No'] = range(1, 7)
for i, row in necta_sample.iterrows():
    necta_sample.at[i, 'necta_results'] = row['necta_results'][:70] + '...' if len(row['necta_results']) > 70 else row['necta_results']
print(necta_sample[['No', 'full_name', 'necta_results']].to_string(index=False))

# -------------------------------
# 9. UPDATE DATABASE
# -------------------------------
print("\n💾 9. DATABASE UPDATE PROCESS")
print("=" * 60)

update_columns = [
    'necta_results', 'subject_count', 'total_marks', 
    'points', 'division', 'position', 'out_of'
]

for col in valid_subject_cols:
    update_columns.append(f"{col}_grade")
    if f"{col}_pos" in df.columns:
        update_columns.extend([f"{col}_pos", f"{col}_out_of"])

print(f"🔄 UPDATING {len(update_columns)} COLUMNS ACROSS {len(df):,} RECORDS")

set_clause = ", ".join([f"{col} = ?" for col in update_columns])
update_sql = f"UPDATE tbl_student_exam_results SET {set_clause} WHERE result_id = ?"

success_count = 0
for _, row in tqdm(df.iterrows(), total=len(df), desc="📤 Updating Database"):
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
        cursor.execute(update_sql, values)
        success_count += 1
    except Exception as e:
        print(f"❌ Error updating {row['full_name']}: {e}")
        break

conn.commit()
print(f"✅ DATABASE UPDATE COMPLETE: {success_count:,}/{len(df):,} records")

# -------------------------------
# 10. COMPETENCY TABLE
# -------------------------------
if UPDATE_COMPETENCY:
    print("\n📊 10. COMPETENCY ANALYSIS UPDATE")
    print("=" * 60)
    
    cursor.execute("DELETE FROM tbl_competency WHERE exam_id = ?", EXAM_ID)
    
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
    for _, subject in subjects_df.iterrows():
        subject_id = int(subject['subject_id'])
        
        col_name = None
        for col, info in subject_column_map.items():
            if info['subject_id'] == subject_id:
                col_name = col
                break
        if not col_name: continue
        
        grade_col = f"{col_name}_grade"
        counts = df[grade_col].value_counts()
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
        
        cursor.execute("""
            INSERT INTO tbl_competency (exam_id, subject_id, A_s, B_s, C_s, D_s, F_s, gpa, competency_level)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, EXAM_ID, subject_id, A, B, C, D, F, gpa, level)
    
    conn.commit()
    
    print("📈 SUBJECT COMPETENCY ANALYSIS:")
    comp_df = pd.DataFrame(competency_data)
    print(comp_df[['No', 'Subject', 'A', 'B', 'C','D','F', 'GPA', 'Level']].head(10).to_string(index=False))

# -------------------------------
# 11. FINAL SUMMARY REPORT
# -------------------------------
print("\n📋 11. COMPREHENSIVE SUMMARY REPORT")
print("=" * 60)

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
        f"{len(df):,}",
        f"{len(valid_subject_cols)}",
        f"{df['total_marks'].mean():.1f}",
        f"{len(df[df['division'] == 'I']):,}",
        f"{valid_students['position'].nunique():,}",
        f"{success_count:,}/{len(df):,}"
    ],
    'Status': [
        '✅ COMPLETE',
        '✅ CONFIGURED',
        '📊 CALCULATED',
        '🎯 CLASSIFIED',
        '🏆 RANKED',
        '💾 UPDATED' if success_count == len(df) else '❌ ISSUE'
    ]
}

summary_df = pd.DataFrame(summary_data)
print(summary_df.to_string(index=False))

print(f"\n🎯 RANKING METHODOLOGY:")
print(f"   • All valid students ranked by: Points ASC, Avg Marks DESC, Subject Count DESC")
print(f"   • Students with NULL points assigned worst ranking position")
print(f"   • ABS students: Not ranked")
print(f"   • Subject ranking: Marks DESC (not null students only)")

print(f"\n📅 PROCESS COMPLETED AT: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")

conn.close()

print("\n" + "=" * 60)
print("🎉 PROCESS COMPLETED SUCCESSFULLY!")
print("=" * 60)