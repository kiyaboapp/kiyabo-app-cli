"""
Exam Results Processor - Python Implementation
Converts VBA exam completion logic to pandas-based processing
Author: Converted from VBA
Date: November 15, 2025
"""

import pandas as pd
import numpy as np
import json
import pyodbc
from typing import Dict, List, Optional, Tuple
from tqdm import tqdm
from colorama import Fore, Style, init
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.layout import Layout
from rich import box
from tabulate import tabulate

# Initialize colorama and rich
init(autoreset=True)
console = Console()


class ExamProcessor:
    """
    Main class for processing exam results
    """
    
    def __init__(self, db_path: str, exam_id: str):
        """
        Initialize with database path and exam ID
        
        Args:
            db_path: Path to MS Access database
            exam_id: Exam identifier to process
        """
        self.db_path = db_path
        self.exam_id = exam_id
        self.conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            f'DBQ={db_path};'
        )
        
        # Display header
        self._display_header()
        
        # Load configuration from database
        console.print("\n[cyan]📡 Connecting to database...[/cyan]")
        
        # Ensure result_json field exists
        self._ensure_result_json_field()
        
        self.class_id = self._get_class_from_exam_id()
        self.subject_columns = self._get_subject_columns()
        self.subject_mapping = self._get_subject_mapping()
        self.marks_style = self._get_marks_style()
        self.grades_df = self._load_grades()
        
        # Display configuration
        self._display_configuration()
    
    def _ensure_result_json_field(self):
        """Ensure result_json field exists in tbl_pupil_exam_results"""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        try:
            # Check if result_json column exists
            cursor.execute("SELECT TOP 1 result_json FROM tbl_pupil_exam_results")
            console.print("[green]✓[/green] result_json field exists")
        except:
            # Field doesn't exist, create it
            console.print("[yellow]⚠[/yellow] result_json field not found, creating it...")
            try:
                cursor.execute("ALTER TABLE tbl_pupil_exam_results ADD COLUMN result_json MEMO")
                conn.commit()
                console.print("[green]✓[/green] result_json field created successfully")
            except Exception as e:
                console.print(f"[red]❌ Error creating result_json field: {str(e)}[/red]")
                console.print("[yellow]→[/yellow] Will continue without result_json")
        
        conn.close()
    
    def _display_header(self):
        """Display beautiful header"""
        console.print("\n")
        console.print("╔" + "═" * 68 + "╗", style="bold magenta")
        console.print("║" + " " * 15 + "EXAM RESULTS PROCESSING SYSTEM" + " " * 23 + "║", style="bold magenta")
        console.print("║" + " " * 20 + "Kiyabo App Backend v3.0.0" + " " * 23 + "║", style="bold magenta")
        console.print("╚" + "═" * 68 + "╝", style="bold magenta")
        """Display beautiful header"""
        console.print("\n")
        console.print("╔" + "═" * 68 + "╗", style="bold magenta")
        console.print("║" + " " * 15 + "EXAM RESULTS PROCESSING SYSTEM" + " " * 23 + "║", style="bold magenta")
        console.print("║" + " " * 20 + "Kiyabo App Backend v3.0.0" + " " * 23 + "║", style="bold magenta")
        console.print("╚" + "═" * 68 + "╝", style="bold magenta")
    
    def _display_configuration(self):
        """Display configuration details in a beautiful table"""
        config_table = Table(title="📋 Configuration Details", 
                            box=box.ROUNDED, 
                            show_header=True,
                            header_style="bold cyan")
        
        config_table.add_column("Parameter", style="yellow", width=20)
        config_table.add_column("Value", style="green", width=40)
        
        config_table.add_row("Exam ID", self.exam_id)
        config_table.add_row("Class ID", self.class_id)
        config_table.add_row("Marks Style", self.marks_style)
        config_table.add_row("Total Subjects", str(len(self.subject_columns)))
        config_table.add_row("Grade Levels", str(len(self.grades_df)))
        
        console.print("\n")
        console.print(config_table)
        
        # Display subjects
        subjects_info = ", ".join([f"{self.subject_mapping[col]}" 
                                  for col in self.subject_columns[:10]])
        if len(self.subject_columns) > 10:
            subjects_info += f"... (+{len(self.subject_columns) - 10} more)"
        
        console.print(f"\n[cyan]📚 Subjects:[/cyan] {subjects_info}\n")
    
    def _get_connection(self):
        """Create database connection"""
        return pyodbc.connect(self.conn_str)
    
    def _get_class_from_exam_id(self) -> str:
        """Get class_id from exam_id"""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        query = "SELECT class_id FROM tbl_pupil_exams WHERE exam_id = ?"
        cursor.execute(query, (self.exam_id,))
        row = cursor.fetchone()
        
        conn.close()
        
        if row:
            return row[0]
        else:
            console.print(f"[red]❌ Error: Exam ID '{self.exam_id}' not found in database[/red]")
            raise ValueError(f"Exam ID '{self.exam_id}' not found in database")
    
    def _get_subject_columns(self) -> List[str]:
        """Get subject columns for the class"""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        query = """
        SELECT tbl_school_subjects.subject_number, tbl_class_subjects.ID
        FROM tbl_school_subjects 
        INNER JOIN tbl_class_subjects 
            ON tbl_school_subjects.subject_number = tbl_class_subjects.subject_id
        WHERE tbl_school_subjects.is_present = True 
            AND tbl_class_subjects.class_id = ?
        ORDER BY tbl_class_subjects.ID ASC
        """
        
        cursor.execute(query, (self.class_id,))
        rows = cursor.fetchall()
        
        conn.close()
        
        return [f"sub{int(row[0]):02d}" for row in rows]
    
    def _get_subject_mapping(self) -> Dict[str, str]:
        """Get mapping of subject columns to short names"""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        query = """
        SELECT tbl_school_subjects.subject_number, 
               tbl_school_subjects.subject_short
        FROM tbl_school_subjects 
        INNER JOIN tbl_class_subjects 
            ON tbl_school_subjects.subject_number = tbl_class_subjects.subject_id
        WHERE tbl_school_subjects.is_present = True 
            AND tbl_class_subjects.class_id = ?
        ORDER BY tbl_class_subjects.ID ASC
        """
        
        cursor.execute(query, (self.class_id,))
        rows = cursor.fetchall()
        
        conn.close()
        
        return {f"sub{int(row[0]):02d}": row[1] for row in rows}
    
    def _get_marks_style(self) -> str:
        """Get marks style for the class"""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        query = "SELECT marks_style FROM tbl_classes WHERE class_id = ?"
        cursor.execute(query, (self.class_id,))
        row = cursor.fetchone()
        
        conn.close()
        
        return row[0] if row and row[0] else "100 Marks"
    
    def _load_grades(self) -> pd.DataFrame:
        """Load grading configuration from database"""
        conn = self._get_connection()
        
        query = """
        SELECT grade, starting_50, ending_50, starting_100, ending_100
        FROM tbl_grades
        ORDER BY starting_100 DESC
        """
        
        df = pd.read_sql(query, conn)
        conn.close()
        
        return df
    
    def _load_results(self) -> pd.DataFrame:
        """Load exam results from database"""
        conn = self._get_connection()
        
        subject_cols = ', '.join(self.subject_columns)
        
        try:
            grade_cols = ', '.join([f"{col}_grade" for col in self.subject_columns])
            query = f"""
            SELECT exam_id, pupil_id, {subject_cols}, {grade_cols},
                   total_marks, subject_count, avg_marks, avg_grade,
                   pos, out_of, pos_stream, out_of_stream
            FROM tbl_pupil_exam_results
            WHERE exam_id = ?
            """
            df = pd.read_sql(query, conn, params=(self.exam_id,))
        except:
            query = f"""
            SELECT exam_id, pupil_id, {subject_cols}
            FROM tbl_pupil_exam_results
            WHERE exam_id = ?
            """
            df = pd.read_sql(query, conn, params=(self.exam_id,))
        
        student_query = """
        SELECT pupil_id, first_name, middle_name, surname, sex, section_id
        FROM tbl_pupil_academic_info
        """
        students_df = pd.read_sql(student_query, conn)
        
        conn.close()
        
        df = df.merge(students_df, on='pupil_id', how='left')
        
        # Create full name
        df['full_name'] = df.apply(
            lambda row: f"{row.get('first_name', '')} {row.get('middle_name', '')} {row.get('surname', '')}".strip(),
            axis=1
        )
        
        return df
    
    def _display_data_preview(self, df: pd.DataFrame):
        """Display comprehensive preview of loaded data"""
        console.print("\n" + "="*100)
        console.print("[bold cyan]📊 DATA PREVIEW - LOADED STUDENTS[/bold cyan]")
        console.print("="*100 + "\n")
        
        # Show first 15 students with all details
        preview_data = []
        for idx, row in df.head(15).iterrows():
            subjects_taken = sum(pd.notna(row[col]) for col in self.subject_columns)
            total = row.get('total_marks', 0) if pd.notna(row.get('total_marks')) else 0
            avg = row.get('avg_marks', 0) if pd.notna(row.get('avg_marks')) else 0
            
            # Get individual subject marks
            subject_marks = []
            for subj_col in self.subject_columns[:4]:  # Show first 4 subjects
                mark = row.get(subj_col, '')
                if pd.notna(mark):
                    subject_marks.append(f"{int(mark)}")
                else:
                    subject_marks.append("-")
            
            preview_data.append([
                row['pupil_id'],
                row.get('full_name', 'N/A')[:25],
                row.get('sex', 'N/A'),
                row.get('section_id', '-'),
                f"{subjects_taken}/{len(self.subject_columns)}",
                f"{int(total)}" if total > 0 else "-",
                f"{avg:.1f}" if avg > 0 else "-",
                row.get('avg_grade', '-'),
                " | ".join(subject_marks)
            ])
        
        # Get first 4 subject names for header
        subject_headers = [self.subject_mapping.get(col, col) for col in self.subject_columns[:4]]
        
        headers = ["Pupil ID", "Full Name", "Sex", "Stream", "Subj", "Total", "Avg", "Grade", " | ".join(subject_headers)]
        table_str = tabulate(preview_data, headers=headers, tablefmt="fancy_grid", 
                           colalign=("left", "left", "center", "center", "center", "right", "right", "center", "left"))
        
        print(table_str)
        
        # Summary statistics
        with_results = df[df[self.subject_columns].notna().any(axis=1)]
        console.print(f"\n[cyan]📈 Quick Stats:[/cyan]")
        console.print(f"  • Total records: [bold]{len(df)}[/bold]")
        console.print(f"  • With results: [bold]{len(with_results)}[/bold]")
        console.print(f"  • Male: {(df['sex'] == 'M').sum()} | Female: {(df['sex'] == 'F').sum()}")
        if 'section_id' in df.columns:
            streams = df['section_id'].value_counts()
            if len(streams) > 0:
                stream_str = " | ".join([f"{s}: {c}" for s, c in streams.items()])
                console.print(f"  • Streams: {stream_str}")
        console.print()
    
    def _display_top_bottom_performers(self, df: pd.DataFrame):
        """Display top and bottom performers side by side"""
        with_results = df[df['subject_count'] > 0].copy()
        
        if len(with_results) == 0:
            return
        
        console.print("\n" + "="*100)
        console.print("[bold cyan]🏆 PERFORMANCE HIGHLIGHTS[/bold cyan]")
        console.print("="*100 + "\n")
        
        # Top 10 performers
        console.print("[bold green]TOP 10 PERFORMERS[/bold green]\n")
        top_10 = with_results.nsmallest(10, 'pos')
        top_data = []
        
        for _, row in top_10.iterrows():
            # Get best 3 subjects
            subject_scores = []
            for subj_col in self.subject_columns:
                if pd.notna(row.get(subj_col)):
                    subj_short = self.subject_mapping.get(subj_col, subj_col)
                    subject_scores.append((subj_short, row[subj_col]))
            
            subject_scores.sort(key=lambda x: x[1], reverse=True)
            best_subjects = " | ".join([f"{s}:{int(m)}" for s, m in subject_scores[:3]])
            
            top_data.append([
                int(row['pos']),
                row['pupil_id'],
                row.get('full_name', 'N/A')[:25],
                row.get('sex', '-'),
                row.get('section_id', '-'),
                f"{row['avg_marks']:.1f}",
                row.get('avg_grade', '-'),
                best_subjects
            ])
        
        headers = ["Pos", "Pupil ID", "Full Name", "Sex", "Stream", "Avg", "Grade", "Best 3 Subjects"]
        print(tabulate(top_data, headers=headers, tablefmt="fancy_grid",
                      colalign=("center", "left", "left", "center", "center", "right", "center", "left")))
        
        # Bottom 10 performers
        console.print("\n[bold red]BOTTOM 10 PERFORMERS (Need Support)[/bold red]\n")
        bottom_10 = with_results.nlargest(10, 'pos')
        bottom_data = []
        
        for _, row in bottom_10.iterrows():
            # Get weakest 3 subjects
            subject_scores = []
            for subj_col in self.subject_columns:
                if pd.notna(row.get(subj_col)):
                    subj_short = self.subject_mapping.get(subj_col, subj_col)
                    subject_scores.append((subj_short, row[subj_col]))
            
            subject_scores.sort(key=lambda x: x[1])
            weak_subjects = " | ".join([f"{s}:{int(m)}" for s, m in subject_scores[:3]])
            
            bottom_data.append([
                int(row['pos']),
                row['pupil_id'],
                row.get('full_name', 'N/A')[:25],
                row.get('sex', '-'),
                row.get('section_id', '-'),
                f"{row['avg_marks']:.1f}",
                row.get('avg_grade', '-'),
                weak_subjects
            ])
        
        print(tabulate(bottom_data, headers=headers, tablefmt="fancy_grid",
                      colalign=("center", "left", "left", "center", "center", "right", "center", "left")))
        console.print()
    
    def get_grade(self, marks: float) -> str:
        """Get grade based on marks"""
        if pd.isna(marks) or not isinstance(marks, (int, float)):
            return ""
        
        marks = int(marks)
        
        if self.marks_style == "50 Marks":
            grade_row = self.grades_df[
                (self.grades_df['starting_50'] <= marks) & 
                (self.grades_df['ending_50'] >= marks)
            ]
        else:
            grade_row = self.grades_df[
                (self.grades_df['starting_100'] <= marks) & 
                (self.grades_df['ending_100'] >= marks)
            ]
        
        if not grade_row.empty:
            return grade_row.iloc[0]['grade']
        return ""
    
    def assign_basic_details(self, results_df: pd.DataFrame) -> pd.DataFrame:
        """Calculate totals, averages, and grades"""
        console.print("\n" + "="*80)
        console.print("[bold white]STEP 1: CALCULATING AVERAGES & GRADES[/bold white]")
        console.print("="*80 + "\n")
        
        df = results_df.copy()
        
        with console.status("[bold green]Computing totals and averages...", spinner="dots"):
            df['total_marks'] = df[self.subject_columns].sum(axis=1, skipna=True)
            df['subject_count'] = df[self.subject_columns].notna().sum(axis=1)
            df['avg_marks'] = df.apply(
                lambda row: row['total_marks'] / row['subject_count'] 
                if row['subject_count'] > 0 else 0,
                axis=1
            )
            df['out_of'] = (df['subject_count'] > 0).sum()
        
        console.print("[green]✓[/green] Totals and averages calculated")
        
        # Calculate grades
        console.print("\n[yellow]📊 Assigning grades...[/yellow]")
        with tqdm(total=len(self.subject_columns) + 1, 
                 desc=f"{Fore.YELLOW}Grading",
                 bar_format='{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}]',
                 colour='yellow') as pbar:
            
            for subject in self.subject_columns:
                grade_col = f"{subject}_grade"
                df[grade_col] = df[subject].apply(self.get_grade)
                pbar.update(1)
            
            df['avg_grade'] = df['avg_marks'].apply(self.get_grade)
            pbar.update(1)
        
        console.print("\n[green]✅ All grades assigned successfully![/green]\n")
        return df
    
    def assign_student_positions(self, results_df: pd.DataFrame) -> pd.DataFrame:
        """Rank students based on average marks"""
        console.print("="*80)
        console.print("[bold white]STEP 2: ASSIGNING STUDENT POSITIONS[/bold white]")
        console.print("="*80 + "\n")
        
        df = results_df.copy()
        has_results = df['subject_count'] > 0
        
        console.print("[yellow]🏆 Overall ranking...[/yellow]")
        df.loc[has_results, 'pos'] = df.loc[has_results, 'avg_marks'].rank(
            method='min', ascending=False
        ).astype('Int64')
        df.loc[has_results, 'out_of'] = has_results.sum()
        console.print(f"[green]✓[/green] Ranked {has_results.sum()} students overall")
        
        if 'sex' in df.columns:
            console.print("\n[yellow]👥 Gender-based ranking...[/yellow]")
            sex_counts = {}
            for sex in df['sex'].dropna().unique():
                mask = has_results & (df['sex'] == sex)
                if mask.any():
                    df.loc[mask, 'pos_sex'] = df.loc[mask, 'avg_marks'].rank(
                        method='min', ascending=False
                    ).astype('Int64')
                    df.loc[mask, 'out_of_sex'] = mask.sum()
                    sex_counts[sex] = mask.sum()
            
            for sex, count in sex_counts.items():
                console.print(f"[green]✓[/green] Ranked {count} students ({sex})")
        
        if 'section_id' in df.columns:
            console.print("\n[yellow]🎓 Stream-based ranking...[/yellow]")
            stream_counts = {}
            for stream in df['section_id'].dropna().unique():
                mask = has_results & (df['section_id'] == stream)
                if mask.any():
                    df.loc[mask, 'pos_stream'] = df.loc[mask, 'avg_marks'].rank(
                        method='min', ascending=False
                    ).astype('Int64')
                    df.loc[mask, 'out_of_stream'] = mask.sum()
                    stream_counts[stream] = mask.sum()
            
            for stream, count in stream_counts.items():
                console.print(f"[green]✓[/green] Ranked {count} students (Stream: {stream})")
        
        console.print("\n[green]✅ Student positions assigned![/green]\n")
        return df
    
    def assign_subject_positions(self, results_df: pd.DataFrame) -> pd.DataFrame:
        """Rank students by individual subject performance"""
        console.print("="*80)
        console.print("[bold white]STEP 3: ASSIGNING SUBJECT POSITIONS[/bold white]")
        console.print("="*80 + "\n")
        
        df = results_df.copy()
        
        with tqdm(total=len(self.subject_columns), 
                 desc=f"{Fore.CYAN}Processing subjects",
                 bar_format='{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}]',
                 colour='cyan') as pbar:
            
            for subject in self.subject_columns:
                has_marks = df[subject].notna()
                if has_marks.any():
                    pos_col = f"{subject}_pos"
                    out_of_col = f"{subject}_out_of"
                    
                    df.loc[has_marks, pos_col] = df.loc[has_marks, subject].rank(
                        method='min', ascending=False
                    ).astype('Int64')
                    df.loc[has_marks, out_of_col] = has_marks.sum()
                    
                    if 'sex' in df.columns:
                        for sex in df['sex'].dropna().unique():
                            mask = has_marks & (df['sex'] == sex)
                            if mask.any():
                                df.loc[mask, f"{subject}_pos_sex"] = df.loc[mask, subject].rank(
                                    method='min', ascending=False
                                ).astype('Int64')
                                df.loc[mask, f"{subject}_sex_out_of"] = mask.sum()
                    
                    if 'section_id' in df.columns:
                        for stream in df['section_id'].dropna().unique():
                            mask = has_marks & (df['section_id'] == stream)
                            if mask.any():
                                df.loc[mask, f"{subject}_pos_stream"] = df.loc[mask, subject].rank(
                                    method='min', ascending=False
                                ).astype('Int64')
                                df.loc[mask, f"{subject}_stream_out_of"] = mask.sum()
                
                pbar.update(1)
        
        console.print("\n[green]✅ Subject positions assigned![/green]\n")
        return df
    

    def create_result_json(self, row: pd.Series) -> str:
        """Create JSON representation of exam results as a string"""
        
        # The root dictionary
        root = {}
        
        root['exam_id'] = row.get('exam_id', '')
        root['avg_marks'] = float(row['avg_marks']) if pd.notna(row['avg_marks']) else 0
        root['avg_grade'] = row.get('avg_grade', '')
        
        # Positions dictionary
        positions = {}
        
        overall = {}
        overall['pos'] = int(row['pos']) if pd.notna(row.get('pos')) else None
        overall['out_of'] = int(row['out_of']) if pd.notna(row.get('out_of')) else None
        positions['overall'] = overall
        
        if 'pos_sex' in row and pd.notna(row.get('pos_sex')):
            sex = {}
            sex['pos'] = int(row['pos_sex'])
            sex['out_of'] = int(row['out_of_sex']) if pd.notna(row.get('out_of_sex')) else None
            positions['sex'] = sex
        
        if 'pos_stream' in row and pd.notna(row.get('pos_stream')):
            stream = {}
            stream['pos'] = int(row['pos_stream'])
            stream['out_of'] = int(row['out_of_stream']) if pd.notna(row.get('out_of_stream')) else None
            positions['stream'] = stream
        
        root['position'] = positions
        
        # Subjects dictionary
        subjects = {}
        
        for subject_col in self.subject_columns:
            if pd.notna(row.get(subject_col)):
                subject_short = self.subject_mapping.get(subject_col, subject_col)
                
                subj_data = {}
                subj_data['marks'] = float(row[subject_col])
                subj_data['grade'] = row.get(f"{subject_col}_grade", '')
                
                subj_pos = {}
                
                if pd.notna(row.get(f"{subject_col}_pos")):
                    subj_overall = {}
                    subj_overall['pos'] = int(row[f"{subject_col}_pos"])
                    subj_overall['out_of'] = int(row[f"{subject_col}_out_of"])
                    subj_pos['overall'] = subj_overall
                
                if pd.notna(row.get(f"{subject_col}_pos_sex")):
                    subj_sex = {}
                    subj_sex['pos'] = int(row[f"{subject_col}_pos_sex"])
                    subj_sex['out_of'] = int(row[f"{subject_col}_sex_out_of"])
                    subj_pos['sex'] = subj_sex
                
                if pd.notna(row.get(f"{subject_col}_pos_stream")):
                    subj_stream = {}
                    subj_stream['pos'] = int(row[f"{subject_col}_pos_stream"])
                    subj_stream['out_of'] = int(row[f"{subject_col}_stream_out_of"])
                    subj_pos['stream'] = subj_stream
                
                subj_data['position'] = subj_pos
                subjects[subject_short] = subj_data
        
        root['subjects'] = subjects
        
        # Convert the final dictionary to a JSON string
        return json.dumps(root)
    
    def save_results(self, results_df: pd.DataFrame):
        """Save processed results back to database"""
        console.print("="*80)
        console.print("[bold white]STEP 5: SAVING TO DATABASE[/bold white]")
        console.print("="*80 + "\n")
        
        conn = self._get_connection()
        cursor = conn.cursor()
        
        # Check if result_json field exists in dataframe
        has_result_json = 'result_json' in results_df.columns
        
        # Fields that exist in database (NO avg_marks - it's calculated, NO sex position fields)
        # avg_marks is auto-calculated from total_marks/subject_count in Access
        update_fields = ['total_marks', 'subject_count', 'avg_grade',
                        'pos', 'out_of', 'pos_stream', 'out_of_stream']
        
        # Add result_json only if it exists
        if has_result_json:
            update_fields.append('result_json')
        
        # Add subject-related fields (NO sex position fields)
        for subject in self.subject_columns:
            update_fields.extend([
                f"{subject}_grade",
                f"{subject}_pos",
                f"{subject}_out_of",
                f"{subject}_pos_stream",
                f"{subject}_stream_out_of"
            ])
        
        success_count = 0
        error_count = 0
        
        with tqdm(total=len(results_df), 
                 desc=f"{Fore.GREEN}Updating database",
                 bar_format='{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}]',
                 colour='green') as pbar:
            
            for _, row in results_df.iterrows():
                # Only include fields that exist in the row AND in our update list
                set_parts = []
                values = []
                
                for field in update_fields:
                    if field in row.index and pd.notna(row[field]):
                        set_parts.append(f"[{field}] = ?")
                        values.append(row[field])
                
                if set_parts:
                    query = f"""
                    UPDATE tbl_pupil_exam_results
                    SET {', '.join(set_parts)}
                    WHERE exam_id = ? AND pupil_id = ?
                    """
                    values.extend([self.exam_id, row['pupil_id']])
                    
                    try:
                        cursor.execute(query, values)
                        success_count += 1
                    except Exception as e:
                        error_count += 1
                        if error_count <= 3:  # Only show first 3 errors
                            console.print(f"\n[red]Error updating {row['pupil_id']}: {str(e)}[/red]")
                
                pbar.update(1)
        
        conn.commit()
        conn.close()
        
        console.print(f"\n[green]✅ Database updated: {success_count} records saved successfully![/green]")
        if error_count > 0:
            console.print(f"[yellow]⚠ {error_count} records had errors[/yellow]\n")
        else:
            console.print()
    
    def _display_final_summary(self, results_df: pd.DataFrame):
        """Display beautiful final summary"""
        console.print("\n" + "="*100)
        console.print("[bold magenta]📊 FINAL SUMMARY & STATISTICS[/bold magenta]")
        console.print("="*100 + "\n")
        
        with_results = results_df[results_df['subject_count'] > 0]
        
        # Overall statistics table
        stats_data = [
            ["Total Students", str(len(results_df))],
            ["Students with Results", str(len(with_results))],
            ["Average Score (Class)", f"{with_results['avg_marks'].mean():.2f}"],
            ["Highest Score", f"{with_results['avg_marks'].max():.2f}"],
            ["Lowest Score", f"{with_results['avg_marks'].min():.2f}"],
            ["Median Score", f"{with_results['avg_marks'].median():.2f}"],
        ]
        
        print(tabulate(stats_data, headers=["Metric", "Value"], tablefmt="fancy_grid",
                      colalign=("left", "right")))
        
        # Grade distribution
        console.print("\n[bold cyan]📈 Grade Distribution[/bold cyan]\n")
        grade_dist = with_results['avg_grade'].value_counts().sort_index()
        grade_data = [[grade, count, f"{(count/len(with_results)*100):.1f}%"] 
                     for grade, count in grade_dist.items()]
        
        print(tabulate(grade_data, headers=["Grade", "Count", "Percentage"], 
                      tablefmt="fancy_grid", colalign=("center", "right", "right")))
        
        # Subject performance analysis
        console.print("\n[bold cyan]📚 Subject Performance Analysis[/bold cyan]\n")
        subject_data = []
        
        for subj_col in self.subject_columns:
            subj_marks = results_df[subj_col].dropna()
            if len(subj_marks) > 0:
                subj_name = self.subject_mapping.get(subj_col, subj_col)
                subject_data.append([
                    subj_name,
                    len(subj_marks),
                    f"{subj_marks.mean():.1f}",
                    f"{subj_marks.max():.0f}",
                    f"{subj_marks.min():.0f}",
                    f"{subj_marks.std():.1f}"
                ])
        
        print(tabulate(subject_data, 
                      headers=["Subject", "Students", "Avg", "Max", "Min", "Std Dev"],
                      tablefmt="fancy_grid",
                      colalign=("left", "right", "right", "right", "right", "right")))
        
        # Display top and bottom performers
        self._display_top_bottom_performers(results_df)
        
        # Sample JSON
        console.print("\n" + "="*100)
        console.print("[bold cyan]📄 SAMPLE RESULT JSON (Top Student)[/bold cyan]")
        console.print("="*100 + "\n")
        
        top_student = with_results.nsmallest(1, 'pos').iloc[0]
        sample_json = top_student['result_json'].replace("''", "'")
        sample_data = json.loads(sample_json)
        
        console.print(Panel(json.dumps(sample_data, indent=2), 
                          border_style="cyan",
                          padding=(1, 2)))
        
        console.print("\n")
    
    def complete_exam(self):
        """Complete all exam processing steps"""
        # Load results
        console.print("\n[cyan]📂 Loading exam results from database...[/cyan]")
        df = self._load_results()
        console.print(f"[green]✓[/green] Loaded [bold]{len(df)}[/bold] student records\n")
        
        # Display data preview
        self._display_data_preview(df)
        
        # Step 1: Basic details
        df = self.assign_basic_details(df)
        
        # Step 2: Student positions
        df = self.assign_student_positions(df)
        
        # Step 3: Subject positions
        df = self.assign_subject_positions(df)
        
        # Step 4: Create JSON
        console.print("="*80)
        console.print("[bold white]STEP 4: GENERATING RESULT JSON[/bold white]")
        console.print("="*80 + "\n")
        
        tqdm.pandas(desc=f"{Fore.MAGENTA}Creating JSON", 
                   bar_format='{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}]')
        
        df['result_json'] = df.progress_apply(
            lambda row: self.create_result_json(row),
            axis=1
        )
        
        console.print("\n[green]✅ JSON records generated![/green]\n")
        
        # Step 5: Save to database
        self.save_results(df)
        
        # Display final summary
        self._display_final_summary(df)
        
        console.print("[bold green]🎉 EXAM PROCESSING COMPLETED SUCCESSFULLY! 🎉[/bold green]\n")
        
        return df


# Main execution
if __name__ == "__main__":
    db_path = r"C:\Kiyabo App\backend\Kiyabo App Backend v3.0.0.accdb"
    exam_id = "ANN420251117"
    
    try:
        processor = ExamProcessor(db_path=db_path, exam_id=exam_id)
        result_df = processor.complete_exam()
        
    except Exception as e:
        console.print(f"\n[red]{'='*70}[/red]")
        console.print(f"[red]  ❌ ERROR OCCURRED[/red]")
        console.print(f"[red]{'='*70}[/red]")
        console.print(f"[red]{str(e)}[/red]\n")
        raise