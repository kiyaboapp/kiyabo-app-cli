import pandas as pd
import pyodbc
from tqdm import tqdm
import time
import argparse
import re
from datetime import datetime
from tabulate import tabulate

class Colors:
    RED = '\033[91m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    MAGENTA = '\033[95m'
    CYAN = '\033[96m'
    WHITE = '\033[97m'
    BOLD = '\033[1m'
    END = '\033[0m'

class ExamDataImporter:
    def __init__(self):
        self.conn = None
        self.cursor = None

    def print_header(self, text):
        print(f"\n{Colors.CYAN}{Colors.BOLD}{'='*80}{Colors.END}")
        print(f"{Colors.CYAN}{Colors.BOLD}{text:^80}{Colors.END}")
        print(f"{Colors.CYAN}{Colors.BOLD}{'='*80}{Colors.END}")

    def print_subheader(self, text):
        print(f"\n{Colors.YELLOW}{Colors.BOLD}{text}{Colors.END}")
        print(f"{Colors.YELLOW}{'-'*60}{Colors.END}")

    def is_numeric(self, value):
        try:
            float(value)
            return True
        except (ValueError, TypeError):
            return False

    def get_subject_ids(self, db_path, use_subject_shorts_instead=False, is_for_display=False, table_name="tbl_student_subjects", class_id=""):
        """Python conversion of VBA GetSubjectIDs function with curriculum handling"""

        def val(s):
            """VBA-like Val()"""
            if s is None:
                return 0
            match = re.match(r"[-+]?\d*\.?\d+", str(s).strip())
            return float(match.group()) if match else 0

        def is_numeric(s):
            """VBA-like IsNumeric()"""
            if s is None:
                return False
            try:
                float(str(s).strip())
                return True
            except ValueError:
                return False

        def vba_bool(b):
            """VBA Boolean: True=-1, False=0"""
            return -1 if b else 0

        try:
            conn_str = f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_path};"
            self.conn = pyodbc.connect(conn_str)
            self.cursor = self.conn.cursor()
            
            # Build SQL based on curriculum
            sql = f"SELECT subject_serial, subject_short, is_present, is_core, sorter, subject_short_display FROM {table_name} WHERE "
            
            # Form VI old curriculum: include subject 31 regardless of is_present, exclude 30 & 34
            if class_id == "VI" and datetime.now() < datetime(2026, 7, 1):
                sql += "(is_present=True OR subject_serial=31) AND subject_serial NOT IN (30,34)"
            else:
                # New curriculum or other classes: normal behavior
                sql += "is_present=True"
            
            self.cursor.execute(sql)
            results = self.cursor.fetchall()
            total_subjects = len(results)
            
            if total_subjects == 0:
                print(f"{Colors.RED}No subjects found!{Colors.END}")
                return []
            
            subject_data = [[None] * 4 for _ in range(total_subjects)]
            
            i = 0
            for row in results:
                subject_serial, subject_short, is_present, is_core, sorter, subject_short_display = row
                
                if use_subject_shorts_instead and is_for_display:
                    subject_data[i][0] = subject_short_display
                else:
                    subject_data[i][0] = subject_short if use_subject_shorts_instead else subject_serial
                
                subject_data[i][1] = is_present
                subject_data[i][2] = vba_bool(is_core if is_core is not None else False)
                subject_data[i][3] = sorter if sorter is not None else "Z"
                i += 1
            
            # EXACT VBA bubble sort
            for i in range(total_subjects - 1):
                for j in range(i + 1, total_subjects):
                    condition1 = subject_data[i][2] > subject_data[j][2]
                    condition2 = (
                        subject_data[i][2] == subject_data[j][2]
                        and is_numeric(subject_data[i][3])
                        and is_numeric(subject_data[j][3])
                        and val(subject_data[i][3]) > val(subject_data[j][3])
                    )
                    condition3 = (
                        subject_data[i][2] == subject_data[j][2]
                        and not is_numeric(subject_data[i][3])
                        and is_numeric(subject_data[j][3])
                    )
                    
                    if condition1 or condition2 or condition3:
                        temp = subject_data[i][:]
                        subject_data[i] = subject_data[j][:]
                        subject_data[j] = temp
            
            subject_ids = [subject_data[i][0] for i in range(total_subjects)]
            
            return subject_ids
            
        except Exception as e:
            print(f"{Colors.RED}Database Error: {str(e)}{Colors.END}")
            import traceback
            traceback.print_exc()
            return []
        finally:
            if getattr(self, "cursor", None):
                self.cursor.close()
            if getattr(self, "conn", None):
                self.conn.close()

    def get_class_from_exam_id(self, db_path, exam_id):
        """Extract class ID from exam record in database"""
        try:
            conn_str = f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_path};"
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()
            
            sql = "SELECT class_id FROM tbl_exams WHERE exam_id = ?"
            cursor.execute(sql, (exam_id,))
            result = cursor.fetchone()
            
            cursor.close()
            conn.close()
            
            if result:
                return result[0]
            return ""
            
        except Exception as e:
            # Try to extract class from exam_id format (e.g., "MID620260124" -> "VI")
            if exam_id and len(exam_id) >= 4:
                # Check if exam_id contains VI or V pattern
                if 'VI' in exam_id.upper():
                    extracted = 'VI'
                elif exam_id[3:5] == '62':  # Form VI pattern in your system
                    extracted = 'VI'
                elif exam_id[3:5] == '61':  # Form V pattern
                    extracted = 'V'
                else:
                    extracted = ''
                
                if extracted:
                    return extracted
            
            return ""

    def check_existing_results(self, db_path, exam_id):
        """Check if results already exist for this exam_id"""
        try:
            conn_str = f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_path};"
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()
            
            sql = "SELECT COUNT(*) FROM tbl_student_exam_results WHERE exam_id = ?"
            cursor.execute(sql, (exam_id,))
            result = cursor.fetchone()
            count = result[0] if result else 0
            
            cursor.close()
            conn.close()
            
            return count
            
        except Exception as e:
            print(f"{Colors.YELLOW}[DEBUG] Could not check existing results: {str(e)}{Colors.END}")
            return 0

    def delete_existing_results(self, db_path, exam_id):
        """Delete all existing results for this exam_id"""
        try:
            conn_str = f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_path};"
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()
            
            sql = "DELETE FROM tbl_student_exam_results WHERE exam_id = ?"
            cursor.execute(sql, (exam_id,))
            deleted_count = cursor.rowcount
            conn.commit()
            
            cursor.close()
            conn.close()
            
            return deleted_count
            
        except Exception as e:
            print(f"{Colors.RED}Error deleting existing results: {str(e)}{Colors.END}")
            return -1

    def show_modern_data_preview(self, df, subject_ids):
        """Modern table display for data preview"""
        self.print_subheader("📊 DATA PREVIEW")
        
        # Show basic info in a table
        info_data = [
            ["Total Students", len(df)],
            ["Total Subjects", len(subject_ids)],
            ["Excel Shape", f"{df.shape[0]} rows × {df.shape[1]} columns"],
            ["Data Columns", f"7 to {df.shape[1]-1}"]
        ]
        
        print(tabulate(info_data, 
                      headers=["Metric", "Value"], 
                      tablefmt="grid",
                      colalign=("left", "right")))

        # Show first 5 students in a modern table
        student_data = []
        for i in range(min(5, len(df))):
            student_id = str(df.iloc[i, 1]).split('.')[0]
            first_name = df.iloc[i, 2] or ""
            middle_name = df.iloc[i, 3] or ""
            last_name = df.iloc[i, 4] or ""
            full_name = f"{first_name} {middle_name} {last_name}".strip()
            comb = df.iloc[i, 6] or "N/A"
            
            student_data.append([
                i + 1,
                student_id,
                full_name[:25],
                comb
            ])
        
        print(f"\n{Colors.GREEN}First 5 Students:{Colors.END}")
        print(tabulate(student_data, 
                      headers=["#", "Student ID", "Full Name", "Combination"], 
                      tablefmt="grid",
                      colalign=("right", "left", "left", "center")))

        # Show subject mapping in a compact table
        mapping_data = []
        for i, subject in enumerate(subject_ids):
            marks_col = 7 + (i * 2)
            if marks_col < len(df.columns):
                sample = df.iloc[0, marks_col] if len(df) > 0 else "N/A"
                sample_str = f"{sample:.1f}" if sample != "N/A" and self.is_numeric(sample) else str(sample)
                mapping_data.append([
                    i + 1,
                    subject,
                    f"Column {marks_col}",
                    sample_str
                ])
        
        print(f"\n{Colors.GREEN}Subject Mapping:{Colors.END}")
        print(tabulate(mapping_data, 
                      headers=["#", "Subject", "Excel Position", "Sample"], 
                      tablefmt="grid",
                      colalign=("right", "center", "left", "center")))

    def show_modern_insert_preview(self, df, import_data, subject_ids):
        """Modern table display for insert preview - shows 10 records with ALL subjects, names, and combinations"""
        self.print_subheader("📋 INSERT PREVIEW - First 10 Records")
        
        # Prepare data for compact horizontal display
        table_data = []
        for i in range(min(10, len(import_data))):
            record = import_data[i]
            
            # Get name and combination from original DataFrame
            first_name = df.iloc[i, 2] or ""
            middle_name = df.iloc[i, 3] or ""
            last_name = df.iloc[i, 4] or ""
            full_name = f"{first_name} {middle_name} {last_name}".strip()[:20]
            comb = df.iloc[i, 6] or "N/A"
            
            row = [i + 1, record['student_id'], full_name, comb]
            
            # Add ALL subjects in the row
            for subject in subject_ids:
                value = record.get(subject)
                if value is not None:
                    # Color code based on marks
                    if value >= 80:
                        colored_value = f"{Colors.GREEN}{value:.0f}{Colors.END}"
                    elif value >= 60:
                        colored_value = f"{Colors.YELLOW}{value:.0f}{Colors.END}"
                    elif value >= 40:
                        colored_value = f"{Colors.BLUE}{value:.0f}{Colors.END}"
                    else:
                        colored_value = f"{Colors.RED}{value:.0f}{Colors.END}"
                    row.append(colored_value)
                else:
                    row.append(f"{Colors.RED}—{Colors.END}")
            
            table_data.append(row)
        
        # Create headers with colors
        headers = [
            f"{Colors.CYAN}#{Colors.END}", 
            f"{Colors.CYAN}Student ID{Colors.END}",
            f"{Colors.CYAN}Name{Colors.END}",
            f"{Colors.CYAN}Comb{Colors.END}"
        ]
        for subject in subject_ids:
            headers.append(f"{Colors.MAGENTA}{subject}{Colors.END}")
        
        # Display the table
        print(tabulate(table_data, 
                      headers=headers, 
                      tablefmt="simple_grid",
                      colalign=("right", "left", "left", "center") + ("center",) * len(subject_ids)))
        
        print(f"\n{Colors.BLUE}Showing {len(table_data)} of {len(import_data)} records with {len(subject_ids)} subjects each{Colors.END}")

    def show_modern_statistics(self, import_data, subject_ids):
        """Modern table display for statistics"""
        self.print_subheader("📈 IMPORT STATISTICS")
        
        import_df = pd.DataFrame(import_data)
        stats_data = []
        total_marks = 0
        marks_count = 0
        
        for subject in subject_ids:
            marks = import_df[subject].dropna()
            if len(marks) > 0:
                avg = marks.mean()
                max_val = marks.max()
                min_val = marks.min()
                count = len(marks)
                total_marks += marks.sum()
                marks_count += count
                
                # Color code average
                if avg >= 80:
                    avg_color = Colors.GREEN
                elif avg >= 60:
                    avg_color = Colors.YELLOW
                elif avg >= 40:
                    avg_color = Colors.BLUE
                else:
                    avg_color = Colors.RED
                
                stats_data.append([
                    subject,
                    count,
                    f"{avg_color}{avg:5.1f}{Colors.END}",
                    f"{max_val:.1f}",
                    f"{min_val:.1f}"
                ])
        
        print(tabulate(stats_data, 
                      headers=["Subject", "Count", "Average", "Max", "Min"], 
                      tablefmt="grid",
                      colalign=("center", "center", "center", "center", "center")))
        
        if marks_count > 0:
            overall_avg = total_marks / marks_count
            overall_color = Colors.GREEN if overall_avg >= 60 else Colors.YELLOW if overall_avg >= 40 else Colors.RED
            print(f"\n{Colors.CYAN}Overall average: {overall_color}{overall_avg:.2f}{Colors.CYAN} | Total marks entered: {Colors.WHITE}{marks_count}{Colors.END}")

    def import_exam_data(self, exam_id, excel_path, db_path):
        """Main import function with curriculum handling"""
        try:
            self.print_header("EXAM DATA IMPORT")
            print(f"{Colors.WHITE}Exam ID: {exam_id}{Colors.END}")
            print(f"{Colors.WHITE}Excel: {excel_path}{Colors.END}")
            print(f"{Colors.WHITE}Database: {db_path}{Colors.END}")

            # Check for existing results
            self.print_subheader("CHECKING EXISTING RESULTS")
            existing_count = self.check_existing_results(db_path, exam_id)
            
            if existing_count > 0:
                print(f"{Colors.RED}⚠ WARNING: Found {existing_count} existing records for exam ID '{exam_id}'{Colors.END}")
                print(f"{Colors.YELLOW}These records will need to be deleted before importing new data.{Colors.END}")
                print(f"\n{Colors.CYAN}Options:{Colors.END}")
                print(f"  {Colors.GREEN}[D]{Colors.END} Delete existing records and continue with import")
                print(f"  {Colors.RED}[Q]{Colors.END} Quit without importing")
                
                choice = input(f"\n{Colors.BOLD}Your choice (D/Q): {Colors.END}").strip().upper()
                
                if choice == 'Q':
                    print(f"\n{Colors.YELLOW}Import cancelled by user.{Colors.END}")
                    return False
                elif choice == 'D':
                    print(f"\n{Colors.YELLOW}Deleting existing records...{Colors.END}")
                    deleted = self.delete_existing_results(db_path, exam_id)
                    if deleted >= 0:
                        print(f"{Colors.GREEN}✓ Deleted {deleted} existing records{Colors.END}")
                    else:
                        print(f"{Colors.RED}✗ Failed to delete existing records{Colors.END}")
                        return False
                else:
                    print(f"\n{Colors.RED}Invalid choice. Import cancelled.{Colors.END}")
                    return False
            else:
                print(f"{Colors.GREEN}✓ No existing records found for this exam ID{Colors.END}")

            # Get class ID from exam
            self.print_subheader("DETERMINING CLASS/CURRICULUM")
            class_id = self.get_class_from_exam_id(db_path, exam_id)
            if class_id:
                print(f"{Colors.GREEN}✓ Class: {class_id}{Colors.END}")
                if class_id == "VI" and datetime.now() < datetime(2026, 7, 1):
                    print(f"{Colors.MAGENTA}✓ Old curriculum mode (Form VI before July 2026){Colors.END}")
            else:
                print(f"{Colors.YELLOW}⚠ Class not detected - using standard curriculum{Colors.END}")

            # Read Excel
            self.print_subheader("READING EXCEL FILE")
            df = pd.read_excel(excel_path, header=None, skiprows=13, engine='openpyxl')
            print(f"{Colors.GREEN}✓ Loaded {len(df)} student records{Colors.END}")

            # Get subjects with curriculum handling
            self.print_subheader("GETTING SUBJECTS FROM DATABASE")
            subject_ids = self.get_subject_ids(db_path, use_subject_shorts_instead=True, class_id=class_id)
            if not subject_ids:
                return False
            print(f"{Colors.GREEN}✓ Subjects ({len(subject_ids)}): {subject_ids}{Colors.END}")

            # Show modern preview
            self.show_modern_data_preview(df, subject_ids)

            # Prepare data
            self.print_subheader("PREPARING IMPORT DATA")
            import_data = []
            for idx in range(len(df)):
                record = {
                    'student_id': str(df.iloc[idx, 1]).split('.')[0],
                    'exam_id': str(exam_id)
                }
                
                for i, subject_id in enumerate(subject_ids):
                    marks_col = 7 + (i * 2)
                    if marks_col < len(df.columns):
                        marks_value = df.iloc[idx, marks_col]
                        record[subject_id] = float(marks_value) if pd.notna(marks_value) else None
                
                import_data.append(record)

            # Show modern insert preview
            self.show_modern_insert_preview(df, import_data, subject_ids)

            # Auto proceed
            print(f"\n{Colors.GREEN}🚀 Auto-proceeding with import...{Colors.END}")

            # Import to database
            self.print_subheader("IMPORTING TO DATABASE")
            conn_str = f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_path};"
            self.conn = pyodbc.connect(conn_str)
            self.cursor = self.conn.cursor()

            success_count = 0
            for record in tqdm(import_data, desc="Importing", ncols=70):
                try:
                    insert_data = {'student_id': record['student_id'], 'exam_id': record['exam_id']}
                    
                    for subject_id in subject_ids:
                        insert_data[subject_id] = record.get(subject_id)
                    
                    columns = ', '.join([f'[{col}]' for col in insert_data.keys()])
                    placeholders = ', '.join(['?' for _ in insert_data])
                    sql = f"INSERT INTO tbl_student_exam_results ({columns}) VALUES ({placeholders})"
                    
                    self.cursor.execute(sql, list(insert_data.values()))
                    success_count += 1
                    
                except Exception as e:
                    continue

            self.conn.commit()

            # Show modern results
            self.print_header("IMPORT COMPLETE")
            print(f"{Colors.GREEN}✓ Successfully imported: {success_count}/{len(import_data)} records{Colors.END}")
            
            # Show modern statistics
            self.show_modern_statistics(import_data, subject_ids)

            return True

        except Exception as e:
            print(f"{Colors.RED}Import failed: {str(e)}{Colors.END}")
            import traceback
            traceback.print_exc()
            return False
        finally:
            if self.cursor:
                self.cursor.close()
            if self.conn:
                self.conn.close()

def main():
    parser = argparse.ArgumentParser(description='Import exam data from Excel to Access database')
    parser.add_argument('--exam_id', required=True, help='Exam ID')
    parser.add_argument('--excel_path', required=True, help='Path to Excel file')
    parser.add_argument('--db_path', required=True, help='Path to Access database')
    
    args = parser.parse_args()
    
    importer = ExamDataImporter()
    success = importer.import_exam_data(args.exam_id, args.excel_path, args.db_path)
    
    if success:
        print(f"\n{Colors.GREEN}{Colors.BOLD}COMPLETED SUCCESSFULLY{Colors.END}")
        # process_exam(args.exam_id, args.db_path)
    else:
        print(f"\n{Colors.RED}{Colors.BOLD}FAILED{Colors.END}")

if __name__ == "__main__":
    main()