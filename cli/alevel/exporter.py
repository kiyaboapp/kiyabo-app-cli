"""
Student Exam Results Exporter - Enhanced Version
Python implementation using pywin32 for Access database and Excel automation
"""

import argparse
from datetime import datetime
import time
import win32com.client
import pythoncom
import pandas as pd
from pathlib import Path
from typing import Dict, List, Tuple, Optional
from tqdm import tqdm
import sys
import os
import traceback

# Color codes for beautiful console output
class Colors:
    RED = '\033[91m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    MAGENTA = '\033[95m'
    CYAN = '\033[96m'
    WHITE = '\033[97m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    END = '\033[0m'




def print_header(text):
    """Print a beautiful header"""
    width = 80
    print(f"\n{Colors.CYAN}{Colors.BOLD}{'=' * width}{Colors.END}")
    print(f"{Colors.CYAN}{Colors.BOLD}{text:^{width}}{Colors.END}")
    print(f"{Colors.CYAN}{Colors.BOLD}{'=' * width}{Colors.END}")

def print_success(text):
    """Print success message"""
    print(f"{Colors.GREEN}{Colors.BOLD}✓ {text}{Colors.END}")

def print_error(text):
    """Print error message"""
    print(f"{Colors.RED}{Colors.BOLD}✗ {text}{Colors.END}")

def print_warning(text):
    """Print warning message"""
    print(f"{Colors.YELLOW}{Colors.BOLD}⚠ {text}{Colors.END}")

def print_info(text):
    """Print info message"""
    print(f"{Colors.BLUE}{Colors.BOLD}ℹ {text}{Colors.END}")

def safe_print_table(title, data, headers=None, max_rows=15):
    """Safely print table without modifying data - show more records for end users"""
    try:
        if data is None or len(data) == 0:
            print_warning(f"No data available for {title}")
            return
        
        print_header(f"📊 {title}")
        
        # Convert to simple list format for safe display
        if headers is None and len(data) > 0:
            headers = [f"{i+1}" for i in range(len(data[0]))]
        
        # Display comprehensive info
        total_rows = len(data)
        total_cols = len(headers) if headers else (len(data[0]) if data else 0)
        print(f"{Colors.MAGENTA}📈 Total Records: {total_rows:,} | Columns: {total_cols}{Colors.END}")
        
        # Display first 15+ rows for better user experience
        display_data = data[:max_rows]
        
        if not display_data:
            return
            
        # Calculate column widths - more generous for better readability
        col_widths = []
        for i in range(len(display_data[0])):
            max_content = max(len(str(row[i])) for row in display_data) if display_data else 0
            header_len = len(headers[i]) if headers and i < len(headers) else 10
            # Allow wider columns for better data visibility
            col_widths.append(min(max(max_content, header_len), 35))
        
        # Print top border
        border_color = Colors.CYAN
        header_color = Colors.YELLOW + Colors.BOLD
        data_color = Colors.WHITE
        
        top_border = border_color + "┌"
        for i, width in enumerate(col_widths):
            top_border += "─" * (width + 2)
            if i < len(col_widths) - 1:
                top_border += "┬"
        top_border += "┐" + Colors.END
        print(top_border)
        
        # Print header if available
        if headers:
            header_line = border_color + "│" + Colors.END
            for i, (header, width) in enumerate(zip(headers, col_widths)):
                header_line += header_color + f" {str(header)[:width]:<{width}} " + Colors.END
                header_line += border_color + "│" + Colors.END
            print(header_line)
            
            # Print separator
            sep_line = border_color + "├"
            for i, width in enumerate(col_widths):
                sep_line += "─" * (width + 2)
                if i < len(col_widths) - 1:
                    sep_line += "┼"
            sep_line += "┤" + Colors.END
            print(sep_line)
        
        # Print data rows with alternating colors for better readability
        for idx, row in enumerate(display_data):
            data_line = border_color + "│" + Colors.END
            row_color = data_color if idx % 2 == 0 else Colors.CYAN
            for i, (value, width) in enumerate(zip(row, col_widths)):
                cell_value = str(value)[:width] if value is not None else ""
                data_line += row_color + f" {cell_value:<{width}} " + Colors.END
                data_line += border_color + "│" + Colors.END
            print(data_line)
        
        # Print bottom border
        bottom_border = border_color + "└"
        for i, width in enumerate(col_widths):
            bottom_border += "─" * (width + 2)
            if i < len(col_widths) - 1:
                bottom_border += "┴"
            bottom_border += "┘" + Colors.END
        print(bottom_border)
        
        # Show comprehensive truncation info
        if len(data) > max_rows:
            remaining = len(data) - max_rows
            print(f"{Colors.MAGENTA}📋 Displaying {max_rows} of {total_rows:,} records ({remaining:,} more records available){Colors.END}")
            
    except Exception as e:
        print_warning(f"Could not display table for {title}: {str(e)}")
        # Just show basic info without fancy formatting
        print(f"{Colors.MAGENTA}📦 Data Summary: {len(data):,} rows × {len(data[0]) if data else 0} columns{Colors.END}")

def print_students_table(title, students_df, max_rows=15):
    """Specialized function to display student data in a beautiful format"""
    try:
        if students_df.empty:
            print_warning(f"No student data available for {title}")
            return
        
        print_header(f"🎓 {title}")
        
        # Display comprehensive student statistics
        total_students = len(students_df)
        valid_students = len(students_df[students_df['avg_marks'] >= 0])
        avg_marks = students_df['avg_marks'].mean()
        max_marks = students_df['avg_marks'].max()
        min_marks = students_df['avg_marks'].min()
        
        print(f"{Colors.MAGENTA}📊 Student Statistics: {Colors.END}")
        print(f"{Colors.MAGENTA}   • Total Students: {total_students:,}{Colors.END}")
        print(f"{Colors.MAGENTA}   • Valid Records: {valid_students:,}{Colors.END}")
        print(f"{Colors.MAGENTA}   • Average Marks: {avg_marks:.2f}{Colors.END}")
        print(f"{Colors.MAGENTA}   • Highest Marks: {max_marks:.2f}{Colors.END}")
        print(f"{Colors.MAGENTA}   • Lowest Marks: {min_marks:.2f}{Colors.END}")
        
        # Prepare display data - show most relevant columns
        display_df = students_df.head(max_rows)
        display_data = []
        
        for _, student in display_df.iterrows():
            display_data.append([
                student['student_id'],
                student['first_name'] or '',
                student['surname'] or '',
                student['sex'] or '',
                student['comb_id'] or '',
                f"{student['avg_marks']:.2f}" if pd.notna(student['avg_marks']) else '-',
                student['division'] or '-',
                f"{student['points']:.1f}" if pd.notna(student['points']) else '-'
            ])
        
        headers = ["Student ID", "First Name", "Surname", "Sex", "Comb", "Avg Marks", "Division", "Points"]
        
        # Calculate column widths
        col_widths = [12, 12, 12, 6, 8, 10, 8, 8]
        
        # Print table
        border_color = Colors.CYAN
        header_color = Colors.YELLOW + Colors.BOLD
        
        top_border = border_color + "┌"
        for i, width in enumerate(col_widths):
            top_border += "─" * (width + 2)
            if i < len(col_widths) - 1:
                top_border += "┬"
        top_border += "┐" + Colors.END
        print(top_border)
        
        # Header
        header_line = border_color + "│" + Colors.END
        for i, (header, width) in enumerate(zip(headers, col_widths)):
            header_line += header_color + f" {str(header)[:width]:<{width}} " + Colors.END
            header_line += border_color + "│" + Colors.END
        print(header_line)
        
        # Separator
        sep_line = border_color + "├"
        for i, width in enumerate(col_widths):
            sep_line += "─" * (width + 2)
            if i < len(col_widths) - 1:
                sep_line += "┼"
        sep_line += "┤" + Colors.END
        print(sep_line)
        
        # Data rows with alternating colors
        for idx, row in enumerate(display_data):
            data_line = border_color + "│" + Colors.END
            row_color = Colors.WHITE if idx % 2 == 0 else Colors.CYAN
            for i, (value, width) in enumerate(zip(row, col_widths)):
                data_line += row_color + f" {str(value)[:width]:<{width}} " + Colors.END
                data_line += border_color + "│" + Colors.END
            print(data_line)
        
        # Bottom border
        bottom_border = border_color + "└"
        for i, width in enumerate(col_widths):
            bottom_border += "─" * (width + 2)
            if i < len(col_widths) - 1:
                bottom_border += "┴"
        bottom_border += "┘" + Colors.END
        print(bottom_border)
        
        # Show truncation info
        if len(students_df) > max_rows:
            remaining = len(students_df) - max_rows
            print(f"{Colors.MAGENTA}🎯 Displaying {max_rows} of {total_students:,} students ({remaining:,} more students in database){Colors.END}")
            
    except Exception as e:
        print_warning(f"Could not display student table for {title}: {str(e)}")
        print(f"{Colors.MAGENTA}📦 Student Data: {len(students_df):,} students loaded{Colors.END}")

def print_subjects_table(title, subjects_df):
    """Specialized function to display subjects in a beautiful format"""
    try:
        if subjects_df.empty:
            print_warning(f"No subject data available for {title}")
            return
        
        print_header(f"📚 {title}")
        
        total_subjects = len(subjects_df)
        print(f"{Colors.MAGENTA}📖 Total Subjects: {total_subjects:,}{Colors.END}")
        
        # Display all subjects in a compact format
        display_data = []
        for _, subject in subjects_df.iterrows():
            display_data.append([
                subject['subject_serial'],
                subject['subject_name'],
                subject['subject_short'],
                subject['subject_user_short']
            ])
        
        headers = ["ID", "Subject Name", "Code", "Display Code"]
        col_widths = [6, 20, 8, 12]
        
        border_color = Colors.CYAN
        header_color = Colors.YELLOW + Colors.BOLD
        
        top_border = border_color + "┌"
        for i, width in enumerate(col_widths):
            top_border += "─" * (width + 2)
            if i < len(col_widths) - 1:
                top_border += "┬"
        top_border += "┐" + Colors.END
        print(top_border)
        
        # Header
        header_line = border_color + "│" + Colors.END
        for i, (header, width) in enumerate(zip(headers, col_widths)):
            header_line += header_color + f" {str(header)[:width]:<{width}} " + Colors.END
            header_line += border_color + "│" + Colors.END
        print(header_line)
        
        # Separator
        sep_line = border_color + "├"
        for i, width in enumerate(col_widths):
            sep_line += "─" * (width + 2)
            if i < len(col_widths) - 1:
                sep_line += "┼"
        sep_line += "┤" + Colors.END
        print(sep_line)
        
        # Data rows
        for idx, row in enumerate(display_data):
            data_line = border_color + "│" + Colors.END
            row_color = Colors.WHITE if idx % 2 == 0 else Colors.CYAN
            for i, (value, width) in enumerate(zip(row, col_widths)):
                data_line += row_color + f" {str(value)[:width]:<{width}} " + Colors.END
                data_line += border_color + "│" + Colors.END
            print(data_line)
        
        # Bottom border
        bottom_border = border_color + "└"
        for i, width in enumerate(col_widths):
            bottom_border += "─" * (width + 2)
            if i < len(col_widths) - 1:
                bottom_border += "┴"
        bottom_border += "┘" + Colors.END
        print(bottom_border)
        
    except Exception as e:
        print_warning(f"Could not display subjects table: {str(e)}")

class StudentExamExporter:
    """Exports student exam results from Access database to Excel with rankings and formatting"""

    # Fixed sorting rules - EXACT column names from database
    SORT_ASCENDING = {
        'points': True,              # Lower is better
        'avg_marks': False,          # Higher is better
        'subject_count_all': False   # Higher is better (ALWAYS USE THIS!)
    }
    
    def __init__(self, exam_id: str, db_path: str = r"C:\Kiyabo App\backend\Kiyabo App Backend v4.0.0.accdb", 
                include_comb_sheets: bool = False, order_by: str = "position",
                top_n: int = 10, bottom_n: int = 10, 
                paper_size: str = "A4", orientation: str = None,
                sort_columns: list = None):
        """
        Initialize the exporter
        
        Args:
            exam_id: Exam identifier
            db_path: Path to Access database
            include_comb_sheets: If True, create separate sheets for each combination
            order_by: "position" (default) or "name" or "sex_name"
            top_n: Number of top students to show in summary table (default 10)
            bottom_n: Number of bottom students to show in summary table (default 10)
            paper_size: Paper size - "A4", "LTR" (Letter), or "A3" (default "A4")
            orientation: Page orientation - "portrait" or "landscape" (auto-detect if None)
            sort_columns: List of EXACT database column names to sort by
                        (default: ["points", "avg_marks", "subject_count_all"])
                        
                        VALID COLUMNS:
                        - "points"             (lower is better)
                        - "avg_marks"          (higher is better)
                        - "subject_count_all"  (higher is better) ⚠️ USE THIS ONE since priotize even student who optionally seect subject for hobby
                        
                        Examples:
                        - ["points", "avg_marks", "subject_count_all"]  # Default
                        - ["avg_marks", "points", "subject_count_all"]  # Prioritize avg_marks
                        - ["points", "avg_marks"]                        # Just these two
        """
        self.exam_id = exam_id
        self.db_path = db_path
        self.include_comb_sheets = include_comb_sheets
        self.order_by = order_by.lower()
        self.top_n = top_n
        self.bottom_n = bottom_n
        self.paper_size = paper_size.upper()
        self.orientation = orientation.lower() if orientation else None
        
        # NEW: Configurable sort columns - DEFAULT USES subject_count_all
        self.sort_columns = sort_columns or ["avg_marks","points", "subject_count_all"]
        
        # Validate ALL columns are valid
        for col in self.sort_columns:
            if col not in self.SORT_ASCENDING:
                valid_cols = list(self.SORT_ASCENDING.keys())
                raise ValueError(
                    f"Invalid sort column: '{col}'\n"
                    f"Valid columns are: {valid_cols}\n"
                    f"⚠️  Use 'subject_count_all' (NOT 'subject_count')"
                )
        
        # Build ascending list based on fixed rules
        self.sort_ascending = [self.SORT_ASCENDING[col] for col in self.sort_columns]
        
        self.start_row = None
        self.excel_app = None
        self.workbook = None
        self.worksheet = None
        self.conn = None

        # 🔍 Print all arguments for debugging
        print("\n" + "=" * 80)
        print("🎯 StudentExamExporter Initialized with Arguments:")
        print("=" * 80)
        for key, value in {
            "exam_id": exam_id,
            "db_path": db_path,
            "include_comb_sheets": include_comb_sheets,
            "order_by": order_by,
            "top_n": top_n,
            "bottom_n": bottom_n,
            "paper_size": paper_size,
            "orientation": orientation,
            "sort_columns": self.sort_columns,
            "sort_ascending": self.sort_ascending,
        }.items():
            print(f"• {key:<20}: {value}")
        print("=" * 80 + "\n")
        
        # Validate inputs
        if self.order_by not in ["position", "name", "sex_name"]:
            raise ValueError(f"order_by must be 'position', 'name', or 'sex_name', got '{order_by}'")
        
        if self.paper_size not in ["A4", "LTR", "A3"]:
            raise ValueError(f"paper_size must be 'A4', 'LTR', or 'A3', got '{paper_size}'")
        
        if self.orientation and self.orientation not in ["portrait", "landscape"]:
            raise ValueError(f"orientation must be 'portrait' or 'landscape', got '{orientation}'")
        
        print_header("🎓 STUDENT EXAM EXPORTER INITIALIZATION")
        print_info(f"📝 Exam ID: {exam_id}")
        print_info(f"🗃️ Database: {db_path}")
        print_info(f"📑 Include combination sheets: {include_comb_sheets}")
        print_info(f"🔢 Order by: {self.order_by}")
        print_info(f"🏆 Top/Bottom N: {self.top_n}/{self.bottom_n}")
        print_info(f"📄 Paper: {self.paper_size}, Orientation: {self.orientation or 'auto'}")
        print_info(f"📊 Sort by: {' → '.join(self.sort_columns)}")
        print_info(f"📈 Sort order: {['ASC' if asc else 'DESC' for asc in self.sort_ascending]}")

        
    def list_tables(self):
        """Helper function to list all tables in the database for debugging"""
        print_header("🗃️ DATABASE TABLES INSPECTION")
        
        if not self.conn:
            self._connect_db()
        
        try:
            catalog = win32com.client.Dispatch("ADOX.Catalog")
            catalog.ActiveConnection = self.conn
            
            tables = []
            for table in catalog.Tables:
                table_name = table.Name
                table_type = table.Type
                if not table_name.startswith("MSys") and table_type == "TABLE":
                    tables.append(table_name)
            
            print_info(f"📊 Found {len(tables):,} user tables in database")
            
            # Display comprehensive tables list
            table_data = [(idx, table_name) for idx, table_name in enumerate(sorted(tables), 1)]
            safe_print_table("ALL DATABASE TABLES", table_data, headers=["#", "Table Name"], max_rows=20)
            
            # Show tables starting with 'tbl_' with more details
            tbl_tables = [t for t in tables if t.startswith("tbl_")]
            tbl_data = [(idx, table_name) for idx, table_name in enumerate(sorted(tbl_tables), 1)]
            safe_print_table("STUDENT-RELATED TABLES (tbl_*)", tbl_data, headers=["#", "Table Name"], max_rows=20)
                
            return tables
            
        except Exception as e:
            print_error(f"Failed to list tables: {str(e)}")
            return []
    
    def inspect_table_structure(self, table_name: str):
        """Helper function to show table structure"""
        print_header(f"🔍 TABLE STRUCTURE: {table_name}")
        
        if not self.conn:
            self._connect_db()
        
        try:
            sql = f"SELECT TOP 1 * FROM {table_name}"
            rs = win32com.client.Dispatch("ADODB.Recordset")
            rs.Open(sql, self.conn, 1, 3)
            
            field_info = []
            for i in range(rs.Fields.Count):
                field = rs.Fields(i)
                field_info.append({
                    'name': field.Name,
                    'type': field.Type,
                    'size': field.DefinedSize
                })
            
            # Display comprehensive field information
            field_data = [(idx, field['name'], field['type'], field['size']) 
                         for idx, field in enumerate(field_info, 1)]
            safe_print_table(f"FIELDS IN {table_name.upper()}", field_data, 
                           headers=["#", "Field Name", "Data Type", "Size"], max_rows=20)
            
            rs.Close()
            return field_info
            
        except Exception as e:
            print_error(f"Failed to inspect table {table_name}: {str(e)}")
            return []
        
    def open_excel_file(self, file_path: str, start_row: int = 1):
        """
        Open or create Excel file and set worksheet
        
        Args:
            file_path: Path to Excel file
            start_row: Starting row number for data
        """
        self.start_row = start_row
        print_header("📂 EXCEL FILE SETUP")
        print_info(f"📁 File path: {file_path}")
        print_info(f"📍 Start row: {start_row}")
        
        pythoncom.CoInitialize()
        self.excel_app = win32com.client.Dispatch("Excel.Application")
        self.excel_app.Visible = True
        self.excel_app.DisplayAlerts = False
        
        if Path(file_path).exists():
            print_success("📖 Opening existing Excel file")
            self.workbook = self.excel_app.Workbooks.Open(file_path)
        else:
            print_success("🆕 Creating new Excel file")
            self.workbook = self.excel_app.Workbooks.Add()
            
        self.worksheet = self.workbook.Worksheets(1)
        print_success("✅ Worksheet ready and configured")
        
        self._connect_db()
        
    def _connect_db(self):
        """Establish connection to Access database"""
        print_header("🗄️ DATABASE CONNECTION")
        print_info(f"🔗 Database path: {self.db_path}")
        
        if not Path(self.db_path).exists():
            raise FileNotFoundError(f"Database file not found: {self.db_path}")
        
        conn_str = (
            r'Provider=Microsoft.ACE.OLEDB.12.0;'
            f'Data Source={self.db_path};'
        )
        
        try:
            self.conn = win32com.client.Dispatch("ADODB.Connection")
            self.conn.Open(conn_str)
            print_success("✅ Database connection established successfully")
        except Exception as e:
            print_error(f"Database connection failed: {str(e)}")
            raise
        
    def _execute_query(self, sql: str, query_name: str = "query") -> List[Tuple]:
        """Execute SQL query and return results - ORIGINAL LOGIC PRESERVED"""
        print_info(f"🔄 Executing {query_name}...")

        try:
            rs = win32com.client.Dispatch("ADODB.Recordset")
            rs.Open(sql, self.conn, 1, 3)
            
            results = []
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    row = tuple(rs.Fields(i).Value for i in range(rs.Fields.Count))
                    results.append(row)
                    rs.MoveNext()
            
            rs.Close()
            print_success(f"✅ {query_name}: {len(results):,} rows returned")
            
            # Enhanced preview with more records
            if results and len(results) > 0:
                preview_data = results[:15]  # Show 15 rows for better user experience
                headers = [f"{i+1}" for i in range(len(results[0]))]
                safe_print_table(f"QUERY RESULTS: {query_name.upper()}", preview_data, headers=headers, max_rows=15)
            
            return results
            
        except Exception as e:
            print_error(f"Query failed ({query_name}): {str(e)}")
            print_error(f"SQL: {sql[:200]}...")
            raise
    
    def _get_exam_info(self) -> Dict:
        """Get exam information - ORIGINAL LOGIC PRESERVED"""
        sql = f"""
        SELECT exam_id, exam_name, exam_start, exam_end 
        FROM tbl_student_exams 
        WHERE exam_id = '{self.exam_id}'
        """
        results = self._execute_query(sql, "exam_info")
        
        if not results:
            raise ValueError(f"Exam {self.exam_id} not found in tbl_student_exams!")
            
        exam_info = {
            'exam_id': results[0][0],
            'exam_name': results[0][1],
            'exam_start': results[0][2],
            'exam_end': results[0][3]
        }
        
        # Enhanced display of exam info
        print_header("📋 EXAM INFORMATION")
        print(f"{Colors.YELLOW}🎯 Exam ID: {Colors.WHITE}{exam_info['exam_id']}{Colors.END}")
        print(f"{Colors.YELLOW}📝 Exam Name: {Colors.WHITE}{exam_info['exam_name']}{Colors.END}")
        print(f"{Colors.YELLOW}📅 Start Date: {Colors.WHITE}{exam_info['exam_start']}{Colors.END}")
        print(f"{Colors.YELLOW}📅 End Date: {Colors.WHITE}{exam_info['exam_end']}{Colors.END}")
        
        return exam_info
    
    def _get_subjects(self) -> pd.DataFrame:
        """Get all active subjects - ORIGINAL LOGIC PRESERVED"""
        sql = """
        SELECT subject_serial, subject_name, subject_short, subject_user_short 
        FROM tbl_student_subjects 
        WHERE is_present = True 
        ORDER BY subject_serial
        """
        results = self._execute_query(sql, "subjects")
        
        if not results:
            raise ValueError("No active subjects found in tbl_student_subjects")
        
        df = pd.DataFrame(results, columns=[
            'subject_serial', 'subject_name', 'subject_short', 'subject_user_short'
        ])
        
        print_info(f"📚 Loaded {len(df):,} active subjects")
        
        # Display comprehensive subjects table
        print_subjects_table("ACTIVE SUBJECTS FOR EXAM", df)
        
        return df
    
    def _get_students_with_results(self, comb_filter: str = None) -> pd.DataFrame:
        """
        Get students with their exam results - ORIGINAL LOGIC PRESERVED
        
        Args:
            comb_filter: Optional combination ID to filter students
        """
        comb_condition = f"AND s.comb_id = '{comb_filter}'" if comb_filter else ""
        
        sql = f"""
        SELECT DISTINCT 
            s.student_id, s.first_name, s.middle_name, s.surname, s.sex,
            s.comb_id, s.class_id, r.result_id, r.avg_marks, r.avg_grade, 
            r.division, r.points, r.subject_count_all
        FROM tbl_student_academic_info s 
        INNER JOIN tbl_student_exam_results r ON s.student_id = r.student_id
        WHERE r.exam_id = '{self.exam_id}' {comb_condition}
        """
        query_name = f"students_with_results{f'[{comb_filter}]' if comb_filter else ''}"
        results = self._execute_query(sql, query_name)
        
        if not results:
            if comb_filter:
                print_warning(f"❌ No students found for combination {comb_filter}")
                return pd.DataFrame()
            else:
                raise ValueError(f"No students found with results for exam {self.exam_id}")
        
        df = pd.DataFrame(results, columns=[
            'student_id', 'first_name', 'middle_name', 'surname', 'sex',
            'comb_id', 'class_id', 'result_id', 'avg_marks', 'avg_grade',
            'division', 'points', 'subject_count_all'
        ])
        
        df['avg_marks'] = pd.to_numeric(df['avg_marks'], errors='coerce')
        df['points'] = pd.to_numeric(df['points'], errors='coerce')
        df['subject_count_all'] = pd.to_numeric(df['subject_count_all'], errors='coerce')
        
        # Add full_name column for sorting
        df['full_name'] = (
            df['first_name'].fillna('') + ' ' + 
            df['middle_name'].fillna('') + ' ' + 
            df['surname'].fillna('')
        ).str.strip()
        
        comb_info = f" for combination {comb_filter}" if comb_filter else ""
        print_success(f"✅ Loaded {len(df):,} students{comb_info}")
        
        # Display comprehensive student data
        table_title = f"STUDENT RESULTS{comb_info.upper()}"
        print_students_table(table_title, df)
        
        return df

    def _get_all_combinations(self) -> List[str]:
        """Get all combinations that have students with results - ORIGINAL LOGIC PRESERVED"""
        sql = f"""
        SELECT DISTINCT s.comb_id 
        FROM tbl_student_academic_info s 
        INNER JOIN tbl_student_exam_results r ON s.student_id = r.student_id
        WHERE r.exam_id = '{self.exam_id}' AND s.comb_id IS NOT NULL
        ORDER BY s.comb_id
        """
        results = self._execute_query(sql, "all_combinations")
        combinations = [row[0] for row in results]
        
        safe_print_table("AVAILABLE COMBINATIONS", 
                        [(idx, comb) for idx, comb in enumerate(combinations, 1)], 
                        headers=["#", "Combination ID"])
        return combinations
    
    def _get_all_comb_subjects(self) -> Dict[str, List[int]]:
        """Get subjects for all combinations in one query - OPTIMIZED - ORIGINAL LOGIC PRESERVED"""
        sql = """
        SELECT comb_id, subject_id 
        FROM tbl_student_comb_subjects 
        ORDER BY comb_id, subject_id
        """
        results = self._execute_query(sql, "all_comb_subjects")
        
        comb_subjects = {}
        for comb_id, subject_id in results:
            if comb_id not in comb_subjects:
                comb_subjects[comb_id] = []
            comb_subjects[comb_id].append(int(subject_id))
        
        print_success(f"✅ Loaded subjects for {len(comb_subjects):,} combinations in optimized query")
        
        # Enhanced display of combination subjects
        sample_items = list(comb_subjects.items())[:10]  # Show first 10 combinations
        sample_data = [(comb, f"{len(subjects):,} subjects") for comb, subjects in sample_items]
        safe_print_table("COMBINATION SUBJECTS DISTRIBUTION", sample_data, 
                        headers=["Combination", "Total Subjects"])
        
        return comb_subjects

    def _get_all_student_marks(self, student_ids: List[str]) -> Dict[str, Dict]:
        """Get marks for all students in one query - OPTIMIZED - ORIGINAL LOGIC PRESERVED"""
        try:
            if not student_ids:
                return {}
            
            ids_str = "','".join(student_ids)
            sql = f"""
            SELECT * 
            FROM tbl_student_exam_results 
            WHERE student_id IN ('{ids_str}') AND exam_id = '{self.exam_id}'
            """
            
            rs = win32com.client.Dispatch("ADODB.Recordset")
            rs.Open(sql, self.conn, 1, 3)
            
            all_marks = {}
            if not rs.EOF:
                rs.MoveFirst()
                while not rs.EOF:
                    marks_dict = {}
                    student_id = None
                    for i in range(rs.Fields.Count):
                        field_name = rs.Fields(i).Name
                        field_value = rs.Fields(i).Value
                        marks_dict[field_name] = field_value
                        if field_name == 'student_id':
                            student_id = field_value
                    
                    if student_id:
                        all_marks[student_id] = marks_dict
                    rs.MoveNext()
        
            rs.Close()
            print_success(f"✅ Fetched marks for {len(all_marks):,} students in optimized query")
            
            # Enhanced preview of marks data
            if all_marks:
                sample_students = list(all_marks.keys())[:20]  # Show first 3 students
                for student_id in sample_students:
                    sample_items = list(all_marks[student_id].items())[:8]  # Show first 8 fields
                    safe_print_table(f"MARKS SAMPLE - STUDENT {student_id}", 
                                   [(k, str(v)[:40] + "..." if len(str(v)) > 40 else str(v)) for k, v in sample_items],
                                   headers=["Field", "Value"], max_rows=10)
            
            return all_marks
        except Exception as e:
            print_error(f"Failed to get all student marks: {str(e)}")
            return {}

    def _calculate_rankings(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Calculate all rankings with PROPER TIE HANDLING and CONFIGURABLE SORTING.
        Rankings use Olympic-style: tied students get same rank, next rank is skipped.
        Example: 1, 1, 3, 4 (not 1, 2, 3, 4)
        
        ✅ SYNCHRONIZED with AlevelProcessor to produce IDENTICAL rankings
        """
        print_info(f"🏆 Calculating rankings for {len(df):,} students...")
        print_info(f"📊 Sort criteria: {list(zip(self.sort_columns, ['↑ ASC' if asc else '↓ DESC' for asc in self.sort_ascending]))}")
        
        # ============================================================
        # CRITICAL: IDENTICAL FILTERING LOGIC FOR BOTH FILES
        # ============================================================
        print(f"{Colors.CYAN}🔍 DEBUG: Total students before filtering: {len(df):,}{Colors.END}")
        
        # Filter criteria MUST match AlevelProcessor exactly
        invalid_mask = (
            (df['division'] == 'ABS') |
            (df['points'].isna()) |
            (df['avg_marks'].isna())
        )
        
        df_valid = df[~invalid_mask].copy()
        invalid_students = df[invalid_mask]
        
        print(f"{Colors.CYAN}🔍 DEBUG: Students with division='ABS': {(df['division'] == 'ABS').sum()}{Colors.END}")
        print(f"{Colors.CYAN}🔍 DEBUG: Students with NaN points: {df['points'].isna().sum()}{Colors.END}")
        print(f"{Colors.CYAN}🔍 DEBUG: Students with NaN avg_marks: {df['avg_marks'].isna().sum()}{Colors.END}")
        print(f"{Colors.CYAN}🔍 DEBUG: Total invalid students: {len(invalid_students):,}{Colors.END}")
        print(f"{Colors.CYAN}🔍 DEBUG: Total valid students: {len(df_valid):,}{Colors.END}")
        
        if df_valid.empty:
            print_warning("⚠ No valid students to rank")
            return df
        
        print_info(f"🎯 Ranking {len(df_valid):,} valid students")
        
        # ============================================================
        # CRITICAL: IDENTICAL NaN HANDLING FOR BOTH FILES
        # ============================================================
        print_info(f"🔧 Preparing sort columns: {self.sort_columns}")
        
        for col in self.sort_columns:
            before_fill = df_valid[col].isna().sum()
            
            if col == 'points':
                df_valid[col] = df_valid[col].fillna(999999)
            elif col == 'avg_marks':
                df_valid[col] = df_valid[col].fillna(-1)
            elif col in ['subject_count', 'subject_count_all']:
                df_valid[col] = df_valid[col].fillna(0)
            
            after_fill = df_valid[col].isna().sum()
            print(f"{Colors.CYAN}- {col}: filled {before_fill} NaN values → {after_fill} remaining{Colors.END}")
        
        # ⚠️ CRITICAL: Round to 4 decimal places to avoid floating point comparison issues
        for col in ['points', 'avg_marks']:
            if col in df_valid.columns:
                df_valid[col] = df_valid[col].round(4)
                print(f"{Colors.CYAN}- {col}: rounded to 4 decimal places{Colors.END}")
        
        # ============================================================
        # OVERALL RANKING (SCHOOL-WIDE) - IDENTICAL TO PROCESSOR
        # ============================================================
        print(f"\n{Colors.YELLOW}{Colors.BOLD}🏆 School-Wide Ranking{Colors.END}")
        print(f"{Colors.CYAN}- Sort criteria: {list(zip(self.sort_columns, ['↑ ASC' if asc else '↓ DESC' for asc in self.sort_ascending]))}{Colors.END}")
        
        # STEP 1: Sort by configured columns
        df_sorted = df_valid.sort_values(
            by=self.sort_columns,
            ascending=self.sort_ascending
        ).copy()
        
        # STEP 2: Create sort tuple AFTER sorting
        df_sorted['_sort_tuple_overall'] = df_sorted[self.sort_columns].apply(
            lambda row: tuple(row), axis=1
        )
        
        # STEP 3: Get unique tuples in sorted order
        unique_tuples = df_sorted['_sort_tuple_overall'].unique()
        print(f"{Colors.CYAN}🔍 DEBUG: Found {len(unique_tuples):,} unique performance levels{Colors.END}")
        
        # STEP 4: Map each unique tuple to sequential ranks
        rank_map = {tuple_val: idx + 1 for idx, tuple_val in enumerate(unique_tuples)}
        
        # STEP 5: Assign ranks
        df_sorted['overall_rank'] = df_sorted['_sort_tuple_overall'].map(rank_map)
        
        # Validation
        max_rank = df_sorted['overall_rank'].max()
        print_success(f"✅ Overall ranking complete (ranks 1-{max_rank})")
        
        # Count ties
        rank_counts = df_sorted['overall_rank'].value_counts()
        ties = rank_counts[rank_counts > 1]
        if len(ties) > 0:
            total_tied = ties.sum()
            print_warning(f"⚠ Ties detected: {len(ties)} rank positions have ties ({total_tied} students total)")
        
        # 🔍 DEBUG: Export first 20 students for verification
        debug_df = df_sorted.head(20)[['student_id', 'first_name', 'surname', 'points', 'avg_marks',
                                        'subject_count_all', '_sort_tuple_overall', 'overall_rank']]
        print(f"\n{Colors.MAGENTA}{Colors.BOLD}🔍 DEBUG: First 20 ranked students (for verification){Colors.END}")
        for idx, row in debug_df.iterrows():
            full_name = f"{row['first_name']} {row['surname']}"
            print(f"{Colors.CYAN} {row['overall_rank']:3d}. {row['student_id']:15s} | "
                f"pts:{row['points']:6.2f} avg:{row['avg_marks']:6.2f} cnt:{row['subject_count_all']:2.0f} | "
                f"{full_name[:30]}{Colors.END}")
        
        # ============================================================
        # COMB OVERALL RANKING - IDENTICAL TO PROCESSOR
        # ============================================================
        def rank_comb_overall(group):
            """Rank students within combination - IDENTICAL to processor logic"""
            group['_comb_sort_tuple'] = group[self.sort_columns].apply(
                lambda row: tuple(row), axis=1
            )
            unique_tuples = group['_comb_sort_tuple'].unique()
            rank_map = {tuple_val: idx + 1 for idx, tuple_val in enumerate(unique_tuples)}
            group['comb_overall_rank'] = group['_comb_sort_tuple'].map(rank_map)
            return group.drop(columns=['_comb_sort_tuple'])
        
        df_sorted = df_sorted.groupby('comb_id', group_keys=False).apply(rank_comb_overall)
        print_success("✅ Comb overall ranking complete")
        
        # ============================================================
        # COMB SEX RANKING
        # ============================================================
        def rank_comb_sex(group):
            """Rank students within combination by sex"""
            group['_comb_sex_tuple'] = group[self.sort_columns].apply(
                lambda row: tuple(row), axis=1
            )
            unique_tuples = group['_comb_sex_tuple'].unique()
            rank_map = {tuple_val: idx + 1 for idx, tuple_val in enumerate(unique_tuples)}
            group['comb_sex_rank'] = group['_comb_sex_tuple'].map(rank_map)
            return group.drop(columns=['_comb_sex_tuple'])
        
        df_sorted = df_sorted.groupby(['comb_id', 'sex'], group_keys=False).apply(rank_comb_sex)
        print_success("✅ Comb sex ranking complete")
        
        # ============================================================
        # CLASS OVERALL RANKING
        # ============================================================
        def rank_class_overall(group):
            """Rank students within class"""
            group['_class_sort_tuple'] = group[self.sort_columns].apply(
                lambda row: tuple(row), axis=1
            )
            unique_tuples = group['_class_sort_tuple'].unique()
            rank_map = {tuple_val: idx + 1 for idx, tuple_val in enumerate(unique_tuples)}
            group['class_overall_rank'] = group['_class_sort_tuple'].map(rank_map)
            return group.drop(columns=['_class_sort_tuple'])
        
        df_sorted = df_sorted.groupby('class_id', group_keys=False).apply(rank_class_overall)
        print_success("✅ Class overall ranking complete")
        
        # ============================================================
        # CLASS SEX RANKING
        # ============================================================
        def rank_class_sex(group):
            """Rank students within class by sex"""
            group['_class_sex_tuple'] = group[self.sort_columns].apply(
                lambda row: tuple(row), axis=1
            )
            unique_tuples = group['_class_sex_tuple'].unique()
            rank_map = {tuple_val: idx + 1 for idx, tuple_val in enumerate(unique_tuples)}
            group['class_sex_rank'] = group['_class_sex_tuple'].map(rank_map)
            return group.drop(columns=['_class_sex_tuple'])
        
        df_sorted = df_sorted.groupby(['class_id', 'sex'], group_keys=False).apply(rank_class_sex)
        print_success("✅ Class sex ranking complete")
        
        # Clean up temporary columns
        df_sorted = df_sorted.drop(columns=['_sort_tuple_overall'], errors='ignore')
        
        # ============================================================
        # MERGE BACK WITH ORIGINAL DF
        # ============================================================
        df_result = df.merge(
            df_sorted[['student_id', 'overall_rank', 'comb_overall_rank', 'comb_sex_rank',
                    'class_overall_rank', 'class_sex_rank']],
            on='student_id',
            how='left'
        )
        
        print_success(f"✅ All rankings calculated with synchronized logic")
        
        return df_result


    def _calculate_total_counts(self, df: pd.DataFrame) -> Dict:
        """Calculate total counts for each grouping (where avg_marks >= 0) - ORIGINAL LOGIC PRESERVED"""
        print_info("📊 Calculating total counts for position denominators...")
        
        # Filter only students with valid avg_marks
        df_valid = df[df['division'] !='ABS'].copy()
        
        total_counts = {}
        
        # Comb sex totals
        comb_sex_counts = df_valid.groupby(['comb_id', 'sex']).size()
        for (comb_id, sex), count in comb_sex_counts.items():
            total_counts[f"comb_sex_{comb_id}_{sex}"] = count
        
        # Comb overall totals
        comb_overall_counts = df_valid.groupby('comb_id').size()
        for comb_id, count in comb_overall_counts.items():
            total_counts[f"comb_overall_{comb_id}"] = count
        
        # Class sex totals
        class_sex_counts = df_valid.groupby(['class_id', 'sex']).size()
        for (class_id, sex), count in class_sex_counts.items():
            total_counts[f"class_sex_{class_id}_{sex}"] = count
        
        # Class overall totals
        class_overall_counts = df_valid.groupby('class_id').size()
        for class_id, count in class_overall_counts.items():
            total_counts[f"class_overall_{class_id}"] = count
        
        print_success(f"✅ Total counts calculated for {len(total_counts):,} groups")
        return total_counts

    def _apply_ordering(self, df: pd.DataFrame) -> pd.DataFrame:
        """Apply ordering based on order_by setting - ORIGINAL LOGIC PRESERVED"""
        if self.order_by == "position":
            print_info("🔢 Using position order (already sorted)")
            return df
        elif self.order_by == "name":
            df_ordered = df.sort_values(by=['full_name'], ascending=[True])
            print_success("✅ Sorted by full name")
            return df_ordered
        elif self.order_by == "sex_name":
            df_ordered = df.sort_values(by=['sex', 'full_name'], ascending=[True, True])
            print_success("✅ Sorted by sex, then full name")
            return df_ordered
        else:
            return df
    
    def _merge_cells(self, r1: int, c1: int, r2: int, c2: int):
        """Merge cells and center align - ORIGINAL LOGIC PRESERVED"""
        cell_range = self.worksheet.Range(
            self.worksheet.Cells(r1, c1),
            self.worksheet.Cells(r2, c2)
        )
        cell_range.Merge()
        cell_range.HorizontalAlignment = -4108
        cell_range.VerticalAlignment = -4108
        
    def _create_header(self, subjects_df: pd.DataFrame):
        """Create formatted header rows with rotated text for subjects - ORIGINAL LOGIC PRESERVED"""
        print_info("📋 Creating Excel header...")
        r = self.start_row
        subject_count = len(subjects_df)
        
        # S/N column
        self._merge_cells(r, 1, r + 2, 1)
        self.worksheet.Cells(r, 1).Value = "S/N"
        self.worksheet.Cells(r, 1).Orientation = 90
        self.worksheet.Cells(r, 1).HorizontalAlignment = -4108
        self.worksheet.Cells(r, 1).VerticalAlignment = -4108
        self.worksheet.Cells(r, 1).WrapText = True
        
        # Student info columns (2-6)
        headers = ["FIRST NAME", "MIDDLE NAME", "SURNAME", "SEX", "COMB"]
        for i, header in enumerate(headers, start=2):
            self._merge_cells(r, i, r + 2, i)
            cell = self.worksheet.Cells(r, i)
            cell.Value = header
            cell.HorizontalAlignment = -4108
            cell.VerticalAlignment = -4108
            cell.WrapText = True
        
        # Subject columns
        col = 7
        for _, subj in subjects_df.iterrows():
            self._merge_cells(r, col, r + 2, col + 1)
            cell = self.worksheet.Cells(r, col)
            cell.Value = subj['subject_user_short']
            cell.Orientation = 90
            cell.HorizontalAlignment = -4108
            cell.VerticalAlignment = -4108
            cell.WrapText = True
            col += 2
        
        # Summary section
        self._merge_cells(r, col, r, col + 3)
        summary_header = self.worksheet.Cells(r, col)
        summary_header.Value = "SUMMARY"
        summary_header.HorizontalAlignment = -4108
        summary_header.VerticalAlignment = -4108
        summary_header.WrapText = True
        
        # Summary sub-headers - merge rows r+1 and r+2
        summary_headers = ["DIV", "PTS", "AVG", "GRD"]
        for i, header in enumerate(summary_headers):
            self._merge_cells(r + 1, col + i, r + 2, col + i)
            cell = self.worksheet.Cells(r + 1, col + i)
            cell.Value = header
            cell.HorizontalAlignment = -4108
            cell.VerticalAlignment = -4108
            cell.WrapText = True
        col += 4
        
        # COMB Position
        self._merge_cells(r, col, r, col + 1)
        comb_header = self.worksheet.Cells(r, col)
        comb_header.Value = "COMB"
        comb_header.HorizontalAlignment = -4108
        comb_header.VerticalAlignment = -4108
        comb_header.WrapText = True
        
        # SEX and OVERALL - merge rows r+1 and r+2
        self._merge_cells(r + 1, col, r + 2, col)
        sex_cell = self.worksheet.Cells(r + 1, col)
        sex_cell.Value = "SEX"
        sex_cell.HorizontalAlignment = -4108
        sex_cell.VerticalAlignment = -4108
        sex_cell.WrapText = True
        
        self._merge_cells(r + 1, col + 1, r + 2, col + 1)
        overall_cell = self.worksheet.Cells(r + 1, col + 1)
        overall_cell.Value = "OVERALL"
        overall_cell.HorizontalAlignment = -4108
        overall_cell.VerticalAlignment = -4108
        overall_cell.WrapText = True
        col += 2
        
        # CLASS Position
        self._merge_cells(r, col, r, col + 1)
        class_header = self.worksheet.Cells(r, col)
        class_header.Value = "CLASS"
        class_header.HorizontalAlignment = -4108
        class_header.VerticalAlignment = -4108
        class_header.WrapText = True
        
        # SEX and OVERALL - merge rows r+1 and r+2
        self._merge_cells(r + 1, col, r + 2, col)
        sex_cell2 = self.worksheet.Cells(r + 1, col)
        sex_cell2.Value = "SEX"
        sex_cell2.HorizontalAlignment = -4108
        sex_cell2.VerticalAlignment = -4108
        sex_cell2.WrapText = True
        
        self._merge_cells(r + 1, col + 1, r + 2, col + 1)
        overall_cell2 = self.worksheet.Cells(r + 1, col + 1)
        overall_cell2.Value = "OVERALL"
        overall_cell2.HorizontalAlignment = -4108
        overall_cell2.VerticalAlignment = -4108
        overall_cell2.WrapText = True
        
        # Center columns
        for c in range(5, col + 2):
            self.worksheet.Columns(c).HorizontalAlignment = -4108
        
        print_success(f"✅ Excel header created with {subject_count} subjects")

    def _fill_student_row(self, row: int, sn: int, student: pd.Series, 
                          subjects_df: pd.DataFrame, marks_dict: Dict, 
                          comb_subjects: List[int], total_counts: Dict):
        """Fill a single student row with data and formatting - PERFORMANCE OPTIMIZED - ORIGINAL LOGIC PRESERVED"""
        # PERFORMANCE: Build row data as list, write once at end
        row_values = []
        shade_ranges = []
        
        # S/N
        row_values.append((row, 1, sn))
        
        # Student info
        row_values.append((row, 2, student['first_name'] or ""))
        row_values.append((row, 3, student['middle_name'] or ""))
        row_values.append((row, 4, student['surname'] or ""))
        row_values.append((row, 5, student['sex'] or ""))
        row_values.append((row, 6, student['comb_id'] or ""))
        
        # Subject marks and grades
        col = 7
        for _, subj in subjects_df.iterrows():
            subject_short = subj['subject_short']
            subject_serial = int(subj['subject_serial'])
            
            mark_value = marks_dict.get(subject_short)
            grade_value = marks_dict.get(f"{subject_short}_grade")
            
            mark_display = ""
            if mark_value is not None:
                try:
                    mark_numeric = float(mark_value)
                    if mark_numeric >= 0:
                        mark_display = round(mark_numeric, 2)
                except (ValueError, TypeError):
                    mark_display = mark_value
            
            row_values.append((row, col, mark_display))
            row_values.append((row, col + 1, grade_value if grade_value is not None else ""))
            
            # Track shading ranges
            is_not_in_comb = (subject_serial not in comb_subjects)
            has_real_marks = (mark_value is not None) and isinstance(mark_value, (int, float)) and (mark_value >= 0)
            should_shade = is_not_in_comb and not has_real_marks
            
            if should_shade:
                shade_ranges.append((row, col, row, col + 1))
            
            col += 2
        
        # Summary data
        row_values.append((row, col, student['division'] or ""))
        col += 1
        
        pts_value = student['points']
        row_values.append((row, col, round(pts_value, 2) if pd.notna(pts_value) else ""))
        col += 1
        
        avg_value = student['avg_marks']
        row_values.append((row, col, round(avg_value, 2) if pd.notna(avg_value) else ""))
        col += 1
        
        row_values.append((row, col, student['avg_grade'] or ""))
        col += 1
        
        # Rankings with "out of" format
        comb_id = student['comb_id']
        class_id = student['class_id']
        sex = student['sex']
        
        # COMB SEX position
        if pd.notna(student['comb_sex_rank']):
            comb_sex_rank = int(student['comb_sex_rank'])
            comb_sex_total = total_counts.get(f"comb_sex_{comb_id}_{sex}", comb_sex_rank)
            row_values.append((row, col, f"'{comb_sex_rank}/{comb_sex_total}"))
        else:
            row_values.append((row, col, ""))
        col += 1
        
        # COMB OVERALL position
        if pd.notna(student['comb_overall_rank']):
            comb_overall_rank = int(student['comb_overall_rank'])
            comb_overall_total = total_counts.get(f"comb_overall_{comb_id}", comb_overall_rank)
            row_values.append((row, col, f"'{comb_overall_rank}/{comb_overall_total}"))
        else:
            row_values.append((row, col, ""))
        col += 1
        
        # CLASS SEX position
        if pd.notna(student['class_sex_rank']):
            class_sex_rank = int(student['class_sex_rank'])
            class_sex_total = total_counts.get(f"class_sex_{class_id}_{sex}", class_sex_rank)
            row_values.append((row, col, f"'{class_sex_rank}/{class_sex_total}"))
        else:
            row_values.append((row, col, ""))
        col += 1
        
        # CLASS OVERALL position
        if pd.notna(student['class_overall_rank']):
            class_overall_rank = int(student['class_overall_rank'])
            class_overall_total = total_counts.get(f"class_overall_{class_id}", class_overall_rank)
            row_values.append((row, col, f"'{class_overall_rank}/{class_overall_total}"))
        else:
            row_values.append((row, col, ""))
        
        # PERFORMANCE: Write all values at once
        for r, c, val in row_values:
            self.worksheet.Cells(r, c).Value = val
        
        # Apply shading
        for r1, c1, r2, c2 in shade_ranges:
            shade_range = self.worksheet.Range(
                self.worksheet.Cells(r1, c1),
                self.worksheet.Cells(r2, c2)
            )
            shade_range.Interior.ColorIndex = 15

    def _create_summary_table(self, students_df: pd.DataFrame, subjects_df: pd.DataFrame,
                             all_marks: Dict, all_comb_subjects: Dict, 
                             start_row: int, title: str, total_counts: Dict, last_col: int) -> int:
        """Create a summary table (Top N or Bottom N) - ORIGINAL LOGIC PRESERVED"""
        print_info(f"📊 Creating {title} table at row {start_row}")
        
        # Title - MERGE ACROSS ALL COLUMNS
        title_cell = self.worksheet.Cells(start_row, 1)
        title_cell.Value = title
        title_cell.Font.Bold = True
        title_cell.Font.Size = 12
        
        # Merge title across all columns
        title_range = self.worksheet.Range(
            self.worksheet.Cells(start_row, 1),
            self.worksheet.Cells(start_row, last_col)
        )
        title_range.Merge()
        title_range.HorizontalAlignment = -4108
        title_range.VerticalAlignment = -4108
        
        start_row += 1
        
        # Create header
        old_start_row = self.start_row
        self.start_row = start_row
        self._create_header(subjects_df)
        self.start_row = old_start_row
        
        # Fill data
        current_row = start_row + 3
        for idx, (_, student) in enumerate(students_df.iterrows(), start=1):
            student_id = student['student_id']
            comb_id = student['comb_id']
            comb_subjects = all_comb_subjects.get(comb_id, [])
            marks_dict = all_marks.get(student_id, {})
            
            self._fill_student_row(
                current_row, idx, student, subjects_df,
                marks_dict, comb_subjects, total_counts
            )
            current_row += 1
        
        # Format the summary table
        last_row = current_row - 1
        
        summary_range = self.worksheet.Range(
            self.worksheet.Cells(start_row, 1),
            self.worksheet.Cells(last_row, last_col)
        )
        summary_range.Borders.LineStyle = 1
        
        # Bold headers
        header_range = self.worksheet.Range(
            self.worksheet.Cells(start_row, 1),
            self.worksheet.Cells(start_row + 2, last_col)
        )
        header_range.Font.Bold = True
        
        # AutoFit after summary table
        self.worksheet.Columns.AutoFit()
        self.worksheet.Rows.AutoFit()
        
        print_success(f"✅ {title} table completed ({len(students_df)} students)")
        return current_row + 2

    def _setup_page_settings(self, subject_count: int):
        """Setup page orientation and paper size - ORIGINAL LOGIC PRESERVED"""
        # Determine orientation
        if self.orientation:
            orientation = self.orientation
        else:
            orientation = "portrait" if subject_count <= 10 else "landscape"
        
        # Paper size constants
        paper_sizes = {
            "A4": 9,
            "LTR": 1,
            "A3": 8
        }
        
        # Orientation constants
        orientations = {
            "portrait": 1,
            "landscape": 2
        }
        
        print_info(f"📄 Setting up: {self.paper_size} paper, {orientation} orientation")
        
        page_setup = self.worksheet.PageSetup
        page_setup.PaperSize = paper_sizes[self.paper_size]
        page_setup.Orientation = orientations[orientation]
        page_setup.Zoom = False
        page_setup.FitToPagesWide = 1
        page_setup.FitToPagesTall = False
        page_setup.LeftMargin = self.excel_app.InchesToPoints(0.5)
        page_setup.RightMargin = self.excel_app.InchesToPoints(0.5)
        page_setup.TopMargin = self.excel_app.InchesToPoints(0.75)
        page_setup.BottomMargin = self.excel_app.InchesToPoints(0.75)
        
        print_success("✅ Page settings applied - fit to 1 page wide")
    
    def _format_worksheet(self, subject_count: int, last_row: int):
        """Apply final formatting to worksheet - ORIGINAL LOGIC PRESERVED"""
        print_info("🎨 Applying final Excel formatting...")
        
        last_col = 6 + (subject_count * 2) + 8
        
        used_range = self.worksheet.Range(
            self.worksheet.Cells(self.start_row, 1),
            self.worksheet.Cells(last_row, last_col)
        )
        used_range.Font.Name = "Arial"
        used_range.Font.Size = 10
        
        header_range = self.worksheet.Range(
            self.worksheet.Cells(self.start_row, 1),
            self.worksheet.Cells(self.start_row + 2, last_col)
        )
        header_range.Font.Bold = True
        
        used_range.Borders.LineStyle = 1
        
        # AutoFit columns and rows
        self.worksheet.Columns.AutoFit()
        self.worksheet.Rows.AutoFit()
        
        self._setup_page_settings(subject_count)
        
        print_success(f"✅ Formatting complete (columns 1-{last_col}, rows {self.start_row}-{last_row})")
        print_success("✅ AutoFit applied to all columns and rows")
    
    def export_exam_results(self, output_path: Optional[str] = None):
        """Main export function - ORIGINAL LOGIC PRESERVED"""
        print_header("🚀 STARTING MAIN EXPORT PROCESS")
        print_info(f"📤 Exporting exam: {self.exam_id}")
        
        try:
            exam_info = self._get_exam_info()
            print_success(f"🎯 Exam: {exam_info['exam_name']}")
            
            subjects_df = self._get_subjects()
            all_comb_subjects = self._get_all_comb_subjects()
            
            print_header("📊 CREATING MAIN SHEET: ALL STUDENTS")
            self.worksheet.Name = "ALL STUDENTS"
            self._export_sheet(subjects_df, all_comb_subjects, comb_filter=None, sheet_name="ALL STUDENTS")
            
            if self.include_comb_sheets:
                combinations = self._get_all_combinations()
                print_header(f"📑 CREATING {len(combinations)} COMBINATION SHEETS")
                
                for comb_id in combinations:
                    print_info(f"🔄 Creating sheet for combination: {comb_id}")
                    try:
                        new_ws = self.workbook.Worksheets.Add()
                        safe_name = str(comb_id)[:31].replace("/", "-").replace("\\", "-").replace("*", "").replace("?", "").replace("[", "").replace("]", "")
                        new_ws.Name = safe_name
                        self.worksheet = new_ws
                        
                        self._export_sheet(subjects_df, all_comb_subjects, comb_filter=comb_id, sheet_name=safe_name)
                    except Exception as e:
                        print_error(f"❌ Failed to create sheet for {comb_id}: {str(e)}")
                        traceback.print_exc()
            
            if output_path is None:
                current_dir = Path(r"C:\Kiyabo App\Results")
                current_dir.mkdir(exist_ok=True)
                output_path = current_dir / f"Exam_Results_{self.exam_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            
            print_header("💾 SAVING EXCEL FILE")
            print_info(f"💾 Saving to: {output_path}")
            self.workbook.SaveAs(str(output_path))
            
            # Close workbook and quit Excel
            self.workbook.Close(SaveChanges=False)
            self.excel_app.Quit()
            
            # Open the generated file with default program
            os.startfile(str(output_path))
            
            print_header("🎉 EXPORT COMPLETED SUCCESSFULLY!")
            print_success(f"📁 Output file: {output_path}")
            if self.include_comb_sheets:
                print_success(f"📑 Sheets created: 1 main + {len(combinations)} combinations")
            else:
                print_success("📑 Sheets created: 1 main sheet")
            
        except Exception as e:
            print_header("💥 EXPORT FAILED")
            print_error(f"💥 Export failed: {str(e)}")
            traceback.print_exc()
            raise

    def _export_sheet(self, subjects_df: pd.DataFrame, all_comb_subjects: Dict[str, List[int]], 
                    comb_filter: str = None, sheet_name: str = "Sheet"):
        """Export data to a specific sheet - FIXED: Use same rankings as main sheet"""
        try:
            # Get students for this specific sheet (filtered by combination if needed)
            students_df = self._get_students_with_results(comb_filter)
            
            if students_df.empty:
                print_warning(f"⚠ No students found for {sheet_name}, skipping")
                return
            
            # 🚀 CRITICAL FIX: Get ALL students to calculate consistent rankings across all sheets
            all_students_df = self._get_students_with_results()  # No filter - all students
            
            # Calculate rankings ONCE using all students data (consistent across all sheets)
            all_students_ranked_df = self._calculate_rankings(all_students_df)
            
            # 🎯 KEY FIX: Merge the pre-calculated rankings from ALL students into current sheet's students
            # This ensures combination sheets use the same rankings as main sheet
            students_df = students_df.merge(
                all_students_ranked_df[['student_id', 'overall_rank', 'comb_overall_rank', 'comb_sex_rank', 
                                    'class_overall_rank', 'class_sex_rank']],
                on='student_id',
                how='left'
            )
            
            print_success(f"✅ Applied consistent rankings from main dataset to {sheet_name}")
            
            # Calculate total counts using ALL students for consistent denominators
            total_counts = self._calculate_total_counts(all_students_ranked_df)
            
            # Apply ordering to current sheet's students
            students_df_ordered = self._apply_ordering(students_df)
            
            # Get marks for current sheet's students
            student_ids = students_df_ordered['student_id'].tolist()
            all_marks = self._get_all_student_marks(student_ids)
            
            # Filter subjects for combination sheets (if needed)
            if comb_filter:
                comb_subjects_list = all_comb_subjects.get(comb_filter, [])
                opted_serials = set()
                
                for _, student in students_df_ordered.iterrows():
                    marks_dict = all_marks.get(student['student_id'], {})
                    for _, subj in subjects_df.iterrows():
                        mark_value = marks_dict.get(subj['subject_short'])
                        if mark_value is not None:
                            try:
                                mark_numeric = float(mark_value)
                                if mark_numeric >= 0:
                                    opted_serials.add(int(subj['subject_serial']))
                            except (ValueError, TypeError):
                                pass
                
                all_serials = set(comb_subjects_list) | opted_serials
                filtered_subjects_df = subjects_df[subjects_df['subject_serial'].isin(all_serials)].sort_values('subject_serial')
                print_success(f"✅ Filtered to {len(filtered_subjects_df)} subjects for {sheet_name}")
            else:
                filtered_subjects_df = subjects_df
            
            # Create header with appropriate subjects
            self._create_header(filtered_subjects_df)
            
            # Fill student data
            sheet_desc = f"📝 Filling {sheet_name}"
            print_info(f"{sheet_desc}...")
            current_row = self.start_row + 3
            
            with tqdm(total=len(students_df_ordered), desc=sheet_desc, unit="student", 
                    bar_format="{l_bar}%s{bar}%s{r_bar}" % (Colors.BLUE, Colors.END)) as pbar:
                for idx, (_, student) in enumerate(students_df_ordered.iterrows(), start=1):
                    student_id = student['student_id']
                    comb_id = student['comb_id']
                    comb_subjects = all_comb_subjects.get(comb_id, [])
                    marks_dict = all_marks.get(student_id, {})
                    
                    self._fill_student_row(
                        current_row, idx, student, filtered_subjects_df, 
                        marks_dict, comb_subjects, total_counts
                    )
                    
                    current_row += 1
                    pbar.update(1)
            
            print_success(f"✅ All {len(students_df_ordered):,} students processed for {sheet_name}")
            
            last_row_main = current_row - 1
            last_col = 6 + (len(filtered_subjects_df) * 2) + 8
            self._format_worksheet(len(filtered_subjects_df), last_row_main)
            
            # Create Top N and Bottom N tables using the SAME rankings as main sheet
            students_df_ranked_valid = students_df[students_df['avg_marks'] >= 0].copy()
            
            # 🎯 FIX: Use consistent ranking logic for Top/Bottom N across all sheets
            if comb_filter:
                # For combination sheets: sort by comb_overall_rank (from main sheet calculation)
                students_df_ranked_valid = students_df_ranked_valid.sort_values('comb_overall_rank')
                rank_type = "comb_overall_rank"
            else:
                # For main sheet: sort by class_overall_rank (from main sheet calculation)  
                students_df_ranked_valid = students_df_ranked_valid.sort_values('class_overall_rank')
                rank_type = "class_overall_rank"
            
            print_info(f"🎯 Top/Bottom N sorting by: {rank_type}")
            
            # Create Top N table
            if self.top_n > 0 and len(students_df_ranked_valid) >= self.top_n:
                next_row = last_row_main + 3
                top_students = students_df_ranked_valid.head(self.top_n)
                next_row = self._create_summary_table(
                    top_students, filtered_subjects_df, all_marks, all_comb_subjects,
                    next_row, f"TOP {self.top_n} STUDENTS", total_counts, last_col
                )
                last_row_main = next_row
            
            # Create Bottom N table
            if self.bottom_n > 0 and len(students_df_ranked_valid) >= self.bottom_n:
                next_row = last_row_main + 1
                # we reverse the order so that the last stident should be in the first position..
                bottom_students = students_df_ranked_valid.tail(self.bottom_n).iloc[::-1].reset_index(drop=True)
                next_row = self._create_summary_table(
                    bottom_students, filtered_subjects_df, all_marks, all_comb_subjects,
                    next_row, f"BOTTOM {self.bottom_n} STUDENTS", total_counts, last_col
                )
            
            # Final AutoFit
            print_info(f"🔧 Final AutoFit for {sheet_name}")
            self.worksheet.Columns.AutoFit()
            self.worksheet.Rows.AutoFit()
            
        except Exception as e:
            print_error(f"❌ Failed to export sheet {sheet_name}: {str(e)}")
            traceback.print_exc()
            raise



    def close(self):
        """Clean up resources - ORIGINAL LOGIC PRESERVED"""
        print_header("🧹 CLEANING UP RESOURCES")
        try:
            if self.conn:
                self.conn.Close()
                print_success("✅ Database connection closed")
            if self.workbook:
                self.workbook.Close(SaveChanges=False)
                print_success("✅ Excel workbook closed")
            if self.excel_app:
                self.excel_app.Quit()
                print_success("✅ Excel application closed")
        except Exception as e:
            print_warning(f"⚠ Cleanup warning: {str(e)}")
        finally:
            pythoncom.CoUninitialize()
        print_success("✅ Cleanup completed")
    
    def __enter__(self):
        """Context manager entry"""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit"""
        self.close()

    def run(self):
        """
        Execute the full export process using the current instance configuration.
        Assumes open_excel_file() has NOT been called yet.
        Saves output to: C:\\Kiyabo App\\Results\\Exam_Results_{exam_id}.xlsx
        """
        from pathlib import Path
        import os

        print_header("🚀 STARTING EXPORT VIA .run()")

        # Ensure output directory exists
        results_dir = Path(r"C:\Kiyabo App\Results")
        results_dir.mkdir(parents=True, exist_ok=True)

        # Generate output path with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")  # e.g., 20251111_143205
        output_path = results_dir / f"Exam_Results_Export_{self.exam_id}_{timestamp}.xlsx"

        # Open Excel file
        self.open_excel_file(str(output_path), start_row=1)

        # Perform full export (this internally handles all sheets, rankings, saving, etc.)
        self.export_exam_results(str(output_path))

        # The file is auto-opened by export_exam_results() via os.startfile()
        return str(output_path)


def main():
    parser = argparse.ArgumentParser(description="🎓 Student Exam Results Exporter - Enhanced Version")

    parser.add_argument("--exam-id", required=True, help="Exam ID (e.g., MNT520251027)")
    parser.add_argument("--db-path", default=r"C:\Kiyabo App\backend\Kiyabo App Backend v4.0.0.accdb", help="Path to the Access database file (.accdb)")
    parser.add_argument("--include-comb-sheets", action="store_true", help="Include combination sheets in export")
    parser.add_argument("--order-by", default="position", help="Order results by field (default: position)")
    parser.add_argument("--top-n", type=int, default=10, help="Number of top students to include (default: 10)")
    parser.add_argument("--bottom-n", type=int, default=10, help="Number of bottom students to include (default: 10)")
    parser.add_argument("--paper-size", default="A4", help="Paper size for export (default: A4)")
    parser.add_argument("--orientation", choices=["portrait", "landscape"], help="Page orientation")

    args = parser.parse_args()

    start_time = time.time()

    print_header("🎓 STUDENT EXAM RESULTS EXPORTER - ENHANCED VERSION")
    print_info("🚀 Starting comprehensive export process...")

    exporter = StudentExamExporter(
        args.exam_id,
        args.db_path,
        include_comb_sheets=args.include_comb_sheets,
        order_by=args.order_by,
        top_n=args.top_n,
        bottom_n=args.bottom_n,
        paper_size=args.paper_size,
        orientation=args.orientation,
    )

    
    try:
        results_dir = Path(r"C:\Kiyabo App\Results")
        results_dir.mkdir(exist_ok=True)

        excel_file = results_dir / f"Exam_Results_{args.exam_id}.xlsx"

        
        exporter.open_excel_file(str(excel_file), start_row=1)
        
        print_header("🔍 DATABASE EXPLORATION")
        exporter.list_tables()
        
        print_header("📋 KEY TABLES STRUCTURE INSPECTION")
        exporter.inspect_table_structure("tbl_student_comb_subjects")
        exporter.inspect_table_structure("tbl_student_combs")
        exporter.inspect_table_structure("tbl_student_subjects")
        
        exporter.export_exam_results()
        
        execution_time = time.time() - start_time
        print_header("✅ ALL OPERATIONS COMPLETED SUCCESSFULLY!")
        print_success(f"⏱️ Total execution time: {execution_time:.2f} seconds")
        print_success(f"📊 Process completed at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
    except Exception as e:
        print_header("❌ FATAL ERROR OCCURRED")
        print_error(f"💥 Fatal Error: {str(e)}")
        traceback.print_exc()
        sys.exit(1)
        
    finally:
        exporter.close() 


if __name__ == "__main__":
    main()