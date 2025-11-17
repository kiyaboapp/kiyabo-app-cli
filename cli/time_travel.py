from .alevel.ranking import process_exam as AlevelProcessor
from .olevel.processDS import OlevelProcessor
from .primary.process import ExamProcessor as PrimaryProcessor
import pyodbc
from datetime import datetime

class Color:
    """Beautiful color codes for terminal output"""
    RESET = '\033[0m'
    BOLD = '\033[1m'
    BLUE = '\033[94m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    CYAN = '\033[96m'
    MAGENTA = '\033[95m'
    ORANGE = '\033[38;5;214m'

class TimeTravelProcessor:
    def __init__(self, level, db_path):
        self.level = level.lower()
        self.db_path = db_path
        
    def into_the_future(self):
        """Main processing method"""
        print(f"\n{Color.CYAN}{Color.BOLD}🚀 Starting time travel for {Color.MAGENTA}{self.level.upper()}{Color.CYAN} level{Color.RESET}")
        print(f"{Color.BLUE}📁 Database: {Color.YELLOW}{self.db_path}{Color.RESET}")
        print(f"{Color.BLUE}⏰ Started:  {Color.YELLOW}{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}{Color.RESET}")
        
        exam_data_list = self._fetch_exam_data()
        
        if not exam_data_list:
            print(f"{Color.RED}❌ No exams found to process.{Color.RESET}")
            return
            
        print(f"{Color.GREEN}🎯 Found {len(exam_data_list)} exam(s) to process{Color.RESET}")
        
        for index, exam_data in enumerate(exam_data_list, 1):
            try:
                print(f"\n{Color.ORANGE}{Color.BOLD}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━{Color.RESET}")
                print(f"{Color.CYAN}{Color.BOLD}📋 [{index}/{len(exam_data_list)}] Processing:{Color.RESET}")
                print(f"   {Color.BLUE}🎯 Exam:     {Color.GREEN}{exam_data.exam_name}{Color.RESET}")
                print(f"   {Color.BLUE}🆔 ID:       {Color.YELLOW}{exam_data.exam_id}{Color.RESET}")
                print(f"   {Color.BLUE}📅 Period:   {Color.MAGENTA}{exam_data.exam_start} {Color.BLUE}to{Color.MAGENTA} {exam_data.exam_end}{Color.RESET}")
                
                self._process_single_exam(exam_data)
                print(f"   {Color.GREEN}✅ Successfully processed{Color.RESET}")
                
            except Exception as e:
                print(f"   {Color.RED}❌ Failed: {e}{Color.RESET}")
                continue
        
        print(f"\n{Color.GREEN}{Color.BOLD}🎉 Time travel completed!{Color.RESET}")
    
    def _fetch_exam_data(self):
        """Fetch exam data from database"""
        connection_string = f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={self.db_path};"
        
        with pyodbc.connect(connection_string) as conn:
            cursor = conn.cursor()
            
            query = """
                SELECT DISTINCT 
                    se.exam_id,
                    se.exam_name,
                    se.exam_start,
                    se.exam_end
                FROM tbl_student_exams se
                INNER JOIN tbl_student_exam_results ser ON se.exam_id = ser.exam_id
                ORDER BY se.exam_start DESC
            """
            if self.level=="primary":
                query=query.replace("student","pupil")
                
            cursor.execute(query)
            results = cursor.fetchall()
            
            exam_data_list = []
            for row in results:
                exam_data = type('ExamData', (), {
                    'exam_id': str(row[0]),
                    'exam_name': str(row[1]),
                    'exam_start': row[2].strftime('%Y-%m-%d') if row[2] else 'N/A',
                    'exam_end': row[3].strftime('%Y-%m-%d') if row[3] else 'N/A'
                })()
                exam_data_list.append(exam_data)
            
            return exam_data_list
    
    def _process_single_exam(self, exam_data):
        """Process a single exam"""
        if self.level == "olevel":
            processor = OlevelProcessor(exam_id=exam_data.exam_id, db_path=self.db_path)
            processor.run()
        elif self.level == "alevel":
            processor = AlevelProcessor(exam_id=exam_data.exam_id, dbpath=self.db_path)
            processor.time_travel_export()
        elif self.level == "primary":
            processor=PrimaryProcessor(db_path=self.db_path,exam_id=exam_data.exam_id)
            processor.complete_exam()
        else:
            raise ValueError(f"Invalid level: {self.level}")


def into_the_future(level, db_path):
    """Facade function"""
    processor = TimeTravelProcessor(level, db_path)
    processor.into_the_future()