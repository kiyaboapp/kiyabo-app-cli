import pyodbc
import os
from pathlib import Path
from colorama import init, Fore, Back, Style
from tabulate import tabulate

# Initialize colorama for Windows color support
init(autoreset=True)

# Configuration
DB_PATH = r"C:\Kiyabo App\backend\Kiyabo App Backend v3.0.0.accdb"
DROP_TABLES_IF_EXIST = True  # Set to False to skip dropping existing tables

# Table definitions in creation order (respecting dependencies)
TABLES = {
    'tbl_academic_years': {
        'sql': """
            CREATE TABLE tbl_academic_years (
                [academic_year] INTEGER PRIMARY KEY,
                [is_active] BIT NOT NULL
            )
        """,
        'fields': [
            ['academic_year', 'INTEGER', 'PRIMARY KEY', 'Academic year number'],
            ['is_active', 'BIT', 'NOT NULL', 'Currently active year']
        ]
    },
    
    'tbl_fee_apply': {
        'sql': """
            CREATE TABLE tbl_fee_apply (
                [apply_to] VARCHAR(100)
            )
        """,
        'fields': [
            ['apply_to', 'VARCHAR(100)', '', 'Fee application category']
        ]
    },
    
    'tbl_fee_type': {
        'sql': """
            CREATE TABLE tbl_fee_type (
                [fee_type_id] COUNTER PRIMARY KEY,
                [fee_type] VARCHAR(100),
                [apply_to] VARCHAR(100),
                [is_adhoc] BIT NOT NULL,
                [note] VARCHAR(255)
            )
        """,
        'fields': [
            ['fee_type_id', 'COUNTER', 'PRIMARY KEY', 'Auto-increment ID'],
            ['fee_type', 'VARCHAR(100)', '', 'Type of fee'],
            ['apply_to', 'VARCHAR(100)', '', 'Who fee applies to'],
            ['is_adhoc', 'BIT', 'NOT NULL', 'Is ad-hoc fee'],
            ['note', 'VARCHAR(255)', '', 'Additional notes']
        ]
    },
    
    'tbl_scholarhip_type': {
        'sql': """
            CREATE TABLE tbl_scholarhip_type (
                [scholarship_type_id] COUNTER PRIMARY KEY,
                [scholarship_type] VARCHAR(255)
            )
        """,
        'fields': [
            ['scholarship_type_id', 'COUNTER', 'PRIMARY KEY', 'Auto-increment ID'],
            ['scholarship_type', 'VARCHAR(255)', '', 'Scholarship category']
        ]
    },
    
    'tbl_fee_structure': {
        'sql': """
            CREATE TABLE tbl_fee_structure (
                [structure_id] COUNTER PRIMARY KEY,
                [fee_type_id] INTEGER,
                [class_id] VARCHAR(20),
                [academic_year] INTEGER,
                [amount] CURRENCY,
                [notes] VARCHAR(255)
            )
        """,
        'fields': [
            ['structure_id', 'COUNTER', 'PRIMARY KEY', 'Auto-increment ID'],
            ['fee_type_id', 'INTEGER', 'FK', 'References tbl_fee_type'],
            ['class_id', 'VARCHAR(20)', 'FK', 'References tbl_classes'],
            ['academic_year', 'INTEGER', 'FK', 'References tbl_academic_years'],
            ['amount', 'CURRENCY', '', 'Fee amount'],
            ['notes', 'VARCHAR(255)', '', 'Additional notes']
        ]
    },
    
    'tbl_fee_collection': {
        'sql': """
            CREATE TABLE tbl_fee_collection (
                [collection_id] COUNTER PRIMARY KEY,
                [pupil_id] VARCHAR(255),
                [structure_id] INTEGER,
                [amount] CURRENCY,
                [date_filled] DATETIME,
                [date_paid] DATETIME,
                [notes] VARCHAR(255)
            )
        """,
        'fields': [
            ['collection_id', 'COUNTER', 'PRIMARY KEY', 'Auto-increment ID'],
            ['pupil_id', 'VARCHAR(255)', 'FK', 'References tbl_pupil_academic_info'],
            ['structure_id', 'INTEGER', 'FK', 'References tbl_fee_structure'],
            ['amount', 'CURRENCY', '', 'Amount collected'],
            ['date_filled', 'DATETIME', '', 'Date invoice filled'],
            ['date_paid', 'DATETIME', '', 'Date payment received'],
            ['notes', 'VARCHAR(255)', '', 'Payment notes']
        ]
    },
    
    'tbl_scholarships': {
        'sql': """
            CREATE TABLE tbl_scholarships (
                [scholarship_id] COUNTER PRIMARY KEY,
                [pupil_id] VARCHAR(255),
                [structure_id] INTEGER,
                [scholarship_type_id] INTEGER,
                [academic_year] INTEGER,
                [amount] CURRENCY,
                [percentage] DOUBLE
            )
        """,
        'fields': [
            ['scholarship_id', 'COUNTER', 'PRIMARY KEY', 'Auto-increment ID'],
            ['pupil_id', 'VARCHAR(255)', 'FK', 'References tbl_pupil_academic_info'],
            ['structure_id', 'INTEGER', 'FK', 'References tbl_fee_structure'],
            ['scholarship_type_id', 'INTEGER', 'FK', 'References tbl_scholarhip_type'],
            ['academic_year', 'INTEGER', 'FK', 'References tbl_academic_years'],
            ['amount', 'CURRENCY', '', 'Scholarship amount'],
            ['percentage', 'DOUBLE', '', 'Discount percentage']
        ]
    }
}

# Foreign key constraints
FOREIGN_KEYS = [
    {'name': 'FK_fee_structure_fee_type', 'table': 'tbl_fee_structure',
     'sql': 'ALTER TABLE [tbl_fee_structure] ADD CONSTRAINT [FK_fee_structure_fee_type] FOREIGN KEY ([fee_type_id]) REFERENCES [tbl_fee_type]([fee_type_id])'},
    {'name': 'FK_fee_structure_class', 'table': 'tbl_fee_structure',
     'sql': 'ALTER TABLE [tbl_fee_structure] ADD CONSTRAINT [FK_fee_structure_class] FOREIGN KEY ([class_id]) REFERENCES [tbl_classes]([class_id])'},
    {'name': 'FK_fee_structure_academic_year', 'table': 'tbl_fee_structure',
     'sql': 'ALTER TABLE [tbl_fee_structure] ADD CONSTRAINT [FK_fee_structure_academic_year] FOREIGN KEY ([academic_year]) REFERENCES [tbl_academic_years]([academic_year])'},
    {'name': 'FK_fee_collection_pupil', 'table': 'tbl_fee_collection',
     'sql': 'ALTER TABLE [tbl_fee_collection] ADD CONSTRAINT [FK_fee_collection_pupil] FOREIGN KEY ([pupil_id]) REFERENCES [tbl_pupil_academic_info]([pupil_id])'},
    {'name': 'FK_fee_collection_structure', 'table': 'tbl_fee_collection',
     'sql': 'ALTER TABLE [tbl_fee_collection] ADD CONSTRAINT [FK_fee_collection_structure] FOREIGN KEY ([structure_id]) REFERENCES [tbl_fee_structure]([structure_id])'},
    {'name': 'FK_scholarships_pupil', 'table': 'tbl_scholarships',
     'sql': 'ALTER TABLE [tbl_scholarships] ADD CONSTRAINT [FK_scholarships_pupil] FOREIGN KEY ([pupil_id]) REFERENCES [tbl_pupil_academic_info]([pupil_id])'},
    {'name': 'FK_scholarships_structure', 'table': 'tbl_scholarships',
     'sql': 'ALTER TABLE [tbl_scholarships] ADD CONSTRAINT [FK_scholarships_structure] FOREIGN KEY ([structure_id]) REFERENCES [tbl_fee_structure]([structure_id])'},
    {'name': 'FK_scholarships_type', 'table': 'tbl_scholarships',
     'sql': 'ALTER TABLE [tbl_scholarships] ADD CONSTRAINT [FK_scholarships_type] FOREIGN KEY ([scholarship_type_id]) REFERENCES [tbl_scholarhip_type]([scholarship_type_id])'},
    {'name': 'FK_scholarships_academic_year', 'table': 'tbl_scholarships',
     'sql': 'ALTER TABLE [tbl_scholarships] ADD CONSTRAINT [FK_scholarships_academic_year] FOREIGN KEY ([academic_year]) REFERENCES [tbl_academic_years]([academic_year])'}
]

def print_header(text):
    """Print a colored header"""
    print(f"\n{Fore.CYAN}{Style.BRIGHT}{'=' * 80}")
    print(f"{Fore.CYAN}{Style.BRIGHT}{text.center(80)}")
    print(f"{Fore.CYAN}{Style.BRIGHT}{'=' * 80}{Style.RESET_ALL}\n")

def print_section(text):
    """Print a section header"""
    print(f"\n{Fore.YELLOW}{Style.BRIGHT}▶ {text}{Style.RESET_ALL}")
    print(f"{Fore.YELLOW}{'─' * 80}{Style.RESET_ALL}")

def print_success(text):
    """Print success message"""
    print(f"{Fore.GREEN}✓ {text}{Style.RESET_ALL}")

def print_error(text):
    """Print error message"""
    print(f"{Fore.RED}✗ {text}{Style.RESET_ALL}")

def print_warning(text):
    """Print warning message"""
    print(f"{Fore.YELLOW}⚠ {text}{Style.RESET_ALL}")

def print_info(text):
    """Print info message"""
    print(f"{Fore.BLUE}ℹ {text}{Style.RESET_ALL}")

def print_table_structure(table_name, fields):
    """Print table structure in a beautiful format"""
    print(f"\n{Fore.MAGENTA}{Style.BRIGHT}📋 Table: {table_name}{Style.RESET_ALL}")
    headers = [f"{Fore.CYAN}Field Name{Style.RESET_ALL}", 
               f"{Fore.CYAN}Data Type{Style.RESET_ALL}", 
               f"{Fore.CYAN}Constraints{Style.RESET_ALL}", 
               f"{Fore.CYAN}Description{Style.RESET_ALL}"]
    print(tabulate(fields, headers=headers, tablefmt="fancy_grid"))

def get_connection(db_path):
    """Create connection to Access database"""
    if not os.path.exists(db_path):
        raise FileNotFoundError(f"Database file not found: {db_path}")
    
    drivers = [
        'Microsoft Access Driver (*.mdb, *.accdb)',
        'Microsoft Access Driver (*.accdb)',
        'Microsoft Access Driver (*.mdb)',
    ]
    
    for driver in drivers:
        try:
            conn_str = f'Driver={{{driver}}};DBQ={db_path};'
            conn = pyodbc.connect(conn_str)
            print_success(f"Connected using driver: {driver}")
            return conn
        except pyodbc.Error:
            continue
    
    raise Exception("Could not connect with any driver")

def table_exists(cursor, table_name):
    """Check if a table exists in the database"""
    try:
        cursor.execute(f"SELECT TOP 1 * FROM [{table_name}]")
        return True
    except pyodbc.Error:
        return False

def get_existing_tables(cursor):
    """Get list of existing tables"""
    existing = []
    for table_name in TABLES.keys():
        if table_exists(cursor, table_name):
            existing.append(table_name)
    return existing

def drop_table(cursor, table_name):
    """Drop a table if it exists"""
    try:
        cursor.execute(f"DROP TABLE [{table_name}]")
        print_success(f"Dropped table: {table_name}")
        return True
    except pyodbc.Error as e:
        error_msg = str(e).lower()
        if "does not exist" in error_msg or "could not find" in error_msg:
            print_info(f"Table {table_name} does not exist")
        else:
            print_warning(f"Could not drop {table_name}: {e}")
        return False

def drop_foreign_key(cursor, table_name, fk_name):
    """Drop a foreign key constraint"""
    try:
        cursor.execute(f"ALTER TABLE [{table_name}] DROP CONSTRAINT [{fk_name}]")
        return True
    except pyodbc.Error:
        return False

def create_schema():
    """Main function to create database schema"""
    try:
        print_header("🎓 FEE MANAGEMENT SYSTEM - DATABASE SCHEMA CREATOR")
        
        print(f"{Fore.WHITE}Database Path: {Fore.GREEN}{DB_PATH}{Style.RESET_ALL}")
        print(f"{Fore.WHITE}Drop Existing: {Fore.GREEN if DROP_TABLES_IF_EXIST else Fore.RED}{DROP_TABLES_IF_EXIST}{Style.RESET_ALL}\n")
        
        conn = get_connection(DB_PATH)
        cursor = conn.cursor()
        
        # Check existing tables
        existing_tables = get_existing_tables(cursor)
        
        if existing_tables and not DROP_TABLES_IF_EXIST:
            print_section("📊 EXISTING TABLES SUMMARY")
            print_info(f"Found {len(existing_tables)} existing tables:")
            for table_name in existing_tables:
                print(f"  {Fore.GREEN}✓{Style.RESET_ALL} {table_name}")
                print_table_structure(table_name, TABLES[table_name]['fields'])
            
            print_warning("\nTables already exist and DROP_TABLES_IF_EXIST is False")
            print_info("Skipping table creation. Set DROP_TABLES_IF_EXIST=True to recreate.")
            
            cursor.close()
            conn.close()
            return
        
        # Drop foreign keys first
        if DROP_TABLES_IF_EXIST and existing_tables:
            print_section("🗑️  DROPPING FOREIGN KEY CONSTRAINTS")
            for fk in FOREIGN_KEYS:
                if drop_foreign_key(cursor, fk['table'], fk['name']):
                    print_success(f"Dropped FK: {fk['name']}")
            conn.commit()
        
        # Drop tables
        if DROP_TABLES_IF_EXIST:
            print_section("🗑️  DROPPING EXISTING TABLES")
            for table_name in reversed(list(TABLES.keys())):
                drop_table(cursor, table_name)
            conn.commit()
        
        # Create tables
        print_section("🏗️  CREATING TABLES")
        tables_created = 0
        
        for table_name, table_info in TABLES.items():
            try:
                cursor.execute(table_info['sql'])
                print_success(f"Created table: {table_name}")
                print_table_structure(table_name, table_info['fields'])
                tables_created += 1
            except pyodbc.Error as e:
                print_error(f"Error creating {table_name}: {e}")
                raise
        
        conn.commit()
        print_success("All tables committed successfully!")
        
        # Create foreign keys
        print_section("🔗 CREATING FOREIGN KEY RELATIONSHIPS")
        fk_success = []
        fk_failed = []
        
        for fk in FOREIGN_KEYS:
            try:
                cursor.execute(fk['sql'])
                print_success(f"Created FK: {fk['name']}")
                fk_success.append(fk['name'])
            except pyodbc.Error as e:
                print_error(f"Failed FK: {fk['name']}")
                print(f"  {Fore.RED}└─ {str(e)[:100]}{Style.RESET_ALL}")
                fk_failed.append(fk['name'])
        
        conn.commit()
        
        # Final summary
        print_header("📊 SCHEMA CREATION SUMMARY")
        
        summary_data = [
            [f"{Fore.GREEN}Tables Created{Style.RESET_ALL}", f"{Fore.GREEN}{Style.BRIGHT}{tables_created}{Style.RESET_ALL}"],
            [f"{Fore.GREEN}Foreign Keys Created{Style.RESET_ALL}", f"{Fore.GREEN}{Style.BRIGHT}{len(fk_success)}{Style.RESET_ALL}"],
            [f"{Fore.RED}Foreign Keys Failed{Style.RESET_ALL}", f"{Fore.RED}{Style.BRIGHT}{len(fk_failed)}{Style.RESET_ALL}"]
        ]
        print(tabulate(summary_data, tablefmt="fancy_grid"))
        
        if fk_failed:
            print(f"\n{Fore.YELLOW}Failed Foreign Keys:{Style.RESET_ALL}")
            for fk_name in fk_failed:
                print(f"  {Fore.RED}✗{Style.RESET_ALL} {fk_name}")
            print_warning("\nSome FKs failed because referenced tables may not exist yet")
            print_info("Create tbl_classes and tbl_pupil_academic_info first if needed")
        else:
            print(f"\n{Fore.GREEN}{Style.BRIGHT}🎉 ALL OPERATIONS COMPLETED SUCCESSFULLY!{Style.RESET_ALL}")
        
        cursor.close()
        conn.close()
        
    except FileNotFoundError as e:
        print_error(f"Database file not found: {e}")
    except pyodbc.Error as e:
        print_error(f"Database error: {e}")
        print_info("\nTroubleshooting:")
        print("  1. Ensure Access ODBC driver is installed")
        print("  2. Close the database if open in Access")
        print("  3. Check file permissions")
    except Exception as e:
        print_error(f"Unexpected error: {e}")
        raise

if __name__ == "__main__":
    print_header("🎓 FEE MANAGEMENT SYSTEM")
    print(f"{Fore.CYAN}MS Access Database Schema Creator{Style.RESET_ALL}")
    print(f"{Fore.CYAN}Version 1.0.0{Style.RESET_ALL}\n")
    
    print(f"{Fore.WHITE}Configuration:{Style.RESET_ALL}")
    print(f"  📁 Database: {Fore.YELLOW}{DB_PATH}{Style.RESET_ALL}")
    print(f"  🗑️  Drop Tables: {Fore.GREEN if DROP_TABLES_IF_EXIST else Fore.RED}{DROP_TABLES_IF_EXIST}{Style.RESET_ALL}\n")
    
    response = input(f"{Fore.CYAN}Proceed with schema creation? (yes/no): {Style.RESET_ALL}")
    if response.lower() in ['yes', 'y']:
        create_schema()
    else:
        print_warning("Operation cancelled by user")



"""
works fine but few changes:

apply_to first values are all,day,boarding
also apply_to in tbl_fee_type should be foreign key to tbl_fee_apply.apply_to
academi year values fornthe first time should be last year, current year,next year
"""