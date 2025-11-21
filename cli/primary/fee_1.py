import pyodbc
import os
from pathlib import Path

# Configuration
DB_PATH = r"C:\Kiyabo App\backend\Kiyabo App Backend v3.0.0.accdb"
DROP_TABLES_IF_EXIST = True  # Set to False to skip dropping existing tables

# Table definitions in creation order (respecting dependencies)
# Using ODBC-compatible data types: COUNTER, INTEGER, VARCHAR, BIT, CURRENCY, DATETIME, DOUBLE
# IMPORTANT: Square brackets around ALL column names to avoid reserved word conflicts
TABLES = {
    'tbl_academic_years': """
        CREATE TABLE tbl_academic_years (
            [academic_year] INTEGER PRIMARY KEY,
            [is_active] BIT NOT NULL
        )
    """,
    
    'tbl_fee_apply': """
        CREATE TABLE tbl_fee_apply (
            [apply_to] VARCHAR(100)
        )
    """,
    
    'tbl_fee_type': """
        CREATE TABLE tbl_fee_type (
            [fee_type_id] COUNTER PRIMARY KEY,
            [fee_type] VARCHAR(100),
            [apply_to] VARCHAR(100),
            [is_adhoc] BIT NOT NULL,
            [note] VARCHAR(255)
        )
    """,
    
    'tbl_scholarhip_type': """
        CREATE TABLE tbl_scholarhip_type (
            [scholarship_type_id] COUNTER PRIMARY KEY,
            [scholarship_type] VARCHAR(255)
        )
    """,
    
    'tbl_fee_structure': """
        CREATE TABLE tbl_fee_structure (
            [structure_id] COUNTER PRIMARY KEY,
            [fee_type_id] INTEGER,
            [class_id] VARCHAR(20),
            [academic_year] INTEGER,
            [amount] CURRENCY,
            [notes] VARCHAR(255)
        )
    """,
    
    'tbl_fee_collection': """
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
    
    'tbl_scholarships': """
        CREATE TABLE tbl_scholarships (
            [scholarship_id] COUNTER PRIMARY KEY,
            [pupil_id] VARCHAR(255),
            [structure_id] INTEGER,
            [scholarship_type_id] INTEGER,
            [academic_year] INTEGER,
            [amount] CURRENCY,
            [percentage] DOUBLE
        )
    """
}

# Foreign key constraints (added after tables are created)
FOREIGN_KEYS = [
    # tbl_fee_structure relationships
    {
        'name': 'FK_fee_structure_fee_type',
        'table': 'tbl_fee_structure',
        'sql': """
        ALTER TABLE [tbl_fee_structure] 
        ADD CONSTRAINT [FK_fee_structure_fee_type] 
        FOREIGN KEY ([fee_type_id]) REFERENCES [tbl_fee_type]([fee_type_id])
        """
    },
    {
        'name': 'FK_fee_structure_class',
        'table': 'tbl_fee_structure',
        'sql': """
        ALTER TABLE [tbl_fee_structure] 
        ADD CONSTRAINT [FK_fee_structure_class] 
        FOREIGN KEY ([class_id]) REFERENCES [tbl_classes]([class_id])
        """
    },
    {
        'name': 'FK_fee_structure_academic_year',
        'table': 'tbl_fee_structure',
        'sql': """
        ALTER TABLE [tbl_fee_structure] 
        ADD CONSTRAINT [FK_fee_structure_academic_year] 
        FOREIGN KEY ([academic_year]) REFERENCES [tbl_academic_years]([academic_year])
        """
    },
    
    # tbl_fee_collection relationships
    {
        'name': 'FK_fee_collection_pupil',
        'table': 'tbl_fee_collection',
        'sql': """
        ALTER TABLE [tbl_fee_collection] 
        ADD CONSTRAINT [FK_fee_collection_pupil] 
        FOREIGN KEY ([pupil_id]) REFERENCES [tbl_pupil_academic_info]([pupil_id])
        """
    },
    {
        'name': 'FK_fee_collection_structure',
        'table': 'tbl_fee_collection',
        'sql': """
        ALTER TABLE [tbl_fee_collection] 
        ADD CONSTRAINT [FK_fee_collection_structure] 
        FOREIGN KEY ([structure_id]) REFERENCES [tbl_fee_structure]([structure_id])
        """
    },
    
    # tbl_scholarships relationships
    {
        'name': 'FK_scholarships_pupil',
        'table': 'tbl_scholarships',
        'sql': """
        ALTER TABLE [tbl_scholarships] 
        ADD CONSTRAINT [FK_scholarships_pupil] 
        FOREIGN KEY ([pupil_id]) REFERENCES [tbl_pupil_academic_info]([pupil_id])
        """
    },
    {
        'name': 'FK_scholarships_structure',
        'table': 'tbl_scholarships',
        'sql': """
        ALTER TABLE [tbl_scholarships] 
        ADD CONSTRAINT [FK_scholarships_structure] 
        FOREIGN KEY ([structure_id]) REFERENCES [tbl_fee_structure]([structure_id])
        """
    },
    {
        'name': 'FK_scholarships_type',
        'table': 'tbl_scholarships',
        'sql': """
        ALTER TABLE [tbl_scholarships] 
        ADD CONSTRAINT [FK_scholarships_type] 
        FOREIGN KEY ([scholarship_type_id]) REFERENCES [tbl_scholarhip_type]([scholarship_type_id])
        """
    },
    {
        'name': 'FK_scholarships_academic_year',
        'table': 'tbl_scholarships',
        'sql': """
        ALTER TABLE [tbl_scholarships] 
        ADD CONSTRAINT [FK_scholarships_academic_year] 
        FOREIGN KEY ([academic_year]) REFERENCES [tbl_academic_years]([academic_year])
        """
    }
]

def get_connection(db_path):
    """Create connection to Access database"""
    if not os.path.exists(db_path):
        raise FileNotFoundError(f"Database file not found: {db_path}")
    
    # Try different Access drivers
    drivers = [
        'Microsoft Access Driver (*.mdb, *.accdb)',
        'Microsoft Access Driver (*.accdb)',
        'Microsoft Access Driver (*.mdb)',
    ]
    
    conn = None
    last_error = None
    
    for driver in drivers:
        try:
            conn_str = f'Driver={{{driver}}};DBQ={db_path};'
            conn = pyodbc.connect(conn_str)
            print(f"✓ Connected using driver: {driver}")
            return conn
        except pyodbc.Error as e:
            last_error = e
            continue
    
    if conn is None:
        raise Exception(f"Could not connect with any driver. Last error: {last_error}")

def drop_table(cursor, table_name):
    """Drop a table if it exists"""
    try:
        cursor.execute(f"DROP TABLE [{table_name}]")
        print(f"✓ Dropped table: {table_name}")
        return True
    except pyodbc.Error as e:
        error_msg = str(e).lower()
        if "does not exist" in error_msg or "could not find" in error_msg or "unknown" in error_msg:
            print(f"  Table {table_name} does not exist, skipping drop")
        else:
            print(f"  Warning dropping {table_name}: {e}")
        return False

def drop_foreign_key(cursor, table_name, fk_name):
    """Drop a foreign key constraint"""
    try:
        cursor.execute(f"ALTER TABLE [{table_name}] DROP CONSTRAINT [{fk_name}]")
        print(f"✓ Dropped foreign key: {fk_name} from {table_name}")
        return True
    except pyodbc.Error as e:
        error_msg = str(e).lower()
        if "does not exist" in error_msg or "could not find" in error_msg or "unknown" in error_msg:
            pass  # FK doesn't exist
        else:
            print(f"  Warning dropping FK {fk_name}: {e}")
        return False

def drop_all_foreign_keys(cursor):
    """Drop all foreign key constraints"""
    print("\n=== Dropping existing foreign keys ===")
    for fk in FOREIGN_KEYS:
        drop_foreign_key(cursor, fk['table'], fk['name'])

def create_schema():
    """Main function to create database schema"""
    try:
        print(f"Connecting to database: {DB_PATH}")
        conn = get_connection(DB_PATH)
        cursor = conn.cursor()
        
        # Drop tables if requested
        if DROP_TABLES_IF_EXIST:
            drop_all_foreign_keys(cursor)
            conn.commit()
            
            print("\n=== Dropping existing tables ===")
            # Drop in reverse order to respect dependencies
            for table_name in reversed(list(TABLES.keys())):
                drop_table(cursor, table_name)
            conn.commit()
        
        # Create tables
        print("\n=== Creating tables ===")
        for table_name, create_sql in TABLES.items():
            try:
                cursor.execute(create_sql)
                print(f"✓ Created table: {table_name}")
            except pyodbc.Error as e:
                print(f"✗ Error creating {table_name}: {e}")
                raise
        
        conn.commit()
        print("✓ All tables committed successfully")
        
        # Create foreign keys
        print("\n=== Creating foreign key relationships ===")
        fk_success = 0
        fk_failed = 0
        failed_fks = []
        
        for fk in FOREIGN_KEYS:
            try:
                cursor.execute(fk['sql'])
                print(f"✓ Created foreign key: {fk['name']}")
                fk_success += 1
            except pyodbc.Error as e:
                print(f"✗ Error creating {fk['name']}: {e}")
                failed_fks.append(fk['name'])
                fk_failed += 1
                # Continue with other FKs even if one fails
        
        conn.commit()
        print("✓ All foreign keys committed successfully")
        
        print("\n" + "=" * 60)
        print("Schema creation completed!")
        print("=" * 60)
        print(f"\nTables created: {len(TABLES)}")
        print(f"Foreign keys created: {fk_success}")
        
        if fk_failed > 0:
            print(f"Foreign keys failed: {fk_failed}")
            print(f"Failed FKs: {', '.join(failed_fks)}")
            print("\n⚠ Note: Some foreign keys may have failed because")
            print("referenced tables (tbl_classes, tbl_pupil_academic_info)")
            print("don't exist yet. Create them first if needed.")
        else:
            print("\n✓ All operations completed successfully!")
        
        cursor.close()
        conn.close()
        
    except FileNotFoundError as e:
        print(f"\n✗ Error: {e}")
    except pyodbc.Error as e:
        print(f"\n✗ Database error: {e}")
        print("\nTroubleshooting:")
        print("1. Ensure the Access ODBC driver is installed")
        print("2. Check if the database file exists")
        print("3. Make sure the database is not open in Access")
        print("4. Verify you have write permissions to the file")
    except Exception as e:
        print(f"\n✗ Unexpected error: {e}")
        raise

if __name__ == "__main__":
    print("=" * 60)
    print("Fee Management System - Database Schema Creator")
    print("MS Access via pyodbc ODBC Driver")
    print("=" * 60)
    print(f"\nDatabase: {DB_PATH}")
    print(f"Drop existing tables: {DROP_TABLES_IF_EXIST}")
    print("\nODBC-Compatible Data Types:")
    print("  - COUNTER (AutoNumber)")
    print("  - INTEGER (Long Integer)")
    print("  - VARCHAR (Text)")
    print("  - BIT (Yes/No)")
    print("  - CURRENCY (Currency)")
    print("  - DATETIME (Date/Time)")
    print("  - DOUBLE (Double)")
    print("\n" + "=" * 60)
    
    response = input("\nProceed with schema creation? (yes/no): ")
    if response.lower() in ['yes', 'y']:
        create_schema()
    else:
        print("Operation cancelled.")