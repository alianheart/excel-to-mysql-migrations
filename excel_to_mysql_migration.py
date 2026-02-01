"""
Excel to MySQL Migration Script
================================
Migrates 21 tables from Excel (.xlsm) to MySQL with proper data types
and special character preservation (§, em dash, etc.)

Requirements:
    pip install pandas openpyxl mysql-connector-python sqlalchemy

Author: Data Migration Script
Date: 2025-01-30
"""

import pandas as pd
import mysql.connector
from sqlalchemy import create_engine, types
import warnings
import os
from datetime import datetime

warnings.filterwarnings('ignore')


class ExcelToMySQLMigrator:
    """
    Migrates Excel tables to MySQL with proper data type mapping
    """
    
    def __init__(self, excel_file, mysql_config, column_types_mapping=None):
        """
        Initialize the migrator
        
        Args:
            excel_file (str): Path to .xlsm file
            mysql_config (dict): MySQL connection configuration
            column_types_mapping (dict): Optional column name to MySQL type mapping
        """
        self.excel_file = excel_file
        self.mysql_config = mysql_config
        self.column_types_mapping = column_types_mapping or {}
        self.connection = None
        self.engine = None
        
    def connect_mysql(self):
        """Establish MySQL connection"""
        try:
            # Create SQLAlchemy engine
            connection_string = (
                f"mysql+mysqlconnector://{self.mysql_config['user']}:"
                f"{self.mysql_config['password']}@{self.mysql_config['host']}:"
                f"{self.mysql_config['port']}/{self.mysql_config['database']}"
                f"?charset=utf8mb4"  # Important for special characters
            )
            self.engine = create_engine(connection_string, echo=False)
            
            # Create direct connection for DDL operations
            self.connection = mysql.connector.connect(**self.mysql_config)
            
            print(f"✓ Connected to MySQL database: {self.mysql_config['database']}")
            return True
            
        except Exception as e:
            print(f"✗ MySQL connection failed: {e}")
            return False
    
    def read_excel_sheets(self):
        """Read all sheets from Excel file"""
        try:
            # Read Excel file with all sheets
            excel_file = pd.ExcelFile(self.excel_file, engine='openpyxl')
            sheet_names = excel_file.sheet_names
            
            print(f"\n✓ Found {len(sheet_names)} sheets in Excel file:")
            for idx, sheet in enumerate(sheet_names, 1):
                print(f"   {idx}. {sheet}")
            
            return excel_file, sheet_names
            
        except Exception as e:
            print(f"✗ Failed to read Excel file: {e}")
            return None, []
    
    def infer_mysql_type(self, series, column_name):
        """
        Infer MySQL data type from pandas Series
        
        Args:
            series: pandas Series
            column_name: Name of the column
            
        Returns:
            SQLAlchemy type object
        """
        # Check if type is explicitly defined in mapping
        if column_name in self.column_types_mapping:
            return self.column_types_mapping[column_name]
        
        # Auto-detect data type
        dtype = series.dtype
        
        # Integer types
        if pd.api.types.is_integer_dtype(dtype):
            max_val = series.max() if len(series) > 0 else 0
            if max_val < 127:
                return types.SMALLINT
            elif max_val < 32767:
                return types.INTEGER
            else:
                return types.BIGINT
        
        # Float types
        elif pd.api.types.is_float_dtype(dtype):
            return types.DECIMAL(precision=15, scale=4)
        
        # Boolean types
        elif pd.api.types.is_bool_dtype(dtype):
            return types.BOOLEAN
        
        # DateTime types
        elif pd.api.types.is_datetime64_any_dtype(dtype):
            return types.DATETIME
        
        # String types
        else:
            # Calculate max length for VARCHAR
            max_length = series.astype(str).str.len().max()
            if pd.isna(max_length) or max_length == 0:
                max_length = 255
            
            # Use TEXT for long strings
            if max_length > 5000:
                return types.TEXT
            elif max_length > 1000:
                return types.String(5000)
            else:
                # Add buffer to max_length
                return types.String(min(int(max_length * 1.5), 1000))
    
    def create_dtype_dict(self, df):
        """
        Create dtype dictionary for to_sql method
        
        Args:
            df: pandas DataFrame
            
        Returns:
            dict: Column name to SQLAlchemy type mapping
        """
        dtype_dict = {}
        for column in df.columns:
            dtype_dict[column] = self.infer_mysql_type(df[column], column)
        
        return dtype_dict
    
    def clean_table_name(self, sheet_name):
        """
        Convert sheet name to valid MySQL table name
        
        Args:
            sheet_name: Original sheet name
            
        Returns:
            str: Valid MySQL table name
        """
        # Replace spaces and special characters with underscores
        table_name = sheet_name.strip()
        table_name = table_name.replace(' ', '_')
        table_name = table_name.replace('-', '_')
        table_name = ''.join(c if c.isalnum() or c == '_' else '_' for c in table_name)
        
        # Ensure it starts with letter or underscore
        if table_name[0].isdigit():
            table_name = 'table_' + table_name
        
        # Convert to lowercase (MySQL convention)
        table_name = table_name.lower()
        
        # Truncate to 64 characters (MySQL limit)
        return table_name[:64]
    
    def migrate_sheet_to_table(self, excel_file, sheet_name):
        """
        Migrate a single Excel sheet to MySQL table
        
        Args:
            excel_file: ExcelFile object
            sheet_name: Name of the sheet to migrate
            
        Returns:
            bool: Success status
        """
        try:
            # Read sheet
            df = pd.read_excel(excel_file, sheet_name=sheet_name, engine='openpyxl')
            
            # Clean table name
            table_name = self.clean_table_name(sheet_name)
            
            # Get row and column count
            rows, cols = df.shape
            
            print(f"\nMigrating: '{sheet_name}' → '{table_name}'")
            print(f"   Rows: {rows:,} | Columns: {cols}")
            
            # Handle empty DataFrames
            if rows == 0:
                print(f"   ⚠ Skipping empty sheet")
                return False
            
            # Clean column names
            df.columns = [col.strip().replace(' ', '_').replace('-', '_') 
                         for col in df.columns]
            
            # Create dtype dictionary
            dtype_dict = self.create_dtype_dict(df)
            
            # Print column types
            print(f"   Column types:")
            for col, dtype in dtype_dict.items():
                print(f"      - {col}: {dtype}")
            
            # Write to MySQL
            df.to_sql(
                name=table_name,
                con=self.engine,
                if_exists='replace',
                index=False,
                dtype=dtype_dict,
                chunksize=1000,  # Insert in batches for large tables
                method='multi'   # Use executemany for better performance
            )
            
            print(f"   ✓ Successfully migrated {rows:,} rows")
            return True
            
        except Exception as e:
            print(f"   ✗ Migration failed: {e}")
            return False
    
    def migrate_all(self):
        """
        Migrate all Excel sheets to MySQL
        
        Returns:
            dict: Migration statistics
        """
        stats = {
            'total_sheets': 0,
            'successful': 0,
            'failed': 0,
            'total_rows': 0,
            'start_time': datetime.now()
        }
        
        print("="*70)
        print("EXCEL TO MYSQL MIGRATION")
        print("="*70)
        
        # Connect to MySQL
        if not self.connect_mysql():
            return stats
        
        # Read Excel file
        excel_file, sheet_names = self.read_excel_sheets()
        if not excel_file:
            return stats
        
        stats['total_sheets'] = len(sheet_names)
        
        # Migrate each sheet
        print("\n" + "="*70)
        print("MIGRATING SHEETS")
        print("="*70)
        
        for sheet_name in sheet_names:
            success = self.migrate_sheet_to_table(excel_file, sheet_name)
            if success:
                stats['successful'] += 1
            else:
                stats['failed'] += 1
        
        # Calculate duration
        stats['end_time'] = datetime.now()
        stats['duration'] = stats['end_time'] - stats['start_time']
        
        # Print summary
        self.print_summary(stats)
        
        # Close connections
        if self.connection:
            self.connection.close()
        
        return stats
    
    def print_summary(self, stats):
        """Print migration summary"""
        print("\n" + "="*70)
        print("MIGRATION SUMMARY")
        print("="*70)
        print(f"Total sheets:       {stats['total_sheets']}")
        print(f"Successful:         {stats['successful']} ✓")
        print(f"Failed:             {stats['failed']} ✗")
        print(f"Duration:           {stats['duration']}")
        print("="*70)
    
    def verify_migration(self):
        """Verify data was migrated correctly"""
        print("\n" + "="*70)
        print("VERIFICATION")
        print("="*70)
        
        try:
            cursor = self.connection.cursor()
            
            # Get all tables
            cursor.execute("SHOW TABLES")
            tables = cursor.fetchall()
            
            print(f"\nTables created in MySQL:")
            for idx, (table,) in enumerate(tables, 1):
                cursor.execute(f"SELECT COUNT(*) FROM `{table}`")
                count = cursor.fetchone()[0]
                print(f"   {idx}. {table}: {count:,} rows")
            
            cursor.close()
            
        except Exception as e:
            print(f"✗ Verification failed: {e}")


def create_column_type_mapping():
    """
    Create explicit column type mapping
    
    This function should be customized based on the client's requirements
    
    Returns:
        dict: Column name to SQLAlchemy type mapping
    """
    mapping = {
        # Example mappings (customize based on client's specification)
        # 'column_name': types.INTEGER,
        # 'price': types.DECIMAL(10, 2),
        # 'description': types.TEXT,
        # 'created_date': types.DATE,
        # 'is_active': types.BOOLEAN,
    }
    
    return mapping


def main():
    """Main execution function"""
    
    # =========================================================================
    # CONFIGURATION - UPDATE THESE VALUES
    # =========================================================================
    
    # Excel file path
    EXCEL_FILE = "data.xlsm"  # Update with actual file path
    
    # MySQL configuration
    MYSQL_CONFIG = {
        'host': 'localhost',
        'port': 3306,
        'user': 'root',          # Update with your MySQL username
        'password': 'password',   # Update with your MySQL password
        'database': 'survey_db',  # Update with your database name
        'charset': 'utf8mb4',     # Important for special characters (§, em dash)
        'collation': 'utf8mb4_unicode_ci'
    }
    
    # Optional: Explicit column type mapping
    # If client provides specific field types, add them here
    COLUMN_TYPES = create_column_type_mapping()
    
    # =========================================================================
    # EXECUTE MIGRATION
    # =========================================================================
    
    # Check if Excel file exists
    if not os.path.exists(EXCEL_FILE):
        print(f"✗ Excel file not found: {EXCEL_FILE}")
        print("Please update EXCEL_FILE path in the script")
        return
    
    # Create migrator instance
    migrator = ExcelToMySQLMigrator(
        excel_file=EXCEL_FILE,
        mysql_config=MYSQL_CONFIG,
        column_types_mapping=COLUMN_TYPES
    )
    
    # Run migration
    stats = migrator.migrate_all()
    
    # Verify migration
    if migrator.connection:
        migrator.verify_migration()
    
    print("\n✓ Migration complete!")
    print(f"✓ Check your MySQL database: {MYSQL_CONFIG['database']}")


if __name__ == "__main__":
    main()
