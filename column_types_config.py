# Column Type Mapping Configuration
# ===================================
# This file defines explicit MySQL data types for each column
# Format: 'table_name.column_name': 'MYSQL_TYPE'

# Instructions:
# 1. Fill in based on client's field type specification
# 2. Import this in the main script
# 3. Uncomment and customize as needed

from sqlalchemy import types

COLUMN_TYPE_MAPPING = {
    # Example format:
    # 'table1.id': types.INTEGER,
    # 'table1.name': types.String(100),
    # 'table1.price': types.DECIMAL(10, 2),
    # 'table1.description': types.TEXT,
    # 'table1.created_date': types.DATE,
    # 'table1.created_datetime': types.DATETIME,
    # 'table1.is_active': types.BOOLEAN,
    # 'table1.amount': types.FLOAT,
    
    # Add your mappings here based on client specification
    # Format: 'column_name': types.TYPE
}

# MySQL Data Type Reference:
# ==========================
# types.INTEGER              - INT
# types.BIGINT              - BIGINT  
# types.SMALLINT            - SMALLINT
# types.String(255)         - VARCHAR(255)
# types.TEXT                - TEXT
# types.DECIMAL(10, 2)      - DECIMAL(10,2)
# types.FLOAT               - FLOAT
# types.DOUBLE              - DOUBLE
# types.BOOLEAN             - BOOLEAN/TINYINT(1)
# types.DATE                - DATE
# types.DATETIME            - DATETIME
# types.TIME                - TIME
# types.TIMESTAMP           - TIMESTAMP
