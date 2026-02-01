# Excel to MySQL Migration Script

Complete solution for migrating 21 Excel tables to MySQL with proper data type handling and special character preservation.

## 📋 Features

✅ Migrates all 21 tables from Excel (.xlsm) to MySQL  
✅ Automatic data type inference (INT, VARCHAR, TEXT, DECIMAL, etc.)  
✅ Manual data type mapping support  
✅ Preserves special characters (§, em dash, unicode)  
✅ Handles tables from 9 to 32,090 rows  
✅ Handles columns from 3 to 23 columns  
✅ UTF-8 encoding for all special characters  
✅ Progress tracking and verification  

---

## 🚀 Quick Start

### Step 1: Install Dependencies
```bash
pip install -r requirements.txt
```

### Step 2: Prepare MySQL Database
```sql
CREATE DATABASE survey_db CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
```

### Step 3: Configure Script
Edit `excel_to_mysql_migration.py` - Update EXCEL_FILE path and MYSQL_CONFIG

### Step 4: Run Migration
```bash
python excel_to_mysql_migration.py
```

---

## 📊 Example Output

```
======================================================================
EXCEL TO MYSQL MIGRATION
======================================================================
✓ Connected to MySQL database: survey_db
✓ Found 21 sheets in Excel file

Migrating: 'Customer_Survey' → 'customer_survey'
   Rows: 32,090 | Columns: 23
   ✓ Successfully migrated 32,090 rows

======================================================================
MIGRATION SUMMARY
======================================================================
Total sheets:       21
Successful:         21 ✓
Failed:             0 ✗
Duration:           0:02:15
======================================================================
```

---

## 🎯 Custom Data Type Mapping

Edit `create_column_type_mapping()` function:

```python
def create_column_type_mapping():
    mapping = {
        'customer_id': types.INTEGER,
        'amount': types.DECIMAL(10, 2),
        'comments': types.TEXT,
    }
    return mapping
```

---

## 📤 Creating .mwb File

After migration:
1. Open MySQL Workbench
2. Database → Reverse Engineer
3. Select your database (survey_db)
4. Choose all 21 tables
5. File → Save Model As → `migration_model.mwb`

---

## 🔧 Troubleshooting

**Special characters not displaying:**
```sql
ALTER DATABASE survey_db CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
```

**Memory error with large tables:**
Script uses chunking automatically (1000 rows per batch)

**Excel file not found:**
Use absolute path: `/full/path/to/data.xlsm`

---

## ✅ Verification

```sql
-- Count tables
SELECT COUNT(*) FROM information_schema.tables WHERE table_schema = 'survey_db';

-- Check special characters
SELECT * FROM customer_survey WHERE comments LIKE '%§%' OR comments LIKE '%—%';
```

---

**Compatible with:** MySQL 5.7, 8.0, 9.6  
**Python Version:** 3.8+
