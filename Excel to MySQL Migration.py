import pandas as pd
import mysql.connector
from pathlib import Path
import re

# Configuration
CSV_FOLDER = 'C:/Users/User/Desktop/TensorFlow/Excell-Batch/Batches'
DATABASE_NAME = 'my_excel_data'

def sanitize_name(name, max_length=64):
    """Clean and truncate name for MySQL table/column"""
    # Remove special characters
    name = re.sub(r'[^a-zA-Z0-9_]', '_', name)
    name = re.sub(r'_+', '_', name).strip('_')
    name = name.lower()
    
    # Truncate if too long (MySQL limit is 64 chars)
    if len(name) > max_length:
        # Keep first 30 and last 30 chars with separator
        name = name[:30] + '_' + name[-(max_length-31):]
    
    return name

def detect_delimiter(file_path):
    """Detect the delimiter used in CSV file"""
    with open(file_path, 'r', encoding='utf-8-sig') as f:
        first_line = f.readline()
        
        # Count occurrences of common delimiters
        delimiters = {
            ',': first_line.count(','),
            ';': first_line.count(';'),
            '\t': first_line.count('\t'),
            '|': first_line.count('|')
        }
        
        # Return delimiter with most occurrences
        return max(delimiters, key=delimiters.get)

def create_column_mapping(df):
    """Create mapping of original to sanitized column names"""
    mapping = {}
    used_names = set()
    
    for i, col in enumerate(df.columns):
        # Clean the column name
        clean_name = sanitize_name(col, 60)  # Leave room for uniqueness
        
        # Ensure uniqueness
        if clean_name in used_names:
            clean_name = f"{clean_name[:55]}_{i}"  # Add index for uniqueness
        
        used_names.add(clean_name)
        mapping[col] = clean_name
    
    return mapping

print("="*60)
print("Fixed CSV Import to MySQL")
print("="*60)

# Connect to MySQL
try:
    conn = mysql.connector.connect(
        host='localhost',
        user='root',
        password='',
        allow_local_infile=True
    )
    cursor = conn.cursor()
    
    # Use database
    cursor.execute(f"USE {DATABASE_NAME}")
    print(f"Using database: {DATABASE_NAME}\n")
    
    # Get all CSV files
    csv_files = list(Path(CSV_FOLDER).glob('*.csv'))
    
    success_count = 0
    
    for csv_file in csv_files:
        print(f"\nProcessing: {csv_file.name}")
        print("-" * 40)
        
        try:
            # Detect delimiter
            delimiter = detect_delimiter(csv_file)
            delimiter_name = 'comma' if delimiter == ',' else 'semicolon' if delimiter == ';' else 'tab' if delimiter == '\t' else delimiter
            print(f"Detected delimiter: {delimiter_name}")
            
            # Read CSV with detected delimiter
            df = pd.read_csv(
                csv_file, 
                delimiter=delimiter,
                encoding='utf-8-sig',  # Handle BOM if present
                on_bad_lines='skip',   # Skip problematic lines
                engine='python'        # More flexible parser
            )
            
            print(f"✓ Read {len(df)} rows, {len(df.columns)} columns")
            
            # Clean table name
            table_name = sanitize_name(csv_file.stem, 64)
            
            # Create column mapping
            col_mapping = create_column_mapping(df)
            
            # Drop existing table
            cursor.execute(f"DROP TABLE IF EXISTS `{table_name}`")
            
            # Create table with proper column names
            columns = []
            for orig_col, clean_col in col_mapping.items():
                # Add comment with original name if truncated
                if len(orig_col) > 64:
                    columns.append(f"`{clean_col}` TEXT COMMENT '{orig_col[:250]}'")
                else:
                    columns.append(f"`{clean_col}` TEXT")
            
            create_query = f"""
            CREATE TABLE `{table_name}` (
                id INT AUTO_INCREMENT PRIMARY KEY,
                {', '.join(columns)}
            ) COMMENT='Imported from {csv_file.name}'
            """
            
            cursor.execute(create_query)
            print(f"✓ Created table: {table_name}")
            
            # Prepare data for insertion
            clean_cols = list(col_mapping.values())
            cols_str = ', '.join([f"`{col}`" for col in clean_cols])
            placeholders = ', '.join(['%s'] * len(clean_cols))
            
            # Insert data in batches
            batch_size = 100
            inserted = 0
            
            for start_idx in range(0, len(df), batch_size):
                end_idx = min(start_idx + batch_size, len(df))
                batch_df = df.iloc[start_idx:end_idx]
                
                batch_data = []
                for _, row in batch_df.iterrows():
                    values = [str(v)[:65535] if pd.notna(v) else None for v in row.values]  # Limit text length
                    batch_data.append(tuple(values))
                
                insert_query = f"INSERT INTO `{table_name}` ({cols_str}) VALUES ({placeholders})"
                cursor.executemany(insert_query, batch_data)
                inserted += len(batch_data)
                
                print(f"  Inserted batch: {inserted}/{len(df)} rows", end='\r')
            
            conn.commit()
            print(f"\n✓ Successfully inserted all {inserted} rows")
            
            # Show column mapping if any were truncated
            truncated = [orig for orig in col_mapping.keys() if len(orig) > 64]
            if truncated:
                print(f"  Note: {len(truncated)} column names were truncated")
            
            success_count += 1
            
        except Exception as e:
            print(f"✗ Error: {str(e)[:200]}")
            conn.rollback()
    
    # Show final summary
    cursor.execute("SHOW TABLES")
    tables = cursor.fetchall()
    
    print("\n" + "="*60)
    print("IMPORT COMPLETE!")
    print(f"Database: {DATABASE_NAME}")
    print(f"Successfully imported: {success_count}/{len(csv_files)} files")
    print(f"Tables in database: {len(tables)}")
    
    if tables:
        print("\nTable Summary:")
        for table in tables:
            cursor.execute(f"SELECT COUNT(*) FROM `{table[0]}`")
            count = cursor.fetchone()[0]
            
            cursor.execute(f"SELECT COUNT(*) FROM information_schema.columns WHERE table_schema = '{DATABASE_NAME}' AND table_name = '{table[0]}'")
            col_count = cursor.fetchone()[0] - 1  # Subtract 1 for id column
            
            print(f"  ✓ {table[0]}: {count} rows, {col_count} columns")
    
    print("="*60)
    print("\nView your data in phpMyAdmin:")
    print(f"http://localhost/phpmyadmin → {DATABASE_NAME}")
    
    cursor.close()
    conn.close()
    
except Exception as e:
    print(f"\n✗ Connection Error: {e}")
    print("\nTroubleshooting:")
    print("1. Make sure XAMPP MySQL is running")
    print("2. Check if database exists")