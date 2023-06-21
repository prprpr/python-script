import argparse
import pandas as pd
import os

def get_excel_files(file_path):
    if os.path.isdir(file_path):
        # If file path is a directory, get all Excel files in the directory
        excel_files = [os.path.join(file_path, f) for f in os.listdir(file_path) if f.endswith(('.xls', '.xlsx'))]
    else:
        # If file path is a file, convert it to a list
        excel_files = [file_path]
    return excel_files

def generate_sql_file(excel_file, args):
    if not os.path.exists(excel_file):
        print(f"File not found: {excel_file}")
        return

    # Read Excel file
    df = pd.read_excel(excel_file, sheet_name=args.sheet, header=args.row - 1)

    # Generate SQL statements and write to file
    output_dir = args.output or os.path.dirname(excel_file)
    sql_file_path = os.path.join(output_dir, os.path.splitext(os.path.basename(excel_file))[0] + ".sql")
    with open(sql_file_path, "w", encoding="utf-8") as f:
        for _, row in df.iterrows():
            field1 = row["st_name"]
            field2 = row["studentid"]
            field3 = row["mobileno"]
            # sql = f"INSERT INTO {args.table} (field1, field2, field3) VALUES ('{field1}', {field2}, '{field3}');\n"
            f.write(sql)
            sql = f"INSERT INTO sys_user (field1, field2, field3) VALUES ('{field1}', {field2}, '{field3}');\n"
            f.write(sql)

    print(f"Generated SQL file for {excel_file}: {sql_file_path}")

# Define command line arguments
parser = argparse.ArgumentParser()
parser.add_argument("file_paths", nargs='*', help="List of Excel file paths or directory containing Excel files, e.g. C:\\test.xlsx D:\\data.xlsx or C:\\excelpath\\")
parser.add_argument("-s", "--sheet", help="Excel file sheet name, default is Sheet1", default="Sheet1")
#parser.add_argument("-t", "--table", help="Target table name, default is tablename", default="tablename")
parser.add_argument("-r", "--row", type=int, help="Row number for the column names, default is 2", default=1)
parser.add_argument("-o", "--output", help="Output directory for SQL files, default is the same as input file directory")
args = parser.parse_args()

# Generate SQL statements for each file
for file_path in args.file_paths:
    excel_files = get_excel_files(file_path)
    for excel_file in excel_files:
        generate_sql_file(excel_file, args)
