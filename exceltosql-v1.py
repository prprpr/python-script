import xlrd

excel_file_path = 'test.xls'
sql_template = "INSERT INTO tablename (field1, field2, field3) VALUES ('{field1_value}', '{field2_value}', {field3_value});"

workbook = xlrd.open_workbook(excel_file_path)
sheet = workbook.sheet_by_index(0)
rows = sheet.nrows

with open('data.sql', 'w') as sql_file:
    for row in range(1, rows):
        row_data = sheet.row_values(row)
        sql_str = sql_template.format(field1_value=row_data[5], field2_value=row_data[6], field3_value=row_data[8])
        sql_file.write(sql_str + '\n')

print('SQL脚本生成完毕！')
