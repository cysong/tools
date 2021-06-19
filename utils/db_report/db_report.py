import os
import yaml
import mysql.connector
import openpyxl
from datetime import datetime

from utils.utils.excel import write_sheet

config = 'db.yml'
out_file = 'db_report_'
ymlFile = open(config, encoding='UTF8', mode='r')
configs = yaml.load(ymlFile)

# 文件保存路径
excel_path = os.path.abspath(
    os.path.join(os.getcwd(), out_file + datetime.now().strftime('%Y%m%d%H%M%S') + '.xlsx'))
# 查询table结构sql
t_sql = '''
SELECT t.TABLE_SCHEMA,t.TABLE_NAME,t.TABLE_ROWS,t.CREATE_TIME,t.`AUTO_INCREMENT`,t.TABLE_COMMENT FROM information_schema.TABLES t 
WHERE t.TABLE_SCHEMA = %s  
ORDER BY t.TABLE_SCHEMA,t.TABLE_NAME;
'''
# table sheet 表头
t_headers = ('TABLE_SCHEMA', 'TABLE_NAME', 'TABLE_ROWS', 'CREATE_TIME', 'AUTO_INCREMENT', 'TABLE_COMMENT')

# 查询column结构sql
c_sql = '''
SELECT t.TABLE_SCHEMA,t.TABLE_NAME,t.COLUMN_NAME,t.ORDINAL_POSITION,t.COLUMN_TYPE,t.IS_NULLABLE,t.COLUMN_DEFAULT,t.COLUMN_COMMENT FROM information_schema.`COLUMNS` t
WHERE t.TABLE_SCHEMA = %s
ORDER BY t.TABLE_SCHEMA,t.TABLE_NAME,t.ORDINAL_POSITION;
'''
# column sheet 表头
c_headers = (
    'TABLE_SCHEMA', 'TABLE_NAME', 'COLUMN_NAME', 'ORDINAL_POSITION', 'COLUMN_TYPE', 'IS_NULLABLE', 'COLUMN_DEFAULT',
    'COLUMN_COMMENT')

sheet_names = ['tables', 'columns']


def fetch_tables_def(host, port, user, password, database, schema, tables, columns):
    conn = mysql.connector.connect(host=host, port=port, user=user, password=password, database=database)

    cursor = conn.cursor()
    cursor.execute(t_sql, (schema,))
    values = cursor.fetchall()
    tables.extend(values)
    cursor.close()

    cursor = conn.cursor()
    cursor.execute(c_sql, (schema,))
    values = cursor.fetchall()
    columns.extend(values)
    cursor.close()

    conn.close()


def write_excel(tables, columns):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = sheet_names[0]
    write_sheet(sheet, t_headers, tables)

    sheet = workbook.create_sheet(sheet_names[1], 1)
    write_sheet(sheet, c_headers, columns)

    workbook.save(excel_path)


def main():
    tables = []
    columns = []
    for config in configs:
        print(config)
        host = config['DB_IP'].split(':')[0]
        port = config['DB_IP'].split(':')[1].split('/')[0]
        database = config['DB_IP'].split('/')[1]
        fetch_tables_def(host=host, port=port, user=config['DB_User'], password=config['DB_Password'],
                         database=database, schema=config['stress_Schema'], tables=tables, columns=columns)
    # print(tables[0])
    # print(columns[0])
    write_excel(tables, columns)


if __name__ == '__main__':
    main()
