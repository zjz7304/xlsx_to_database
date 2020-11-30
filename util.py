import openpyxl,os,shutil
import pandas as pd
import pymssql

from sqlalchemy import create_engine

backupPath='files-backup'
def get_xlsx_to_dataframe(fliename):
    wb = openpyxl.load_workbook(fliename)
    sheets = wb.sheetnames
    ws = wb.get_sheet_by_name(sheets[0])

    df = pd.read_excel(fliename)

    if not os.path.exists(backupPath):
        os.mkdir(backupPath)
    shutil.move(fliename,os.path.join(backupPath,os.path.basename(fliename)))
    return df


def dataframe_to_mssql(df, tablename):
    engine = create_engine('mssql+pymssql://sa:123456@127.0.0.1/database?tds_version=7.0')
    df.to_sql(tablename, engine, if_exists='append', index=False)
