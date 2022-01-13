import os
import shutil
import traceback
import re
import pandas as pd
import openpyxl
import xlrd

def sheet2cnt():
    path = 'C:\\Users\\admin\\Downloads\\SQL_DIAG_20220105\\컬럼개수차이\\'
    result_path = 'C:\\Users\\admin\\Downloads\\SQL_DIAG_20220105\\notable\\'
    t1_path = 'C:\\Users\\admin\\Downloads\\SQL_DIAG_20220105\\0105\\'
    filelist = os.listdir(path)
    csv_df = pd.read_csv('./id_organ_col.csv')
    for file in filelist:
        print(file)
        try:
            if file.split('_')[1].startswith('F1'):
                fileid = file.split('_')[1]
            elif file.split('_')[2].startswith('F1'):
                fileid = file.split('_')[2]
        except IndexError:
            fileid = file.split('.')[0]

        try:
            if file.endswith('.xlsx') or file.endswith('.xlsm'):
                wb_data = openpyxl.load_workbook(path + file, read_only=True)
                sheets = wb_data.sheetnames
                if len(sheets) == 21:
                    wb_data.close()
                    shutil.move(path + file, t1_path + file)
                    continue
                for sheet in sheets:
                    if not sheet.startswith('C'):
                        wb_data.close()
                        ex_df = pd.read_excel(path + file,sheet_name=sheet)
                        try:
                            ex_colcnt = len(ex_df['테이블명'].tolist())
                        except:
                            shutil.move(path + file, result_path + file)
                            break
                        is_same_id = csv_df['id'] == fileid
                        colcnt = csv_df[is_same_id]['col'].tolist()[0]
                        if not ex_colcnt==int(colcnt):
                            shutil.move(path + file, t1_path + file)
                        break



            elif file.endswith('.xls'):
                workbook = xlrd.open_workbook(path + file)
                sheets = workbook.sheet_names()
                if len(sheets) == 21:
                    shutil.move(path + file, t1_path + file)
                    continue
                for sheet in sheets:
                    if not sheet.startswith('C'):
                        ex_df = pd.read_excel(path + file, sheet_name=sheet)
                        try:
                            ex_colcnt = len(ex_df['테이블명'].tolist())
                        except:
                            shutil.move(path + file, result_path + file)
                            break
                        is_same_id = csv_df['id'] == fileid
                        colcnt = csv_df[is_same_id]['col'].tolist()[0]
                        if not ex_colcnt == int(colcnt):
                            shutil.move(path + file, t1_path + file)
                        break
        except:
            traceback.print_exc()
            print(file)

sheet2cnt()