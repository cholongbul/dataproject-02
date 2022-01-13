import os
import shutil
import traceback
import re
import pandas as pd
import openpyxl
import xlrd

def sheet2cnt():
    path = 'C:\\Users\\admin\\Documents\\aa2\\D3014247\\'
    oddpath = 'C:\\Users\\admin\\Documents\\aa2\\시트이상\\'
    filelist = os.listdir(path)
    for file in filelist:

        try:
            if file.endswith('.xlsx') or file.endswith('.xlsm'):
                wb_data = openpyxl.load_workbook(path + file, read_only=True)
                sheets = wb_data.sheetnames

                for sheet in sheets:
                    if not sheet.startswith('C'):
                        pass
                    elif not re.match('^C[0-9]{1,}$',str(sheet)):
                        print(file)
                        print(str(sheet))
                        wb_data.close()
                        shutil.move(path + file,oddpath+file)
                        break
                    else:
                        wb_data.close()


            elif file.endswith('.xls'):
                workbook = xlrd.open_workbook(path + file)
                sheets = workbook.sheet_names()
                for sheet in sheets:
                    if not sheet.startswith('C'):
                        pass
                    elif not re.match('^C[0-9]{1,}$',str(sheet)):
                        print(file)
                        print(sheet)
                        shutil.move(path + file,oddpath+file)
                        break
        except:
            traceback.print_exc()
            print(file)

sheet2cnt()