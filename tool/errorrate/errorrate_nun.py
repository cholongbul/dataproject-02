import pandas as pd
import os
import win32com.client
import traceback
import cx_Oracle
from sqlalchemy import create_engine
import openpyxl
path = "C:\\Users\\admin\\Documents\\1.API\\2.오픈API종합보고서\\새 폴더\\"
error_list = open('./errorlog1.txt','r',encoding='utf-8').readlines()
report_list = os.listdir(path)
result_df = pd.DataFrame(index=range(0,0), columns=['1'])
tatalcnt = 0
normalcnt = 0
for report in report_list:
    log = open('./api수준평가미대상_문서없음.csv', 'a', encoding='utf8')
    # win32com #excel 사용할 수 있게 설정
    excel = win32com.client.Dispatch("Excel.Application")
    #임시 Workbook 객체 생성 및 엑셀 열기
    temp_wb = excel.Workbooks.Open(path + report)
    #저장
    temp_wb.Save()
    #excel 종료
    excel.quit()
    print(report)
    organnm = report.rstrip('_종합.xlsx')
    wb_data = openpyxl.load_workbook(path + report,data_only=True)

    ws_data = wb_data['00.수준평가']
    organnm = ws_data['C5'].value
    tatalcnt = tatalcnt + ws_data['D18'].value
    normalcnt = normalcnt + ws_data['E18'].value


    errorcnt = tatalcnt - normalcnt
    errorrate = errorcnt/tatalcnt
log.write(organnm + ','+str(tatalcnt) + ','+str(errorcnt)+','+str(errorrate)+'\n')
