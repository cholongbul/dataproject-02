import pandas as pd
import os
import win32com.client
import traceback
import cx_Oracle
from sqlalchemy import create_engine
import openpyxl
path = "C:\\Users\\admin\\Documents\\2.개방데이터\\13.파일진단보고서\\"
report_list = os.listdir(path)
result_df = pd.DataFrame(index=range(0,0), columns=['1'])

for report in report_list:
    print(report)
    log = open('./제외파일오류율.csv', 'a', encoding='utf8')
    # win32com #excel 사용할 수 있게 설정
    excel = win32com.client.Dispatch("Excel.Application")
    #임시 Workbook 객체 생성 및 엑셀 열기
    temp_wb = excel.Workbooks.Open(path + report)
    #저장
    temp_wb.Save()
    #excel 종료
    excel.quit()
    wb_data = openpyxl.load_workbook(path + report,data_only=True)

    ws_data = wb_data['개방데이터(파일) 값 진단 결과보고서']
    organnm = ws_data['C5'].value
    tatalcnt = ws_data['E20'].value
    errorcnt = ws_data['F20'].value



    log.write(organnm + ','+str(tatalcnt) + ','+str(errorcnt)+'\n')
    log.close()
