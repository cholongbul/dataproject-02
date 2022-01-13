import pandas as pd
import os
import win32com.client
import traceback
import cx_Oracle
from sqlalchemy import create_engine
import openpyxl
path = "C:\\Users\\admin\\Documents\\1.API\\2.오픈API종합보고서\\11.9월신규종합\\"
error_list = open('./errorlog1.txt','r',encoding='utf-8').readlines()
report_list = os.listdir(path)
result_df = pd.DataFrame(index=range(0,0), columns=['1'])
for report in report_list:
    log = open('./api9월신규_항목별.csv', 'a', encoding='utf8')
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
    ws_data = wb_data['서비스진단대상']
    ex_cnt = 0
    s1target = 0
    s1error = 0
    s2target = 0
    s2error = 0
    s3target = 0
    s3error = 0
    s4target = 0
    s4error = 0
    s5target = 0
    s5error = 0
    s6target = 0
    s6error = 0
    target = 0
    error = 0
    while True:
        if ws_data['A'+ str(6+ex_cnt)].value == None:
            break
        else:
            print(ws_data['A'+ str(6+ex_cnt)].value)
            print(ws_data['T'+ str(6+ex_cnt)].value)
            apinm = ws_data['A'+ str(6+ex_cnt)].value.replace(',','')
            s1target =  ws_data['B'+ str(6+ex_cnt)].value
            s1error = ws_data['C'+ str(6+ex_cnt)].value
            s2target =  ws_data['E' + str(6 + ex_cnt)].value
            s2error = ws_data['F' + str(6 + ex_cnt)].value
            s3target = ws_data['H' + str(6 + ex_cnt)].value
            s3error =  ws_data['I' + str(6 + ex_cnt)].value
            s4target = ws_data['K' + str(6 + ex_cnt)].value
            s4error = ws_data['L' + str(6 + ex_cnt)].value
            s5target =  ws_data['N' + str(6 + ex_cnt)].value
            s5error =  ws_data['O' + str(6 + ex_cnt)].value
            if ws_data['Q' + str(6 + ex_cnt)].value != None:
                s6target = ws_data['Q' + str(6 + ex_cnt)].value
                s6error =  ws_data['R' + str(6 + ex_cnt)].value
            else:
                s6target = 0
                s6error = 0
            target = target + ws_data['T' + str(6 + ex_cnt)].value
            try:
                s1errorcnt = s1target - s1error
                s2errorcnt = s2target - s2error
                s3errorcnt = s3target - s3error
                s4errorcnt = s4target - s4error
                s5errorcnt = s5target - s5error
                s6errorcnt = s6target - s6error
                ex_cnt = ex_cnt + 1
            except:
                s1errorcnt = 0
                s2errorcnt = 0
                s3errorcnt = 0
                s4errorcnt = 0
                s5errorcnt = 0
                s6errorcnt = 0
                ex_cnt = ex_cnt + 1



            log.write('"'+apinm + '",'+str(s1target) + ','+str(s1error)+','+str(s1errorcnt)
                      + ','+str(s2target) + ','+str(s2error)+','+str(s2errorcnt)
                      + ','+str(s3target) + ','+str(s3error)+','+str(s3errorcnt)
                      + ','+str(s4target) + ','+str(s4error)+','+str(s4errorcnt)
                      + ','+str(s5target) + ','+str(s5error)+','+str(s5errorcnt)
                      + ','+str(s6target) + ','+str(s6error)+','+str(s6errorcnt)
                      +'\n')
