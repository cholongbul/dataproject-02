##api리소스가 정상적인 오퍼레이션을 가지고 존재하는지 체크 후 결과를 정리하여 csv로 출력
import re
import openpyxl
import traceback, os
import pandas as pd

try:
    templet = 'api_confirm2.csv'
    tem_df = pd.read_csv(templet)
    path = 'Z:\\104.DB적재소스\\report_numbering\\'
    resourcepath = 'Z:\\104.DB적재소스\\리소스모음\\리소스\\'
    fname_list = os.listdir(path)
    for fname in fname_list:
        try:
            wb_data = openpyxl.load_workbook(path + fname)
            ws_data = wb_data['00.수준평가']
            sheet_len = len(wb_data.sheetnames)
            link = ws_data['C11'].value
            numbers = re.sub(r'[^0-9]', '', link)
            print(numbers)
            ##오퍼레이션명이 같은 행 찾기
            is_opertaion = tem_df['오픈APIID'] == int(numbers)
            # is_opercnt = tem_df['오퍼레이션갯수'] == sheet_len
            for ind in tem_df[is_opertaion].index:
                print(ind)
                tem_df.loc[ind,'chk'] = '일치'
            wb_data.close()
        except Exception as ie:
            continue
        tem_df.to_csv('./api_confirm.csv', index=False)

except Exception as e:
    traceback.print_exc()
