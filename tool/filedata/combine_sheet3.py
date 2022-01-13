import os

import pandas as pd
path = 'C:\\Users\\admin\\Documents\\2.개방데이터\\22.개방데이터정리\\300.10월11월진단결과보고서종합(신규)\\'
files = os.listdir(path)
result_df = pd.DataFrame({'번호':[],
                   '개방데이터ID':[],
                   '개방데이터파일명':[],
                   '컬럼순번':[],'컬럼명':[],'검증유형':[],'검증유형 상세':[],'전체건수':[]
                    ,'오류건수':[],'오류율(%)':[]})
for file in files:
    print(file)
    excell_df = pd.read_excel(path + file,sheet_name=2, na_values='')
    result_df =pd.concat([result_df,excell_df])
    print(result_df)


result_df.to_csv('./1011월신규_진단규칙 및 오류목록_종합.csv')
