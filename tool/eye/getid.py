import os

import pandas as pd

path = 'C:\\Users\\admin\\Documents\\1.재수정육안\\'
files = os.listdir(path)
result_df = pd.DataFrame({'개방데이터ID':[],
                   '개방데이터목록명':[],
                   '개방데이터파일명':[]})
for file in files:
    print(file)
    ex_df = pd.read_excel(path+file,sheet_name=0,header=3)
    print(ex_df)
    result_df = pd.concat([result_df,ex_df])


result_df.to_csv('./육안ID리스트.csv')