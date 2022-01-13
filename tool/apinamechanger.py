##코드명으로 정리한 리소스를 이름으로 변환
import os
import pandas as pd
import shutil
import re

path = 'C:\\Users\\admin\\Documents\\리소스\\'
result_path = 'C:\\Users\\admin\\Documents\\리소스\\name\\'
orilist = os.listdir(path)
id_df = pd.read_csv('../document/apiidmatch.csv')##입력데이터
for ori in orilist:
    if ori == 'name':
        continue
    id = ori.rstrip('.xlsx')
    print(id)
    is_sameid = id_df['apicd'] == int(id)

    print(id_df[is_sameid])
    name = id_df[is_sameid]['apiname'].tolist()[0]
    name = re.sub('[\/:*?"<>|]','',name)
    name = name.replace('\t','').replace('\n','').replace('\r','')
    shutil.copy(path + ori, result_path + str(name).replace('\/','') + "_리소스.xlsx")
