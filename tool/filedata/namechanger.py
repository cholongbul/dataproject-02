import os
import shutil

import pandas as pd
path = 'C:\\Users\\admin\\Documents\\2.개방데이터\\14.미대상파일-미대상기관-육안진단\\'
result_path = 'C:\\Users\\admin\\Documents\\2.개방데이터\\14.미대상파일-미대상기관-육안진단(이름변환)\\'
data = pd.read_csv('./id_data.csv')
list = os.listdir(path)
for file in list:
    file_id = file.split('_')[0]
    file_type = file.split('.')[1]
    is_same_id = data['INTERN_SYS_ID'] == file_id
    real_fileid = data[is_same_id]['FILE_ID'].tolist()[0]
    datanm = data[is_same_id]['DATA_NM'].tolist()[0]
    datatype= data[is_same_id]['INTERN_CD'].tolist()[0]
    dataresult = data[is_same_id]['상태'].tolist()[0]
    shutil.copy(path + file,result_path+real_fileid+'_'+str(file_id)+'_'+datanm+'_'+datatype+'_'+dataresult+'.'+file_type)
