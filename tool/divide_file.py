import os
import module.module as md
import pandas as pd
import shutil
csv_df = pd.read_csv('./new_file23.csv')
path = 'Z:\\304.파일명변환(09.02)\\'
result_path = 'C:\\Users\\admin\\Documents\\2.개방데이터\\11.신규기관별\\'
filelist = os.listdir(path)
for file in filelist:
    try:
        fileid = file.replace('.csv','')
        is_sameid = csv_df['파일ID'] == fileid
        orgnnm = csv_df[is_sameid]['기관명'].tolist()[0]
        md.createFolder(result_path+orgnnm)
        shutil.copy(path+file,result_path+'(전체)'+'\\'+file)
    except:
        pass
