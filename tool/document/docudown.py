import os
import pandas as pd
from urllib.request import urlretrieve


path = '../.././document/참고문서.csv'
result_path = 'C:\\Users\\admin\\Documents\\1.API\\4.API참고문서\\'
csv_df = pd.read_csv(path)
for apicd in csv_df['apicd'].tolist():
    doculink = str(csv_df[csv_df['apicd']==apicd]['doculink'].iloc[0])
    url = doculink
    docuname = str(csv_df[csv_df['apicd']==apicd]['docunm'].iloc[0]).replace('\n','')
    try:
        if docuname.split('.')[-1] == 'hwp':
            urlretrieve(url, result_path + '1.hwp\\' + str(apicd) + ".hwp")
        elif docuname.split('.')[-1] == 'docx':
            urlretrieve(url, result_path + '2.docx\\' + str(apicd) + ".docx")
        elif docuname.split('.')[-1] == 'zip':
            urlretrieve(url, result_path + '3.zip\\' + str(apicd) + ".zip")
        elif docuname.split('.')[-1] == 'doc':
            urlretrieve(url, result_path + '4.doc\\' + str(apicd) + ".doc")
        elif docuname.split('.')[-1] == 'pdf':
            urlretrieve(url, result_path + '5.pdf\\' + str(apicd) + ".pdf")
        else:
            urlretrieve(url, result_path + docuname)



    except Exception as e:
        f = open("미다운참고문서.csv", 'a', encoding='utf8')
        f.write(url + ',' + str(apicd) + "\n")
        f.close()
        pass
