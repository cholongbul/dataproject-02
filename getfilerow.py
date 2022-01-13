import traceback
import pandas as pd

idfile = open('./idlist.txt', 'r')
id_list = idfile.readlines()
path = 'Z:\\115.신규파일리스트(10.28)(UTF8BOM)\\'
for id in id_list:
    try:
        filenm = id.replace('\n','') + '.csv'
        csv_file = open(path + filenm,'r',encoding='utf8')
        row = len(csv_file.readlines())
        resultfile =open('./result.csv', 'a',encoding='utf-8')
        resultfile.write(id.replace('\n','')+','+str(row) + '\n')
        resultfile.close()
    except Exception as e:
        traceback.print_exc()
        errorfile = open('./error.csv', 'a', encoding='utf-8')
        errorfile.write(id+','+str(e))
        errorfile.close()