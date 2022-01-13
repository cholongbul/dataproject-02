import os
import re

path='C:\\Users\\admin\\Documents\\1.API\\3.리소스\\4.수준평가미대상리소스\\'
filelist = os.listdir(path)
idfile = open('./nidflist.txt', 'r', encoding='utf8')
idlist = idfile.readlines()
for id in idlist:
    for file in filelist:
        file_num = re.sub(r'[^0-9]', '', file)
        id_num = re.sub(r'[^0-9]', '', id)
        if file_num == id_num:
            os.remove(path+file)


