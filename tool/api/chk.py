import os
import re

path = 'C:\\Users\\admin\\Documents\\1.API\\3.리소스\\4.수준평가미대상리소스\\'
idfile = open('./idlist.txt', 'r', encoding='utf8')
idfilelist =idfile.readlines()
filelist = os.listdir(path)
num_idlist = []
for file in filelist:
    numbers = re.sub(r'[^0-9]', '', file)
    num_idlist.append(numbers)
for id in idfilelist:
    if id.replace('\n','') in num_idlist:
        num_idlist.remove(id.replace('\n',''))
print(len(num_idlist))
print(num_idlist)
for id in num_idlist:
    wfile = open('./nidflist.txt', 'a', encoding='utf8')
    wfile.write(id + '\n')
    wfile.close()