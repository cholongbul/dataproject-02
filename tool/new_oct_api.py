import os

filelist = os.listdir('C:\\Users\\admin\\Documents\\1.API\\3.리소스\\8.10월신규\\')
txt = open('./rnewoctfile.txt', 'r', encoding='utf8')
txtlist = txt.readlines()
for txt in txtlist:
    os.remove('C:\\Users\\admin\\Documents\\1.API\\3.리소스\\8.10월신규\\'+str(txt).rstrip('\n')+'.xlsx')

