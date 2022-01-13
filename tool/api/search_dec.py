import os
import shutil
path = 'C:\\Users\\admin\\Documents\\1.API\\3.리소스\\'
file = open('./12월신규.txt', 'r', encoding='utf-8')
dec_id_list = file.readlines()
for dec_id in dec_id_list:
    dec_id = dec_id.replace('\n','')
    for (root, dirs, files) in os.walk(path):
            print("# root : " + root)
            for file in files:
                if root != 'C:\\Users\\admin\\Documents\\1.API\\3.리소스\\14.12월신규모음' and file.replace('.xlsx','')==dec_id:
                    shutil.copy(root +'\\'+ file,'C:\\Users\\admin\\Documents\\1.API\\3.리소스\\14.12월신규모음\\'+file)
