import os
import openpyxl
import shutil
def chklist():
    path = 'C:\\Users\\admin\\Documents\\1.API\\2.오픈API종합보고서\\11.9월신규(7,8번삭제)\\'
    list = os.listdir(path)

    for file in list:
        print(file)
        wb_data = openpyxl.load_workbook(path + file)
        ws_data = wb_data['00.수준평가']
        url = ws_data['C11'].value
        apiid = str(url).replace('https://www.data.go.kr/data/','').replace('/openapi.do','')
        wirtefile = open('09list.txt', 'a',encoding='utf8')
        wirtefile.write(apiid + '\n')
        wirtefile.close()

def delfile():
    path = 'C:\\Users\\admin\\Documents\\1.API\\3.리소스\\8.9월신규\\'
    list = os.listdir(path)
    delfile = open('./09dellist.txt', 'r', encoding='utf8')
    dellist = delfile.readlines()
    for file in list:
        for del_id in dellist:
            print(file.replace('.xlsx',''))
            print(del_id.replace('\n',''))
            if file.replace('.xlsx','')==del_id.replace('\n',''):
                shutil.move(path + file, 'C:\\Users\\admin\\Documents\\1.API\\3.리소스\\15.9월미신규파일\\'+file)

delfile()