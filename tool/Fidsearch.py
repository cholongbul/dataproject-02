import openpyxl
import os

path='C:\\Users\\admin\\Documents\\2.개방데이터\\2.fid작업\\'
filelist= os.listdir(path)
i = 0
for file in filelist:
    print(i)
    result_file = open('./fileid.csv', 'a', encoding='utf8')
    wb_data = openpyxl.load_workbook(path+file)
    ws_data = wb_data['대상파일목록']
    id_index = 2
    while True:
        fileid = ws_data['B'+str(id_index)].value
        print(fileid)
        if fileid == None:
            break
        else:
            result_file.write(fileid + '\n')
            id_index = id_index + 1
    i = i + 1
    wb_data.close()
    result_file.close()