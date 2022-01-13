import openpyxl
import os
import shutil

def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print('Error: Creating directory. ' + directory)


path = './report/'
filelist = os.listdir(path)
for file in filelist:
    wb_data = openpyxl.load_workbook(path + file)
    ws_data = wb_data['00.수준평가']
    organnm = str(ws_data['C5'].value).replace('/','')
    createFolder('./report_name/' + organnm)
    shutil.copy(path + file, './report_name/' + organnm + '/'+file)
