import os
import shutil
import traceback
import re
import pandas as pd
import openpyxl
import xlrd
def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print('Error: Creating directory. ' + directory)


def div_file():
    path = 'C:\\Users\\admin\\Downloads\\SQL_DIAG_20220105\\0105\\'
    result_path = 'C:\\Users\\admin\\Downloads\\SQL_DIAG_20220105\\기관별\\'
    csv_df = pd.read_csv('./id_organ.csv')
    filelist = os.listdir(path)
    resultlist = os.listdir(result_path)
    for file in filelist:
        try:
            file_dotname = '.'+file.split('.')[-1]
            if file.split('_')[1].startswith('F1'):
                fileid = file.split('_')[1]
            elif file.split('_')[2].startswith('F1'):
                fileid = file.split('_')[2]
        except IndexError:
            fileid = file.split('.')[0]
        if file.endswith('.xls') or file.endswith('.xlsx') or file.endswith('.xlsm'):
            is_same_id = csv_df['id'] == fileid
            organnm = csv_df[is_same_id]['organ'].tolist()[0]
            createFolder(result_path + organnm)
            shutil.copy(path + file, result_path + organnm + '\\' + fileid + file_dotname)
        else:
            os.remove(path + file)

def div_file2():
    path = 'C:\\Users\\admin\\Downloads\\SQL_DIAG_20220105\\D3014247\\'
    result_path = 'C:\\Users\\admin\\Downloads\\SQL_DIAG_20220105\\기관별\\'
    csv_df = pd.read_csv('./id_organ.csv')
    filelist = os.listdir(path)
    resultlist = os.listdir(result_path)

    for file in filelist:
        if '육안진단결과보고서' in file:
            continue
        try:
            file_dotname = '.'+file.split('.')[-1]
            if file.split('_')[0].startswith('F1'):
                fileid = file.split('_')[0]
            elif file.split('_')[1].startswith('F1'):
                fileid = file.split('_')[1]
            elif file.split('_')[2].startswith('F1'):
                fileid = file.split('_')[2]
        except IndexError:
            traceback.print_exc()
            fileid = file.split('.')[0]
        if file.endswith('.xls') or file.endswith('.xlsx') or file.endswith('.xlsm'):

            is_same_id = csv_df['id'] == fileid
            organnm = csv_df[is_same_id]['organ'].tolist()[0]
            createFolder(result_path + organnm)
            shutil.copy(path + file, result_path + organnm + '\\' + fileid + file_dotname)
            print(file)
        else:
            os.remove(path + file)

def filelist():
    path ='C:\\Users\\admin\\Downloads\\SQL_DIAG_20220105\\기관별\\'

    for (root, dirs, files) in os.walk(path):
        for file in files:
            log=open('file.csv','a',encoding='utf8')
            log.write(file+'\n')
            log.close()

def div_file3():
    path = 'C:\\Users\\admin\\Downloads\\01-11\\'
    result_path = 'C:\\Users\\admin\\Downloads\\01-11\\기관별분류'
    for (root, dirs, files) in os.walk(path):
        for file in files:
            if str(root).endswith('10.육안진단(01.10)'):
                createFolder(result_path+'\\'+file.split('_')[0]+'_개방데이터')
                shutil.copy(root+'\\'+file,result_path+'\\'+file.split('_')[0]+'_개방데이터\\'+file)
            elif str(root).endswith('20.미대상기관(01.10)'):
                createFolder(result_path+'\\'+file.split('_')[1]+'_개방데이터')
                shutil.copy(root+'\\'+file,result_path+'\\'+file.split('_')[1]+'_개방데이터\\'+file)
            elif str(root).endswith('30.9월신규파일(01.10)'):
                createFolder(result_path + '\\' + file.split('_')[1] + '_개방데이터')
                shutil.copy(root + '\\' + file, result_path + '\\' + file.split('_')[1] + '_개방데이터\\' + file.replace('종합진단결과보고서','종합진단결과보고서(9월)'))
            elif str(root).endswith('40.10월11월신규파일(01.10)'):
                createFolder(result_path + '\\' + file.split('_')[1] + '_개방데이터')
                shutil.copy(root + '\\' + file, result_path + '\\' + file.split('_')[1] + '_개방데이터\\' + file.replace('종합진단결과보고서','종합진단결과보고서(10월11월)'))
            elif str(root).endswith('50.LINK진단결과보고서(01.10)'):
                createFolder(result_path + '\\' + file.split('_')[1] + '_개방데이터')
                shutil.copy(root + '\\' + file, result_path + '\\' + file.split('_')[1] + '_개방데이터\\' + file.replace('종합진단결과보고서','종합진단결과보고서(LINK)'))
            print(root)
    # filelist = os.listdir(path)
    # resultlist = os.listdir(result_path)
    #
    # for file in filelist:
    #     if '육안진단결과보고서' in file:
    #         continue
    #     try:
    #         file_dotname = '.'+file.split('.')[-1]
    #         if file.split('_')[0].startswith('F1'):
    #             fileid = file.split('_')[0]
    #         elif file.split('_')[1].startswith('F1'):
    #             fileid = file.split('_')[1]
    #         elif file.split('_')[2].startswith('F1'):
    #             fileid = file.split('_')[2]
    #     except IndexError:
    #         traceback.print_exc()
    #         fileid = file.split('.')[0]
    #     if file.endswith('.xls') or file.endswith('.xlsx') or file.endswith('.xlsm'):
    #
    #         is_same_id = csv_df['id'] == fileid
    #         organnm = csv_df[is_same_id]['organ'].tolist()[0]
    #         createFolder(result_path + organnm)
    #         shutil.copy(path + file, result_path + organnm + '\\' + fileid + file_dotname)
    #         print(file)
    #     else:
    #         os.remove(path + file)


div_file3()