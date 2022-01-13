import openpyxl
import pandas as pd
import os
import shutil

addlinelist = pd.read_csv('./document/addlinelist.csv')
report_path= './report/'
result_templete = './document/templete.xlsx'
report_file_list = os.listdir(report_path)

for report in report_file_list:
    wb_data = openpyxl.load_workbook(report_path + report)
    s0 = wb_data['표지']
    s0['E18'].value = '진단기간 : 2021-10-25 ~ 2021-10-31'
    s1 = wb_data['서비스진단대상']
    i1 = 6
    s1['W4'].value = '1차 진단여부'
    s1['X4'].value = '2차 진단여부'
    s1['Y4'].value = '이의신청 내용'
    print(report)
    while True:
        apiname = s1['A' + str(i1)].value
        if apiname == None:
            break
        if str(apiname).endswith(' - 표준데이터셋'):
            apiname = str(apiname).rstrip(' - 표준데이터셋')
        if str(apiname).endswith(' -표준데이터셋'):
            apiname = str(apiname).rstrip(' -표준데이터셋')
        if str(apiname).endswith(' - 표준데이터셋(LINK)'):
            apiname = str(apiname).rstrip(' - 표준데이터셋(LINK)')
        if str(apiname)=='국립암센터_대장암 라이브러리 환자수':
            apiname = '국립암센터_대장암 라이브러리  환자수'
        if '∙' in str(apiname):
            apiname = apiname.replace('∙','?')

        apinmlist = addlinelist['API_NM'] == apiname
        try:
            if addlinelist[apinmlist]['1차진단'].tolist()[0] == '1차진단':
                s1['W'+ str(i1)].value = '진단'
            else:
                s1['W' + str(i1)].value = '제외'
        except IndexError:
            print(apiname)
            s1['W' + str(i1)].value = '제외'
        try:
            if addlinelist[apinmlist]['2차진단대상'].tolist()[0] == '진단대상':
                s1['X'+ str(i1)].value = '진단'
            else:
                s1['X'+ str(i1)].value = '제외'
        except IndexError:
            print(apiname)
            s1['X'+ str(i1)].value = '진단'
        i1 = i1 + 1

    wb_data.save(report_path + report)

