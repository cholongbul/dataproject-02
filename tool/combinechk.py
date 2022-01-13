import shutil
import os
import openpyxl

path = '../report_combine/'
result_path = 'Z:\\110.기관제출보고서\\수준평가미대상기관(10.15)\\'
reportlist = os.listdir(path)
for report in reportlist:
    wb_data = openpyxl.load_workbook(path + report, read_only=False, data_only=True)
    ws_data = wb_data['서비스진단대상']
    organnm = ws_data['A2'].value
    execapi = ws_data['G2'].value
    if execapi == None:
        execapi = 0
    apicnt = ws_data['D2'].value-execapi
    s={1:[],2:[],3:[],4:[],5:[],6:[]}
    for i in range(0, apicnt):
        if ws_data['B' + str(i+6)].value == ws_data['C' + str(i+6)].value:
            s[1].append(1)
        else:
            s[1].append(0)
        if ws_data['E' + str(i+6)].value == ws_data['F' + str(i+6)].value:
            s[2].append(1)
        else:
            s[2].append(0)
        if ws_data['H' + str(i + 6)].value == ws_data['I' + str(i + 6)].value:
            s[3].append(1)
        else:
            s[3].append(0)
        if ws_data['K' + str(i + 6)].value == ws_data['L' + str(i + 6)].value:
            s[4].append(1)
        else:
            s[4].append(0)
        if ws_data['N' + str(i + 6)].value == ws_data['O' + str(i + 6)].value:
            s[5].append(1)
        else:
            s[5].append(0)
        if ws_data['Q' + str(i + 6)].value != None:
            if ws_data['Q' + str(i + 6)].value == ws_data['R' + str(i + 6)].value:
                s[6].append(1)
            else:
                s[6].append(0)

    api_s_len=len(s[1]) + len(s[2]) + len(s[3]) + len(s[4]) + len(s[5]) + len(s[6])
    api_acount = sum(s[1])+sum(s[2])+sum(s[3])+sum(s[4])+sum(s[5])+sum(s[6])
    answer=api_acount/api_s_len/apicnt*2
    print(answer)
    print(ws_data['A2'].value)
    if answer == 2.0:
        shutil.copy(path + report, result_path+organnm+'.xlsx')