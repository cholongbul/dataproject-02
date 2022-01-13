import openpyxl
import os
import xlrd

path = 'C:\\Users\\admin\\Documents\\2.개방데이터\\17.데이터추출\\'
for (root, dirs, files) in os.walk(path):
    for file in files:
        print(root + file)
        if os.path.isfile(root +'\\'+file):
            if file.endswith('.xlsx'):
                wb_data = openpyxl.load_workbook(root +'\\'+ file, read_only=True)
                firstsheet=wb_data.sheetnames[0]
                ws_data = wb_data[firstsheet]
                organnm = ws_data['C5'].value
                money_target = ws_data['D11'].value
                money_all = ws_data['E11'].value
                money_error = ws_data['F11'].value
                num_target = ws_data['D12'].value
                num_all = ws_data['E12'].value
                num_error = ws_data['F12'].value
                rate_target = ws_data['D13'].value
                rate_all = ws_data['E13'].value
                rate_error = ws_data['F13'].value
                yn_target = ws_data['D14'].value
                yn_all = ws_data['E14'].value
                yn_error = ws_data['F14'].value
                date_target = ws_data['D15'].value
                date_all = ws_data['E15'].value
                date_error = ws_data['F15'].value
                code_target = ws_data['D16'].value
                code_all = ws_data['E16'].value
                code_error = ws_data['F16'].value
                time_target = ws_data['D17'].value
                time_all = ws_data['E17'].value
                time_error = ws_data['F17'].value
                relation_target = ws_data['D18'].value
                relation_all = ws_data['E18'].value
                relation_error = ws_data['F18'].value
                cal_target = ws_data['D19'].value
                cal_all = ws_data['E19'].value
                cal_error = ws_data['F19'].value
            elif file.endswith('.xls'):
                wb_data = xlrd.open_workbook(root + '\\' + file)
                firstsheet = wb_data.sheet_names()[0]
                ws_data = wb_data[firstsheet]
                organnm = str(ws_data.cell_value(4,2))
                money_target = str(ws_data.cell_value(14,3))
                money_all = str(ws_data.cell_value(14,4))
                money_error = str(ws_data.cell_value(14,5))
                num_target = str(ws_data.cell_value(15,3))
                num_all = str(ws_data.cell_value(15,4))
                num_error = str(ws_data.cell_value(15,5))
                rate_target = str(ws_data.cell_value(16,3))
                rate_all = str(ws_data.cell_value(16,4))
                rate_error = str(ws_data.cell_value(16,5))
                yn_target = str(ws_data.cell_value(17,3))
                yn_all = str(ws_data.cell_value(17,4))
                yn_error = str(ws_data.cell_value(17,5))
                date_target = str(ws_data.cell_value(18,3))
                date_all = str(ws_data.cell_value(18,4))
                date_error = str(ws_data.cell_value(18,5))
                code_target = str(ws_data.cell_value(19,3))
                code_all = str(ws_data.cell_value(19,4))
                code_error = str(ws_data.cell_value(19,5))
                time_target = str(ws_data.cell_value(20,3))
                time_all = str(ws_data.cell_value(20,4))
                time_error = str(ws_data.cell_value(20,5))
                relation_target = str(ws_data.cell_value(21,3))
                relation_all = str(ws_data.cell_value(21,4))
                relation_error = str(ws_data.cell_value(21,5))
                cal_target = str(ws_data.cell_value(22,3))
                cal_all = str(ws_data.cell_value(22,4))
                cal_error = str(ws_data.cell_value(22,5))

            log = open('./파일데이터_통계'+root.split('\\')[-1]+'.csv','a', encoding='utf8')
            log.write(organnm +","+ money_target+','+money_all+','+money_error+','+num_target+','+num_all+','+num_error+','+rate_target+','+rate_all+','+rate_error+','+
                      yn_target+','+yn_all+','+yn_error+','+date_target+','+date_all+','+date_error+','+code_target+','+code_all+','+code_error
            +','+time_target+','+time_all+','+time_error+','+relation_target+','+relation_all+','+relation_error+','+cal_target+','+cal_all+','+cal_error+'\n')
            log.close()