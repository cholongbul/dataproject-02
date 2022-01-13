import openpyxl
import shutil
import module.module as md
def create_resource():
    courosr = md.dbcursor_module()
    path= 'C:\\Users\\admin\\Documents\\1.API\\3.리소스\\12.12월신규\\'
    dec_file = open('./12월신규.txt','r',encoding='utf8')
    dec_id_list = dec_file.readlines()
    resource_templet = './리소스템플릿.xlsx'
    for dec_id in dec_id_list:
        dec_id = dec_id.replace('\n','')
        shutil.copy(resource_templet, path + dec_id + '.xlsx')
        wb_data = openpyxl.load_workbook(path + dec_id + '.xlsx')
        opernm_list = md.opernm_list(courosr,dec_id)
        for i in range(0,len(opernm_list)):
            if i == 0:
                ws_data=wb_data['1']
                ws_data['A2'].value = str(opernm_list[i])
                ws_data['L2'].value = 'https://www.data.go.kr/data/'+str(dec_id)+'/openapi.do'
            else:
                target = wb_data.copy_worksheet(ws_data)
                target.title = str(i+1)
                target['A2'].value = opernm_list[i]

        wb_data.save(path + dec_id + '.xlsx')

create_resource()