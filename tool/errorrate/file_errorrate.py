import os
import traceback

import openpyxl
link_path = "C:\\Users\\admin\\Documents\\2.개방데이터\\8.파일데이터종합\\101.링크ZIP\\"
nun_path = "C:\\Users\\admin\\Documents\\2.개방데이터\\8.파일데이터종합\\102.수준평가미대상\\"
sep_path = "C:\\Users\\admin\\Documents\\2.개방데이터\\8.파일데이터종합\\103.9월신규\\"
octnov_path = "C:\\Users\\admin\\Documents\\2.개방데이터\\8.파일데이터종합\\104.1011신규\\"

link_report_list = os.listdir(link_path)
nun_report_list = os.listdir(nun_path)
sep_report_list = os.listdir(sep_path)
octnov_report_list = os.listdir(octnov_path)

for link_report in octnov_report_list:
    log = open('./1011월신규.csv', 'a', encoding='utf8')
    print(link_report)
    try:
        wb_data = openpyxl.load_workbook(octnov_path + link_report, read_only=True)
        ws_data = wb_data['개방데이터(파일) 값 진단 결과보고서']
        organnm = ws_data['C5'].value
        total_data = ws_data['E20'].value
        error_data = ws_data['F20'].value
        log.write(organnm + "," + str(total_data) + "," + str(error_data) + '\n')
        log.close()
        wb_data.close()
    except Exception as e:
        traceback.print_exc()
