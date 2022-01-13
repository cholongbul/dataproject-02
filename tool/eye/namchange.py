import openpyxl
import os
import win32com.client

path = 'C:\\Users\\admin\\Documents\\8.육안\\1.재수정육안\\'
result_path = 'C:\\Users\\admin\\Documents\\8.육안\\2.이름수정\\'
files = os.listdir(path)
for file in files:
    # win32com #excel 사용할 수 있게 설정
    excel = win32com.client.Dispatch("Excel.Application")
    # 임시 Workbook 객체 생성 및 엑셀 열기
    temp_wb = excel.Workbooks.Open(path + file)
    # 저장
    temp_wb.Save()
    # excel 종료
    excel.quit()
    print(file)
    wb_data = openpyxl.load_workbook(path + file)
    ws1 = wb_data['개방데이터(파일) 값 진단 결과보고서']
    ws2 = wb_data['진단규칙 및 오류목록']
    ws2['B1'].value = '개방데이터파일명'
    cnt = 0
    cnt2 = 0
    filename_dict = {}
    while True:
        if ws1['B'+str(5+cnt)].value == None:
            break
        else:
            filename_dict[str(ws1['B'+str(5+cnt)].value)]=str(ws1['C'+str(5+cnt)].value)
            cnt = cnt +1

    while True:
        if ws2['B'+str(2+cnt2)].value == None:
            break
        else:
            ws2['B' + str(2 + cnt2)].value = filename_dict[str(ws2['B' + str(2 + cnt2)].value)]
            cnt2 = cnt2 + 1
    wb_data.save(result_path + file)
