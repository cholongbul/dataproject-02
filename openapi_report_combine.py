import openpyxl
import os

##보고서 읽기
class read_report:
    def __init__(self):
        self.path = './report'
        self.report_list = os.listdir(self.path)
        for self.report in self.report_list:
            self.wb_data = openpyxl.load_workbook(self.path + self.report)

    def sheet(self, sheetname):
        ws_data = self.wb_data[sheetname]
##수준평가 시트 합연산

##추가진단 시트 합연산

##S1시트 합치기
###시트 길이가 다름을 고려
###길이가 비대칭이기도 함


##