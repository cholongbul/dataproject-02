import openpyxl
import math
file = './1차진단결과-01.xlsx'
wb_data = openpyxl.load_workbook(file)
ws_data = wb_data['Data']
cnt = 0
while True:
    print(cnt)
    if ws_data['C'+str(cnt+2)].value == None:
        break

    if ws_data['C'+str(cnt+2)].value == ws_data['E'+str(cnt+2)].value and ws_data['D'+str(cnt+2)].value == ws_data['F'+str(cnt+2)].value:
        cnt = cnt + 1
        continue

    if ws_data['C'+str(cnt+2)].value != ws_data['E'+str(cnt+2)].value:
        ##두번쨰 토탈에서 첫번째 토탈을 뺀 값
        gap = int(ws_data['E' + str(cnt + 2)].value) - int(ws_data['C'+str(cnt+2)].value)
        div_cnt = 0
        alpalist = ['G','I','K','M','O','Q','S','U','W']
        totcntlist = []
        alpa_gap_dict = {}
        div_gap_list = []
        div_gap_total = 0
        ##뺀 값을 비율대로 나누어야 함
        for alpa in alpalist:
            if ws_data[alpa + str(cnt + 2)].value != 0:
                ##비율은 항목값/두번째 토탈
                alpa_gap_dict[alpa]=math.floor(gap*int(ws_data[alpa + str(cnt + 2)].value)/int(ws_data['E' + str(cnt + 2)].value))
                div_gap_total = div_gap_total + math.floor(gap * int(ws_data[alpa + str(cnt + 2)].value) / int(ws_data['E' + str(cnt + 2)].value))


        for alpa in alpa_gap_dict.keys():
            print(gap, div_gap_total)
            if gap != div_gap_total and alpalist.index(alpa) == 0:
                if gap > div_gap_total:
                    div_gap_gap = gap - div_gap_total
                else:
                    div_gap_gap = gap - div_gap_total
                print(div_gap_gap)
                ws_data[alpa + str(cnt + 2)].value = int(ws_data[alpa + str(cnt + 2)].value) - (alpa_gap_dict[alpa] + div_gap_gap)
            else:
                print(alpa_gap_dict)
                ws_data[alpa + str(cnt + 2)].value = int(ws_data[alpa + str(cnt + 2)].value) - alpa_gap_dict[alpa]

        ws_data['E' + str(cnt + 2)].value = ws_data['C'+str(cnt+2)].value

    if ws_data['D'+str(cnt+2)].value != ws_data['F'+str(cnt+2)].value:
        gap = abs(int(ws_data['D' + str(cnt + 2)].value) - int(ws_data['F'+str(cnt+2)].value))
        div_cnt = 0
        alpalist = ['H','J','L','N','P','R','T','V','X']
        totcntlist = []
        alpa_gap_dict = {}
        div_gap_list = []
        div_gap_total = 0
        for alpa in alpalist:
            if ws_data[alpa + str(cnt + 2)].value != 0:
                ##비율은 항목값/두번째 토탈
                alpa_gap_dict[alpa] = math.floor(
                    gap * int(ws_data[alpa + str(cnt + 2)].value) / int(ws_data['F' + str(cnt + 2)].value))
                div_gap_total = div_gap_total + math.floor(
                    gap * int(ws_data[alpa + str(cnt + 2)].value) / int(ws_data['F' + str(cnt + 2)].value))

        for alpa in alpa_gap_dict.keys():
            print(gap, div_gap_total)
            if gap != div_gap_total and alpalist.index(alpa) == 0:
                if gap > div_gap_total:
                    div_gap_gap = gap - div_gap_total
                else:
                    div_gap_gap = gap - div_gap_total
                print(div_gap_gap)
                ws_data[alpa + str(cnt + 2)].value = int(ws_data[alpa + str(cnt + 2)].value) - (
                            alpa_gap_dict[alpa] + div_gap_gap)
            else:
                print(alpa_gap_dict)
                ws_data[alpa + str(cnt + 2)].value = int(ws_data[alpa + str(cnt + 2)].value) - alpa_gap_dict[alpa]

        ws_data['F' + str(cnt + 2)].value = ws_data['D' + str(cnt + 2)].value
    cnt = cnt +1
wb_data.save('./1차진단결과-02.xlsx')