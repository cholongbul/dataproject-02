import os
path = 'C:\\Users\\admin\\Downloads\\01-11\\기관별분류\\'
files = os.listdir(path)

for file in files:
    log = open('기관명.csv','a',encoding='utf8')
    log.write(file.replace('_개방데이터.zip','') + '\n')
    log.close()
