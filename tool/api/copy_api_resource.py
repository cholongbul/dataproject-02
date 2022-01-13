import shutil
import os

list_file = open('C:\\Users\\admin\\Desktop\\openapi\\document\\식품의약품안전처.txt', 'r', encoding='utf8')
resource_path = 'C:\\Users\\admin\\Documents\\1.API\\3.리소스\\3.수준평가대상기관리소스\\'
resource_list = os.listdir(resource_path)
idlist = list_file.readlines()
for id in idlist:
    id = id.rstrip('\n')
    print(id)
    for resource in resource_list:
        if resource.startswith(id):
            shutil.copy(resource_path + resource, 'C:\\Users\\admin\\Documents\\1.API\\3.리소스\\6.식품의약품안전처\\' + resource)
