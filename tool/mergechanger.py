##병합된 셀을 해제
import os
import module.module as md

path = 'C:\\Users\\admin\\Desktop\\리소스오류정리\\셀병합오류\\'
resource_list = os.listdir(path)

for resource in resource_list:
    print(resource)
    md.mergechange(path,resource)