import json
import traceback

import requests
from bs4 import BeautifulSoup
import pandas as pd
import sys
import cx_Oracle
import re
import json
from urllib.request import urlopen


rootlink = 'https://www.data.go.kr/tcs/dss/selectDataSetList.do?dType=FILE&keyword=&detailKeyword=&publicDataPk=&recmSe=&detailText=&relatedKeyword=&commaNotInData=&commaAndData=&commaOrData=&must_not=&tabId=&dataSetCoreTf=&coreDataNm=&sort=&relRadio=&orgFullName=&orgFilter=&org=&orgSearch=&currentPage=1&perPage=971&brm=&instt=&svcType=&kwrdArray=&extsn=CSV&coreDataNmArray=&pblonsipScopeCode='
fileid_txt = open('./csvfileid.csv','r',encoding='utf8')
filelinklist = fileid_txt.readlines()
csv_df = pd.read_csv('./new_file.csv')
for filelink in filelinklist:
    try:
        filelink2 = filelink.split(',')[0]
        fileid = re.search('https://www.data.go.kr/data/(.*)/fileData.do', filelink2).group(1)
        print(fileid)
        metalink = 'https://www.data.go.kr/catalog/'+fileid+'/fileData.json'
        id = filelink.split(',')[1].replace('\n','')
        portal_response = urlopen(metalink).read().decode('utf-8')
        print(portal_response)
        resjson = json.loads(portal_response)
        url = resjson.get("distribution")[0].get("contentUrl")
        print(url)
        # date = soup.select_one("#fileDetailTableArea tr.bg-skyblue:nth-of-type(9) td:nth-of-type(1)").text
        # organnm = soup.select_one('#fileDetailTableArea tr:nth-of-type(2) td:nth-of-type(2)').text
        # listname = soup.select_one('.tit-area p').text
        # id = filelink.replace('/fileData.do','').replace('/data/','')
        # buseo = soup.select_one('.file-meta-table-pc tr:nth-of-type(3) td.td:nth-of-type(1)').text
        # buseo_bunho = soup.select_one('.file-meta-table-pc div#telNo').text
        # bunru = soup.select_one('.file-meta-table-pc tr:nth-of-type(2) td.td:nth-of-type(1)').text


        log = open('./new_file3.csv','a',encoding='utf-8')
        log.write(filelink2+','+id+','+url+'\n')
        log.close()
    except:
        traceback.print_exc()


# apiidlist = apiidfile.readlines()
# for apiid in apiidlist:
#     print(apiid)
#     try:
#         conn = cx_Oracle.connect('DQM_01/DQM_01@localhost:1521/orcl')
#         cs = conn.cursor()
#         sql = "insert into api (APICD,APINM,ORGANNM,DATAFORMMAT,DOCXNAME,APITYPE,DETAILPK,DOCULINK) values (:1,:2,:3,:4,:5,:6,:7,:8)"
#         apicd = apiid.replace('\n','')
#         portal_response = requests.get(apilink_head + apiid + apilink_tail)
#         soup = BeautifulSoup(portal_response.text, 'html.parser')
#         if '<br class="just-mb"/>요청하신 페이지를 찾을 수 없습니다.' in str(soup):
#             try:
#                 sql = "delete from api where apicd = " + apicd
#                 sql2 = "delete from operation_req where apicd = " + apicd
#                 sql3 = "delete from operation_res where apicd = "+ apicd
#                 sql4 = "insert into apiexception(apicd, reason) values(:1,:2)"
#                 cs.execute(sql4, (apicd, '포털서비스종료'))
#                 cs.execute(sql)
#                 cs.execute(sql2)
#                 cs.execute(sql3)
#                 cs.close()
#                 conn.commit()
#             except:
#                 cs.close()
#                 conn.commit()
#             continue
#         else:
#             apinm = soup.select_one('#contents > div.data-search-view > div.data-set-title.open-api > div.tit-area > p').text
#             print(apinm)
#             dataformmat = soup.select_one('tr:nth-of-type(3) td.td:nth-of-type(2)').text
#             docxname = soup.select_one(".file-meta-table-pc tr:contains('참고문서') a").text
#             apitype = soup.select_one('tr.bg-skyblue:nth-of-type(3) td:nth-of-type(1)').text
#             detailpk = soup.select_one('#publicDataDetailPk')['value']
#             organnm = soup.select_one('tr:nth-of-type(1) td.td:nth-of-type(2)').text
#             operation_list = soup.select('#open_api_detail_select > option')
#             doculink = soup.select_one(".file-meta-table-pc tr:contains('참고문서') a")
#             re.match('fn_fileDownload', doculink['onclick'])
#             doculink = 'https://www.data.go.kr/cmm/cmm/fileDownload.do?atchFileId=' + \
#                         str(doculink['onclick']).split("'")[1] + '&fileDetailSn=' + str(doculink['onclick']).split("'")[
#                             3]
#             print(doculink)
#             cs.execute(sql, (str(apicd), str(apinm),
#                              str(organnm),
#                              str(dataformmat), str(docxname), str(apitype), str(detailpk), str(doculink)))
#             cs.close()
#             conn.commit()
#     except Exception as e:
#         cs.close()
#         conn.commit()
#         if str(e).startswith('ORA-00001'):
#             try:
#                 cs = conn.cursor()
#                 sql = "update api set APINM = :1,ORGANNM =:2,DATAFORMMAT= :3,DOCXNAME= :4,APITYPE = :5,DETAILPK =:6,DOCULINK =:7" \
#                        "where apicd = :8"
#                 print(sql)
#                 cs.execute(sql,
#                             (str(apinm),
#                              str(organnm),
#                              str(dataformmat), str(docxname), str(apitype), str(detailpk), str(doculink), str(apicd)))
#                 cs.close()
#                 conn.commit()
#             except:
#                 traceback.print_exc()
#         else:
#
#             traceback.print_exc()
#     try:
#         for operation in operation_list:
#             operationcd = str(operation['value'])
#             opertaionnm = operation.text
#             print(opertaionnm)
#             operurl = operlink + optionseqno + operationcd + publicDataDetailPk +detailpk + publicDataPk + apiid
#             print(operurl)
#             oper_response = requests.get(operurl).text
#             oper_message_df = pd.read_html(oper_response)
#             oper_soup = BeautifulSoup(oper_response, 'html.parser')
#             req_url = str(oper_soup.select_one('li:nth-of-type(3)').getText).replace('<strong>요청주소</strong>','').replace('\n','').replace('</li>>','').replace('\t','').replace(' ','').replace('\r','').replace('<boundmethodTag.get_textof<li>','')
#             survice_url = str(oper_soup.select_one('li:nth-of-type(4)').getText).replace('<strong>서비스URL</strong>','').replace('\n','').replace('</li>>','').replace('\t','').replace(' ','').replace('\r','').replace('<boundmethodTag.get_textof<li>','')
#             if req_url == '':
#                 req_url = 'nan'
#             if survice_url == '':
#                 survice_url= 'nan'
#             print(req_url)
#             print(survice_url)
#             oper_req_df = oper_message_df[0]
#             oper_res_df = oper_message_df[1]
#             if oper_req_df.empty:
#                 try:
#                     colnmkr = 'nan'
#                     colnmen = 'nan'
#                     colnmsize = 'nan'
#                     coltype = 'nan'
#                     sample = 'nan'
#                     explan = 'nan'
#                     cs2 = conn.cursor()
#                     sql2 = "insert into operation_req (APICD,OPERCD,OPERNM,COLLINE,COLNMKR,COLNMEN,COLSIZE,COLTYPE,SAMPLE,EXPLAN,SERVICE_URL,REQ_URL)" \
#                            " values (:1,:2,:3,:4,:5,:6,:7,:8,:9,:10,:11,:12)"
#                     cs2.execute(sql2,
#                                 (str(apicd), str(operationcd), str(opertaionnm), str(colline), str(colnmkr), str(colnmen),
#                                  str(colnmsize), str(coltype), str(sample), str(explan),survice_url,req_url))
#                     cs2.close()
#                     conn.commit()
#                 except Exception as e:
#                     cs2.close()
#                     conn.commit()
#                     print(e)
#                     if str(e).startswith('ORA-00001'):
#                         try:
#                             cs2 = conn.cursor()
#                             sql2 = "update operation_req set " \
#                                    "COLNMKR = :1," \
#                                    "COLNMEN = :2," \
#                                    "COLSIZE = :3," \
#                                    "COLTYPE = :4," \
#                                    "SAMPLE = :5," \
#                                    "EXPLAN = :6," \
#                                    "SERVICE_URL = :7," \
#                                    "REQ_URL = :8 " \
#                                    "where apicd = :9 and opercd = :10 and colline = :11"
#                             print(sql2)
#                             cs2.execute(sql2,
#                                         (str(colnmkr),
#                                          str(colnmen),
#                                          str(colnmsize), str(coltype), str(sample), str(explan), survice_url, req_url,
#                                          str(apicd), str(operationcd), '0'))
#                             cs2.close()
#                             conn.commit()
#                         except:
#                             traceback.print_exc()
#                     else:
#                         traceback.print_exc()
#             else:
#                 for colline in range(0,len(oper_req_df['항목명(국문)'].tolist())):
#                     try:
#                         colnmkr = oper_req_df['항목명(국문)'][colline]
#                         colnmen = oper_req_df['항목명(영문)'][colline]
#                         colnmsize = oper_req_df['항목크기'][colline]
#                         coltype = oper_req_df['항목구분'][colline]
#                         sample = oper_req_df['샘플데이터'][colline]
#                         explan = oper_req_df['항목설명'][colline]
#                         cs2 = conn.cursor()
#                         sql2 = "insert into operation_req (APICD,OPERCD,OPERNM,COLLINE,COLNMKR,COLNMEN,COLSIZE,COLTYPE,SAMPLE,EXPLAN,SERVICE_URL,REQ_URL)" \
#                            " values (:1,:2,:3,:4,:5,:6,:7,:8,:9,:10,:11,:12)"
#                         cs2.execute(sql2,
#                                     (str(apicd), str(operationcd), str(opertaionnm), str(colline), str(colnmkr),
#                                      str(colnmen),
#                                      str(colnmsize), str(coltype), str(sample), str(explan), survice_url, req_url))
#                         cs2.close()
#                         conn.commit()
#                     except Exception as e:
#                         cs2.close()
#                         conn.commit()
#                         print(e)
#                         if str(e).startswith('ORA-00001'):
#                             try:
#                                 cs2 = conn.cursor()
#                                 sql2 = "update operation_req set " \
#                                        "COLNMKR = :1," \
#                                        "COLNMEN = :2," \
#                                        "COLSIZE = :3," \
#                                        "COLTYPE = :4," \
#                                        "SAMPLE = :5," \
#                                        "EXPLAN = :6," \
#                                        "SERVICE_URL = :7," \
#                                        "REQ_URL = :8 " \
#                                        "where apicd = :9 and opercd = :10 and colline = :11"
#                                 print(sql2)
#                                 cs2.execute(sql2,
#                                             (str(colnmkr),
#                                              str(colnmen),
#                                              str(colnmsize), str(coltype), str(sample), str(explan), survice_url, req_url,
#                                              str(apicd), str(operationcd), str(colline)))
#                                 cs2.close()
#                                 conn.commit()
#                             except:
#                                 traceback.print_exc()
#                         else:
#                             traceback.print_exc()
#             if oper_res_df.empty:
#                 try:
#                     colnmkr = 'nan'
#                     colnmen = 'nan'
#                     colnmsize = 'nan'
#                     coltype = 'nan'
#                     sample = 'nan'
#                     explan = 'nan'
#                     cs3 = conn.cursor()
#                     sql2 = "insert into operation_res (APICD,OPERCD,OPERNM,COLLINE,COLNMKR,COLNMEN,COLSIZE,COLTYPE,SAMPLE,EXPLAN)" \
#                            " values (:1,:2,:3,:4,:5,:6,:7,:8,:9,:10)"
#
#                     cs3.execute(sql2, (str(apicd), str(operationcd), str(opertaionnm), '0', str(colnmkr), str(colnmen),
#                                        str(colnmsize),str(coltype),str(sample),str(explan)))
#                     cs3.close()
#                     conn.commit()
#                 except:
#                     traceback.print_exc()
#             else:
#                 for colline in range(0,len(oper_res_df['항목명(국문)'].tolist())):
#                     try:
#                         colnmkr = oper_res_df['항목명(국문)'][colline]
#                         colnmen = oper_res_df['항목명(영문)'][colline]
#                         colnmsize = oper_res_df['항목크기'][colline]
#                         coltype = oper_res_df['항목구분'][colline]
#                         sample = oper_res_df['샘플데이터'][colline]
#                         explan = oper_res_df['항목설명'][colline]
#                         cs3 = conn.cursor()
#                         sql2 = "insert into operation_res (APICD,OPERCD,OPERNM,COLLINE,COLNMKR,COLNMEN,COLSIZE,COLTYPE,SAMPLE,EXPLAN)" \
#                                " values (:1,:2,:3,:4,:5,:6,:7,:8,:9,:10)"
#
#                         cs3.execute(sql2, (str(apicd), str(operationcd), str(opertaionnm), str(colline), str(colnmkr), str(colnmen),
#                                            str(colnmsize),str(coltype),str(sample),str(explan)))
#                         cs3.close()
#                         conn.commit()
#                     except:
#                         traceback.print_exc()
#
#         conn.close()
#     except:
#         conn.close()
#         traceback.print_exc()
#         log.write(apiid + '\n')


