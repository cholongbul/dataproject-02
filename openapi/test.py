from bs4 import BeautifulSoup
import module.module as md
import traceback
import ssl
from urllib import parse
from urllib import request
from urllib.error import HTTPError
import pandas as pd

servicekey_tail = 'WBaXX3pce9C9AKfYTQc5%2FXVYPXYJWfHVzWNaird%2Fv0f8C0zKhPFhjY10Tuf2QuiA83hfkGLzHknlOz5FWPbaDQ%3D%3D'

try:
    cursor = md.dbcursor_module()
except:
    traceback.print_exc()
apiidfile = open('apiid2.txt', 'r', encoding='utf8')
apiidlist = apiidfile.readlines()
result_df = pd.DataFrame(
    columns=['apicd', 'apinm', 'opernm', 'dataformmat', 'xmlurl', 'xml', 'jsonurl', 'json', 's7', 's8'])
index_cnt = 0

for apiid in apiidlist:
    print(apiid)
    apiid = str(apiid).rstrip('\n')
    try:
        ##sql문 api목록 요청, 판다스에 담기
        apilist_df = md.apilist_df_module(cursor, apiid)
        dataformmat = apilist_df.iloc[0]['DATAFORMMAT']
        docxname = apilist_df.iloc[0]['DOCXNAME']
        api_name = apilist_df['APINM'][0]
        api_type = apilist_df['APITYPE'][0]
        organ_name = apilist_df['ORGANNM'][0]
        api_portal_url = 'https://www.data.go.kr/data/' + apiid + '/openapi.do'
    except:
        traceback.print_exc()
    try:
        opercode_list = md.oper_req_df_A_B(cursor, apiid)
    except:
        log = open('./error.txt', 'a', encoding='utf8')
        log.write(apiid + '\n')
    for opercode in opercode_list:

        oper_req_df = md.oper_req_df_B(cursor, apiid, str(opercode))
        opernm = oper_req_df.iloc[0]['OPERNM']
        service_url = oper_req_df.iloc[0]['SERVICE_URL']
        oper_req_url = oper_req_df.iloc[0]['REQ_URL']
        print(oper_req_df['REQ_URL'])

        if dataformmat == 'XML':
            try:
                paramlist = oper_req_df['COLNMEN'].tolist()
                samplelist = oper_req_df['SAMPLE'].tolist()
                query = {}
                for i in range(0, len(paramlist)):
                    if str(paramlist[i]).lower() == 'servicekey' or str(paramlist[i]).lower() == 'authapikey' or str(
                            paramlist[i]).lower() == 'nan' or str(samplelist[i]).lower() == 'nan':
                        continue
                    else:
                        query[paramlist[i]] = samplelist[i]
                context = ssl._create_unverified_context()
                print(oper_req_url)
                if not str(oper_req_url) == 'nan':
                    api_req_link = str(oper_req_url) + '?' + parse.urlencode(query,
                                                                             doseq=True) + '&serviceKey=' + servicekey_tail
                else:
                    api_req_link = str(service_url) + '?' + parse.urlencode(query,
                                                                            doseq=True) + '&serviceKey=' + servicekey_tail
                api_req_link2 = ''
                headers = {'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 9_3_2 like Mac OS X) AppleWebKit/601.1.46 (KHTML, like Gecko) Version/9.0 Mobile/13F69 Safari/601.1', 'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8', 'Accept-Charset': 'ISO-8859-1,utf-8;q=0.7,*;q=0.3', 'Accept-Encoding': 'none', 'Accept-Language': 'en-US,en;q=0.8', 'Connection': 'keep-alive'}
                api_req = request.Request(api_req_link, headers=headers)

                try:
                    req_html = request.urlopen(api_req, data=None, context=context, timeout=30)
                    requrl = api_req.get_full_url()
                    requrl2 = ''
                except HTTPError as e:
                    req_html = e
                req_html = req_html.read().decode("utf-8")
                req_html = BeautifulSoup(req_html, 'html.parser')
                req_html = str(req_html)
                req_html2 = ''
            except Exception as e:
                traceback.print_exc()
                if str(e).startswith("'utf-8' codec can't decode"):
                    req_html = '이미지'
                    req_html2 = ''
                else:
                    req_html = "응답없음"
                    req_html2 = ''

        elif dataformmat == 'JSON':
            try:
                paramlist = oper_req_df['COLNMEN'].tolist()
                samplelist = oper_req_df['SAMPLE'].tolist()
                query = {}
                for i in range(0, len(paramlist)):
                    if str(paramlist[i]).lower() == 'servicekey' or str(paramlist[i]).lower() == 'authapikey' or str(
                            paramlist[i]).lower() == 'nan' or str(samplelist[i]).lower() == 'nan':
                        continue
                    else:
                        query[paramlist[i]] = samplelist[i]

                context = ssl._create_unverified_context()
                if not str(oper_req_url) == 'nan':
                    api_req_link2 = str(oper_req_url) + '?' + parse.urlencode(query,
                                                                              doseq=True) + '&serviceKey=' + servicekey_tail
                else:
                    api_req_link2 = str(service_url) + '?' + parse.urlencode(query,
                                                                             doseq=True) + '&serviceKey=' + servicekey_tail
                api_req_link = ''
                headers = {
                    'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 9_3_2 like Mac OS X) AppleWebKit/601.1.46 (KHTML, like Gecko) Version/9.0 Mobile/13F69 Safari/601.1',
                    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                    'Accept-Charset': 'ISO-8859-1,utf-8;q=0.7,*;q=0.3', 'Accept-Encoding': 'none',
                    'Accept-Language': 'en-US,en;q=0.8', 'Connection': 'keep-alive'}
                api_req = request.Request(api_req_link2, headers=headers)

                try:
                    req_html2 = request.urlopen(api_req, data=None, context=context, timeout=30)
                    requrl2 = api_req.get_full_url()
                    requrl = ''
                except HTTPError as e:
                    req_html2 = e
                req_html2 = req_html2.read().decode("utf-8")
                req_html2 = BeautifulSoup(req_html2, 'html.parser')
                req_html2 = str(req_html2)
                req_html = ''
            except Exception as e:
                traceback.print_exc()
                if str(e).startswith("'utf-8' codec can't decode"):
                    req_html2 = '이미지'
                    req_html = ''
                else:
                    req_html2 = "응답없음"
                    req_html = ''
        elif dataformmat == 'JSON+XML':
            returnparam_list = ['_returntype', '_type', 'act', 'apitype', 'contenttype', 'dataformat',
                                'dataformat', 'datatype', 'datetype'
                , 'format', 'output', 'resulttype', 'retuntype', 'returntype', 'service_type', 'type',
                                'viewtype', 'viewtype']
            try:
                paramlist = oper_req_df['COLNMEN'].tolist()
                samplelist = oper_req_df['SAMPLE'].tolist()
                query = {}
                for i in range(0, len(paramlist)):
                    if str(paramlist[i]).lower() == 'servicekey' or str(paramlist[i]).lower() == 'authapikey' or str(
                            paramlist[i]).lower() == 'nan' or str(samplelist[i]).lower() == 'nan':
                        continue
                    elif paramlist[i].lower() in returnparam_list:
                        query[paramlist[i]] = 'xml'
                    else:
                        query[paramlist[i]] = samplelist[i]

                context = ssl._create_unverified_context()
                if not str(oper_req_url) == 'nan':
                    api_req_link = str(oper_req_url) + '?' + parse.urlencode(query,
                                                                             doseq=True) + '&serviceKey=' + servicekey_tail
                else:
                    api_req_link = str(service_url) + '?' + parse.urlencode(query,
                                                                            doseq=True) + '&serviceKey=' + servicekey_tail
                headers = {
                    'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 9_3_2 like Mac OS X) AppleWebKit/601.1.46 (KHTML, like Gecko) Version/9.0 Mobile/13F69 Safari/601.1',
                    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                    'Accept-Charset': 'ISO-8859-1,utf-8;q=0.7,*;q=0.3', 'Accept-Encoding': 'none',
                    'Accept-Language': 'en-US,en;q=0.8', 'Connection': 'keep-alive'}
                api_req = request.Request(api_req_link, headers=headers)

                try:
                    req_html = request.urlopen(api_req, data=None, context=context, timeout=30)
                    requrl = api_req.get_full_url()
                except HTTPError as e:
                    req_html = e
                req_html = req_html.read().decode("utf-8")
                req_html = BeautifulSoup(req_html, 'html.parser')
                req_html = str(req_html)
            except Exception as e:
                traceback.print_exc()
                if str(e).startswith("'utf-8' codec can't decode"):
                    req_html = '이미지'
                else:
                    req_html = "응답없음"
            try:
                paramlist = oper_req_df['COLNMEN'].tolist()
                samplelist = oper_req_df['SAMPLE'].tolist()
                paramlowerlist = []
                for param in paramlist:
                    paramlowerlist.append(str(param).lower())
                if not 'servicekey' in paramlowerlist or not 'authapikey' in paramlowerlist:
                    paramlist.append('serviceKey')
                    samplelist.append(servicekey_tail)
                query = {}
                for i in range(0, len(paramlist)):
                    if str(paramlist[i]).lower() == 'servicekey' or str(paramlist[i]).lower() == 'authapikey' or str(
                            paramlist[i]).lower() == 'nan' or str(samplelist[i]).lower() == 'nan':
                        continue
                    elif paramlist[i].lower() in returnparam_list:
                        query[paramlist[i]] = 'json'
                    else:
                        query[paramlist[i]] = samplelist[i]

                context = ssl._create_unverified_context()
                if not str(oper_req_url) == 'nan':
                    api_req_link2 = str(oper_req_url) + '?' + parse.urlencode(query,
                                                                              doseq=True) + '&serviceKey=' + servicekey_tail
                else:
                    api_req_link2 = str(service_url) + '?' + parse.urlencode(query,
                                                                             doseq=True) + '&serviceKey=' + servicekey_tail
                headers = {
                    'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 9_3_2 like Mac OS X) AppleWebKit/601.1.46 (KHTML, like Gecko) Version/9.0 Mobile/13F69 Safari/601.1',
                    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                    'Accept-Charset': 'ISO-8859-1,utf-8;q=0.7,*;q=0.3', 'Accept-Encoding': 'none',
                    'Accept-Language': 'en-US,en;q=0.8', 'Connection': 'keep-alive'}
                api_req = request.Request(api_req_link2, headers=headers)

                try:
                    req_html2 = request.urlopen(api_req, data=None, context=context, timeout=30)
                    requrl2 = api_req.get_full_url()
                except HTTPError as e:
                    req_html2 = e
                req_html2 = req_html2.read().decode("utf-8")
                req_html2 = BeautifulSoup(req_html2, 'html.parser')
                req_html2 = str(req_html2)
            except Exception as e:
                traceback.print_exc()
                if str(e).startswith("'utf-8' codec can't decode"):
                    req_html2 = '이미지'
                else:
                    req_html2 = "응답없음"

        if '이미지' in req_html:
            s7 = '정상'
            s8 = "정상"
        elif dataformmat == 'XML' and '</' in req_html[:1000]:
            s7 = '정상'

        elif dataformmat == 'JSON' and '{' in req_html2[:1000]:
            s7 = '정상'
        elif dataformmat == 'JSON' and not '{' in req_html2[:1000]:
            if 'wms' in req_html2.lower() or 'wfs' in req_html2.lower():
                s7 = "정상"
                s8 = "정상"
            else:
                s7 = "오류"
        elif dataformmat == 'JSON+XML' and '</' in req_html[:1000] and '{' in req_html2[:1000]:
            s7 = '정상'

        else:
            if 'wms' in req_html.lower() or 'wms' in req_html2.lower() or 'wfs' in req_html.lower() or 'wfs' in req_html2.lower() or '이미지' in req_html or '이미지' in req_html2:
                s8 = "정상"
                s7 = "정상"
            else:
                s7 = '오류'

        if dataformmat == 'XML':
            if 'normal service' in req_html.lower():
                s8 = "정상"
            elif 'wms' in req_html.lower():
                s8 = "정상"
            elif '<result_msg>ok</result_msg>' in req_html.replace(' ', '').replace('\n', '').lower():
                s8 = "정상"
            elif 'normal_service' in req_html.lower():
                s8 = "정상"
            elif 'non_error' in req_html.lower():
                s8 = "정상"
            elif req_html == '응답없음':
                s8 = "오류"
            elif 'success' in req_html.lower():
                s8 = "정상"
            elif '정상' in req_html:
                s8 = "정상"
            elif '성공' in req_html:
                s8 = "정상"
            elif '이미지' in req_html:
                s8 = "정상"
            elif 'success_info' in req_html.lower():
                s8 = "정상"
            elif '<successYN>N</successYN>' in req_html.lower():
                s8 = "오류"
            elif 'normal' in req_html.lower():
                s8 = '정상'
            elif 'nomal' in req_html.lower():
                s8 = '정상'
            elif '<successYN>Y</successYN>' in req_html.lower():
                s8 = '정상'
            elif '<resultcode>00' in req_html.lower():
                s8 = '정상'
            elif 'wfs' in req_html.lower():
                s8 = '정상'
            elif 'soapenv' in req_html.lower():
                s8 = "오류"
            elif 'errorcode' in req_html.lower():
                s8 = "오류"
            elif 'service error' in req_html.lower():
                s8 = "오류"
            elif 'SERVICE KEY IS NOT' in req_html.upper():
                s8 = "오류"
            elif 'APPLICATION ERROR' in req_html.upper():
                s8 = "오류"
            elif 'DB ERROR' in req_html.upper():
                s8 = "오류"
            elif 'NODATA ERROR' in req_html.upper():
                s8 = "오류"
            elif 'HTTP ERROR' in req_html.upper():
                s8 = "오류"
            elif 'SERVICETIMEOUT ERROR' in req_html.upper():
                s8 = "오류"
            elif 'NO OPENAPI SERVICE ERROR' in req_html.upper():
                s8 = "오류"
            elif 'SERVICE ACCESS DENIED ERROR' in req_html.upper():
                s8 = "오류"
            elif 'LIMITED NUMBER OF SERVICE REQUESTS EXCEEDS ERROR' in req_html.upper():
                s8 = "오류"
            elif 'SERVICE KEY IS NOT REGISTERED ERROR' in req_html.upper():
                s8 = "오류"
            elif 'DEADLINE HAS EXPIRED ERROR' in req_html.upper():
                s8 = "오류"
            elif 'UNREGISTERED IP ERROR' in req_html.upper():
                s8 = "오류"
            else:
                s8 = "정상"
        elif dataformmat == "JSON":
            if 'normal service' in req_html2.lower():
                s8 = "정상"
            elif 'wms' in req_html2.lower():
                s8 = "정상"
            elif '<result_msg>ok</result_msg>' in req_html2.replace(' ', '').replace('\n', '').lower():
                s8 = "정상"
            elif 'normal_service' in req_html2.lower():
                s8 = "정상"
            elif 'non_error' in req_html2.lower():
                s8 = "정상"
            elif req_html2 == '응답없음':
                s8 = "오류"
            elif 'success' in req_html2.lower():
                s8 = "정상"
            elif '정상' in req_html2:
                s8 = "정상"
            elif '성공' in req_html2:
                s8 = "정상"
            elif '이미지' in req_html2:
                s8 = "정상"
            elif 'success_info' in req_html2.lower():
                s8 = "정상"
            elif '<successYN>N</successYN>' in req_html2.lower():
                s8 = "오류"
            elif 'normal' in req_html2.lower():
                s8 = '정상'
            elif 'nomal' in req_html2.lower():
                s8 = '정상'
            elif '<successYN>Y</successYN>' in req_html2.lower():
                s8 = '정상'
            elif '<resultcode>00' in req_html2.lower():
                s8 = '정상'
            elif 'wfs' in req_html2.lower():
                s8 = '정상'
            elif 'soapenv' in req_html2.lower():
                s8 = "오류"
            elif 'errorcode' in req_html2.lower():
                s8 = "오류"
            elif 'service error' in req_html2.lower():
                s8 = "오류"
            elif 'SERVICE KEY IS NOT' in req_html2.upper():
                s8 = "오류"
            elif 'APPLICATION ERROR' in req_html2.upper():
                s8 = "오류"
            elif 'DB ERROR' in req_html2.upper():
                s8 = "오류"
            elif 'NODATA ERROR' in req_html2.upper():
                s8 = "오류"
            elif 'HTTP ERROR' in req_html2.upper():
                s8 = "오류"
            elif 'SERVICETIMEOUT ERROR' in req_html2.upper():
                s8 = "오류"
            elif 'NO OPENAPI SERVICE ERROR' in req_html2.upper():
                s8 = "오류"
            elif 'SERVICE ACCESS DENIED ERROR' in req_html2.upper():
                s8 = "오류"
            elif 'LIMITED NUMBER OF SERVICE REQUESTS EXCEEDS ERROR' in req_html2.upper():
                s8 = "오류"
            elif 'SERVICE KEY IS NOT REGISTERED ERROR' in req_html2.upper():
                s8 = "오류"
            elif 'DEADLINE HAS EXPIRED ERROR' in req_html2.upper():
                s8 = "오류"
            elif 'UNREGISTERED IP ERROR' in req_html2.upper():
                s8 = "오류"
            else:
                s8 = "정상"
        elif dataformmat == 'JSON+XML':
            if 'normal service' in req_html.lower():
                s8 = "정상"
            elif 'wms' in req_html.lower():
                s8 = "정상"
            elif '<result_msg>ok</result_msg>' in req_html.replace(' ', '').replace('\n', '').lower():
                s8 = "정상"
            elif 'normal_service' in req_html.lower():
                s8 = "정상"
            elif 'non_error' in req_html.lower():
                s8 = "정상"
            elif req_html == '응답없음':
                s8 = "오류"
            elif 'success' in req_html.lower():
                s8 = "정상"
            elif '정상' in req_html:
                s8 = "정상"
            elif '성공' in req_html:
                s8 = "정상"
            elif '이미지' in req_html:
                s8 = "정상"
            elif 'success_info' in req_html.lower():
                s8 = "정상"
            elif '<successYN>N</successYN>' in req_html.lower():
                s8 = "오류"
            elif 'normal' in req_html.lower():
                s8 = '정상'
            elif 'nomal' in req_html.lower():
                s8 = '정상'
            elif '<successYN>Y</successYN>' in req_html.lower():
                s8 = '정상'
            elif '<resultcode>00' in req_html.lower():
                s8 = '정상'
            elif 'wfs' in req_html.lower():
                s8 = '정상'
            elif 'soapenv' in req_html.lower():
                s8 = "오류"
            elif 'errorcode' in req_html.lower():
                s8 = "오류"
            elif 'service error' in req_html.lower():
                s8 = "오류"
            elif 'SERVICE KEY IS NOT' in req_html.upper():
                s8 = "오류"
            elif 'APPLICATION ERROR' in req_html.upper():
                s8 = "오류"
            elif 'DB ERROR' in req_html.upper():
                s8 = "오류"
            elif 'NODATA ERROR' in req_html.upper():
                s8 = "오류"
            elif 'HTTP ERROR' in req_html.upper():
                s8 = "오류"
            elif 'SERVICETIMEOUT ERROR' in req_html.upper():
                s8 = "오류"
            elif 'NO OPENAPI SERVICE ERROR' in req_html.upper():
                s8 = "오류"
            elif 'SERVICE ACCESS DENIED ERROR' in req_html.upper():
                s8 = "오류"
            elif 'LIMITED NUMBER OF SERVICE REQUESTS EXCEEDS ERROR' in req_html.upper():
                s8 = "오류"
            elif 'SERVICE KEY IS NOT REGISTERED ERROR' in req_html.upper():
                s8 = "오류"
            elif 'DEADLINE HAS EXPIRED ERROR' in req_html.upper():
                s8 = "오류"
            elif 'UNREGISTERED IP ERROR' in req_html.upper():
                s8 = "오류"
            else:
                s8 = "정상"
            if 'normal service' in req_html2.lower():
                s81 = "정상"
            elif 'wms' in req_html2.lower():
                s81 = "정상"
            elif '<result_msg>ok</result_msg>' in req_html2.replace(' ', '').replace('\n', '').lower():
                s81 = "정상"
            elif 'normal_service' in req_html2.lower():
                s81 = "정상"
            elif 'non_error' in req_html2.lower():
                s81 = "정상"
            elif req_html2 == '응답없음':
                s81 = "오류"
            elif 'success' in req_html2.lower():
                s81 = "정상"
            elif '정상' in req_html2:
                s81 = "정상"
            elif '성공' in req_html2:
                s81 = "정상"
            elif '이미지' in req_html2:
                s81 = "정상"
            elif 'success_info' in req_html2.lower():
                s81 = "정상"
            elif '<successYN>N</successYN>' in req_html2.lower():
                s81 = "오류"
            elif 'normal' in req_html2.lower():
                s81 = '정상'
            elif 'nomal' in req_html2.lower():
                s81 = '정상'
            elif '<successYN>Y</successYN>' in req_html2.lower():
                s81 = '정상'
            elif '<resultcode>00' in req_html2.lower():
                s81 = '정상'
            elif 'wfs' in req_html2.lower():
                s81 = '정상'
            elif 'soapenv' in req_html2.lower():
                s81 = "오류"
            elif 'errorcode' in req_html2.lower():
                s81 = "오류"
            elif 'service error' in req_html2.lower():
                s81 = "오류"
            elif 'SERVICE KEY IS NOT' in req_html2.upper():
                s81 = "오류"
            elif 'APPLICATION ERROR' in req_html2.upper():
                s81 = "오류"
            elif 'DB ERROR' in req_html2.upper():
                s81 = "오류"
            elif 'NODATA ERROR' in req_html2.upper():
                s81 = "오류"
            elif 'HTTP ERROR' in req_html2.upper():
                s81 = "오류"
            elif 'SERVICETIMEOUT ERROR' in req_html2.upper():
                s81 = "오류"
            elif 'NO OPENAPI SERVICE ERROR' in req_html2.upper():
                s81 = "오류"
            elif 'SERVICE ACCESS DENIED ERROR' in req_html2.upper():
                s81 = "오류"
            elif 'LIMITED NUMBER OF SERVICE REQUESTS EXCEEDS ERROR' in req_html2.upper():
                s81 = "오류"
            elif 'SERVICE KEY IS NOT REGISTERED ERROR' in req_html2.upper():
                s81 = "오류"
            elif 'DEADLINE HAS EXPIRED ERROR' in req_html2.upper():
                s81 = "오류"
            elif 'UNREGISTERED IP ERROR' in req_html2.upper():
                s81 = "오류"
            else:
                s81 = "정상"

            if s8 == "정상" or s81 == "정상":
                s8 = "정상"

        if dataformmat == "JSON+XML":
            if 'wms' in req_html.lower() or 'wfs' in req_html.lower() or 'wms' in req_html2.lower() or 'wfs' in req_html2.lower():
                s7 = "정상"
                s8 = "정상"

        result_df.loc[index_cnt] = [apiid, api_name, opernm, dataformmat, api_req_link, req_html, api_req_link2,
                                    req_html2, s7, s8]
        result_df.to_excel("./78result2.xlsx", sheet_name='sheet1')

        index_cnt = index_cnt + 1