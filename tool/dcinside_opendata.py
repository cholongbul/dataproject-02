#-------------------------------------------------------------------------------
# Name:        module2
# Purpose:
#
# Author:      seong
#
# Created:     05-01-2022
# Copyright:   (c) seong 2022
# Licence:     <your licence>
#-------------------------------------------------------------------------------
from selenium import webdriver

from bs4 import BeautifulSoup
import urllib.request
import pandas as pd
import time

def get_page(url):
    request = urllib.request.Request(url)
    request.add_header('User-Agent',
                       'Mozilla/5.0')
    request.add_header('Accept',
                       '*/*')
    request.add_header('Accept-Language',
                       'ko-kr,ko;q=0.8,en-us;q=0.5,en;q=0.3')
    request.add_header('Accept-Charset',
                       'ISO-8859-1,utf-8;q=0.7,*;q=0.7')
    request.add_header('Connection',
                       'keep-alive')
    response = urllib.request.urlopen(request,timeout=20)
    html = response.read()
    return html


data_list = []
page = 0
out = open("test.txt","w", encoding="utf-8")
browser = webdriver.Chrome("D:\\chromedriver.exe")
while 1:
    page +=1
    if page > 118:
        break
    url = "https://gall.dcinside.com/mgallery/board/lists/?id=opendata2021&page="+str(page)

    #html = get_page(url)
    browser.get(url)
    time.sleep(2)
    html = browser.page_source
    soup_body = BeautifulSoup(html,"html.parser",from_encoding="utf-8")

    body_html = soup_body('tr', {'class' : 'ub-content us-post'})
    print(url)
    print(str(len(body_html)))
    for body in body_html:
        
        tmp_list = []
        body =BeautifulSoup(str(body),"html.parser")
        num = str(body('td', {'class' : 'gall_num'})[0].text).strip()
        link = "https://gall.dcinside.com"+body.findAll("a")[0]['href']
        title = str(body.findAll("a")[0].text).strip()
        title = str.join(' ', title.split()).strip()
        nickname = str(body('span', {'class' : 'nickname'})[0].text).strip()
        update = str(body('td', {'class' : 'gall_date'})[0]['title']).strip()

        tmp_html = get_page(link)
        tmp_body = BeautifulSoup(tmp_html,"html.parser",from_encoding="utf-8")
        try :
            body_text = str(tmp_body('div', {'class' : 'write_div'})[0].text).strip()
            body_text = str.join(' ', body_text.split()).strip()
        except:
            doby_text=""

        out.write(num+"\t"+title+"\t"+link + "\t"+nickname+"\t"+update+"\t"+body_text+"\n")
        tmp_list.append(num)
        tmp_list.append(title)
        tmp_list.append(link)
        tmp_list.append(nickname)
        tmp_list.append(update)
        tmp_list.append(body_text)
        data_list.append(tmp_list)
        time.sleep(3)
    time.sleep(60)

out.close()
df = pd.DataFrame(data_list)
df.to_csv("dcinside.csv")
browser.quit()


