import requests
from bs4 import BeautifulSoup
from xml.etree.ElementTree import parse
from urllib import request
import os


url = 'http://apis.data.go.kr/1390804/NihhsTodayFlowerInfo01/selectTodayFlowerView01?serviceKey=WBaXX3pce9C9AKfYTQc5%2FXVYPXYJWfHVzWNaird%2Fv0f8C0zKhPFhjY10Tuf2QuiA83hfkGLzHknlOz5FWPbaDQ%3D%3D&dataNo='
path = 'C:\\Users\\admin\\Pictures\\꽃사진\\'
for cnt in range(87,366):
    datalist = []
    res = requests.get(url+str(cnt)).text.encode('utf-8')
    xmlobj = BeautifulSoup(res,'lxml-xml')
    fMonthlist = xmlobj.findAll('fMonth')
    fDaylist = xmlobj.findAll('fDay')
    flowNmlist = xmlobj.findAll('flowNm')
    flowlanglist = xmlobj.findAll('flowLang')
    flowNm = str(flowNmlist[0]).rstrip('</flowNm>').lstrip('<flowNm>')
    flowlang = str(flowlanglist[0]).rstrip('</flowLang>').lstrip('<flowLang>')
    fMonth = str(fMonthlist[0]).rstrip('</fMonth>').lstrip('<fMonth>')
    fDay = str(fDaylist[0]).rstrip('</fDay>').lstrip('<fDay>')
    if len(fMonth)<2 and len(fDay)<2:
        monthday = '0' + fMonth + '0' + fDay
    elif len(fMonth)==2 and len(fDay)<2:
        monthday = fMonth + '0' + fDay
    elif len(fMonth)<2 and len(fDay)==2:
        monthday = '0' + fMonth +  fDay
    if len(fMonth)==2 and len(fDay)==2:
        monthday = fMonth + fDay

    imgUrl1 = str(xmlobj.findAll('imgUrl1')[0]).lstrip('<imgUrl1>').rstrip('</imgUrl1>')
    imgUrl2 = str(xmlobj.findAll('imgUrl2')[0]).lstrip('<imgUrl2>').rstrip('</imgUrl2>')
    imgUrl3 = str(xmlobj.findAll('imgUrl3')[0]).lstrip('<imgUrl3>').rstrip('</imgUrl3>')
    request.urlretrieve(imgUrl1+'g',path+monthday+'_'+flowNm+'_'+flowlang+'_1.jpg')
    request.urlretrieve(imgUrl2+'g',path+monthday+'_'+flowNm+'_'+flowlang+'_2.jpg')
    request.urlretrieve(imgUrl3+'g',path+monthday+'_'+flowNm+'_'+flowlang+'_3.jpg')
