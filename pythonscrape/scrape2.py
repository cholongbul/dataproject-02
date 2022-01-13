from urllib.request import urlopen
from urllib.error import HTTPError
from bs4 import BeautifulSoup
try:
  html = urlopen("http://www.pythonscraping.com/pages/error.html")
  bs = BeautifulSoup(html.read(),'lxml')
  print(bs.h1)
except HTTPError as e:
  print(e)