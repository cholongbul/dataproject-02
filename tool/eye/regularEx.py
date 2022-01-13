import re

data = """
park-a 800905-1049118
kim-a 700905-1059119
"""

pat = re.compile("(\d{6})[-]\d{7}")
print(pat.sub("\g<1>-*******",data))

#[0-9] == \d
#[^0-9] == \D
#[ \t\n\r\f\v] == \s
#[^ \t\n\r\f\v] == \S
#[a-zA-Z0-9_] == \w
#[^a-zA-Z0-9_] == \W

#dot(.) == \n 제외 모든문자
#*==0부터 무한대까지 반복
#+==1부터 무한대까지 반복
#{m,n} m부터 n까지 반복
#? 는 있어도 되고 없어도 된다
#
# p = re.compile('[a-z]+')
# #match
# m = p.match("python")
# print(m)
#
# p = re.compile('[a-z]+')
# #match
# m = p.match("3 python")
# print(m)
#
# p = re.compile('C[0-9]+')
# m = p.match('2C900')
# if m:
#     print('Match found: ', m.group())
# else:
#     print('No match')

#search
# p = re.compile('[a-z]+')
# m = p.search("python")
# print(m)
#
# m = p.search("3 python")
# print(m)
#
# #findall
# m = p.findall("life is too short")
# print(m)
#
# #finditer
# m = p.finditer("life is too short")
# for r in m: print(r)
#
# m = p.search("3 python")
# print(m.group())
# print(m.start())
# print(m.end())
# print(m.span())
#

# m = re.match('a.b', 'a\nb')
# p = re.compile('a.b', re.DOTALL)
# m = p.match('a\nb')
# print(m)


# p = re.compile('[a-z]+', re.I)
# m = p.match('Python')
# print(m)
#
# p = re.compile("^python\s\w+", re.M)
# data = """python one
# life is too short
# python two
# you need python
# python tree"""
#
# print(p.findall(data))

charref = re.compile(r'&[#](0[0-7]+|[0-9]+|x[0-9a-fA-F]+);')