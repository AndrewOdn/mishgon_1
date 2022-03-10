import pandas as pd
import requests
import re
import csv
import time
import xlsxwriter
start_time = time.time()

from bs4 import BeautifulSoup
# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

header = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.20 Safari/537.36',
    'Upgrade-Insecure-Requests': '1',
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'same - origin',
    'Sec-Fetch-User': '?1',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
    'Accept-Encoding': 'gzip, deflate, br',
    'Accept-Language': 'ru-RU,ru;q=0.9',
    'Cache-Control': 'no-cache',
    'pragma': 'no-cache',
    'cookie': 'tester=%D0%98%D0%BD%D0%BA%D0%BE%D0%B3%D0%BD%D0%B8%D1%82%D0%BE; _ga=GA1.2.1031594214.1646727351; _gid=GA1.2.2084028604.1646727351; _ym_uid=1646727351761286665; _ym_d=1646727351; _ym_isad=2'
}
df = pd.read_csv('sites.csv', encoding='utf-8')
print('Введите название файла')
name = str(input())
workbook = xlsxwriter.Workbook(name+'.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write(0, 0, 'site')
worksheet.write(0, 1, 'tel')
worksheet.write(0, 2, 'mail')
print('Введите 2 часла, с какого - по какой номер сайта парсить')
print('Минимальный номер - 0, максимальный равен колличеству строк с сайтами, место него можно написать -1 для прохода по всем сайтам в .csv')
print('Пример ввода 1:')
print('0')
print('-1')
print('Программа пройдет по всем сайтам в файле')
print('\n')
print('Пример ввода 2:')
print('4')
print('8')
print('Программа пройдет по сайтам под порядковыми номерами 4,5,6,7,8')
print('P.S на каждый сайт уходит примерно 2с, не советую запускать сразу больше 10000')
mink = int(input())
maxk = int(input())
r = requests.Session()
"""first"""
if maxk == -1:
    maxk = len(df['sites'])
for i in range(mink, maxk):
    mail = ''
    tel = []
    try:
        s = r.get('http://' + df['sites'][i], headers=header)
        a = s.status_code
    except:
        a = -1
        print('http://' + df['sites'][i])
        telll = ''
        for asdasd in range(0, len(tel)):
            if asdasd == 0:
                telll = telll + str(tel[asdasd])
            else:
                telll = telll + ', ' + str(tel[asdasd])
        worksheet.write(i + 1, 0, 'http://' + df['sites'][i])
        worksheet.write(i + 1, 1, telll)
        worksheet.write(i + 1, 2, mail)
    if a == 200:
        print('http://' + df['sites'][i])
        text = str(s.text)
        soup = BeautifulSoup(text, 'html.parser')
        tel_txt = soup.text
        if str.find(text, 'mailto') != -1:
            mail = text[str.find(text, 'mailto:')+7:]
            mail = mail[:str.find(mail, '"')]
        print(mail)
        tel = re.findall(r"(?:(?:8|\+7)[\- ]?)?(?:\(?\d{3}\)?[\- ]?)?[\d\- ]{7,16}", tel_txt)
        kkk1 = len(tel)
        kkk2 = 0
        for d in range(0, kkk1):
            tel[d-kkk2] = re.sub("[^+\d]", "", tel[d-kkk2])
            if tel[d-kkk2] == '' or len(tel[d-kkk2]) < 11 or len(tel[d-kkk2]) > 12:
                del tel[d-kkk2]
                kkk1 = kkk1-1
                kkk2 += 1
                d = d - 1
        print(tel)
        worksheet.write(i+1, 0, 'http://' + df['sites'][i])
        telll = ''
        for asdasd in range(0, len(tel)):
            if asdasd == 0:
                telll = telll + str(tel[asdasd])
            else:
                telll = telll +', '+ str(tel[asdasd])
        worksheet.write(i + 1, 1, telll)
        worksheet.write(i + 1, 2, mail)
workbook.close()
print("--- %s seconds ---" % (time.time() - start_time))
input('введите любой символ для выхода')