from bs4 import BeautifulSoup
import urllib.request
import xlsxwriter

download_path = 'C:/Test/Test.xlsx'
workbook = xlsxwriter.Workbook(download_path)
worksheet = workbook.add_worksheet()

row = 1
col = 0
worksheet.write(0, 0, 'CODE')
worksheet.write(0, 1, 'NAME')

for i in range(1, 11):
    page = urllib.request.urlopen('http://finance.daum.net/quote/marketvalue.daum?stype=P&page=' + str(i) + '&col=listprice&order=desc')
    soup = BeautifulSoup(page.read().decode('utf-8'), "html.parser")
    elements = soup.findAll('td', {'class' : 'txt'})
    for e in elements:
        stock_code = e.find('a')['href'][-6:]
        stock_name = e.find('a').contents[0]
        worksheet.write(row, col, stock_code)
        worksheet.write(row, col + 1, stock_name)
        row += 1
    price = soup.findAll('td', {'class' : 'num'})
    for e in price:
        p1 = e.contents[0]
        print(p1)

workbook.close()