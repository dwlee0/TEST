from bs4 import BeautifulSoup
import urllib.request
import xlsxwriter

download_path = 'C:/Test/Test.xlsx'

stock_code_list = list()
stock_name_list = list()
for i in range(1, 11):
    page = urllib.request.urlopen('http://finance.daum.net/quote/marketvalue.daum?stype=P&page=' + str(i) + '&col=listprice&order=desc')
    soup = BeautifulSoup(page.read().decode('utf-8'), "html.parser")
    elements = soup.findAll('td', {'class' : 'txt'})
    for e in elements:
        stock_code = e.find('a')['href'][-6:]
        stock_name = e.find('a').contents[0]
        stock_code_list.append(stock_code)
        stock_name_list.append(stock_name)

workbook = xlsxwriter.Workbook(download_path)
worksheet = workbook.add_worksheet()

row = 1
col = 0

for code, name in zip(stock_code_list, stock_name_list):
    worksheet.write(row, col, code)
    worksheet.write(row, col + 1, name)
    row += 1


workbook.close()