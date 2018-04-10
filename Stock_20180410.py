#-*- coding: utf-8 -*-
from bs4 import BeautifulSoup
import xlsxwriter
import re
import requests
import time


def get_stockinfo(worksheet):

    row = 1

    for i in range(1, 92): #92
        url = 'http://finance.daum.net/quote/marketvalue.daum?stype=KOS&page=%s' % (str(i))
        page = requests.get(url).text
        soup = BeautifulSoup(page, 'lxml')
        elements = soup.findAll('td', {'class': 'txt'})
        price = soup.findAll('td', {'class': 'num'})

        row1 = 1

        for e in elements:
            stock_code = e.find('a')['href'][-6:]
            stock_name = e.find('a').contents[0]

            worksheet.write(row, 0, stock_code)
            worksheet.write(row, 1, stock_name)

            for j in range(0, 6):
                worksheet.write(row, j + 2, price[6 * (row1 - 1) + j].text)

            [per, pbr, bae] = get_detailinfo(stock_code)

            worksheet.write(row, 8, per)
            worksheet.write(row, 9, pbr)
            worksheet.write(row, 10, bae)

            row += 1
            row1 += 1


def get_detailinfo(code):

    url = 'http://wisefn.stock.daum.net/company/cF1001.aspx?cmp_cd=%s&finGubun=MAIN' % code
    page = requests.get(url).text
    soup = BeautifulSoup(page, 'lxml')
    p = re.compile("changeFinData.*?\;", re.DOTALL)
    cfd_list = p.findall(soup.prettify())
    cfd_str = cfd_list[0].replace('\n', '').replace('changeFinData = ', '').replace(';', '')
    cfd = eval(cfd_str)

    per2017 = cfd[3][0][10][3]  # [3][0 연간 1 분기]
    pbr2017 = cfd[3][0][12][3]
    bae2017 = cfd[3][0][14][3]

    return [per2017, pbr2017, bae2017]


def initsheet(worksheet):
    worksheet.write(0, 0, '종목코드')
    worksheet.write(0, 1, '종목명')
    worksheet.write(0, 2, '순위')
    worksheet.write(0, 3, '현재가')
    worksheet.write(0, 4, '전일대비')
    worksheet.write(0, 5, '등락률')
    worksheet.write(0, 6, '시가총액')
    worksheet.write(0, 7, '총주식')
    worksheet.write(0, 8, 'PER')
    worksheet.write(0, 9, 'PBR')
    worksheet.write(0, 10, '배당률')


def main():
    file = 'U:/99_Temp/Test.x'
    workbook = xlsxwriter.Workbook(file)
    worksheet = workbook.add_worksheet('KOS')
    initsheet(worksheet)
    get_stockinfo(worksheet)
    workbook.close()


if __name__ == "__main__":
    tic = time.clock()
    main()
    toc = time.clock()
    print(toc-tic)
    # 331.56
