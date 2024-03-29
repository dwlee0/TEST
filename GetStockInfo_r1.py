#-*- coding: utf-8 -*-
from bs4 import BeautifulSoup
import urllib.request
import xlsxwriter
import time


def get_stock_info(worksheet, kos):
    row = 1
    col = 0
    row1 = 1

    if kos == 'P':
        pages = 47
    else:
        pages = 43

    for i in range(1, pages):
        url = 'http://finance.daum.net/quote/marketvalue.daum?stype=%s&page=%s' % (kos, str(i))
        page = urllib.request.urlopen(url)
        soup = BeautifulSoup(page.read().decode('utf-8'), "html.parser")
        elements = soup.findAll('td', {'class': 'txt'})
        price = soup.findAll('td', {'class': 'num'})

        for e in elements:
            data = e.find('a')
            stock_code = data['href'][-6:]
            stock_name = data.contents[0]
            [per, pbr, bae] = get_detail_info(stock_code)

            worksheet.write(row, 0, stock_code)
            worksheet.write(row, 1, stock_name)
            worksheet.write(row, 8, per)
            worksheet.write(row, 9, pbr)
            worksheet.write(row, 10, bae)
            row += 1

        for e in price:
            col += 1
            worksheet.write(row1, col + 1, e.text)
            if col == 6:
                col = 0
                row1 += 1


def get_detail_info(code):
    url_tmp = 'http://companyinfo.stock.naver.com/v1/company/ajax/cF1001.aspx?cmp_cd=%s&fin_typ=%s&freq_typ=%s'
    url = url_tmp % (code, '4', 'Y')
    page = urllib.request.urlopen(url)
    soup = BeautifulSoup(page.read().decode('utf-8'), "html.parser")

    data = soup.findAll('td', {'class': 'num'})

    per2017 = data[205].text  # 25 * 8 + 5
    pbr2017 = data[221].text  # 27 * 8 + 5
    bae2017 = data[237].text  # 29 * 8 + 5

    return [per2017, pbr2017, bae2017]


def init_sheet(worksheet):
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
    file = 'C:/Test/Test.xlsx'
    workbook = xlsxwriter.Workbook(file)
    worksheet = workbook.add_worksheet('KOSPI')
    init_sheet(worksheet)
    get_stock_info(worksheet, 'P')

    worksheet = workbook.add_worksheet('KOSDAQ')
    init_sheet(worksheet)
    get_stock_info(worksheet, 'Q')

    workbook.close()


if __name__ == "__main__":
    tic = time.clock()
    main()
    toc = time.clock()
    print(toc - tic)

'''
* 2012 2013 2014 2015 2016 2017 2018 2019 : 2017 col 5
* ROW : PER 25 PBR 27 배당 29
'매출액', '영업이익', '세전계속사업이익', '당기순이익', '당기순이익(지배)', 
'당기순이익(비지배)', '자산총계', '부채총계', '자본총계', '자본총계(지배)', 
'자본총계(비지배)', '자본금', '영업활동현금흐름', '투자활동현금흐름', '재무활동현금흐름', 
'CAPEX', 'FCF', '이자발생부채', '영업이익률', '순이익률', 
'ROE(%)','ROA(%)', '부채비율', '자본유보율', 'EPS(원)', 
'PER(배)', 'BPS(원)', 'PBR(배)', '현금DPS(원)', '현금배당수익률', 
'현금배당성향(%)', '발행주식수(보통주)'
* fin_type = '0': 재무제표 종류 (0: 주재무제표, 1: GAAP개별, 2: GAAP연결, 3: IFRS별도, 4:IFRS연결)
* freq_type = 'Y': 기간 (Y:년, Q:분기)
'''