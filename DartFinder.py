import win32com.client
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
import os
import time
import re

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
Datafile = os.getcwd() + "\기업리스트.xlsx"

wb = excel.Workbooks.Open(Datafile)
ws = wb.ActiveSheet

used = ws.UsedRange
nrows = used.Row + used.Rows.Count - 1
ncols = used.Column + used.Columns.Count - 1


inputfile = os.getcwd() + "\chromedriver"
driver = webdriver.Chrome(inputfile)
driver.implicitly_wait(3)


for i in range(1,  nrows + 1):
    if ws.Cells(i, 1).Value is None :
        break
    http = 'http://dart.fss.or.kr/html/search/SearchCompanyIR3_M.html?textCrpNM=' + ws.Cells(i, 1).Value
    driver.get(http)
    driver.execute_script("openSearchCorpWindow();")
    time.sleep(5)

    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')

    prodList = soup.find_all("p", "page_info")

    for s in prodList:
        Test = re.findall("\d+", str(s))

        if not Test:
            ws.Cells(i, 2).Value = 0
            break

        ws.Cells(i, 2).Value = int(Test[2])

        if int(Test[2]) > 0:
            ws.Cells(i, 2).Interior.ColorIndex = 44

            cropList = soup.select('#corpListContents > div > fieldset > div.table_scroll > table > tbody > tr')
            cols = 3
            colorCount = 35
            for tr in cropList:

                crpName = tr.find_all('input')
                for Name in crpName:
                    ws.Cells(i, cols).Value = Name["value"]
                    ws.Cells(i, cols).Interior.ColorIndex = colorCount
                    cols = cols + 1
                    break

                img = tr.find_all('img')
                for g in img:
                    ws.Cells(i, cols).Value = g["alt"]
                    ws.Cells(i, cols).Interior.ColorIndex = colorCount
                    cols = cols + 1

                tds = tr.find_all('td')
                ws.Cells(i, cols).Value = tds[1].string
                ws.Cells(i, cols).Interior.ColorIndex = colorCount
                cols = cols + 1
                ws.Cells(i, cols).Value = tds[2].string
                ws.Cells(i, cols).Interior.ColorIndex = colorCount
                cols = cols + 1
                ws.Cells(i, cols).Value = tds[3].string
                ws.Cells(i, cols).Interior.ColorIndex = colorCount
                cols = cols + 1
                colorCount = ((colorCount + 34) % 4) + 35

print("☆★☆★☆ 다트 검색 완료 ☆★☆★☆ ")

driver.close()

outputfile = os.getcwd() + "\다트검색_기업리스트.xlsx"
wb.SaveAs(outputfile)
excel.Quit()

