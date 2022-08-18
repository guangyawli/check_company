# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


import bs4
import requests
#import json
import pandas
import openpyxl
import datetime
from os.path import exists

FilePath = "sample.xlsx"


def print_hi(name):

    file_exists = exists(FilePath)

    if file_exists is True:
        ExcelWorkbook = openpyxl.load_workbook(FilePath)
        writer = pandas.ExcelWriter(FilePath, engine='openpyxl')
        writer.book = ExcelWorkbook

    tmp_name_target = []
    tmp_address = []
    for j in range(1, 50):
        turl = "https://www.104.com.tw/jobs/search/?keyword=NX%20UG&order=1&jobsource=2018indexpoc&ro=0&page="+str(j)
        htmlfile = requests.get(turl)
        #print(htmlfile.status_code)
        #print(htmlfile.text)
        if "搜尋條件無符合工作機會，建議放寬條件重新查詢" in htmlfile.text:
            #print(j)
            break
        else:
            objssoup = bs4.BeautifulSoup(htmlfile.text, 'lxml')
            objstag = objssoup.select('ul li a')
            #print(objstag)
            #rint(str(objstag[5]).split("公司住址：")[1].split('">')[0])

            for i in range(5, len(objstag)):
                if objstag[i].text.rstrip() not in tmp_name_target:
                    tmp_name_target.append(objstag[i].text.rstrip())
                    tmp_address.append(str(objstag[i]).split("公司住址：")[1].split('">')[0])
                    #print(objstag[i].text.rstrip())
    #print(list(zip(tmp_name_target,tmp_address)))

    final_target = pandas.Series(tmp_address, index=tmp_name_target)
    if file_exists is True:
        final_target.to_excel(writer, sheet_name=str(datetime.date.today()), header=False, startcol=0, startrow=1)

        writer.save()
        writer.close()
    else:
        final_target.to_excel("sample.xlsx", sheet_name=str(datetime.date.today()), header=False, startcol=0, startrow=1)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/