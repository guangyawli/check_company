# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


import bs4
import requests
import pandas
import openpyxl
import datetime
from os.path import exists

FilePath = (str(datetime.date.today()) + "-" + "company.xlsx")


def write_others_info(others):
    # 讀取 Excel 檔案
    wb = openpyxl.load_workbook(FilePath)
    sheet = wb['Sheet1']
    # 以行號、列號指定儲存格
    sheet.cell(row=2, column=3).value = "test"

    print(others)
    wb.save(FilePath)
    wb.close()


def get_company_info(xurl):
    tmp_str = []
    htmlfile2 = requests.get(xurl)
    # print(htmlfile.status_code)
    # print(htmlfile2.text)
    # tmp_str = str(htmlfile2.text).split("profile")[1]
    # print(tmp_str)

    objssoup = bs4.BeautifulSoup(htmlfile2.text, 'lxml')
    objstag = objssoup.select('body div div div div div div div div div div')

    for i in range(0, len(objstag)):

        if "公司網址" and "產業類別" in objstag[i].text:

            if len(objstag[i].text) > 80:
                #print("=========================== start   ========================at" + str(i))
                #print(objstag[i].text)
                objstag[i].text  # .split("產業類別")[1].split("產業描述")[0]
                tmp_str.append(objstag[i].text)

                #print("=========================== end   ===========================")
    return tmp_str


def print_hi():
    file_exists = exists(FilePath)

    if file_exists is True:
        ExcelWorkbook = openpyxl.load_workbook(FilePath)
        writer = pandas.ExcelWriter(FilePath, engine='openpyxl')
        writer.book = ExcelWorkbook

    tmp_name_target = []
    tmp_address = []
    tmp_others = []
    tmp_others_list = []
    for j in range(1, 2):
        print(j)
        # for j in range(1, 3):
        turl = "https://www.104.com.tw/jobs/search/?keyword=NX%20UG&order=1&jobsource=2018indexpoc&ro=0&page=" + str(j)
        htmlfile = requests.get(turl)
        # print(htmlfile.status_code)
        # print(htmlfile.text)
        if "搜尋條件無符合工作機會，建議放寬條件重新查詢" in htmlfile.text:
            # print(j)
            break
        else:
            objssoup = bs4.BeautifulSoup(htmlfile.text, 'lxml')
            objstag = objssoup.select('ul li a')
            # print(objstag)
            # print(str(objstag[5]).split("公司住址：")[1].split('">')[0])
            # print(str(objstag[5]).split('" target=')[0].split('href="')[1])

            for i in range(5, len(objstag)):
                if objstag[i].text.rstrip() not in tmp_name_target:
                    tmp_name_target.append(objstag[i].text.rstrip())
                    tmp_address.append(str(objstag[i]).split("公司住址：")[1].split('">')[0])
                    tmp_others=get_company_info("https:" + str(objstag[i]).split('" target=')[0].split('href="')[1])

                    # get_company_info("https:"+str(objstag[i]).split('" target=')[0].split('href="')[1])
                    # print(objstag[i].text.rstrip())
    # print(list(zip(tmp_name_target,tmp_address)))
            tmp_others_list = tmp_others[0].split("  ")
            #print(tmp_others_list)
            print(len(tmp_name_target))

    final_target = pandas.Series(tmp_address, index=tmp_name_target)
    if file_exists is True:
        # final_target.to_excel(writer, sheet_name=str(datetime.date.today()), header=False, startcol=0, startrow=1)
        final_target.to_excel(writer, header=False, startcol=0, startrow=1)

        writer.save()
        writer.close()
    else:
        final_target.to_excel(FilePath, header=False, startcol=0, startrow=1)

    write_others_info(tmp_others_list)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/