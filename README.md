# 與關鍵字相關徵才的公司

此程式會爬取有與關鍵字相關人才需求的公司，提供資訊：
公司名稱	地址	 產業類別	聯絡人	產業描述	電話	資本額	員工人數，
藉以發掘潛在客戶


## 產生執行檔案
```
pip install virtualenv
virtualenv venv
source   venv/bin /activate   (linux)  
../venv/Script/active  (windows)
pyinstaller -F company_check.py

# 加上版號 加上icon 
pyinstaller --version-file version_info.txt -i setup.ico -F company_check.py

```
## 加上版號 參數差異
打包成單執行檔(適合單檔)：
pyinstaller -F <python file>   

打包成多個文件(多檔，適合框架)：
pyinstaller -D <python file>   

***


