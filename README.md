# 徵 NX UG 人才的公司

此程式會爬取有 NX UG 人才需求的公司，提供資訊：
公司名稱	地址	 產業類別	聯絡人	產業描述	電話	資本額	員工人數，
藉以發掘潛在客戶

## 以Git命令模式複製專案

- Git command:

```
cd projects
git clone http://172.18.1.79/kevinli/check_company.git

```

## 從 Visual Studio 下載專案
- 在歡迎畫面選擇 "複製存放庫"
- 存放庫位置填寫：http://172.18.1.79/kevinli/check_company.git ，並修改專案要放置的路徑後，按下"複製"

## 執行程式
- 打開專案之後，在 Visual Studio 的方案總管中的 "Python 環境"按滑鼠右鍵"新增環境"，系統就會根據 requirements.txt 安裝相關的元件，
- 執行"開始偵錯"(F5)，就會開始收集資料

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


