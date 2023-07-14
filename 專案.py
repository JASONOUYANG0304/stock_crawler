import xmltodict
import requests
import time
from openpyxl import Workbook

def fillSheet(sheet, data, row): #建立一個function名稱裡面放置三種參數
    for column, value in enumerate(data, 1):
        sheet.cell(row = row, column = column, value = value)
        #將資料放置在row行column列上，其格子裡填寫value資料
def returnStrDayList(startYear, startMonth, endYear, endMonth, day = "01"):
    result = []
    if startYear == endYear:
        for month in range(startMonth, endMonth+1):
            month = str(month)
            if len(month)==1:
                month = "0" + month
            result.append(str(startYear)+month+day)   
        return result 
    for year in range(startYear, endYear+1):
        if year == startYear:
            for month in range(startMonth, 13):
                month = str(month)
                if len(month)==1:
                    month = "0" + month
                result.append(str(year)+month+day)
        elif year == endYear:
            for month in range(1,endMonth+1):
                month = str(month)
                if len(month)==1:
                    month = "0" + month
                result.append(str(year)+month+day)
        else:
            for month in range(1,13):
                month = str(month)
                if len(month)==1:
                    month = "0" + month
                result.append(str(year)+month+day)
    return result
# 讀取XML檔案
with open('data1.xml',encoding="UTF-8") as file:
    xml_data = file.read()

# 將XML轉換為字典格式
data_dict = xmltodict.parse(xml_data)
data_dict = data_dict["params"]
print(data_dict)
fields = ["日期","成交股數","成交金額","開盤價","最高價","最低價","收盤價","漲跌價差","成交筆數"]
wb = Workbook() #建立excel檔案
sheet = wb.active #讓excel表格成功啟動，建立第一個工作表格
sheet.title = "fields"
fillSheet(sheet, fields, 1) #執行函式

startYear, startMonth = int(data_dict["startYear"]), int(data_dict["startMonth"])
endYear, endMonth = int(data_dict["endYear"]), int(data_dict["endMonth"])
#上面兩行為讀取字典裡的內容，讀取時變正整數

yearList = returnStrDayList(startYear, startMonth, endYear, endMonth) #執行函式
print(yearList)
row = 2
for YearMonth in yearList:
    rq = requests.get(data_dict["url"], params = {"response":"json", "date":YearMonth, "stockNo":data_dict["stockNo"]})
    jsonData = rq.json()
    dailyPriceList = jsonData.get("data", [])
    for dailyPrice in dailyPriceList:
        fillSheet(sheet, dailyPrice, row)
        row += 1
    time.sleep(3)
name = data_dict["excelname"]
wb.save(name+".xlsx") #存檔

