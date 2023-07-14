# stock_crawler
# 前言
這是一個方便且高效的股票每日數據抓取程式。

使用這個程式，你只需在XML檔案中填入欲查詢的股票代碼、起始年月和終止年月和檔案名稱，然後運行程式。它會根據提供的參數，自動從股票市場數據提供者的API中獲取股票的每日數據。

獲取的數據將保存在一個Excel文件中，每一行代表一天的數據，包括日期、成交股數、成交金額、開盤價、最高價、最低價、收盤價、漲跌價差和成交筆數等。

這個程式的優勢在於它的靈活性和效率。你可以根據自己的需求輕鬆定製化參數，並獲取特定日期範圍內的股票數據。它能夠自動化數據抓取的過程，節省你寶貴的時間和精力。

不論你是個人投資者、分析師還是股票愛好者，這個程式都是你抓取股票每日數據的理想選擇。它簡單易用，讓你輕鬆獲取準確的股票數據，並助你做出明智的投資決策。
# 如何使用
1. 先將exe檔和xml檔下載下來
2. 打開xml檔案
3. 輸入欲查詢的股票代碼至stockNo的欄位內
4. 輸入欲查詢的日期範圍至四個欄位內，分別為startYear、startMonth、endYear、endMonth
5. 在excelname的欄位中輸入最後輸出的excel檔案名稱
6. 輸入完之後存檔(存檔很重要!)關閉xml
7. 執行exe檔
8. 最後即可獲得一個整理好資料的excel檔案
