# 爬取財報狗和 CMoney 的資料並整合至 excel 中

- 爬取財報狗的台股 [市場焦點](https://statementdog.com/market-trend) 前後各 5 名族群和各族群中的前三檔股票的開高低收資料。
- 爬取 CMoney 的 [資金流向](https://www.cmoney.tw/finance/f00018.aspx?o=1&o2=4) 前後各 10 名族群資料中的前三檔成交量的股票的開高低收資料。
- 將資料整合進 excel 中。

## Usage
- 先 clone 專案到 local 中
  ```bash
  git clone https://github.com/weijay0804/stock_crawler.git
  ```
- 進入到專案目錄
  ```bash
  cd stock_crawler
  ```
- 建立虛擬環境  
  範例是使用 `pipenv` 建立虛擬環境，如果習慣用其他方式安裝也可以  

  - 如果還沒安裝 `pipenv` ，請執行：
  ```bash
  pip install pipenv
  ```

  - 在專案目錄下建立 `.venv` 資料夾，或執行以下命令：
  ```bash
  mkdir .venv
  ```

  - 建立虛擬環境：
  ```bash
  pipenv install

  # 啟動虛擬環境
  pipenv shell
  ```

- 安裝瀏覽器 driver  

  - 查看 chrome 版本：
    打開 chrome 後按右上角的三個小點點的按鈕 -> 設定 -> 關於 chrome

  - 下載 driver
    
    如果 chrome 的版本是 `115` 以下，請前往 [這個](https://sites.google.com/chromium.org/driver/downloads?authuser=0) 頁面  
    如果是 `115` 以上，請前往 [這個](https://googlechromelabs.github.io/chrome-for-testing/) 頁面  

    根據你的 chorme 版本選擇對應的 driver，點進去後根據你的作業系統下載對應的檔案。

  *注意*：不要更改下載檔案的檔名，檔名應該會是 `chromedriver`

  下載完成後，將下載的檔案放到專案根目錄下。

- 啟動程式
  ```bash
  python main.py
  ```

執行完畢後，資料會放在 `data` 資料夾中

*注意*：不要更改 `data` 資料夾中的 excel 檔名，程式會根據檔名執行更新。
