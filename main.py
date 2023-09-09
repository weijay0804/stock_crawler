import os
from collections import defaultdict
from datetime import datetime, timedelta

import requests
import openpyxl
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class BaseRequset:
    """
    封裝 request 模組
    """

    @staticmethod
    def get_requset(url: str):
        response = requests.get(url)

        if response.status_code != 200:
            raise RuntimeError(f"Response error, status code: [{response.status_code}]")

        return response


class _BaseCrawler:
    """
    針對 CMoney 和 財報狗爬蟲的 interface
    """

    def _get_group_data(self) -> list:
        """
        取得產業類別的資料

        回傳格式:
        ```
        [
            {"name" : "IC-設計", "url" : "http...."},...
        ]
        ```
        """

        raise NotImplementedError

    def _get_increase_reduce_group_data(self, day_type_arg: str = "1day") -> dict:
        """
        取得增加、減少的產業類別的資料

        Args:

        day_type_arg (str): 指定天數參數 (1day, 1week, 1month, 3months)

        回傳格式:
        ```
        {
        "increase" : [{"name" : "砷化鎵", "url" : "http...."}, ...],
        "reduce" : [{"name" : "家電", "url" : "http...."}, ...]
        }
        ```
        """

        raise NotImplementedError

    def _get_top_3_stock_of_group_data(self, url: str, group_name: str) -> dict:
        """
        取得產業類別的前三名股票資料

        Args:

        url (str): 產業類別的股票詳細頁面 url

        group_name (str): 產業類別名稱

        回傳格式:
        ```
        {
        "group" : "砷化鎵",
        "data" : [["3105", "穩懋"], ["2455", "全新"], ["8086", "宏捷科"]]
        }
        ```
        """

        raise NotImplementedError

    def _get_data(self, day_type_arg: str = "1day") -> dict:
        """
        取得增加、減少的產業類別資料，和股票的資料

        Args:

        day_type_arg (str): 指定天數參數 (1day, 1week, 1month, 3months)

        回傳格式:
        ```
        {
        "increase" : [
            {
                "group" : "砷化鎵",
                "data" : [["3105" : "穩懋"], ...]
            }, ...],
        "reduce" : [
            {
                "group" : "LCD塑膠框",
                "data" : [["2371" : "大同"]. ...]
            }, ...]
        }
        ```
        """

        raise NotImplementedError

    def get_data(
        self, Listed_price_data: dict, OTC_price_data: dict, day_type_arg: str = "1day"
    ) -> dict:
        """
        取得最終處裡完的資料。
        換句話說，就是取得增加、減少的產業資料和股票名稱、代碼，和開高低收資料

        Args:

        Listed_price_data (dict): 上市公司股票交易資料

        OTC_price_data (dict): 上櫃公司股票交易資料

        data_type_arg (str): 指定天數參數 (1day, 1week, 1month, 3months)

        回傳格式:
        ```
        {
        "increase" : [
            {"group" : "砷化鎵", "data" : [
                {"code" : "3105", "name" : "穩懋", "opening_price" : 101.1, "highest_price" : 120.0, "lowest_price" : 100.0, "cloesing_price" : 102.2}, ...]
            }, ...],
        "reduce: [....]
        }
        ```
        """

        result = defaultdict(list)
        meta_data = self._get_data(day_type_arg)

        for k in meta_data:
            for group_data in meta_data[k]:
                tmp = {}
                tmp["group"] = group_data["group"]
                tmp["data"] = []

                for stock in group_data["data"]:
                    if Listed_price_data.get(stock[0]):
                        tmp["data"].append(Listed_price_data[stock[0]])

                    elif OTC_price_data.get(stock[0]):
                        tmp["data"].append(OTC_price_data[stock[0]])

                    else:
                        # 如果找不到股票資料
                        tmp["data"].append(
                            {
                                "code": stock[0],
                                "name": stock[1],
                                "opening_price": None,
                                "highest_price": None,
                                "lowest_price": None,
                                "cloesing_price": None,
                            }
                        )

                result[k].append(tmp)

        return result


class StatementDogCrawler(_BaseCrawler):
    """
    財報狗爬蟲

    使用 API 和解析 HTML 獲得資料
    """

    def _get_increase_reduce_group_data(self, day_type_arg: str = "1day") -> dict:
        response = BaseRequset.get_requset(
            f"https://statementdog.com/api/v1/market-trend/tw/{day_type_arg}"
        )

        datas = response.json()["data"]

        if not datas:
            raise RuntimeError(f"Error: 取得財報狗 [{day_type_arg}] 的增加減少產業族群資料時發生錯誤.")

        sorted_datas = sorted(datas, key=lambda i: i["diff_percentage"], reverse=True)

        get_group_data = lambda s: [{"name": d["name"], "url": d["url"]} for d in sorted_datas[s]]

        increase = get_group_data(slice(0, 5))

        reduce = get_group_data(slice(-1, -6, -1))

        result = {}

        result["increase"] = increase
        result["reduce"] = reduce

        return result

    def _get_top_3_stock_of_group_data(self, url: str, group_name: str) -> dict:
        response = BaseRequset.get_requset(f"{url}?country=tw")

        soup = BeautifulSoup(response.text, "html.parser")

        tbody = soup.find("tbody", id="stock-tags-list-body")

        items = tbody.find_all("td", class_="stock-tags-list-item ticker-name")

        stock_list = []

        for item in items:
            code_name = item.text.replace("\n", "").split(" ")

            # 針對像 [1111, "iphone", "12"] 這樣的資料進行處理
            if len(code_name) > 2:
                code_name = [code_name[0], "".join(code_name[1:])]

            stock_list.append(code_name)

        stock_list = stock_list[:3]

        result = {}

        result["group"] = group_name
        result["data"] = stock_list

        return result

    def _get_data(self, day_type_arg: str = "1day") -> dict:
        result = defaultdict(list)

        group_stock_data = self._get_increase_reduce_group_data(day_type_arg)

        for k in group_stock_data:
            for group in group_stock_data[k]:
                increase_reduce = self._get_top_3_stock_of_group_data(group["url"], group["name"])

                result[k].append(increase_reduce)

        return result


class CMoneyCrawler(_BaseCrawler):
    """

    CMoney 爬蟲

    使用 selenium
    """

    def __init__(self, is_headless: bool = True):
        options = None

        if is_headless:
            options = webdriver.ChromeOptions()
            options.add_argument("--headless=new")

        self.driver = webdriver.Chrome(options=options)

    def _get_group_data(self) -> list:
        result = []

        items = self.driver.find_element(By.ID, "MainContent").find_elements(By.TAG_NAME, "tr")

        # 取前 10 個
        for item in items[1:11]:
            url = item.find_element(By.TAG_NAME, "a").get_attribute("href")

            group_meta_data = item.text.split(" ")

            # 防止名稱出現像 "概念股_iPhone 12" 這種情況
            # 這邊的 7 這個數字是依據 https://www.cmoney.tw/finance/f00018.aspx?o=1&o2=1 這個網頁上的表格欄位數量
            if len(group_meta_data) > 7:
                name = "".join(group_meta_data[0 : len(group_meta_data) - 7 + 1])

            else:
                name = group_meta_data[0]

            tmp = {"name": name, "url": url}

            result.append(tmp)

        return result

    def _get_increase_reduce_group_data(self, day_type_arg: str = "1day") -> dict:
        if day_type_arg == "1day":
            url_arg = 1
        elif day_type_arg == "1week":
            url_arg = 2
        elif day_type_arg == "1month":
            url_arg = 3
        elif day_type_arg == "3months":
            url_arg = 4
        else:
            raise ValueError("Invalid value for 'day_arg'")

        # 增加和減少的 url
        top_url = f"https://www.cmoney.tw/finance/f00018.aspx?o=1&o2={url_arg}"
        last_url = f"https://www.cmoney.tw/finance/f00018.aspx?o=2&o2={url_arg}"

        result = {}

        self.driver.get(top_url)
        locator = (By.CLASS_NAME, "up")
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_all_elements_located(locator),
            f"Error: 取得增加頁面的資料時出現錯誤, day_type_arg: [{day_type_arg}]",
        )
        top_group_data = self._get_group_data()

        self.driver.get(last_url)
        locator = (By.CLASS_NAME, "down")
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_all_elements_located(locator),
            f"Error: 取得減少頁面的資料時出現錯誤, day_type_arg: [{day_type_arg}]",
        )
        last_group_data = self._get_group_data()

        result["top"] = top_group_data
        result["last"] = last_group_data

        return result

    def _get_top_3_stock_of_group_data(self, url: str, group_name: str) -> dict:
        result = {}

        self.driver.get(url)
        locator = (By.CLASS_NAME, "bk-clr")
        WebDriverWait(self.driver, 10).until(
            EC.presence_of_all_elements_located(locator), f"Error: 取得前三名股票頁面的資料時出現錯誤, url: {url}"
        )

        stock_table = self.driver.find_element(By.ID, "table1").find_elements(By.TAG_NAME, "tr")

        result["group"] = group_name
        result["data"] = []

        for i in stock_table[1:4]:
            data = i.text.split(" ")
            code = data[1]
            name = data[2]

            result["data"].append([code, name])

        return result

    def _get_data(self, day_type_arg: str = "1day") -> dict:
        result = defaultdict(list)

        data = self._get_increase_reduce_group_data(day_type_arg)

        for i in data:
            for d in data[i]:
                stock_data = self._get_top_3_stock_of_group_data(d["url"], d["name"])

                result[i].append(stock_data)

        return result

    def close_driver(self):
        """關閉瀏覽器 driver"""

        self.driver.close()


class StockPrice:
    """
    取得股票的每日交易價格相關的類別
    """

    TRADING_DATE = None

    @staticmethod
    def _translate_stock_data(stock_data: dict) -> dict:
        """
        處理由證交所 API 取得的每日交易資訊，只留下必要的資料
        (代碼、名稱、開高低收)

        NOTE: 會給 `StockPrice.TRADING_DATE` 複寫成 API 回傳的資料日期
        """

        result = {}

        # 這邊的處理要看 https://www.twse.com.tw/exchangeReport/STOCK_DAY_ALL 這隻 API 的回傳格式
        for data in stock_data["data"]:
            tmp = {
                "code": data[0],
                "name": data[1],
                "opening_price": float(data[4].replace(",", "")) if data[4] else None,
                "highest_price": float(data[5].replace(",", "")) if data[5] else None,
                "lowest_price": float(data[6].replace(",", "")) if data[6] else None,
                "cloesing_price": float(data[7].replace(",", "")) if data[7] else None,
            }

            result[data[0]] = tmp

        StockPrice.TRADING_DATE = stock_data["date"]

        return result

    @staticmethod
    def _translate_mainborad_data(mainborad_data: dict) -> dict:
        """
        處理由櫃買中心 API 取得的每日交易資訊，只留下必要的資料
        (代碼、名稱、開高低收)
        """

        result = {}

        for data in mainborad_data:
            tmp = {
                "code": data["SecuritiesCompanyCode"],
                "name": data["CompanyName"],
                "opening_price": float(data["Open"]) if data["Open"] != "----" else None,
                "highest_price": float(data["High"]) if data["High"] != "----" else None,
                "lowest_price": float(data["Low"]) if data["Low"] != "----" else None,
                "cloesing_price": float(data["Close"]) if data["Close"] != "----" else None,
            }

            result[data["SecuritiesCompanyCode"]] = tmp

        return result

    @staticmethod
    def get_stock_day_all() -> dict:
        """
        根據證交所的 API 取得上市股票的每日交易資料

        i.e:
        ```
        {
        "3105" : {
        "code" : "3105",
        "name" : "穩懋",
        "opening_price" : 141.0,
        "highest_price" : 146.0,
        "lowest_price" : 139.0,
        "cloesing_price" : 144.0}, ...
        }
        ```
        """

        response = BaseRequset.get_requset("https://www.twse.com.tw/exchangeReport/STOCK_DAY_ALL")

        return StockPrice._translate_stock_data(response.json())

    @staticmethod
    def get_mainborad_day_all() -> dict:
        """
        根據櫃買中心的 API 取得上櫃股票的每日交易資訊

        i.e:
        ```
        {
        "3105" : {
        "code" : "3105",
        "name" : "穩懋",
        "opening_price" : 141.0,
        "highest_price" : 146.0,
        "lowest_price" : 139.0,
        "cloesing_price" : 144.0}, ...
        }
        ```
        """

        response = BaseRequset.get_requset(
            "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_quotes"
        )

        return StockPrice._translate_mainborad_data(response.json())


class ExcelWriter:
    """
    負責處理將資料寫入 excel 的類別
    """

    def __init__(self, base_filename: str, save_filename: str):
        """

        base_filename (str): 基本 excel 模板檔案路徑
        save_filename (str): 要儲存的 excel 檔案路徑
        """

        self.wb = openpyxl.load_workbook(base_filename)
        self.save_name = save_filename

    def _write_data(
        self,
        data: dict,
        group_number: int,
        worksheet_name: str,
        group_col_code: str,
        data_col_start: str,
        data_col_end: str,
    ):
        """將資料寫入至 excel 中

        Args:
            data (dict): 要寫入的資料
            group_number (int): 族群的數量
            worksheet_name (str): 要寫入至哪個分頁的名稱
            group_col_code (str): 族群名稱的欄位代號 i.e: "A"
            data_col_start (str): 資料範圍的開始欄位代號 i.e: "C"
            data_col_end (str): 資料範圍的結束欄位代號 i.e: "H"

        data 的格式要符合：
        ```
        [
            {
                "group" : "砷化鎵",
                "data" : [
                    {
                        "code" : "3105",
                        "name" : "穩懋",
                        "opening_price" : 141.0,
                        "highest_price" : 146.0,
                        "lowest_price" : 139.0,
                        "cloesing_price" : 144.0
                    }, ...]
            }, ....
        ]
        ```
        """

        s1 = self.wb[worksheet_name]

        # 控制族群名稱
        for i in range(group_number):
            # 寫入族群名稱
            s1[f"{group_col_code}{6 + (i*3)}"].value = data[i]["group"]

            if not data[i]["data"]:
                continue

            # 控制每個族群中的三筆資料
            for j in range(3):
                n = i * 3 + j

                cell_range = s1[f"{data_col_start}{6 + n}":f"{data_col_end}{6 + n}"]

                stcok_data = data[i]["data"][j]

                # 寫入各項資料
                cell_range[0][0].value = stcok_data["code"] if stcok_data["code"] else "null"
                cell_range[0][1].value = stcok_data["name"] if stcok_data["name"] else "null"
                cell_range[0][2].value = (
                    stcok_data["opening_price"] if stcok_data["opening_price"] else "null"
                )
                cell_range[0][3].value = (
                    stcok_data["highest_price"] if stcok_data["highest_price"] else "null"
                )
                cell_range[0][4].value = (
                    stcok_data["lowest_price"] if stcok_data["lowest_price"] else "null"
                )
                cell_range[0][5].value = (
                    stcok_data["cloesing_price"] if stcok_data["cloesing_price"] else "null"
                )

        self.wb.save(self.save_name)

    def write_statement_dog_data(self, statement_dog_data: dict, day_type: str):
        """將財報狗的資料寫入 execl

        statement_dog_data (dict): 財報狗的資料，由 ``StatementDog.get_final_data()`` 生成的資料
        day_type (str): 資料的種類 ("1day", "1week", "1month", "3months")
        """

        if day_type == "1day":
            raise_group_col_code = "A"
            raise_data_col_start_code = "C"
            raise_data_col_end_code = "H"

            reduce_group_col_code = "BA"
            reduce_data_col_start_code = "BC"
            recude_data_col_end_code = "BH"

        elif day_type == "1week":
            raise_group_col_code = "N"
            raise_data_col_start_code = "P"
            raise_data_col_end_code = "U"

            reduce_group_col_code = "BN"
            reduce_data_col_start_code = "BP"
            recude_data_col_end_code = "BU"

        elif day_type == "1month":
            raise_group_col_code = "AA"
            raise_data_col_start_code = "AC"
            raise_data_col_end_code = "AH"

            reduce_group_col_code = "CA"
            reduce_data_col_start_code = "CC"
            recude_data_col_end_code = "CH"

        elif day_type == "3months":
            raise_group_col_code = "AN"
            raise_data_col_start_code = "AP"
            raise_data_col_end_code = "AU"

            reduce_group_col_code = "CN"
            reduce_data_col_start_code = "CP"
            recude_data_col_end_code = "CU"

        else:
            raise ValueError("Invalid value for 'day_type'")

        self._write_data(
            statement_dog_data["increase"],
            5,
            "漲跌幅-前五族群前三檔",
            raise_group_col_code,
            raise_data_col_start_code,
            raise_data_col_end_code,
        )
        self._write_data(
            statement_dog_data["reduce"],
            5,
            "漲跌幅-前五族群前三檔",
            reduce_group_col_code,
            reduce_data_col_start_code,
            recude_data_col_end_code,
        )

    def write_cmoney_data(self, cmoney_data: dict, day_type: str = "1day"):
        if day_type == "1day":
            raise_group_col_code = "A"
            raise_data_col_start_code = "C"
            raise_data_col_end_code = "H"

            reduce_group_col_code = "BA"
            reduce_data_col_start_code = "BC"
            recude_data_col_end_code = "BH"

        elif day_type == "1week":
            raise_group_col_code = "N"
            raise_data_col_start_code = "P"
            raise_data_col_end_code = "U"

            reduce_group_col_code = "BN"
            reduce_data_col_start_code = "BP"
            recude_data_col_end_code = "BU"

        elif day_type == "1month":
            raise_group_col_code = "AA"
            raise_data_col_start_code = "AC"
            raise_data_col_end_code = "AH"

            reduce_group_col_code = "CA"
            reduce_data_col_start_code = "CC"
            recude_data_col_end_code = "CH"

        elif day_type == "3months":
            raise_group_col_code = "AN"
            raise_data_col_start_code = "AP"
            raise_data_col_end_code = "AU"

            reduce_group_col_code = "CN"
            reduce_data_col_start_code = "CP"
            recude_data_col_end_code = "CU"

        else:
            raise ValueError("Invalid value for 'day_type'")

        self._write_data(
            cmoney_data["top"],
            10,
            "資金流向-前十族群前三檔",
            raise_group_col_code,
            raise_data_col_start_code,
            raise_data_col_end_code,
        )
        self._write_data(
            cmoney_data["last"],
            10,
            "資金流向-前十族群前三檔",
            reduce_group_col_code,
            reduce_data_col_start_code,
            recude_data_col_end_code,
        )

    def write_date(self, data_date: str, trading_date: str):
        """寫入日期資料

        Args:
            data_date (str): 資料日期
            trading_date (str): 交易日期
        """

        data_date_col = ["A4", "N4", "AA4", "AN4", "BA4", "BN4", "CA4", "CN4"]
        tading_date_col = ["B4", "O4", "AB4", "AO4", "BB4", "BO4", "CB4", "CO4"]

        s1 = self.wb["漲跌幅-前五族群前三檔"]
        s2 = self.wb["資金流向-前十族群前三檔"]

        for col in data_date_col:
            s1[col].value = data_date
            s2[col].value = data_date

        for col in tading_date_col:
            s1[col].value = trading_date
            s2[col].value = trading_date

        self.wb.save(self.save_name)


class ExeclUpdater:
    """
    更新前一天的 execl 檔案的股票資訊
    """

    def __init__(self, filename: str):
        """

        filename (str): 要更新的 excel 檔案路徑
        """

        self.wb = openpyxl.load_workbook(filename)
        self.save_name = filename

    def _write_data(
        self,
        worksheet_name: str,
        data_number: int,
        stock_price_all_day: dict,
        mainborad_price_all_day: dict,
    ):
        """將資料寫入至指定的工作區

        Args:
            worksheet_name (str): 工作區名稱
            data_number (int): 資料的總數
            stock_price_all_day (dict): 上市股票的交易資料
            mainborad_price_all_day (dict): 上櫃股票的交易資料
        """

        s1 = self.wb[worksheet_name]

        stock_code_col = ["C", "P", "AC", "AP", "BC", "BP", "CC", "CP"]
        stock_data_col = [
            ["I", "L"],
            ["V", "Y"],
            ["AI", "AL"],
            ["AV", "AY"],
            ["BI", "BL"],
            ["BV", "BY"],
            ["CI", "CL"],
            ["CV", "CY"],
        ]

        for i, col in enumerate(stock_code_col):
            for j in range(data_number):
                stock_code = s1[f"{col}{6 + j}"].value

                if stock_code:
                    if stock_price_all_day.get(str(stock_code)):
                        data = stokc_price_all_day[str(stock_code)]

                    elif mainborad_price_all_day.get(str(stock_code)):
                        data = mainborad_price_all_day[str(stock_code)]

                    else:
                        data = {
                            "opening_price": None,
                            "highest_price": None,
                            "lowest_price": None,
                            "cloesing_price": None,
                        }

                else:
                    continue

                data_cell_range = s1[
                    f"{stock_data_col[i][0]}{6 + j}":f"{stock_data_col[i][1]}{6 + j}"
                ]

                data_cell_range[0][0].value = (
                    data["opening_price"] if data["opening_price"] else "null"
                )
                data_cell_range[0][1].value = (
                    data["highest_price"] if data["highest_price"] else "null"
                )
                data_cell_range[0][2].value = (
                    data["lowest_price"] if data["lowest_price"] else "null"
                )
                data_cell_range[0][3].value = (
                    data["cloesing_price"] if data["cloesing_price"] else "null"
                )

        self.wb.save(self.save_name)

    def update_file(self, stock_price_all_day: dict, mainborad_price_all_day: dict):
        """更新股票資料

        Args:
            stock_price_all_day (dict): 上市股票的交易資料
            mainborad_price_all_day (dict): 上櫃股票的交易資料
        """

        self._write_data("漲跌幅-前五族群前三檔", 15, stock_price_all_day, mainborad_price_all_day)
        self._write_data("資金流向-前十族群前三檔", 30, stock_price_all_day, mainborad_price_all_day)


if __name__ == "__main__":
    # 取得今天日期
    today_date = datetime.now().strftime("%Y-%m-%d")

    base_dir = os.path.abspath(os.path.dirname(__name__))

    statement_dog_crawler = StatementDogCrawler()
    stokc_price = StockPrice()
    excel = ExcelWriter(
        os.path.join(base_dir, "base.xlsx"), os.path.join(base_dir, "data", f"{today_date}.xlsx")
    )

    print("取得股票交易資料....")
    stokc_price_all_day = stokc_price.get_stock_day_all()
    mainborad_price_all_day = stokc_price.get_mainborad_day_all()
    print("處理完成")

    excel.write_date(
        today_date,
        datetime.strftime(datetime.strptime(stokc_price.TRADING_DATE, "%Y%m%d"), "%Y-%m-%d"),
    )

    day_args_list = ["1day", "1week", "1month", "3months"]

    print(f"{'-' * 5} 爬取財報狗資料 {'-' * 5}")

    for day_arg in day_args_list:
        print(f"取得 [{day_arg}] 資料...")

        statement_dog_data = statement_dog_crawler.get_data(
            stokc_price_all_day, mainborad_price_all_day, day_arg
        )

        print("爬取完成")

        print("寫入 excel...")

        excel.write_statement_dog_data(statement_dog_data, day_arg)

        print("寫入完成")

    print(f"{'-' * 5} 財報狗資料處理完畢 {'-' * 5}")

    print(f"{'-' * 5} 爬取 CMoney 資料 {'-' * 5}")

    cmoney_crawler = CMoneyCrawler(is_headless=False)

    for day_arg in day_args_list:
        print(f"取得 [{day_arg}] 資料...")

        cmoney_data = cmoney_crawler.get_data(stokc_price_all_day, mainborad_price_all_day, day_arg)

        print("爬取完成")

        print("寫入 excel...")

        excel.write_cmoney_data(cmoney_data, day_arg)

        print("寫入完成")

    print(f"{'-' * 5} CMoney 資料處理完畢 {'-' * 5}")

    excel.wb.close()

    # 更新資料
    today_weekday = datetime.today().weekday()

    # 星期一
    if today_weekday == 0:
        subtract_number = 3

    else:
        subtract_number = 1

    pre_date = datetime.now() - timedelta(days=subtract_number)
    pre_date_str = pre_date.strftime("%Y-%m-%d")

    pre_filename = os.path.join(base_dir, "data", f"{pre_date_str}.xlsx")

    if os.path.exists(pre_filename):
        print(f"{'-' * 5} 更新前一天的資料 {'-' * 5}")

        print(f"更新 [{pre_filename}]")
        excel_updater = ExeclUpdater(pre_filename)
        excel_updater.update_file(stokc_price_all_day, mainborad_price_all_day)
        print("更新完成")

    else:
        print(f"找不到 {pre_filename} 因此跳過更新")

    print("程式執行結束")
