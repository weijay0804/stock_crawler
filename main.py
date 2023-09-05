from collections import defaultdict

from bs4 import BeautifulSoup
import requests
import openpyxl


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


class StatementDog:
    """
    處理有關財報狗相關的類別
    """

    @staticmethod
    def get_market_focus(arg: str = "1day") -> dict:
        """
        取得市場焦點的前三名和後三名的產業類別

        args: 要取得的天數 (1day, 1week, 1month, 3months)

        i.e:
        ```
        {
        "top" : [{"name" : "砷化鎵", "diff_percentage" : 3.87, "url" : "http...."}, ...],
        "last" : [{"name" : "家電", "diff_percentage" : -8.23, "url" : "http...."}, ...]
        }
        ```
        """

        response = BaseRequset.get_requset(f"https://statementdog.com/api/v1/market-trend/tw/{arg}")

        datas = response.json()["data"]

        sorted_datas = sorted(datas, key=lambda i: i["diff_percentage"], reverse=True)

        top = sorted_datas[:5]
        last = sorted_datas[-1:-6:-1]

        result = {}

        result["top"] = top
        result["last"] = last

        return result

    @staticmethod
    def get_group_top_last_3(url: str, name: str) -> dict:
        """
        取得產業類別中的前三名股票名稱、代號

        url: 產業詳細頁面 url
        name: 產業名稱

        i.e:
        ```
        {
        "group" : "砷化鎵",
        "data" : [["3105", "穩懋"], ["2455", "全新"], ["8086", "宏捷科"]]
        }
        ```
        """

        response = BaseRequset.get_requset(f"{url}?country=tw")

        soup = BeautifulSoup(response.text, "html.parser")

        tbody = soup.find("tbody", id="stock-tags-list-body")

        items = tbody.find_all("td", class_="stock-tags-list-item ticker-name")

        stock_list = []

        for item in items:
            code_name = item.text.replace("\n", "").split(" ")

            if len(code_name) > 2:
                code_name = [code_name[0], "".join(code_name[1:])]

            stock_list.append(code_name)

        stock_list = stock_list[:3]

        result = {}

        result["group"] = name
        result["data"] = stock_list

        return result

    @staticmethod
    def get_data(arg: str = "1day") -> dict:
        """
        取得市場焦點和每個產業的股票

        args: 要取得的天數 (1day, 1week, 1month, 3months)

        i.e:
        ```
        {
        "top" : [
            {
                "group" : "砷化鎵",
                "data" : [["3105" : "穩懋"], ...]
            }, ...],
        "last" : [
            {
                "group" : "LCD塑膠框",
                "data" : [["2371" : "大同"]. ...]
            }, ...]
        }
        ```
        """

        result = defaultdict(list)

        market_foucs_data = StatementDog.get_market_focus(arg)

        for i in market_foucs_data:
            for data in market_foucs_data[i]:
                top_last_3 = StatementDog.get_group_top_last_3(data["url"], data["name"])

                result[i].append(top_last_3)

        return result

    @staticmethod
    def get_final_data(
        stock_price_data: dict, mainborad_price_data: dict, arg: str = "1day"
    ) -> dict:
        """
        取得最終處理完的資料，換句話說就是取得開高低收等資訊

        args: 要取得的天數 (1day, 1week, 1month, 3months)

        i.e:
        ```
        {
        "top" : [
            {"group" : "砷化鎵", "data" : [
                {"code" : "3105", "name" : "穩懋", "opening_price" : 101.1, "highest_price" : 120.0, "lowest_price" : 100.0, "cloesing_price" : 102.2}, ...]
            }, ...],
        "last: [....]
        }
        ```
        """

        result = defaultdict(list)
        meta_data = StatementDog.get_data(arg)

        for k in meta_data:
            for group_data in meta_data[k]:
                tmp = {}
                tmp["group"] = group_data["group"]
                tmp["data"] = []

                for stock in group_data["data"]:
                    if stock_price_data.get(stock[0]):
                        tmp["data"].append(stock_price_data[stock[0]])

                    elif mainborad_price_data.get(stock[0]):
                        tmp["data"].append(mainborad_price_data[stock[0]])

                    else:
                        tmp["data"].append(None)

                result[k].append(tmp)

        return result


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
            statement_dog_data["top"],
            5,
            "漲跌幅-前五族群前三檔",
            raise_group_col_code,
            raise_data_col_start_code,
            raise_data_col_end_code,
        )
        self._write_data(
            statement_dog_data["last"],
            5,
            "漲跌幅-前五族群前三檔",
            reduce_group_col_code,
            reduce_data_col_start_code,
            recude_data_col_end_code,
        )


if __name__ == "__main__":
    statement_dog = StatementDog()
    stokc_price = StockPrice()
    excel = ExcelWriter("./base.xlsx", "test4.xlsx")

    print("取得股票交易資料....")
    stokc_price_all_day = stokc_price.get_stock_day_all()
    mainborad_price_all_day = stokc_price.get_mainborad_day_all()
    print("處理完成")

    print(f"{'-' * 5} 爬取財報狗資料 {'-' * 5}")

    print("取得 [1day] 資料...")

    statement_dog_1day = statement_dog.get_final_data(
        stokc_price_all_day, mainborad_price_all_day, "1day"
    )

    print("爬取完成")

    print("寫入 excel...")

    excel.write_statement_dog_data(statement_dog_1day, "1day")

    print("寫入完成")

    print("取得 [1week] 資料...")
    statement_dog_1week = statement_dog.get_final_data(
        stokc_price_all_day, mainborad_price_all_day, "1week"
    )
    print("爬取完成")

    print("寫入 execl...")
    excel.write_statement_dog_data(statement_dog_1week, "1week")
    print("寫入完成")

    print("取得 [1month] 資料...")
    statement_dog_1month = statement_dog.get_final_data(
        stokc_price_all_day, mainborad_price_all_day, "1month"
    )
    print("爬取完成")

    print("寫入 excel...")
    excel.write_statement_dog_data(statement_dog_1month, "1month")
    print("寫入完成")

    print("取得 [3months] 資料...")
    statement_dog_3months = statement_dog.get_final_data(
        stokc_price_all_day, mainborad_price_all_day, "3months"
    )

    print("爬取完成")

    print("寫入 excel...")
    excel.write_statement_dog_data(statement_dog_3months, "3months")
    print("寫入完成")

    print(f"{'-' * 5} 財報狗資料處理完畢 {'-' * 5}")

    excel.wb.close()
    print("程式執行結束")
