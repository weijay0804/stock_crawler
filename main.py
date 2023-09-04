from collections import defaultdict

from bs4 import BeautifulSoup
import requests


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
        """

        response = BaseRequset.get_requset(f"https://statementdog.com/api/v1/market-trend/tw/{arg}")

        datas = response.json()["data"]

        sorted_datas = sorted(datas, key=lambda i: i["diff_percentage"], reverse=True)

        top = sorted_datas[:3]
        last = sorted_datas[-1:-4:-1]

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
        """

        response = BaseRequset.get_requset(url)

        soup = BeautifulSoup(response.text, "html.parser")

        tbody = soup.find("tbody", id="stock-tags-list-body")

        items = tbody.find_all("td", class_="stock-tags-list-item ticker-name")

        stock_list = []

        for item in items:
            code_name = item.text.replace("\n", "").split(" ")

            try:
                code_name[0] = int(code_name[0])
                stock_list.append(code_name)

            except ValueError:
                continue

        stock_list = stock_list[:3]

        result = {}

        result[name] = stock_list

        return result

    @staticmethod
    def get_data(arg: str = "1day") -> dict:
        """
        取得市場焦點和每個產業的股票

        args: 要取得的天數 (1day, 1week, 1month, 3months)
        """

        result = defaultdict(list)

        market_foucs_data = StatementDog.get_market_focus(arg)

        for i in market_foucs_data:
            for data in market_foucs_data[i]:
                top_last_3 = StatementDog.get_group_top_last_3(data["url"], data["name"])

                result[i].append(top_last_3)

        return result
