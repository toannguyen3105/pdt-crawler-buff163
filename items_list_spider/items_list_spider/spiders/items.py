import csv
import os.path
import scrapy
import glob

from openpyxl import Workbook
from scrapy.http import Request
from decouple import config

COOKIES = config('COOKIE')
CSGO_URL = "https://buff.163.com/api/market/goods?game=csgo&page_num=1&min_price=50&max_price=30000&sort_by=price.desc"
DOTA_URL = "https://buff.163.com/api/market/goods?game=dota2&page_num=1&sort_by=price.desc"
CSGO_STR = "csgo"
DOTA_STR = "dota2"
# Set the headers here.
HEADERS = {
    'Cookie': COOKIES
}


def generate_url(gameId, page):
    if gameId == 730:
        return 'goods?game={}&page_num={}&min_price=50&max_price=30000&sort_by=price.desc'.format(CSGO_STR, page + 1)
    else:
        return 'goods?game={}&page_num={}&sort_by=price.desc'.format(DOTA_STR, page + 1)


class ItemsSpider(scrapy.Spider):
    name = 'items'
    allowed_domains = ['buff.163.com']

    def start_requests(self):
        url = CSGO_URL if int(self.gameId) == 730 else DOTA_URL
        yield scrapy.Request(url=url, callback=self.parse, headers=HEADERS)

    def parse(self, response):
        data = response.json()['data']
        items = data['items']
        for item in items:
            yield {
                'name': item['name'],
                'price': item['sell_min_price']
            }

        current_page = data['page_num']
        last_page = data['total_page']
        if current_page <= last_page:
            yield Request(
                response.urljoin(generate_url(int(self.gameId), current_page)),
                method="GET",
                headers=HEADERS,
                callback=self.parse
            )

    def close(self, reason):
        csv_file = max(glob.iglob('*csv'), key=os.path.getctime)

        wb = Workbook()
        ws = wb.active

        with open(csv_file, 'r') as f:
            for row in csv.reader(f):
                ws.append(row)

        wb.save(csv_file.replace('.csv', '') + '.xlsx')
