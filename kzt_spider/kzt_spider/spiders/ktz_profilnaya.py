import scrapy
import openpyxl

class KtzSpider(scrapy.Spider):
    name = 'ktz_spider'
    allowed_domains = ['www.ktzholding.com']
    start_urls = ['https://www.ktzholding.com/catalog/truba-profilnaya/?PAGEN_2=1']

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.excel = openpyxl.Workbook()
        self.sheet = self.excel.active
        self.sheet.append(['Продукт', 'Длина', 'Вес за погонный метр', 'Марка стали', 'Цена'])

    def parse(self, response, **kwargs):

        for links in response.css('table.main_table td a::attr(href)'):
            yield response.follow(links, callback=self.parse_product)

        for i in range(2, 22):
            next_page = f'https://www.ktzholding.com/catalog/truba-profilnaya/?PAGEN_2={i}'
            yield response.follow(next_page, callback=self.parse)



    def parse_product(self, response):

        product_name = response.css('h1.bx-title::text').get()
        product_price = response.css('dl.product-item-detail-properties dd::text')[7].get().split(' ')[92]
        product_model = response.css('dl.product-item-detail-properties dd::text')[4].get().split(' ')[92]
        product_weight = response.css('dl.product-item-detail-properties dd::text')[6].get().split(' ')[92]
        product_length = response.css('dl.product-item-detail-properties dd::text')[2].get().split(' ')[92]

        if product_name and product_price:
            self.sheet.append([product_name, product_length, product_weight, product_model, product_price])

    def closed(self, reason):
        self.excel.save('kzt_profilnaya.xlsx')