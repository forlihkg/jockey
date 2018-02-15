# -*- coding: utf-8 -*-
import scrapy
import xlwt

class HorseSpider(scrapy.Spider):
    name = 'horse'
    allowed_domains = ['www.hkjc.com']
    start_urls = ['http://www.hkjc.com/chinese/racing/SelectHorse.asp']

    def __init__(self):
        try:
            self.book = xlwt.Workbook()
            self.sheet = self.book.add_sheet('horse')
            self.row = 0
            self.sheet.write(self.row, 0, '馬匹烙號'.decode('utf-8'))
            self.sheet.write(self.row, 1, '馬名'.decode('utf-8'))
            self.sheet.write(self.row, 2, '馬主'.decode('utf-8'))
        except:
            print 'cannot create excel'

    def parse(self, response):
        relative_urls = response.xpath('//a/@href').extract()
        for relative_url in relative_urls:
            if relative_url.startswith('ListByStable.asp?TrainerCode='):
                absolute_url = response.urljoin(relative_url)
                yield scrapy.Request(url=absolute_url, callback=self.parse_by_stable)

    def parse_by_stable(self, response):
        relative_urls = response.xpath('//a/@href').extract()
        for relative_url in relative_urls:
            if relative_url.startswith('Horse.asp?HorseNo='):
                absolute_url = response.urljoin(relative_url)
                yield scrapy.Request(url=absolute_url, callback=self.parse_by_horse)

    def parse_by_horse(self, response):
        horse_details = response.xpath('//td[@class="subsubheader"]/font/text()').extract()
        if len(horse_details) > 0:
            horse_detail = horse_details[0]
            horse_detail_array = horse_detail.split()

            self.row += 1
            if len(horse_detail_array) > 1:
                horse_name = horse_detail_array[0]
                horse_number = horse_detail_array[1]
                self.sheet.write(self.row, 1, horse_name)
            else:
                horse_number = horse_detail_array[0]
            self.sheet.write(self.row, 0, horse_number)

            horse_owner = unicode(response.xpath('//a[@onclick="goOwnerSearch();"]/text()').extract_first())
            self.sheet.write(self.row, 2, horse_owner)
            self.book.save('horse.xls')
