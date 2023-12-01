"""
This spider crawls ECCV paper abstracts to a json file.
"""

import scrapy


class PaperSpider(scrapy.Spider):
    name = "eccv_paper_spider"

    def __init__(self, year=None, *args, **kwargs):
        super(PaperSpider, self).__init__(*args, **kwargs)
        self.start_urls = [f"https://www.ecva.net/papers.php"]
        self.year = year

    def parse(self, response):  # noqa
        paper_links = None
        if self.year == "2022":
            paper_links = response.xpath("/html/body/main/div[2]/div[1]/div/dl/dt/a/@href").getall()
        elif self.year == "2020":
            paper_links = response.xpath("/html/body/main/div[2]/div[2]/div/dl/dt/a/@href").getall()

        yield from response.follow_all(paper_links, self.parse_paper)

    def parse_paper(self, response):  # noqa
        def extract_with_xpath(query):
            return response.xpath(query).get().strip()

        yield {
            "title": extract_with_xpath("//*[@id='papertitle']/text()"),
            "authors": extract_with_xpath("//*[@id='authors']/b/i/text()"),
            "abstract": extract_with_xpath("//*[@id='abstract']/text()"),
        }
