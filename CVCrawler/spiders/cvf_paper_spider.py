"""
This spider crawls CVF paper abstracts to a json file.
"""

import scrapy


class PaperSpider(scrapy.Spider):
    name = "cvf_paper_spider"

    def __init__(self, conference=None, year=None, *args, **kwargs):
        super(PaperSpider, self).__init__(*args, **kwargs)
        self.start_urls = [f"https://openaccess.thecvf.com/{conference}{year}?day=all"]

    def parse(self, response):  # noqa
        paper_links = response.xpath("//*[@id='content']/dl/dt/a/@href").getall()
        yield from response.follow_all(paper_links, self.parse_paper)

    def parse_paper(self, response):  # noqa
        def extract_with_xpath(query):
            return response.xpath(query).get().strip()

        yield {
            "title": extract_with_xpath("//*[@id='papertitle']/text()"),
            "authors": extract_with_xpath("//*[@id='authors']/b/i/text()"),
            "abstract": extract_with_xpath("//*[@id='abstract']/text()"),
        }
