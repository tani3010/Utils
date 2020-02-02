# -*- coding,  utf-8 -*-

from urllib.parse import urlencode, urlparse, quote_plus, unquote_plus
from urllib.parse import quote
from urllib.request import Request, urlopen
from urllib.request import URLError, HTTPError
from lxml import html

class BaseWebUtil:
    def __init__(self, encoding='utf-8'):
        self.encoding = encoding
        self.url_guard_string = ':/'

    def __del__(self):
        pass

    @staticmethod
    def parse_url(url):
        return urlparse(url)

    def encode_url(self, url):
        return quote(url, encoding=self.encoding, safe=self.url_guard_string)

    @staticmethod
    def encode_string(target_string):
        return quote_plus(target_string)

    @staticmethod
    def decode_string(target_string):
        return unquote_plus(target_string)

    def is_exist_url(self, url):
        is_exist = False
        encoded_url = url
        try:
            web_page = urlopen(encoded_url)
            web_page.close()
            is_exist = True
        except HTTPError or URLError:
            print('Not found:', self.decode_string(url))
        finally:
            return is_exist

    def get_html(self, url):
        encoded_url = self.encode_url(url)
        if not self.is_exist_url(encoded_url):
            print('Not found:', self.decode_string(url))
            return

        dat_raw = urlopen(encoded_url)
        dat_html = dat_raw.read()
        dat_html_obj = html.fromstring(dat_html)
        return dat_html_obj

    def write_string_to_file(self, target_string, file_path):
        with open(file_path, mode='w',
                  encoding=self.encoding) as file_ptr:
            print(target_string, file=file_ptr)

class WikiUtil(BaseWebUtil):
    def __init__(self, encoding='utf-8'):
        super().__init__(encoding)
        self.base_url_wiki \
            = 'https://ja.wikipedia.org/wiki/'
        self.base_url_mediawiki_api \
            = 'https://ja.wikipedia.org/w/api.php?'

    def __del__(self):
        super().__del__()

    @staticmethod
    def set_params(page_title, output_format='xml'):
        params = {
            'action': 'query',
            'prop': 'revisions',
            'rvprop': 'content',
            'titles': page_title,
            'formatversion': 2,
            'format': output_format
        }
        return params

    def build_api_url(self, params):
        return self.base_url_mediawiki_api + self.encode_url(params)

    def build_url(self, page_title):
        return self.base_url_wiki + \
            quote_plus(page_title, encoding=self.encoding)

    def get_xml_by_api(self, page_title, output_format='xml'):
        try:
            params = self.set_params(page_title, output_format)
            url_wiki = self.build_url(page_title)
            if not self.is_exist_url(url_wiki):
                return None
            url_api = self.build_api_url(params)
            with urlopen(url_api) as res:
                res_json = res.read()
                return str(res_json.decode(encoding=self.encoding))
        except HTTPError as e:
            print('HTTPError: {}'.format(e.reason))
        except URLError as e:
            print('URLError: {}'.format(e.reason))
