# -*- coding,  utf-8 -*-

import QuantLib as ql
import pandas as pd
from imm_date import get_next_imm_date, get_next_imm_code
from web_util import BaseWebUtil

def test_imm_date_and_code():
    base_date = ql.Date.todaysDate()
    for iter in range(12):
        print(
            get_next_imm_date(base_date, False).ISO(),
            get_next_imm_code(base_date, False))
        base_date = get_next_imm_date(base_date)

def test_url_util():
    web = BaseWebUtil()
    url = 'https://ja.wikipedia.org/wiki/日本の企業一覧'
    dat = web.get_html(url)

if __name__  ==  '__main__':
    # test_imm_date_and_code()
    web = BaseWebUtil()
    base_url = 'https://ja.wikipedia.org'
    url = 'https://ja.wikipedia.org/wiki/日本の企業一覧'
    dat = web.get_html(url).xpath('//td/a')
    category_list = [iter for iter in dat if '企業の一覧' in iter.text]
    category_urls = [base_url + iter.values()[0] for iter in category_list]

    comp = []
    comp_info = []
    for cate in category_urls:
        dat = web.get_html(web.decode_string(cate))
        # dat = web.get_html(web.decode_string(category_urls[4]))
        tmp = dat.xpath('/html/body/div/div/div/div/ul/li/a')
        for iter in tmp:
            if '企業一覧' in iter.text or '企業の一覧' in iter.text or '日本の' in iter.text:
                if cate != "https://ja.wikipedia.org/wiki/%E6%97%A5%E6%9C%AC%E3%81%AE%E4%BC%81%E6%A5%AD%E4%B8%80%E8%A6%A7_(%E3%82%B5%E3%83%BC%E3%83%93%E3%82%B9)":
                    break
                elif iter.text == '日本企業の一覧':
                    break
            if '#ア行' not in iter.text or 'バス事業者' not in iter.text:
                comp.append(iter.text)
                comp_info.append(iter)

    for iter in comp_info[0:50]:
        if iter.attrib.has_key('href') and iter.attrib.has_key('title'):
            title = iter.attrib['title']
            if '存在しない' not in title:
                param = web.decode_string(iter.attrib['href']).replace('/wiki/', '')
                print(param)

    # comp = pd.DataFrame(comp)
    # comp.to_csv('./test.csv')
