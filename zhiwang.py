# 需求：知网空间文献信息获取
# 增加文献类型（paper_type）、文献来源（paper_source）、导师（mentor）、下载量（download_times）、被引量（refer_times）等字段
# 更新日期：2021/11/06
import requests
from lxml import etree
import os
import time
import xlsxwriter as xw
import random
import re

base_url = 'http://search.cnki.com.cn/Search/ListResult'
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:93.0) Gecko/20100101 Firefox/93.0',
    # 'Content-Type': 'application/x-www-form-urlencoded',
    # 'Content-Length': '463'
}


def get_page_text(url, headers, search_word, page_num):
    data = {
        'searchType': ' MulityTermsSearch',
        'ArticleType': '',
        'ReSearch': '',
        'ParamIsNullOrEmpty': ' false',
        'Islegal': ' false',
        'Content': search_word,
        'Theme': '',
        'Title': '',
        'KeyWd': '',
        'Author': '',
        'SearchFund': '',
        'Originate': '',
        'Summary': '',
        'PublishTimeBegin': '',
        'PublishTimeEnd': '',
        'MapNumber': '',
        'Name': '',
        'Issn': '',
        'Cn': '',
        'Unit': '',
        'Public': '',
        'Boss': '',
        'FirstBoss': '',
        'Catalog': '',
        'Reference': '',
        'Speciality': '',
        'Type': '',
        'Subject': '',
        'SpecialityCode': '',
        'UnitCode': '',
        'Year': '',
        'AuthorFilter': '',
        'BossCode': '',
        'Fund': '',
        'Level': '',
        'Elite': '',
        'Organization': '',
        'Order': ' 1',
        'Page': str(page_num),
        'PageIndex': '',
        'ExcludeField': '',
        'ZtCode': '',
        'Smarts': '',
    }

    response = requests.post(url=url, headers=headers, data=data)
    page_text = response.text
    return page_text


def get_abstract(url):
    response = requests.get(url=url, headers=headers)
    page_text = response.text
    tree = etree.HTML(page_text)
    abstract = tree.xpath('//div[@class="xx_font"]//text()')

    return abstract


def list_to_str(my_list):
    my_str = "".join(my_list)
    return my_str


def parse_page_text(page_text):
    tree = etree.HTML(page_text)
    item_list = tree.xpath('//div[@class="list-item"]')
    page_info = []
    for item in item_list:
        # 标题
        title = list_to_str(item.xpath(
            './p[@class="tit clearfix"]/a[@class="left"]/@title'))
        # 链接
        link = 'https:' +\
            list_to_str(item.xpath(
                './p[@class="tit clearfix"]/a[@class="left"]/@href'))
        # 作者
        author = list_to_str(item.xpath(
            './p[@class="source"]/span[1]/@title'))
        # 导师
        mentor = list_to_str(item.xpath(
            './p[@class="source"]/span[2]/a[1]/text()'))
        # 出版日期
        date = list_to_str(item.xpath(
            './p[@class="source"]/span[last()-1]/text() | ./p[@class="source"]/a[2]/span[1]/text() ')).strip()
        date = date[0:4]
        # 关键词
        keywords = list_to_str(item.xpath(
            './div[@class="info"]/p[@class="info_left left"]/a[1]/@data-key'))
        keys = keywords.split("/")
        keywords = keys[random.randint(0,len(keys)-1)]
        # 摘要
        abstract = list_to_str(get_abstract(url=link))
        # 文献来源
        paper_source = list_to_str(item.xpath(
            './p[@class="source"]/span[last()-2]/text() | ./p[@class="source"]/a[1]/span[1]/text() '))
        # 文献类型
        paper_type = list_to_str(item.xpath(
            './p[@class="source"]/span[last()]/text()'))
        # 下载量
        download_times = list_to_str(item.xpath(
            './div[@class="info"]/p[@class="info_right right"]/span[@class="time1"]/text()'))
        download_times = re.search('\\d+', download_times).group()
        # 被引量
        refer_times = list_to_str(item.xpath(
            './div[@class="info"]/p[@class="info_right right"]/span[@class="time2"]/text()'))

        refer_times = re.search('\\d+', refer_times).group()

        item_info = [i.strip() for i in [title, author, mentor,
                     paper_source, paper_type, date, keywords, abstract, download_times, refer_times, link]]
        page_info.append(item_info)
        # print(page_info)
    return page_info


def write_to_excel(workbook, info,  search_word):

    wb = workbook
    worksheet1 = wb.add_worksheet(search_word)  # 创建子表
    worksheet1.activate()  # 激活表

    title = ['title', 'author', 'mentor',
             'paper_source', 'paper_type', 'date', 'keywords', 'abstract', 'download_times', 'refer_times', 'link', 'type']  # 设置表头
    worksheet1.write_row('A1', title)  # 从A1单元格开始写入表头

    i = 2  # 从第二行开始写入数据
    for j in range(len(info)):
        insert_data = info[j]
        insert_data.append(search_word)
        start_pos = 'A' + str(i)
        # print(insert_data)
        worksheet1.write_row(start_pos, insert_data)
        i += 1
    return True


if __name__ == '__main__':

    # 1、创建一个文件夹
    if not os.path.exists('./paper_info'):
        os.mkdir('./paper_info')
    file_name = './paper_info/文献爬取.xlsx'

    # 2、设置搜索词
    search_words = ['云计算','数据安全','大数据','人工智能']
    # 3、创建工作簿
    workbook = xw.Workbook(filename=file_name)

    # 4、获取每个搜索词的文献内容
    for search_word in search_words:
        infos = []
        # 每个搜索词搜索6-1=5页
        for page_num in range(1, 10):
            print('搜索词：'+search_word+'---正在爬取第【'+str(page_num)+'】页...')
            page_text = get_page_text(
                url=base_url, headers=headers, search_word=search_word, page_num=page_num)
            page_info = parse_page_text(page_text=page_text)
            # 用+合并成一个列表，不是嵌套列表；用append，会形成嵌套列表
            infos += page_info
            time.sleep(5)

        # 5、按照搜索词，依次写入工作簿
        write_to_excel(workbook, infos, search_word)
    # 6、关闭工作簿
    workbook.close()

    print('爬取完成!')