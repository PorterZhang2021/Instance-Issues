"""
在爬取当当的Top500的数据后将其存放进excel当中
主体已经有了 爬取
excel的运用
"""
import re
import requests
import json
import openpyxl

def get_html(url):
    """
    获取到我们所需要的网页
    """
    # 异常捕获方式
    try:
        # 获取响应
        response = requests.get(url=url)
        # 进行响应码的判断
        if response.status_code == 200:
            # 将获取的网页信息返回
            return response.text
    # 未能请求成功
    except requests.RequestException:
        # 返回None
        return None

def get_parase_result(html):
    """解析相关的结果值"""
    # 构建正则表达式
    pattern = re.compile('<li>.*?list_num.*?(\d+).</div>.*?<img src="(.*?)".*?class="name".*?title="(.*?)">.*?class="star">.*?class="tuijian">(.*?)</span>.*?class="publisher_info">.*?target="_blank">(.*?)</a>.*?class="biaosheng">.*?<span>(.*?)</span></div>.*?<p><span\sclass="price_n">&yen;(.*?)</span>.*?</li>', re.S)
    # 进行查找
    items = re.findall(pattern, html)
    # 表达式函数生成
    for item in items:
        yield {
             # 排名
            'range': item[0],
            # 图片地址
            'image': item[1],
            # 书名
            'title': item[2],
            # 推荐
            'recommend': item[3],
            # 作者
            'author': item[4],
            # 五星评分次数
            'times': item[5],
            # 价格
            'price': item[6]
        }

def write_to_excels(item):
    """将数据写入excels"""
    pass


def main(page):
    # 爬取网页
    url = 'http://bang.dangdang.com/books/fivestars/1-' + str(page)
    # 获取html
    html = get_html(url)
    # 解析相关结果
    items = get_parase_result(html)
    # 将结果存入excel中
    for item in items:
        write_to_excels(item)

