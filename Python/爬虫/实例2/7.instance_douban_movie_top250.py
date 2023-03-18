from bs4 import BeautifulSoup
import requests
import xlwt

# 全局变量
n = 1

def request_douban(url):
    """
    请求豆瓣电影，希望其能够返回top250
    """
    try:
        # 请求响应 这里需要进行模拟请求
        headers={
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Host': 'movie.douban.com',
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_6) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.0.5 Safari/605.1.15'
        }
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            return response.text
    except requests.RequestException:
        return None

# 这个不能写在里面，因为这里会执行5次每次会重新更新
# 创建文件
book = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = book.add_sheet('豆瓣电影Top250', cell_overwrite_ok=True)
sheet.write(0, 0, '名称')
sheet.write(0, 1, '图片')
sheet.write(0, 2, '排名')
sheet.write(0, 3, '评分')
sheet.write(0, 4, '作者')
sheet.write(0, 5, '简介')



def save_to_excel(soup):
    """获取到豆瓣的数据"""
    list = soup.find(class_='grid_view').find_all('li')
    

    for item in list:
        # 名称
        item_name = item.find(class_='title').string
        # 图片
        item_img = item.find('a').find('img').get('src')
        # 索引
        item_index = item.find(class_='').string
        # 评分
        item_score = item.find(class_='rating_num').string
        # 作者
        item_author = item.find('p').text
        # 描述
        if(item.find(class_='inq')!=None):
            item_intr = item.find(class_='inq').string
        else:
            item_intr = ''
        print('爬取电影：' + item_index + ' | ' + item_name  +' | ' + item_score  +' | ' + item_intr )
        
        global n

        # 写入数据
        sheet.write(n, 0, item_name)
        sheet.write(n, 1, item_img)
        sheet.write(n, 2, item_index)
        sheet.write(n, 3, item_score)
        sheet.write(n, 4, item_author)
        sheet.write(n, 5, item_intr)
        # n自增
        n = n + 1
    

def main(page):
    # 请求链接
    url = 'https://movie.douban.com/top250?start=' + str(page * 25) + '&filter='
    # 保存html
    html = request_douban(url)
    soup = BeautifulSoup(html, 'lxml')
    save_to_excel(soup)
    

if __name__ == '__main__':
    for i in range(0, 10):
        main(i)


# 数据保存
    book.save(u'豆瓣最受欢迎的250部电影.xls')
