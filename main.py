"""
在当当网上图书频道爬取最热门畅销的500本书，将爬取的内容存入Excel文件，
对Excel中的数据内容进行分析，将分析的结果用图形进行展示。
"""
import jieba
import matplotlib.pyplot as plt
import requests
import wordcloud
import xlwings as xw
from bs4 import BeautifulSoup
from tqdm import tqdm
import os


def request_dangdang(url):  # 用requests库来请求html内容
    try:
        response = requests.get(url)
        if response.status_code == 200:
            html = response.text
            return html
    except requests.RequestException:
        return None


def plot_price_pie(data):  # 当当网top500书籍价格饼状图
    price_list_without_comma = []
    for i in data:
        price_list_without_comma.append(i.replace(',', ''))
    print('price_list_without_comma:', price_list_without_comma)
    price_list_without_comma = [float(x) for x in price_list_without_comma]

    results = {'小于20元': sum(i < 20 for i in price_list_without_comma),
               '20-50元': sum((50 > i >= 20) for i in price_list_without_comma),
               '50-70元': sum((70 > i >= 50) for i in price_list_without_comma),
               '70-100元': sum((100 > i >= 70) for i in price_list_without_comma),
               '100元以上': sum(i >= 100 for i in price_list_without_comma)}

    plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']  # 用来显示中文标签(仅mac os可用苹方)

    plt.pie(list(results.values()), labels=list(results.keys()))

    plt.title("当当网top500书籍价格饼状图")
    plt.grid()
    plt.savefig('bin/price_pie.png')
    plt.show()


def plot_comments_pie(data):  # 当当网top500书籍评论数量饼状图
    data2 = []

    for i in data:  # 有一本书的评论无法爬取, 是空数据, 所以将它设为0
        if i == '':
            data2.append(0)
        else:
            data2.append(int(i))

    results = {'小于10k条评论': sum(i < 10000 for i in data2),
               '10k-30k条评论': sum((30000 > i >= 10000) for i in data2),
               '30k-50k条评论': sum((50000 > i >= 30000) for i in data2),
               '50k-100k条评论': sum((100000 > i >= 50000) for i in data2),
               '100k以上条评论': sum(i >= 100000 for i in data2)}

    plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']  # 用来显示中文标签(仅mac os可用苹方)

    plt.pie(list(results.values()), labels=list(results.keys()))

    plt.title("当当网top500书籍评论数量饼状图")
    plt.grid()
    plt.savefig('bin/comment_pie.png')
    plt.show()


def plot_fivestars_pie(data):  # 当当网top500五星评价数量饼状图
    data2 = []
    for i in data:
        data2.append(int(i))

    results = {'少于1k个五星评价个数': sum(i < 1000 for i in data2),
               '1k-3k个五星评价': sum((3000 > i >= 1000) for i in data2),
               '3k-5k个五星评价': sum((5000 > i >= 3000) for i in data2),
               '5k-7k个五星评价': sum((7000 > i >= 5000) for i in data2),
               '7k以上的五星评价个数': sum(i >= 7000 for i in data2)}

    plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']  # 用来显示中文标签(仅mac os可用苹方)

    plt.pie(list(results.values()), labels=list(results.keys()))

    plt.title("当当网top500五星评价数量饼状图")
    plt.grid()
    plt.savefig('bin/fivestars_pie.png')
    plt.show()


def plot_wordcloud(data):  # 书名词云图
    ls = []
    for i in data:
        ls.extend(jieba.lcut(i))
    # print(ls)
    txt = " ".join('%s' % i for i in ls)
    # print(txt)
    w = wordcloud.WordCloud(width=600, height=600, background_color="white", font_path='bin/PingFang Medium.otf')
    w.generate(txt)
    w.to_file("bin/词云.png")


def main():
    # 数据初始化区
    name_list = []
    link_list = []
    price_list = []
    comment_list = []
    fivestars_list = []
    publisher_list = []
    press_list = []

    for x in tqdm(range(1, 26)):  # 使用tqdm库来实现更美观的进度条
        # print(x * 100 / 25, '%')  # 传统的进度条
        url = 'http://bang.dangdang.com/books/fivestars/01.00.00.00.00.00-recent30-0-0-1-' + str(x)
        html = request_dangdang(url)  # 获取html内容
        soup = BeautifulSoup(html, "html.parser")  # 使用bs4解析至soup

        # 爬取数据
        all_book = soup.find('ul', attrs={'class': 'bang_list clearfix bang_list_mode'})
        each_book = all_book.find_all('li')
        for i in each_book:
            all_name_list = i.find_all('div', attrs={'class': 'name'})  # 书名
            all_price_list = i.find_all('div', attrs={'class': 'price'})  # 价格
            all_comment_list = i.find_all('div', attrs={'class': 'star'})  # 评论数量
            all_fivestars_list = i.find_all('div', attrs={'class': 'biaosheng'})  # 评论数量
            all_publisher_list = i.find_all('div', attrs={'class': 'publisher_info'})  # 作者信息
            for j in all_name_list:
                name_list.append(j.find('a').get('title'))
                link_list.append(j.find('a').get('href'))
            for j in all_price_list:
                temp = j.find('p').find('span').contents[0][1:]
                price_list.append(temp)
            for j in all_comment_list:
                comment_list.append(j.find('a').contents[0][:-3])
            for j in all_fivestars_list:
                fivestars_list.append(j.find('span').contents[0][:-1])
            try:
                for index, j in enumerate(all_publisher_list):
                    if index % 2 == 0:
                        try:
                            publisher_list.append(j.find('a').get('title'))
                        except AttributeError:
                            publisher_list.append('无')
                    else:
                        press_list.append(j.find('a').contents[0])
            except IndexError:  # 此处总是出现异常, 因为当当网数据变化
                pass

    # debug
    print('name_list:', name_list)
    print('link_list', link_list)
    print('price_list:', price_list)
    print('comment_list:', comment_list)
    print('fivestars_list:', fivestars_list)
    print('publisher_list:', publisher_list)
    print('press_list:', press_list)

    # 创建data.xls
    os.remove('data.xls')
    f = open('data.xls', "w")
    f.close()

    # xlwings基本框架
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    app.screen_updating = True
    filepath = r'data.xls'
    wb = app.books.open(filepath)
    she = wb.sheets['data']

    # 填入书本的基本信息到表格
    she.range('A1').value = ['序号', '书名', '书本链接', '价格', '评论数量',
                             '五星评分次数', '作者信息', '出版社']
    she.range('A2').options(transpose=True).value = list(range(1, 501))  # 序号
    she.range('B2').options(transpose=True).value = name_list  # 书名
    she.range('C2').options(transpose=True).value = link_list  # 书本链接
    she.range('D2').options(transpose=True).value = price_list  # 价格
    she.range('E2').options(transpose=True).value = comment_list  # 评论数量
    she.range('F2').options(transpose=True).value = fivestars_list  # 五星评分次数
    she.range('G2').options(transpose=True).value = publisher_list  # 作者信息
    she.range('H2').options(transpose=True).value = press_list  # 出版社

    # 根据爬取到的数据来画数据分析图
    plot_price_pie(price_list)
    plot_comments_pie(comment_list)
    plot_fivestars_pie(fivestars_list)
    plot_wordcloud(name_list)

    # 写入xls文件
    she.pictures.add('bin/price_pie.png', left=she.range('L1').left, top=she.range('L1').top)
    she.pictures.add('bin/comment_pie.png', left=she.range('L30').left, top=she.range('L30').top)
    she.pictures.add('bin/fivestars_pie.png', left=she.range('L60').left, top=she.range('L60').top)
    she.pictures.add('bin/词云.png', left=she.range('L90').left, top=she.range('L90').top)

    wb.save()


if __name__ == '__main__':
    main()
