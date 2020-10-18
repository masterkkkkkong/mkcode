# -*- coding:utf-8 -*-
# author   : MK
# datetime : 2020/10/18 4:17
# software : PyCharm

'''
本篇代码旨在从中国食品药品检定研究院（https://bio.nifdc.org.cn/pqf/search.do?formAction=pqfGs）中抓取各地方批签发公示表，
目前中检院共有7个地方药检所和1个中检院，共8个大目录，每个大目录下有不同月份签发的公示表，存放在小目录，即单独的页面中，
以下所使用的初始网址链接是小目录网址；

8个机构的小目录链接分别是：
中检院：https://bio.nifdc.org.cn/pqf/search.do?formAction=pqfGsByJG&orgId=1
北京药检所：https://bio.nifdc.org.cn/pqf/search.do?formAction=pqfGsByJG&orgId=5b6ea8c91cf9013d011cfdfbda100041
上海药检所：https://bio.nifdc.org.cn/pqf/search.do?formAction=pqfGsByJG&orgId=4028813a1d225be5011d2265474b0004
广东药检所：https://bio.nifdc.org.cn/pqf/search.do?formAction=pqfGsByJG&orgId=4028813a1d225be5011d226a9159001c
四川药检所：https://bio.nifdc.org.cn/pqf/search.do?formAction=pqfGsByJG&orgId=4028813a1d225be5011d226ba310001e
湖北药检所：https://bio.nifdc.org.cn/pqf/search.do?formAction=pqfGsByJG&orgId=4028813a1d225be5011d22697942001a
吉林药检所：https://bio.nifdc.org.cn/pqf/search.do?formAction=pqfGsByJG&orgId=4028813a1d225be5011d226392100002
甘肃药检所：https://bio.nifdc.org.cn/pqf/search.do?formAction=pqfGsByJG&orgId=4028813a1d225be5011d226c637d0020


用to_excel保存的利弊：
1.每张表格数据量相对较小，用Excel保存可以保留首行的合并单元格格式；
2.首列索引列无法去除，index=0会进行报错；
3.第三行会自动生成空白行；
4.CSV可以避免上述2,3问题，但无法实现1；


目前为1.0版本，后续优化功能：
1.针对下载xlsx表格格式进行优化，如去掉首列索引列，第三行空白等；
2.针对下载的多张工作簿表格，一键进行合并，形成一个工作簿，方便用户统计分析；

'''

from urllib.request import urlopen
from bs4 import BeautifulSoup
#用pandas库需先安装xlrd和openpyxl
import pandas as pd
import re


class GetNIFDCData(object):

    def get_name(self,name):
        #定义8大机构网址，放入字典
        url_head = 'https://bio.nifdc.org.cn/pqf/search.do?formAction=pqfGsByJG&orgId='
        url_dic={'中检院':'1',
                '北京药检所':'5b6ea8c91cf9013d011cfdfbda100041',
                '上海药检所':'4028813a1d225be5011d2265474b0004',
                '广东药检所':'4028813a1d225be5011d226a9159001c',
                '四川药检所':'4028813a1d225be5011d226ba310001e',
                '湖北药检所':'4028813a1d225be5011d22697942001a',
                '吉林药检所':'4028813a1d225be5011d226392100002',
                '甘肃药检所':'4028813a1d225be5011d226c637d0020'
        }
        #以用户输入名称为key，将value传给url
        url = url_head + url_dic[name]
        #url传参
        self.get_download_url(url)

    def get_download_url(self, url):
        '''
        请求头用于抓取中检院批签发数据暂且用不上，
        需要时可以用response.add_header('User-Agent'，'Mozilla/5.0....'进行添加
        headers = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 \
        (KHTML, like Gecko) Chrome/72.0.3626.121 Safari/537.36'}
        '''
        #打开列表页网站，返回给response，用美味汤解析得到源码data
        response = urlopen(url).read()
        data = BeautifulSoup(response, 'lxml')
        #用正则表达式取出源码中的Excel文件超链接，返回给re_data
        re_data = re.findall('<a href=\"(search\S+?\.xls)',str(data))
        #遍历re_data，用append方法放入url_list列表
        url_list = []
        for s in re_data:
            s = s.replace('amp;', '')
            url_list.append('https://bio.nifdc.org.cn/pqf/'+s)

        #让用户选择是否全部下载，如回答否，则选择下载最新文件的数量
        get_all = input(f'共有{len(url_list)}个文件，是否全部下载(Y/N)：')
        if get_all.lower() == 'n':
            get_num = input('请输入需下载最新文件数量：')
            get_num = int(get_num)
            #判断用户输入数值是否合法
            while get_num > len(url_list):
                print('输入的数值不得大于文件总数，请重新输入！')
                get_num = input('请输入需下载最新文件数量：')
                get_num = int(get_num)
            #根据用户输入数量，对原超链接列表进行切片
            url_list = url_list[:get_num]
        else:
            print('开始下载全部文件...')
        #将超链接列表url_list作为参数传递
        self.download_data(url_list)

    def download_data(self, url_list):
        #定义一个累加器，用于给文件命名
        i = 1
        #遍历超链接，取出每一个Excel源地址进行内容抓取
        for url in url_list:
            #用pandas对网页Excel表格进行抓取，并保存到G盘目录下，以药检所+编码的形式进行命名，以中文GB18030编码进行保存
            html_data = pd.read_html(url, encoding='utf-8')
            table_data = pd.DataFrame(html_data[1])
            #该表格内容不多，且有合并单元格，用Excel保存即可，如数据量较大，需改用CSV格式
            #table_data.to_csv(f'G:/nifdc/{name}-{i}.csv', index=0,encoding='gb18030')
            table_data.to_excel(f'G:/nifdc/{name}-{i}.xlsx', encoding='gb18030')
            i += 1
            #控制台显示抓取进度
            print(f'共{len(url_list)},已下载{i-1}个，剩余{len(url_list)-i+1}个')
        #全部抓取完成，输出完成信息
        print('批签发数据已下载完成！')

    def main(self, name):
        self.get_name(name)


if __name__ == '__main__':
    spider = GetNIFDCData()
    name = input('中检院、北京药检所、上海药检所、广东药检所、四川药检所、湖北药检所、'
                 '吉林药检所、甘肃药检所\n请选择以上一个机构名称，并输入：')
    spider.main(name)
