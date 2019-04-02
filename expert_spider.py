import xlrd
import redis
import time
from lxml import etree
import re
import requests

pool_redis = redis.ConnectionPool(host='localhost', port=6379, decode_responses=True)
myredis = redis.Redis(connection_pool=pool_redis)

# excel=xlrd.open_workbook('专家表.xlsx')
# table=excel.sheets()[0]
# address=table.col_values(4)
# name=table.col_values(7)
#
# for a,n in zip(address,name):
#     if a and n:
#         person=f'{n}<->{a}'
#         print(person)
#         myredis.sadd('expert',person)


s = requests.session()

header = {'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
          'Accept': '*/*',
          'Host': 'kns.cnki.net',
          'Origin': 'http://kns.cnki.net',
          'Proxy-Connection': 'keep-alive',
          'Referer': 'http://kns.cnki.net/kns/brief/result.aspx?dbprefix=SCDB&crossDbcodes=CJFQ,CDFD,CMFD,CPFD,IPFD,CCND,CCJD',
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36',
          }


def get_parms(name, address):
    parms = {'action': '',
             'NaviCode': '*',
             'ua': '1.21',
             'isinEn': '1',
             'PageName': 'ASP.brief_result_aspx',
             'DbPrefix': 'SCDB',
             'DbCatalog': '中国学术期刊网络出版总库',
             'ConfigFile': 'SCDB.xml',
             'db_opt': 'CJFQ,CDFD,CMFD,CPFD,IPFD,CCND,CCJD',
             'publishdate_from': '2013-01-01',
             'publishdate_to': '2019-01-01',
             'au_1_sel': 'AU',
             'au_1_sel2': 'AF',
             'au_1_value1': name,
             'au_1_value2': address,
             'au_1_special1': '=',
             'au_1_special2': '%',
             'his': '0'
             }
    return parms


def get_procy():
    try:
        # proxy=myredis.srandmember('ip_proxy')
        # proxies = {'https': 'https://' + proxy, 'http': 'http://' + proxy}
        # return proxies
        return None
    except:
        raise


def reserch(name, address):#name与adress
    url = 'http://kns.cnki.net/kns/request/SearchHandler.ashx'
    s.post(url, headers=header, data=get_parms(name, address),  timeout=(10, 10))#proxies=get_procy(),


def get_articl_url(tree):
    urls = tree.xpath('//*[@class="fz14"]/@href')
    print(len(urls))
    if len(urls) > 2:
        for u in urls:
            myredis.sadd('cnki_article_expert_315', u)
    else:
        pass


def get_page_by_pageindex(pageindex, name, address):
    for _ in range(3):
        try:
            r = s.get(
                'http://kns.cnki.net/kns/brief/brief.aspx?curpage=%s&RecordsPerPage=50&QueryID=1&ID=&turnpage=1&tpagemode=L&dbPrefix=SCDB&Fields=&DisplayMode=listmode&PageName=ASP.brief_result_aspx&isinEn=1' % pageindex,
                headers=header, timeout=(10, 10))
            if '请输入验证码' in r.text:
                print('遇到验证码  重新获取 cookies')
                s.cookies.clear()
                reserch(name, address)
                get_page_by_pageindex(pageindex, name, address)
            tree = etree.HTML(r.text)
            print('Get Urls %s' % pageindex)
            get_articl_url(tree)
            break
        except Exception as e:
            print(e)
            pass
            if _ == 2:
                raise Exception('too many times')
    print('success', pageindex, name, address)


def Get_article_first(name, address):
    # 获取查询页第一页信息
    def Get_first_page(url):
        for _ in range(3):
            try:
                x = s.get('http://kns.cnki.net/kns/brief/brief.aspx?pagename=' + url, headers=header, timeout=(10, 10))
                return x.text.replace('&nbsp;', ' ')
            except Exception as e:
                print(e)

    url = 'http://kns.cnki.net/kns/request/SearchHandler.ashx'
    r = s.post(url, headers=header, data=get_parms(name, address), timeout=(10, 10))
    text = Get_first_page(r.text)
    tree = etree.HTML(text)
    article_nums = tree.xpath('//*[@class="pagerTitleCell"]/text()')[0]
    article_num = int(article_nums.split(' ')[2].replace(',', ''))
    max_page = int((article_num - 20) / 50)+1
    get_articl_url(tree)
######-----------转到详情页
    for i in range(1, max_page):
        try:
            get_page_by_pageindex(i, name, address)
        except:
            with open('makesi_faild.txt', 'a')as f:
                f.write('%s\n' % i)


def run():
    while True:
        person = myredis.spop('expert')
        Get_article_first('花向红','武汉大学')
        # try:
        #     if person:
        #         name, address = person.split('<->')
        #         Get_article_first(name, address)
        #     else:
        #         break
        # except:
        #     myredis.sadd('expert',person)


if __name__ == '__main__':
    run()
    # print(myredis.srandmember('expert'))
