#-*- coding-8 -*-
import requests
from bs4 import BeautifulSoup
import xlwt
import time
import urllib
import random

User_Agent = 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:56.0) Gecko/20100101 Firefox/56.0'
uestc_headers = {
            'Host': 'sose.uestc.edu.cn',
            'Connection': 'keep-alive',
            'Cache-Control': 'max-age=0',
            'Upgrade-Insecure-Requests': '1',
            'User-Agent': r'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
            'Referer': 'http://sose.uestc.edu.cn/szdw/szdwjs.htm',#'https://www.tianyancha.com/search?key=%E5%B1%B1%E4%B8%9C%20%E7%A7%91%E6%8A%80',
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'zh-CN,zh;q=0.9',
}

tianyan_headers = {
            'Host': 'www.tianyancha.com',
            'Connection': 'keep-alive',
            'Cache-Control': 'max-age=0',
            'Upgrade-Insecure-Requests': '1',
            'User-Agent': r'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36',
            'Content-Language': 'zh-CN',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
            'Referer': 'https://www.tianyancha.com/',#'https://www.tianyancha.com/search?key=%E5%B1%B1%E4%B8%9C%20%E7%A7%91%E6%8A%80',
            'Accept-Encoding': 'gzip, deflate, br',
            'Accept-Language': 'zh-CN,zh;q=0.9',
            'Cookie': r'TYCID=e4768780a20911ea8b8a45df60a2616f; undefined=e4768780a20911ea8b8a45df60a2616f; ssuid=4506631042; _ga=GA1.2.1861530587.1590797399; tyc-user-phone=%255B%252213717652988%2522%255D; aliyungf_tc=AQAAAKZE0zUXfAMABXjHt6rTjodwt8YD; csrfToken=isJE2joF9E_HPlLjgN3F9l8F; bannerFlag=false; Hm_lvt_e92c8d65d92d534b0fc290df538b4758=1590797398,1590822705,1590887287,1591838353; _gid=GA1.2.335105467.1591838354; _gat_gtag_UA_123487620_1=1; token=68c03603f92a46a197d0f823d18fc45a; _utm=3d73470194fb4884860f5103e545aceb; tyc-user-info=%257B%2522claimEditPoint%2522%253A%25220%2522%252C%2522vipToMonth%2522%253A%2522false%2522%252C%2522explainPoint%2522%253A%25220%2522%252C%2522personalClaimType%2522%253A%2522none%2522%252C%2522integrity%2522%253A%252210%2525%2522%252C%2522state%2522%253A%25220%2522%252C%2522announcementPoint%2522%253A%25220%2522%252C%2522bidSubscribe%2522%253A%2522-1%2522%252C%2522vipManager%2522%253A%25220%2522%252C%2522onum%2522%253A%25220%2522%252C%2522monitorUnreadCount%2522%253A%25220%2522%252C%2522discussCommendCount%2522%253A%25220%2522%252C%2522showPost%2522%253Anull%252C%2522claimPoint%2522%253A%25220%2522%252C%2522token%2522%253A%2522eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiIxMzcxNzY1Mjk4OCIsImlhdCI6MTU5MTgzODM3NiwiZXhwIjoxNjIzMzc0Mzc2fQ.u260n9ahJg4ETYW-0hq1eQMcSDHcJqRGoRe_Lhj6wqBw3AlRPsHDwy1iCjDmwYomZ3GitPM3EOK9wXoWXK2pjA%2522%252C%2522schoolAuthStatus%2522%253A%25222%2522%252C%2522redPoint%2522%253A%25220%2522%252C%2522companyAuthStatus%2522%253A%25222%2522%252C%2522myAnswerCount%2522%253A%25220%2522%252C%2522myQuestionCount%2522%253A%25220%2522%252C%2522signUp%2522%253A%25220%2522%252C%2522nickname%2522%253A%2522%25E4%25BD%2595%25E6%2589%25A7%25E4%25B8%25AD%2522%252C%2522privateMessagePointWeb%2522%253A%25220%2522%252C%2522privateMessagePoint%2522%253A%25220%2522%252C%2522bossStatus%2522%253A%25222%2522%252C%2522isClaim%2522%253A%25220%2522%252C%2522yellowDiamondEndTime%2522%253A%25220%2522%252C%2522yellowDiamondStatus%2522%253A%2522-1%2522%252C%2522pleaseAnswerCount%2522%253A%25220%2522%252C%2522vnum%2522%253A%25220%2522%252C%2522bizCardUnread%2522%253A%25220%2522%252C%2522mobile%2522%253A%252213717652988%2522%257D; auth_token=eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiIxMzcxNzY1Mjk4OCIsImlhdCI6MTU5MTgzODM3NiwiZXhwIjoxNjIzMzc0Mzc2fQ.u260n9ahJg4ETYW-0hq1eQMcSDHcJqRGoRe_Lhj6wqBw3AlRPsHDwy1iCjDmwYomZ3GitPM3EOK9wXoWXK2pjA; Hm_lpvt_e92c8d65d92d534b0fc290df538b4758=1591838378'
}
def craw_name(url, p):
    if p>1:
        url=url+'/'+str(p)+'.htm'
    else:
        url=url+'.htm'
    try:
        response=requests.get(url,headers=uestc_headers)
        if response.status_code != 200:
            response.encoding = 'utf-8'
            print(response.status_code)
            print('ERROR')
        soup = BeautifulSoup(response.text,'lxml')
    except Exception:
        print('request failed...')
    name=soup.body.select('.wrap.clearfix .main-pic-list .text .tit')
    lt=[]
    for l in name:
        lt.append(l.text.encode('iso-8859-1').decode('utf-8'))
    return lt

def search_tianyan(url):
    try:
        response=requests.get(url,headers=tianyan_headers)
        if response.status_code != 200:
            response.encoding = 'utf-8'
            print(response.status_code)
            print('ERROR')
        soup = BeautifulSoup(response.text, 'lxml')
    except Exception:
        print('request failed...')
    info=soup.body.select('.mt74 .slider-human .human.sv-search-company-human .bottom .company')
    print(info)


if __name__ ==  '__main__':
    name=[]
    '''
    for p in range(1,11):
        l=craw_name('http://sose.uestc.edu.cn/szdw/szdwjs',p)
        name=name+l
    '''
    search_tianyan('https://www.tianyancha.com/search?key=刘天')
    '''for nm in name:
        url='https://www.tianyancha.com/search?key='+nm
        search_tianyan(url)'''

    '''
    workbook = xlwt.Workbook()
    sheet1 = workbook.add_sheet('人名', cell_overwrite_ok=True)
    for i in range(0,len(name)):
        sheet1.write(i,0,name[i])
    workbook.save('name.xls')
    '''
