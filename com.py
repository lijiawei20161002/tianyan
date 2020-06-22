#-*- coding-8 -*-
import requests
from bs4 import BeautifulSoup
import xlrd
from xlrd import open_workbook
import xlwt
from urllib import parse
import xlutils.copy

User_Agent = 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:56.0) Gecko/20100101 Firefox/56.0'
headers = {
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
            'Cookie': r'TYCID=e4768780a20911ea8b8a45df60a2616f; undefined=e4768780a20911ea8b8a45df60a2616f; ssuid=4506631042; _ga=GA1.2.1861530587.1590797399; activityTag=20200618; _gid=GA1.2.335080107.1592469073; RTYCID=8c493816e83e45948397f32540493f03; CT_TYCID=951dc080d7f5400389ac816bcb37cdf9; aliyungf_tc=AQAAAFGQbX+UFQwAXXjHt1fz9XUnT3IU; csrfToken=XWAt8cwn7FLom1o3xVS3f0JP; activityIpTag=20200618IP; Hm_lvt_e92c8d65d92d534b0fc290df538b4758=1590887287,1591838353,1592469073,1592539542; bannerFlag=false; tyc-user-phone=%255B%252213716388133%2522%252C%2522137%25201765%25202988%2522%255D; bad_id658cce70-d9dc-11e9-96c6-833900356dc6=d62f70e1-b1e2-11ea-b0c3-bde9615d76ea; nice_id658cce70-d9dc-11e9-96c6-833900356dc6=d62f70e2-b1e2-11ea-b0c3-bde9615d76ea; relatedHumanSearchGraphId=3061432849; relatedHumanSearchGraphId.sig=l4LSEL0d5ivcII1YXtmwClp8OZ9K3Jfs5XZWYFzikT4; tyc-user-info=%257B%2522claimEditPoint%2522%253A%25220%2522%252C%2522vipToMonth%2522%253A%2522false%2522%252C%2522explainPoint%2522%253A%25220%2522%252C%2522personalClaimType%2522%253A%2522none%2522%252C%2522integrity%2522%253A%252210%2525%2522%252C%2522state%2522%253A%25225%2522%252C%2522surday%2522%253A%2522455%2522%252C%2522announcementPoint%2522%253A%25220%2522%252C%2522bidSubscribe%2522%253A%2522-1%2522%252C%2522vipManager%2522%253A%25220%2522%252C%2522monitorUnreadCount%2522%253A%25220%2522%252C%2522discussCommendCount%2522%253A%25220%2522%252C%2522onum%2522%253A%25229%2522%252C%2522showPost%2522%253Anull%252C%2522claimPoint%2522%253A%25220%2522%252C%2522token%2522%253A%2522eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiIxMzcxNjM4ODEzMyIsImlhdCI6MTU5MjU1OTE1MSwiZXhwIjoxNjI0MDk1MTUxfQ.mbyhb11gEfcnCUryqQV8HUId-E0Re39tu0sfrvYhUY7wViKVSflJ0vXKPBv5ut3K3wpy0itnL7lTdFnV7A88aQ%2522%252C%2522schoolAuthStatus%2522%253A%25222%2522%252C%2522vipToTime%2522%253A%25221631847112378%2522%252C%2522redPoint%2522%253A%25220%2522%252C%2522myTidings%2522%253A%25220%2522%252C%2522companyAuthStatus%2522%253A%25222%2522%252C%2522myAnswerCount%2522%253A%25220%2522%252C%2522myQuestionCount%2522%253A%25220%2522%252C%2522signUp%2522%253A%25220%2522%252C%2522privateMessagePointWeb%2522%253A%25220%2522%252C%2522nickname%2522%253A%2522%25E6%25B9%2596%25E7%2595%2594%25E9%2587%258C%25E7%25A8%258B%2522%252C%2522privateMessagePoint%2522%253A%25220%2522%252C%2522bossStatus%2522%253A%25222%2522%252C%2522isClaim%2522%253A%25220%2522%252C%2522yellowDiamondEndTime%2522%253A%25220%2522%252C%2522isExpired%2522%253A%25220%2522%252C%2522yellowDiamondStatus%2522%253A%2522-1%2522%252C%2522pleaseAnswerCount%2522%253A%25220%2522%252C%2522bizCardUnread%2522%253A%25220%2522%252C%2522vnum%2522%253A%25220%2522%252C%2522mobile%2522%253A%252213716388133%2522%257D; auth_token=eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiIxMzcxNjM4ODEzMyIsImlhdCI6MTU5MjU1OTE1MSwiZXhwIjoxNjI0MDk1MTUxfQ.mbyhb11gEfcnCUryqQV8HUId-E0Re39tu0sfrvYhUY7wViKVSflJ0vXKPBv5ut3K3wpy0itnL7lTdFnV7A88aQ; _gat_gtag_UA_123487620_1=1; token=9a73887d15284e8495bcbc1271d9f5ae; _utm=ef0075c348d945e0aa01ffc9343ece60; Hm_lpvt_e92c8d65d92d534b0fc290df538b4758=1592559166; cloud_token=73e4f8ec948f4e3c8dd60a69d9624d56; cloud_utm=631bc94a6bf645fea434026d8547fcda'
}

def craw(url):
    try:
        response = requests.get(url,headers = headers)
        if response.status_code != 200:
            response.encoding = 'utf-8'
            print(response.status_code)
            print('ERROR')
        soup = BeautifulSoup(response.text,'lxml')
        return soup
    except Exception as e:
        print(e)
        return e


if __name__ ==  '__main__':
    workbook = open_workbook(r'company.xls')
    sheet=workbook.sheet_by_index(0)
    name=[]
    for i in range(sheet.nrows):
        name.append(sheet.row_values(i))
    for nm in name[118:len(name)]:
        soup=craw('https://www.tianyancha.com/search?key='+parse.quote(nm[0]))
        link = soup.body.select('#web-content .result-list.sv-search-container .header .name')[0].get('href')
        soup=craw(link)
        list1=[]
        list1.append(nm)
        tab=soup.select('#_container_baseInfo > table.table.-striped-col.-border-top-none.-breakall')
        print(tab[0].select('tr'))
        '''
            wb = xlrd.open_workbook('/Users/lijiawei/desktop/tianyan/company.xls')
            ws = xlutils.copy.copy(wb)
            sheet = ws.get_sheet(0)
            nrows = wb.sheet_by_index(0).nrows
            i = 0
            for c in com:
                sheet.write(nrows + i, 0, c.text)
                i = i + 1
            ws.save('/Users/lijiawei/desktop/tianyan/company.xls')'''