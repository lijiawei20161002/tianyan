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
            'Cookie': r'TYCID=e4768780a20911ea8b8a45df60a2616f; undefined=e4768780a20911ea8b8a45df60a2616f; ssuid=4506631042; _ga=GA1.2.1861530587.1590797399; activityTag=20200618; _gid=GA1.2.335080107.1592469073; RTYCID=8c493816e83e45948397f32540493f03; CT_TYCID=951dc080d7f5400389ac816bcb37cdf9; aliyungf_tc=AQAAAFGQbX+UFQwAXXjHt1fz9XUnT3IU; csrfToken=XWAt8cwn7FLom1o3xVS3f0JP; activityIpTag=20200618IP; Hm_lvt_e92c8d65d92d534b0fc290df538b4758=1590887287,1591838353,1592469073,1592539542; bannerFlag=false; tyc-user-phone=%255B%252213716388133%2522%252C%2522137%25201765%25202988%2522%255D; bad_id658cce70-d9dc-11e9-96c6-833900356dc6=d62f70e1-b1e2-11ea-b0c3-bde9615d76ea; nice_id658cce70-d9dc-11e9-96c6-833900356dc6=d62f70e2-b1e2-11ea-b0c3-bde9615d76ea; cloud_token=a91a20b6410146af944bdcb4aeed9548; token=48af65e6e6ba471c9c37b0584d5b1995; _utm=6a7049e522b341a695bd95ca197c3c66; tyc-user-info=%257B%2522claimEditPoint%2522%253A%25220%2522%252C%2522vipToMonth%2522%253A%2522false%2522%252C%2522explainPoint%2522%253A%25220%2522%252C%2522personalClaimType%2522%253A%2522none%2522%252C%2522integrity%2522%253A%252210%2525%2522%252C%2522state%2522%253A%25225%2522%252C%2522surday%2522%253A%2522455%2522%252C%2522announcementPoint%2522%253A%25220%2522%252C%2522bidSubscribe%2522%253A%2522-1%2522%252C%2522vipManager%2522%253A%25220%2522%252C%2522monitorUnreadCount%2522%253A%25220%2522%252C%2522discussCommendCount%2522%253A%25220%2522%252C%2522onum%2522%253A%25221%2522%252C%2522showPost%2522%253Anull%252C%2522claimPoint%2522%253A%25220%2522%252C%2522token%2522%253A%2522eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiIxMzcxNjM4ODEzMyIsImlhdCI6MTU5MjU0OTg5MiwiZXhwIjoxNjI0MDg1ODkyfQ.Dl-J8OHpHlOyJhTrsKhuOcmUhUeNNK4WFfc5eSZ90ghCznHvFARbhn3-sxnEAcX7-taoxk1J33u1L4rA1YBF2Q%2522%252C%2522schoolAuthStatus%2522%253A%25222%2522%252C%2522vipToTime%2522%253A%25221631847112378%2522%252C%2522redPoint%2522%253A%25220%2522%252C%2522myTidings%2522%253A%25220%2522%252C%2522companyAuthStatus%2522%253A%25222%2522%252C%2522myAnswerCount%2522%253A%25220%2522%252C%2522myQuestionCount%2522%253A%25220%2522%252C%2522signUp%2522%253A%25220%2522%252C%2522privateMessagePointWeb%2522%253A%25220%2522%252C%2522nickname%2522%253A%2522%25E6%25B9%2596%25E7%2595%2594%25E9%2587%258C%25E7%25A8%258B%2522%252C%2522privateMessagePoint%2522%253A%25220%2522%252C%2522bossStatus%2522%253A%25222%2522%252C%2522isClaim%2522%253A%25220%2522%252C%2522yellowDiamondEndTime%2522%253A%25220%2522%252C%2522isExpired%2522%253A%25220%2522%252C%2522yellowDiamondStatus%2522%253A%2522-1%2522%252C%2522pleaseAnswerCount%2522%253A%25220%2522%252C%2522bizCardUnread%2522%253A%25220%2522%252C%2522vnum%2522%253A%25220%2522%252C%2522mobile%2522%253A%252213716388133%2522%257D; auth_token=eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiIxMzcxNjM4ODEzMyIsImlhdCI6MTU5MjU0OTg5MiwiZXhwIjoxNjI0MDg1ODkyfQ.Dl-J8OHpHlOyJhTrsKhuOcmUhUeNNK4WFfc5eSZ90ghCznHvFARbhn3-sxnEAcX7-taoxk1J33u1L4rA1YBF2Q; Hm_lpvt_e92c8d65d92d534b0fc290df538b4758=1592549983; _gat_gtag_UA_123487620_1=1'
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
    workbook = open_workbook(r'name.xls')
    sheet=workbook.sheet_by_index(0)
    name=[]
    for i in range(sheet.nrows):
        name.append(sheet.row_values(i))
    for nm in name[118:len(name)]:
        soup=craw('https://www.tianyancha.com/search?key='+parse.quote(nm[0])+'%20'+parse.quote('成都'))
        links = soup.body.select('.human.sv-search-company-human')
        for l in links:
            soup=craw(l.get('href'))
            com=soup.select('#_container_allCompany .link-click')
            wb=xlrd.open_workbook('/Users/lijiawei/desktop/tianyan/map.xls')
            ws=xlutils.copy.copy(wb)
            sheet = ws.get_sheet(0)
            nrows = wb.sheet_by_index(0).nrows
            sheet.write(nrows,0,nm)
            i=1
            for c in com:
                sheet.write(nrows,i,c.text)
                i=i+1
            ws.save('/Users/lijiawei/desktop/tianyan/map.xls')
            wb = xlrd.open_workbook('/Users/lijiawei/desktop/tianyan/company.xls')
            ws = xlutils.copy.copy(wb)
            sheet = ws.get_sheet(0)
            nrows = wb.sheet_by_index(0).nrows
            i = 0
            for c in com:
                sheet.write(nrows + i, 0, c.text)
                i = i + 1
            ws.save('/Users/lijiawei/desktop/tianyan/company.xls')
