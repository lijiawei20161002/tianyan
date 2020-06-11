#-*- coding-8 -*-
import requests
from bs4 import BeautifulSoup
import xlwt
import time
import urllib
import random

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
            'Cookie': r'aliyungf_tc=AQAAANp7mEWnig0AKHjHt1IFK41mCWiu; csrfToken=6LonsP1DlKE1oxgIA_ljxp5T; TYCID=e4768780a20911ea8b8a45df60a2616f; undefined=e4768780a20911ea8b8a45df60a2616f; ssuid=4506631042; Hm_lvt_e92c8d65d92d534b0fc290df538b4758=1590797398; _ga=GA1.2.1861530587.1590797399; _gid=GA1.2.959696635.1590797399; tyc-user-phone=%255B%252213717652988%2522%255D; RTYCID=51986f61c972441ab3d5b1ad200f3327; CT_TYCID=ccb36e59ad234f9fa8a4802f479266fd; bannerFlag=true; cloud_token=f0aa75cf80a74e3a9bd3339e96417945; token=0a29e53dd21641ec8e480bb2a0006411; _utm=07ab75aaca44406d9232b1fd766a3163; tyc-user-info=%257B%2522claimEditPoint%2522%253A%25220%2522%252C%2522vipToMonth%2522%253A%2522false%2522%252C%2522explainPoint%2522%253A%25220%2522%252C%2522integrity%2522%253A%252210%2525%2522%252C%2522state%2522%253A%25220%2522%252C%2522announcementPoint%2522%253A%25220%2522%252C%2522bidSubscribe%2522%253A%2522-1%2522%252C%2522vipManager%2522%253A%25220%2522%252C%2522onum%2522%253A%25220%2522%252C%2522monitorUnreadCount%2522%253A%25220%2522%252C%2522discussCommendCount%2522%253A%25220%2522%252C%2522token%2522%253A%2522eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiIxMzcxNzY1Mjk4OCIsImlhdCI6MTU5MDgxODQzMywiZXhwIjoxNjIyMzU0NDMzfQ.A5u82pgQTgO8ij218mvKD1JhYDtd0YdM9QcPsHT5Cwiz8LsCtpXCjVPm8tdrlmHjFa8gENjyUrzkD0B4d49cRg%2522%252C%2522claimPoint%2522%253A%25220%2522%252C%2522redPoint%2522%253A%25220%2522%252C%2522myAnswerCount%2522%253A%25220%2522%252C%2522myQuestionCount%2522%253A%25220%2522%252C%2522signUp%2522%253A%25220%2522%252C%2522nickname%2522%253A%2522%25E4%25BD%2595%25E6%2589%25A7%25E4%25B8%25AD%2522%252C%2522privateMessagePointWeb%2522%253A%25220%2522%252C%2522privateMessagePoint%2522%253A%25220%2522%252C%2522isClaim%2522%253A%25220%2522%252C%2522pleaseAnswerCount%2522%253A%25220%2522%252C%2522vnum%2522%253A%25220%2522%252C%2522bizCardUnread%2522%253A%25220%2522%252C%2522mobile%2522%253A%252213717652988%2522%257D; auth_token=eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiIxMzcxNzY1Mjk4OCIsImlhdCI6MTU5MDgxODQzMywiZXhwIjoxNjIyMzU0NDMzfQ.A5u82pgQTgO8ij218mvKD1JhYDtd0YdM9QcPsHT5Cwiz8LsCtpXCjVPm8tdrlmHjFa8gENjyUrzkD0B4d49cRg; Hm_lpvt_e92c8d65d92d534b0fc290df538b4758=1590818444',
}
proxyHost='58.218.214.158'
proxyPort='2383'
proxyMeta="http://%(host)s:%(port)s" %{
    "host": proxyHost,
    "port": proxyPort
}
proxies={
    "http" : proxyMeta,
}

def craw(url,key_word,x,new_num):
    if x == 0:
        re = 'https://www.tianyancha.com/search?key=%E5%9F%BA%E9%87%91&base=sc'
    else:
        re = 'https://www.tianyancha.com/search/p{}?key=%E5%9F%BA%E9%87%91&base=sc'.format(x)
    headers['Referer']=re
    url=re
    try:
        response = requests.get(url,headers = headers,proxies=proxies)
        if response.status_code != 200:
            response.encoding = 'utf-8'
            print(response.status_code)
            print('ERROR')
        soup = BeautifulSoup(response.text,'lxml')
    except Exception:
        print('请求都不让，这天眼查也搞事情吗？？？')
    try:
        com_all_info = soup.body.select('.mt74 .container.-top .container-left .search-block.header-block-container')[0]
        com_all_info_array = com_all_info.select('.search-item.sv-search-company')
        print('开始爬取数据，请勿打开excel')

    except Exception:
        print('好像被拒绝访问了呢...请稍后再试叭...')
    try:
        for i in range(new_num,len(com_all_info_array)):
            try:
                temp_g_name = com_all_info_array[i].select('.content .header .name')[0].text    #获取公司名
                try:
                    temp_g_state = com_all_info_array[i].select('.content .header .tag-common.-normal-bg')[0].text  #获取公司状态
                except Exception:
                    temp_g_state = ''
                try:
                    temp_r_name = com_all_info_array[i].select('.content .legalPersonName.link-click')[0].text    #获取法人名
                except Exception:
                    temp_r_name = ''
                try:
                    temp_g_money = com_all_info_array[i].select('.content .info.row.text-ellipsis div')[1].text.strip('注册资本：')    #获取注册资本
                except Exception:
                    temp_g_money = ''
                try:
                    temp_g_date = com_all_info_array[i].select('.content .info.row.text-ellipsis div')[2].text.strip('成立日期：')    #获取公司注册时间
                except Exception:
                    temp_g_date = ''
                try:
                    try:
                        temp_r_phone = com_all_info_array[i].select('.content .contact.row div script')[0].text.strip('[').strip(']')    #获取法人手机号
                    except Exception:
                        temp_r_phone = com_all_info_array[i].select('.content .contact.row div')[0].select('span span')[0].text #获取法人手机
                except Ellipsis:
                    temp_r_phone = '未找到法人手机号'
                try:
                    try:
                        temp_r_email = com_all_info_array[i].select('.content .contact.row div script')[1].text.strip('[').strip(']')    #获取法人Email
                    except Exception:
                        temp_r_email = com_all_info_array[i].select('.content .contact.row div')[1].select('span')[1].text    #获取法人Email
                except Exception:
                    temp_r_email = '未找到法人邮箱'
    #            temp_g_desc = com_all_info_array[i].select('.content .match.row.text-ellipsis')[0].text    #获取公司备注
                try:
                    temp_g_link = com_all_info_array[i].select('.content .header .name')[0].get('href')
                except Exception:
                    temp_g_link = ''
                g_name_list.append(temp_g_name)
                g_link_list.append(temp_g_link)
                g_state_list.append(temp_g_state)
                r_name_list.append(temp_r_name)
                g_money_list.append(temp_g_money)
                g_date_list.append(temp_g_date)
                r_phone_list.append(temp_r_phone)
                r_email_list.append(temp_r_email)
            except Exception:
                print(temp_g_name+"-信息不完整，>>>>跳过>>下一个")
    except Exception:
        print("这页有毒，换下一页")


def craw2(url):
    try:
        response = requests.get(url, headers=headers,proxies=proxies)
        if response.status_code != 200:
            response.encoding = 'utf-8'
            print(response.status_code)
            print('ERROR')
        soup = BeautifulSoup(response.text, 'lxml')
    except Exception:
        print('请求都不让，这天眼查也搞事情吗？？？')

    try:
        summary=soup.select('.company-warp.-public .tabline .tabline-right .box.-company-box .content .detail .summary')[0]
        summary = summary.select('script')[0]
        summary = str(summary)
        summary = summary.split('>')[1]
        summary = summary.split('<')[0]
        summary = summary.strip()
    except Exception:
        summary = summary.select('span')[1].text
    try:
        site=soup.select('.detail-list .block-data-group .block-data .data-content .table.-striped-col.-border-top-none.-breakall tr')[5].select('td')[3].text
    except Exception:
        site=''
    try:
        rang=soup.select('.detail-list .block-data-group .block-data .data-content .table.-striped-col.-border-top-none.-breakall tr')[10].select('td')[1].text
    except Exception:
        rang=''
    g_sum_list.append(summary)
    g_site_list.append(site)
    g_rang_list.append(rang)

global g_link_list
global g_name_list
global g_state_list
global r_name_list
global g_money_list
global g_date_list
global r_phone_list
global r_email_list
global g_sum_list
global g_site_list
global g_rang_list

g_link_list=[]
g_name_list=[]
g_state_list=[]
r_name_list=[]
g_money_list=[]
g_date_list=[]
r_phone_list=[]
r_email_list=[]
g_sum_list=[]
g_site_list=[]
g_rang_list=[]

if __name__ ==  '__main__':
    key_word = input('请输入您想搜索的关键词：')
    try:
        startx = int(input('请输入您想从第几页检索：'))
    except Exception:
        startx = 0
    try:
        num = int(input('请输入您想检索的次数：'))+1
    except Exception:
        num = 5
    try:
        sleep_time = int(input('请输入每次检索延时的秒数：'))
    except Exception:
        sleep_time = 5

    key_word = urllib.parse.quote(key_word)

    print('正在搜索，请稍后')

    for x in range(startx,num):
        url = r'https://www.tianyancha.com/search?key=%E5%9F%BA%E9%87%91&base=sc'
        s1 = craw(url,key_word,x,0)
        time.sleep(random.randint(1,3))

    for url in g_link_list:
        if url != '':
            s2=craw2(url)
            time.sleep(random.randint(1,3))
        else:
            g_site_list.append('')
            g_rang_list.append('')
            g_sum_list.append('')
    workbook = xlwt.Workbook()
    #创建sheet对象，新建sheet
    sheet1 = workbook.add_sheet('天眼查数据', cell_overwrite_ok=True)
    #---设置excel样式---
    #初始化样式
    style = xlwt.XFStyle()
    #创建字体样式
    font = xlwt.Font()
    font.name = '仿宋'
#    font.bold = True #加粗
    #设置字体
    style.font = font
    #使用样式写入数据
    print('正在存储数据，请勿打开excel')
    #向sheet中写入数据
    name_list = ['公司名字','公司状态','法定法人','注册资本','成立日期','注册机关','简介','经营范围','法人电话','法人邮箱']
    print(len(g_name_list),len(g_site_list))
    for cc in range(0,len(name_list)):
        sheet1.write(0,cc,name_list[cc],style)
    for i in range(0,len(g_name_list)):
        sheet1.write(i+1,0,g_name_list[i],style)#公司名字
        sheet1.write(i+1,1,g_state_list[i],style)#公司状态
        sheet1.write(i+1,2,r_name_list[i],style)#法定法人
        sheet1.write(i+1,3,g_money_list[i],style)#注册资本
        sheet1.write(i+1,4,g_date_list[i],style)#成立日期
        sheet1.write(i+1,5,g_site_list[i],style)#注册机关
        sheet1.write(i+1,6, g_sum_list[i], style)  # 公司简介
        sheet1.write(i+1,7, g_rang_list[i], style)  # 经营范围
        sheet1.write(i+1,8,r_phone_list[i],style)#法人电话
        sheet1.write(i+1,9,r_email_list[i],style)#法人邮箱
    #保存excel文件，有同名的直接覆盖
    workbook.save(r"D:\wyy-tyc-"+time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime()) +".xls")
    print('保存完毕~')
