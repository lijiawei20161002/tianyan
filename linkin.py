#-*- coding-8 -*-
import requests
from bs4 import BeautifulSoup
import xlwt
import xlrd
from xlutils.copy import copy
import time
import urllib
import random
import json
import re

User_Agent = 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:56.0) Gecko/20100101 Firefox/56.0'
headers = {
            'authority': 'www.linkedin.com',
            'method': 'GET',
            'path': '/search/results/all/?keywords=investment%20fund%20TMT&origin=GLOBAL_SEARCH_HEADER',
            'scheme': 'https',
            'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'accept-encoding': 'gzip, deflate, br',
            'accept-language': 'zh-CN,zh;q=0.9',
            'cache-control': 'max-age=0',
            'cookie': 'li_sugr=8f7cae9b-2def-40dc-b26a-176dbf49bf1b; bcookie="v=2&4bb073db-0665-4f7c-82f9-61b4719269e8"; bscookie="v=1&202006020835189bc6c0b1-fdb0-4613-8ebc-3142431bd93eAQEKwZcdcpJcEt20o9UaAp4sGe3kWBxy"; lissc=1; _ga=GA1.2.1638655186.1591839612; li_rm=AQEWsfHPgbbuUAAAAXKhCVvM9WmZ1x4lOqlOUVrFWzv78Fj-Gmx1Glb_TXHrhjwq7FiL-CXmVe1_HZL96OzLcCcWwUDp8Z6eBqU-34ELBlen4peoI1O69BqG; AMCVS_14215E3D5995C57C0A495C55%40AdobeOrg=1; _guid=683a5fa6-f580-4c20-bbd0-aae6233d4dee; li_oatml=AQHzuAwoY5O-UgAAAXKhCYRGfb4_12YLGwPXc4ijMOUB6KT0RodDoDPk-Ux8cCl36kCXFIDB9UX81Hgn0FQngh0J0DaO2dVi; aam_uuid=69423465375104826640944582606735839862; UserMatchHistory=AQJuARpwmkbvJgAAAXKhDIxkSuAOYJf956PDBjqubNokxls8pSFagqd2aoB3ovxl4XOtbuIE17c-sxlcsPljGuCN3WJNP0kQIttbJ73MpOFhwy9poU8qswW39nOkNQ5luHX2Z9HdbA; JSESSIONID="ajax:6867532158528684576"; visit=v=1&M; li_at=AQEDAQW3jBEFwHwKAAABcqEhPB8AAAFyxS3AH1YAEp1w7HSeEE40tpv-zKjPmrH4TYuOii-GTGpOx8dKPQiEr5F9etVkVOeU5y29HlzdGCvr5Ua1-j5viDYcVGXHDcwHirBe4pFpqQdu6xSEDOBlW5FE; liap=true; sl=v=1&Esoi3; li_cc=AQFF1zi2IwlOowAAAXKhIT3B5z3pzZW9Hs7fZhm5zCEoHsRQwTj-914IjXC5Lu-2J__kFhuDUF8; lang=v=2&lang=en-us; AMCV_14215E3D5995C57C0A495C55%40AdobeOrg=-408604571%7CMCIDTS%7C18425%7CMCMID%7C69600365271252503390996698343951458749%7CMCOPTOUT-1591848413s%7CNONE%7CvVersion%7C4.6.0%7CMCAAMLH-1592446013%7C11%7CMCAAMB-1592446013%7Cj8Odv6LonN4r3an7LhD3WZrU1bUpAkFkkiY1ncBR96t2PTI%7CMCCIDH%7C-295798417; spectroscopyId=2a6eeefc-f7f0-478e-8c5f-38ce7ba9c030; UserMatchHistory=AQJBFZT3aSQjoQAAAXKhI0RcH3OnzsNGGoWSZfMuZv1lbTHZ1FVS9ekPiFMSOPdUKw4f2oMiImZofgtItkjIr_hYhN_wW4Qwe1GqWGs12ppn-yNwq9gUHkjFQ_4DXcTM33JJVECsCmt0tkjpYkFcd0y9BWlZCBFTo-RsrjnF5Sn9Yl6JacIeqxOrqciGi3CnuZctrh77a1VX0AhGvhQhYDzjbLNyQsPT8B1fxk1OM3qC_wK3fm92UV-p9P8ZwlRBLXfbkruKEq2zKj4-qT-tsGXawsURZZEicIg; lidc="b=OB25:s=O:r=O:g=2746:u=473:i=1591841344:t=1591925534:v=1:sig=AQG6vUPHzY9M5SkEQf9Azu8li4JRvDJ1',
            'referer': 'https://www.linkedin.com/checkpoint/challenge/AgF8_ywikGfFJAAAAXKhIFYtCj2kImG0VWaMiHRkcAMGrsMrfmPlxEaHCRS_-XvnTzLPemfl7Mauk8pg0i6ZbEtnZtaDiw?ut=2kY5mMErZZQVg1',
            'sec-fetch-dest': 'document',
            'sec-fetch-mode': 'navigate',
            'sec-fetch-site': 'same-origin',
            'sec-fetch-user': '?1',
            'upgrade-insecure-requests': '1',
            'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.97 Safari/537.36',
}

def craw(url,p):
    url=url+'&page='+str(p)
    try:
        response=requests.get(url,headers=headers)
        if response.status_code != 200:
            response.encoding = 'utf-8'
            print(response.status_code)
            print('ERROR')
        soup = BeautifulSoup(response.text, 'lxml')
    except Exception as e:
        print(str(e))

    #extract name
    name=[]
    s = str(soup.body).split('data')[79].split('</code')[0]
    s1 = s
    while (len(s1) > 0):
        ss = s1.split('title', 1)
        if (len(ss) > 1):
            s1 = ss[1]
            sss = ss[1].split('}')[0].split('text":"')
            if((len(sss)>1) and (len(ss[1].split('"type":"'))>1) and ("PROFILE" in ss[1].split('"type":"')[1])):
                sss=sss[1].split('"')[0].encode('iso-8859-1').decode('utf-8')
                name.append(sss)
        else:
            s1 = ''

    # extract position
    pos = []
    s = str(soup.body).split('data')[79].split('</code')[0]
    s1 = s
    while (len(s1) > 0):
        ss = s1.split('headline', 1)
        if (len(ss) > 1):
            s1 = ss[1]
            sss = ss[1].split('}')[0].split('text":"')
            if ((len(sss) > 1)):
                sss = sss[1].split('"')[0].encode('iso-8859-1').decode('utf-8')
                pos.append(sss)
        else:
            s1 = ''

    #extract link
    link=[]
    s=str(soup.body).split('data')[79].split('</code')[0]
    s1=s
    while(len(s1)>0):
        ss=s1.split('navigationUrl',1)
        if(len(ss)>1):
            s1=ss[1]
            sss=ss[1].split(':"')[1].split('"')[0]
            link.append(sss)
        else:
            s1=''

    workbook = xlrd.open_workbook('linkedin.xls')
    sheet1=workbook.sheet_by_name('linkedin')
    oldrow = sheet1.nrows
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)
    for i in range(0,len(link)):
        new_worksheet.write(oldrow+i,0,name[i])
        new_worksheet.write(oldrow+i,1,pos[i])
        new_worksheet.write(oldrow+i,2,link[i])
    new_workbook.save('linkedin.xls')


if __name__ ==  '__main__':
    url = 'https://www.linkedin.com/search/results/all/?keywords=investment%20fund%20TMT&origin=GLOBAL_SEARCH_HEADER'
    for p in range(51,61):
        craw(url,p)
