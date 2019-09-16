#coding=utf-8
import urllib
import pymongo
import requests
import json
import xlwt
import smtplib
import time

from pyquery import PyQuery as pq
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.header import Header


# 排除已存在的
def insandexclude(collectionname, datalist):
    mycol = mydb[collectionname]
    for object in datalist:
        condition = {"title": object['title']}
        if mycol.find_one(condition) is None:
            mycol.insert_one(object)
        else:
            continue


# 关键词筛选
def keywordhandle(keywords, title, date, detailurl, newlist):
    for keyWord in keywords:
        if title.find(keyWord) != -1:
            newlist.append({'title': title,
                            'date': date,
                            'url': detailurl,
                            'isSend': 0})
            break

filePath = "/opt/py/"
myclient = pymongo.MongoClient("mongodb://localhost:27017/")
mydb = myclient["datas"]
# 竞品名称对照表
jingpinMap1 = {
    'wangzhongwang': '网中网',
    'keyun': '科云',
    'zhongjiao': '中教畅享',
    'fusite': '福斯特',
    'xindao': '用友新道',
    'changjietong': '畅捷通'

}

jingpinMap2 = {
    'jindie': '金蝶',
    'haokuaiji': '好会计',
    'bajie': '八戒财税',
    'yunzhangfang': '云账房',
    'zhangwuyou': '账无忧'
}

jingpinMap3 = {
'shandongshuilizhiye': '山东水利职业学院',
    'shandongshangzhi': '山东商职',
    'shandonglaodongzhiye': '山东劳动职业',
    'weihaizhiye': '威海职业学院',
    'rizhaozhiye': '日照职业技术学院',
    'liaochengzhiye': '聊城职业技术学院',
    'dongyingzhiye': '东营职业技术学院',
    'weifangzhiye': '潍坊职业技术学院',
    'sdkejizhiye': '山东科技职业',
    'binzhouzhiye': '滨州职业学院',
    'sdzhiye': '山东职业',
    'taishanzhiye': '泰山职业',
    'jnzhiye': '济南职业',
    'jndianzizhiye': '济南电子职业',
    'yantaiqichegongcheng': '烟台汽车工程',
    'sdshangwuzhiye': '山东商务职业',
    'sdqinggongzhiye': '山东轻工职业学院',
    'sdgongyezhiye': '山东工业职业学院',
    'sdwaimaozhiye': '山东外贸职业学院',
    'jngongchengzhiyejishu': '山东工程职业技术学院',
    'zibozhiye': '淄博职业技术学院',
    'sdjingmao': '山东经贸职业学院',
    'sdshenghan': '山东圣翰商贸学院',
    'weifanggongshangzhiye': '潍坊工商职业学院',
    'sdligongzhiye': '山东理工职业学院',
    'dezhouzhiye': '德州职业技术学院',
    'dongyingkejizhiye': '东营科技职业学院',
    'yantaigongchengzhiye': '烟台工程职业技术学院',
    'zaozhuangzhiye': '枣庄职业',
    'linyizhiye': '临沂职业学院',
    'weihaihaiyangzhiye': '威海海洋职业学院',
    'yantaihuangjin': '烟台黄金职业学院'
}

#关键词
keyWords = ['培训', '产教', '产学', '双师', '实训', '真账', '投标', '竞标', '技能', '就业', '会计技能', '金融会计', '管理会计']



#
#网中网 新闻筛选 begin
collectionName = 'wangzhongwang'
page = urllib.request.urlopen('http://www.netinnet.cn/jf/main/toNewsView')
htmlcode = page.read()
doc = pq(htmlcode)
#随机数
randomCode = doc.find("li:contains('公司新闻')").attr("id")
headers = {"Content-Type":"application/x-www-form-urlencoded"}
url = "http://www.netinnet.cn/jf/main/getNews"
params = {'page': 1, 'pageCount': 20, 'newsType': randomCode}
session = requests.session()
requ = session.post(url, data=params, headers=headers)
res = requ.text
#新闻数据
responseDataList = json.loads(res)
newList = []
#循环处理
for o in responseDataList['data']['list']:
    for keyWord in keyWords:
        if o['title'].find(keyWord) != -1:
            newList.append(o)
            break

mycol = mydb[collectionName]
for o in newList:
    condition = {"id":o['news_id']}
    if mycol.find_one(condition) is None:
        oo = {'id':o['news_id'],
        'date': o['publish_time'][0:8],
        'title': o['title'],
        'url':' http://www.netinnet.cn/jf/main/toNewsDetailView?newsId='+o['news_id'],
        'isSend': 0}
        mycol.insert_one(oo)
    else:
        continue
# 网中网 end.
#
# # 科云 begin
# # 选取8页
collectionName = 'keyun'
monthsMap = {'JAN': '01', 'FEB': '02', 'MAR': '03', 'APR': '04', 'MAY': '05', 'JUN': '06', 'JUL': '07', 'AUG': '08', 'SEP': '09', 'OCT': '10', 'NOV': '11', 'DEC': '12',
             'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06', 'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'}
newList = []
for i in range(8):
    url = "http://www.xmkeyun.com.cn/?mod=info&col_key=news&page=" + str(i)
    page = urllib.request.urlopen(url)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)
    actItems = doc.find(".article-item")
    # 每页8条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = monthsMap[o.find('.article-mon-day').text()[0:3]] + "-" + o.find('.article-mon-day>b').text()
        title = o.find("a[title]").attr('title')
        detailUrl = 'http://www.xmkeyun.com.cn' + o.find("a[title]").attr('href')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 科云 end.
#
# 中教畅想 begin
# 选取3页
collectionName = 'zhongjiao'
newList = []
for i in range(3):
    url = "http://www.itmc.cn/html/tnews/"
    if i > 0:
        url = url + "index_" + str(i + 1) + ".html"
    page = urllib.request.urlopen(url)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)
    actItems = doc.find(".main-m1")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.find(".m1-r dd:first").text() + "-" + o.find(".m1-r dt").text()
        title = o.find("a[title]").attr('title')
        detailUrl = o.find("a[title]").attr('href')
        # 如果以斜线开头 的特殊处理
        if detailUrl.find("/") == 0:
            detailUrl = "http://www.itmc.cn" + detailUrl
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 中教畅想 end.
#
#
# # 福斯特 begin
# # 选取1页 前10条
collectionName = 'fusite'
newList = []

url = "http://www.fstgs.com/user/newsCenter.aspx?module=1"
page = urllib.request.urlopen(url)
htmlcode = page.read()
htmlcode = htmlcode.decode('utf-8')
doc = pq(htmlcode)
actItems = doc.find("#ContentPlaceHolder1_DataList1 table.f-il")
# 每页n条
for ii in range(10):
    o = pq(actItems[ii])
    date = o.find("td.dateStr").text()
    title = o.find("a").text()
    detailUrl = "http://www.fstgs.com/user/" + o.find("a").attr('href')

    # 关键字筛选
    keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 福斯特 end.
#
#
# # 用友新道 begin
# # 选取15页
collectionName = 'xindao'
newList = []
for i in range(15):
    url = "http://www.seentao.com/news/index/cid/12.html?cid=12&page="
    url = url + str(i)
    page = urllib.request.urlopen(url)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)
    actItems = doc.find("div.content_dynamic")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.find("span.content_journalism_txt_p3_span1").text()[0:10]
        title = o.find("div.content_journalism_txt_p1").text()
        detailUrl = o.children("a[href]").attr('href')

        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 用友新道 end.
#
# # 畅捷通 begin
# # 选取3页
collectionName = 'changjietong'
newList = []
for i in range(3):
    url = "http://www.chanjet.com/news/newsList?num="
    url = url + str(i + 1)
    page = urllib.request.urlopen(url)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)
    actItems = doc.find("div.list-group-news li.clearfloat")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.find("div.news-time-box>p.fz12").text() + "-" + o.find("div.news-time-box>p.fz24").text()
        title = o.find("div.list-group-text>p.fz20").text()
        detailUrl = "http://www.chanjet.com" + o.children("a[href]").attr('href')

        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 畅捷通 end.


# 企业咨询汇总
# 企业咨询关键词
keyWords = ['合作', '联盟', '学校', '学院', '高职', '中职', '机构', '平台', '协会']

# # 金蝶 begin
collectionName = 'jindie'
newList = []
for i in range(13):
    url = "http://www.kingdee.com/news/"
    if i > 0:
        url = url + "page/" + str(i + 1)
    page = urllib.request.urlopen(url)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)
    actItems = doc.find("section.mk-blog-container article")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.find("time").text()
        title = o.find("h3>a").text()
        detailUrl = o.find("a[href]:first").attr('href')

        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 金蝶 end

# 好会计 begin
collectionName = 'haokuaiji'
newList = []
for i in range(25):
    url = "https://h.chanjet.com/news?auto=false&page="
    url = url + str(i + 1)
    page = urllib.request.urlopen(url)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)
    actItems = doc.find("div.list-group-news li.clearfloat")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.find("div.news-time-box>p.fz12").text() + "-" + o.find("div.news-time-box>p.fz24").text()
        title = o.find("div.list-group-text>p.fz20").text()
        detailUrl = o.find("a[href]:first").attr('href')

        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 好会计 end

# 云账房 begin
collectionName = 'yunzhangfang'
newList = []
for i in range(10):
    url = "http://www.yunzhangfang.com/category/3?page="
    url = url + str(i + 1)
    page = urllib.request.urlopen(url)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)
    actItems = doc.find("div.item-list div.item")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.find("div.title-fl>i").text() + "-" + o.find("div.title-fl>em").text()[0:2] + "-" + o.find("div.title-fl>em").text()[3:5]
        title = o.find("div.title-fr>h3").text()
        detailUrl = "http://www.yunzhangfang.com" + o.find("a[href]:first").attr('href')

        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 云账房 end

# 账无忧 begin
collectionName = 'zhangwuyou'
newList = []
url = "http://www.kdzwy.com/xinwendongtai.html"
page = urllib.request.urlopen(url)
htmlcode = page.read()
htmlcode = htmlcode.decode('utf-8')
doc = pq(htmlcode)
actItems = doc.find("section.newsList li")
# 每页n条
for ii in range(actItems.length):
    o = pq(actItems[ii])
    date = o.find("span.time>font").text() + "-" + o.find("span.time").text()[0:2]
    title = o.find("h4").text()
    detailUrl = "http://www.kdzwy.com" + o.find("a[href]:first").attr('href')

    # 关键字筛选
    keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 账无忧 end

# 高职院校咨询 begin.

# 山东水利职业学院 begin
collectionName = 'shandongshuilizhiye'
newList = []
url = "http://www.sdwcvc.cn/jjgl/xwzx/"
newsUrl = "http://www.sdwcvc.cn/jjgl"
succ = 'xwdt.htm'

for i in range(3):
    url = url + succ
    page = urllib.request.urlopen(url)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)

    succ = doc.find("a.Next").attr('href')
    if succ.find("xwdt") == -1:
        url = "http://www.sdwcvc.cn/jjgl/xwzx/xwdt/"
    else:
        url = "http://www.sdwcvc.cn/jjgl/xwzx/"
    actItems = doc.find("div.main_conRCb li")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.find("span").text()
        title = o.find("a").text()
        detailUrl = newsUrl + o.find("a").attr('href').replace('..', '')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 山东水利职业学院 end


# 山东商职 begin
collectionName = 'shandongshangzhi'
newList = []
url = "http://kjxy.sict.edu.cn/index/index/art_list/classid/376.html?page="

for i in range(3):
    url = url + str(i + 1)
    page = urllib.request.urlopen(url)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)
    actItems = doc.find("div.gtco-section>div.gtco-container div.row>ul>li")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.find("span.date").text()
        title = o.find("a").text()
        detailUrl = "http://kjxy.sict.edu.cn" + o.find("a").attr('href')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 山东商职 end


# 山东劳动职业 begin
collectionName = 'shandonglaodongzhiye'
newList = []
url = "http://jgx.sdlvtc.cn/xbxw.htm"
headersx = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/34.0.1847.137 Safari/537.36 LBBROWSER'}
req = urllib.request.Request(url=url, headers=headersx)
page = urllib.request.urlopen(req)
htmlcode = page.read()
htmlcode = htmlcode.decode('utf-8')
doc = pq(htmlcode)
actItems = doc.find("div.lstbox ul.listUl>li")
# 每页n条
for ii in range(actItems.length):
    o = pq(actItems[ii])
    date = o.find("span").text()
    title = o.find("a").text()
    detailUrl = "http://jgx.sdlvtc.cn/" + o.find("a").attr('href')
    # 关键字筛选
    keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 山东劳动职业 end

# 威海职业学院 begin
collectionName = 'weihaizhiye'
newList = []
url = "http://www.weihaicollege.com/jjgl2/1566/list"

for i in range(3):
    url1 = url + str(i + 1) + '.htm'
    page = urllib.request.urlopen(url1)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)
    actItems = doc.find("table.wp_article_list_table>tr")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.find("td[align='right']").text()
        title = o.find("a").text()
        detailUrl = "http://www.weihaicollege.com" + o.find("a").attr('href')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 威海职业学院 end


# 日照职业技术学院 begin
collectionName = 'rizhaozhiye'
newList = []
url = "http://kjxy.rzpt.cn/ctnlist.php/mid/12/page/"

for i in range(3):
    url1 = url + str(i + 1)
    page = urllib.request.urlopen(url1)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)
    actItems = doc.find("ul#ctnlist>li")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.find("span.time").text().replace('[', '').replace(']', '')
        title = o.find("a").text()
        detailUrl = "http://kjxy.rzpt.cn" + o.find("a").attr('href')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 日照职业技术学院 end

# 聊城职业技术学院 begin
collectionName = 'liaochengzhiye'
newList = []
url = "http://jfx.lcvtc.edu.cn/xyxw.htm"

page = urllib.request.urlopen(url)
htmlcode = page.read()
htmlcode = htmlcode.decode('utf-8')
doc = pq(htmlcode)
actItems = doc.find("div.main_conRCb>ul>li")
# 每页n条
for ii in range(actItems.length):
    o = pq(actItems[ii])
    date = o.find("span").text()
    title = o.find("em").text()
    detailUrl = "http://jfx.lcvtc.edu.cn/" + o.find("a").attr('href')
    # 关键字筛选
    keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 日照职业技术学院 end

# 东营职业技术学院 begin
collectionName = 'dongyingzhiye'
newList = []
url = "http://www.dyxy.edu.cn/kjxy/news_list.aspx?category_id=58&page="

for i in range(24):
    url1 = url + str(i + 1)
    page = urllib.request.urlopen(url1)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)
    actItems = doc.find("div.mcent2>ul>li")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.find("span.rig").text().replace('【', '').replace('】', '')
        title = o.find("a").text()
        detailUrl = "http://www.dyxy.edu.cn" + o.find("a").attr('href')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 东营职业技术学院 end

# 潍坊职业技术学院 begin
collectionName = 'weifangzhiye'
newList = []
url = "http://jg.sdwfvc.com/xwzx/"
succ = 'xyyw.htm'

for i in range(6):
    url = url + succ
    page = urllib.request.urlopen(url)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)

    succ = doc.find("a.Next").attr('href')
    if succ.find("xyyw") == -1:
        url = "http://jg.sdwfvc.com/xwzx/xyyw/"
    else:
        url = "http://jg.sdwfvc.com/xwzx/"
    actItems = doc.find("ul.listUl>li")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.find("span").text()
        title = o.find("a").text()
        detailUrl = "http://jg.sdwfvc.com" + o.find("a").attr('href').replace('..', '')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 潍坊职业技术学院 end


# 山东科技职业技术学院 begin
collectionName = 'sdkejizhiye'
newList = []
url = "http://jg.sdvcst.edu.cn/index/"
succ = 'xbxw.htm'

for i in range(3):
    url = url + succ
    page = urllib.request.urlopen(url)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)

    succ = doc.find("a.Next").attr('href')
    if succ.find("xbxw") == -1:
        url = "http://jg.sdvcst.edu.cn/index/xbxw/"
    else:
        url = "http://jg.sdvcst.edu.cn/index/"
    actItems = doc.find("span.timestyle120532")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.text()
        title = o.parent().prev().find("a").text()
        detailUrl = "http://jg.sdvcst.edu.cn/" + o.parent().prev().find("a").attr('href').replace('../', '')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 山东科技职业技术学院 end


# 滨州职业学院 begin
collectionName = 'binzhouzhiye'
newList = []
url = "http://kj.bzpt.edu.cn/s/82/t/126/p/11/i/"
succ = "/list.htm"
for i in range(2):
    url1 = url + str(i + 1) + succ
    page = urllib.request.urlopen(url1)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)
    actItems = doc.find("#newslist>table:first>tr")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = ""
        title = o.find("a").text()
        detailUrl = "http://kj.bzpt.edu.cn/" + o.find("a").attr('href')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 滨州职业学院 end.

# 山东职业学院 begin
collectionName = 'sdzhiye'
newList = []
url = "http://www.sdp.edu.cn/glx/"
succ = 'xbxw.htm'

for i in range(6):
    url = url + succ
    page = urllib.request.urlopen(url)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)

    succ = doc.find("a.Next").attr('href')
    if succ.find("xbxw") == -1:
        url = "http://www.sdp.edu.cn/glx/xbxw/"
    else:
        url = "http://www.sdp.edu.cn/glx/"
    actItems = doc.find("span.timestyle202058")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.text()
        title = o.parent().prev().find("a").text()
        detailUrl = "http://www.sdp.edu.cn/" + o.parent().prev().find("a").attr('href')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 山东职业学院 end


# 泰山职业学院 begin
collectionName = 'taishanzhiye'
newList = []
url = "http://www.mtotc.com.cn/caijing/LmflShow.asp?cid=25&Page="
for i in range(2):
    url1 = url + str(i + 1)
    page = urllib.request.urlopen(url1)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('GBK')
    doc = pq(htmlcode)
    actItems = doc.find("tr.artlist")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = ""
        title = o.find("a").text()
        detailUrl = "http://www.mtotc.com.cn/caijing/" + o.find("a").attr('href')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 泰山职业学院 end.

# 济南职业学院 begin
collectionName = 'jnzhiye'
newList = []
url = "http://cj.jnvc.cn/index/"
succ = 'xbkx.htm'

for i in range(6):
    url = url + succ
    page = urllib.request.urlopen(url)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)

    succ = doc.find("a.Next").attr('href')
    if succ.find("xbkx") == -1:
        url = "http://cj.jnvc.cn/index/xbkx/"
    else:
        url = "http://cj.jnvc.cn/index/"
    actItems = doc.find("div.nr li")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.find('span').text().replace('[', '').replace(']', '')
        title = o.find("a").text()
        detailUrl = "http://cj.jnvc.cn/" + o.find("a").attr('href').replace('../', '')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 济南职业学院 end

# 济南电子职业学院 begin
collectionName = 'jndianzizhiye'
newList = []
url = "http://www.sdcet.cn/cjjrx/channels/ch00485/"
page = urllib.request.urlopen(url)
htmlcode = page.read()
htmlcode = htmlcode.decode('utf-8')
doc = pq(htmlcode)

actItems = doc.find("tr.pagedContent")
# 每页n条
for ii in range(actItems.length):
    o = pq(actItems[ii])
    date = o.find('td.zi4').text().replace("/", "-")
    title = o.find("a").text()
    detailUrl = o.find("a").attr("href")
    # 关键字筛选
    keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 济南电子职业学院 end

# 烟台汽车工程 begin
collectionName = 'yantaiqichegongcheng'
newList = []
url = "http://jjglx.ytqcvc.cn/erji.jsp?a9t=36&a9p="
succ = "&a9c=10&urltype=tree.TreeTempUrl&wbtreeid=1469"
for i in range(5):
    url1 = url + str(i + 1) + succ
    page = urllib.request.urlopen(url1)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)
    actItems = doc.find("span.timestyle1331")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.text().replace("/", "-")
        title = o.parent().prev().find("a").text()
        detailUrl = "http://jjglx.ytqcvc.cn/" + o.parent().prev().find("a").attr('href')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 烟台汽车工程 end.

# 山东商务职业学院 begin
# collectionName = 'sdshangwuzhiye'
# newList = []
# url = "http://kj.sdbi.com.cn/"
# succ = 'xwdt.htm'
#
# for i in range(12):
#     url = url + succ
#     page = urllib.request.urlopen(url)
#     htmlcode = page.read()
#     htmlcode = htmlcode.decode('utf-8')
#     doc = pq(htmlcode)
#
#     succ = doc.find("a.Next").attr('href')
#     if succ.find("xwdt") == -1:
#         url = "http://kj.sdbi.com.cn/xwdt/"
#     else:
#         url = "http://kj.sdbi.com.cn/"
#     actItems = doc.find("div.list li")
#     # 每页n条
#     for ii in range(actItems.length):
#         o = pq(actItems[ii])
#         date = o.find('span').text().replace("<", "").replace(">", "").replace("/", "-").replace(" ", "")
#         title = o.find("a").text()
#         detailUrl = "http://kj.sdbi.com.cn/" + o.find("a").attr('href').replace('..', '')
#         # 关键字筛选
#         keywordhandle(keyWords, title, date, detailUrl, newList)
# insandexclude(collectionName, newList)
# 山东商务职业学院 end

# 山东轻工职业学院 begin
collectionName = 'sdqinggongzhiye'
newList = []
url = "http://gongshang.sdlivc.com/"
succ = 'jxky.htm'

for i in range(3):
    url = url + succ
    page = urllib.request.urlopen(url)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)

    succ = doc.find("span.p_next>a").attr('href')
    if succ.find("jxky") == -1:
        url = "http://gongshang.sdlivc.com/jxky/"
    else:
        url = "http://gongshang.sdlivc.com/"
    actItems = doc.find("ul.list>li")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.find('span').text()
        title = o.find("a").text()
        title = title.replace(date, "")
        detailUrl = "http://gongshang.sdlivc.com" + o.find("a").attr('href').replace('..', '')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 山东轻工职业学院 end

# 山东工业职业学院 begin
# 分页错误只查询第一页
collectionName = 'sdgongyezhiye'
newList = []
url = "http://gs.sdivc.edu.cn/"
succ = 'xbdt.htm'
url = url + succ
page = urllib.request.urlopen(url)
htmlcode = page.read()
htmlcode = htmlcode.decode('utf-8')
doc = pq(htmlcode)

actItems = doc.find("div.main_conRCb>ul>li")
# 每页n条
for ii in range(actItems.length):
    o = pq(actItems[ii])
    date = o.find('span').text()
    title = o.find("a").text()
    title = title.replace(date, "")
    detailUrl = "http://gs.sdivc.edu.cn/" + o.find("a").attr('href')
    # 关键字筛选
    keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 山东工业职业学院 end

# 山东外贸职业学院 begin
collectionName = 'sdwaimaozhiye'
newList = []
url = "http://ckjrx.sdwm.edu.cn/index/"
succ = 'cjxw.htm'

for i in range(3):
    url = url + succ
    page = urllib.request.urlopen(url)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)

    succ = doc.find("a.Next").attr('href')
    if succ.find("cjxw") == -1:
        url = "http://ckjrx.sdwm.edu.cn/index/cjxw/"
    else:
        url = "http://ckjrx.sdwm.edu.cn/index/"
    actItems = doc.find("div.main_conRCb>ul>li")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.find('span').text()
        title = o.find("a").text()
        title = title.replace(date, "")
        detailUrl = "http://ckjrx.sdwm.edu.cn/" + o.find("a").attr('href').replace('..', '')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 山东外贸职业学院 end


# 山东工程职业技术学院 begin
collectionName = 'jngongchengzhiyejishu'
newList = []
url = "https://jjglx.jngcxy.edu.cn/"
succ = 'jxgz.htm'

for i in range(3):
    url = url + succ
    requests.packages.urllib3.disable_warnings()
    page = requests.get(url, verify=False)
    htmlcode = page.content
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)

    succ = doc.find("a.Next").attr('href')
    if succ.find("jxgz") == -1:
        url = "https://jjglx.jngcxy.edu.cn/jxgz/"
    else:
        url = "https://jjglx.jngcxy.edu.cn/"
    actItems = doc.find("span.timestyle59182")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.text().replace("/", "-")
        title = o.parent().prev().find("a").text()
        title = title.replace(date, "")
        detailUrl = "https://jjglx.jngcxy.edu.cn/" + o.parent().prev().find("a").attr('href').replace('..', '')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 山东工程职业技术学院 end

# 淄博职业学院 begin
collectionName = 'zibozhiye'
newList = []
url = "https://kjxy.zbvc.edu.cn/index/"
succ = 'xwkd.htm'

for i in range(7):
    url = url + succ
    requests.packages.urllib3.disable_warnings()
    page = requests.get(url, verify=False)
    htmlcode = page.content
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)

    succ = doc.find("a.Next").attr('href')
    if succ.find("xwkd") == -1:
        url = "https://kjxy.zbvc.edu.cn/index/xwkd/"
    else:
        url = "https://kjxy.zbvc.edu.cn/index/"
    actItems = doc.find("div.newslist>ul>li")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.find("span").text()
        title = o.find("a").text()
        title = title.replace(date, "")
        detailUrl = "https://kjxy.zbvc.edu.cn/" + o.find("a").attr('href').replace('../', '')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 淄博职业学院 end

# 山东经贸职业学院 begin
collectionName = 'sdjingmao'
newList = []
url = "http://www.sdecu.com/html/xwzx/xbdt"
succ = ''
headersx = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/34.0.1847.137 Safari/537.36 LBBROWSER'}
for i in range(20):
    if i > 0:
        succ = "/" + str(i + 1) + ".html"
    url1 = url + succ
    req = urllib.request.Request(url=url1, headers=headersx)
    page = urllib.request.urlopen(req)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)

    actItems = doc.find("ul.list a[href]")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.parent().text().replace("【", "").replace("】", "")
        title = o.text()
        date = date.replace(title, "").replace(" ", "")
        detailUrl = o.attr("href")
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 山东经贸职业学院 end

# 山东圣翰财贸职业学院 begin
collectionName = 'sdshenghan'
newList = []
url = "http://www.suu.com.cn/news/yuanbushezhi/caimaoxueyuan/jiaoxueguanli/list_159_"
for i in range(20):
    url1 = url + str(i + 1) + ".html"
    page = urllib.request.urlopen(url1)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('GBK')
    doc = pq(htmlcode)

    actItems = doc.find("div.gjc_title")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.find("span").text()
        title = o.find("a").text()
        detailUrl = "http://www.suu.com.cn" + o.find("a").attr("href")
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 山东圣翰财贸职业学院 end

# 潍坊工商职业 begin
collectionName = 'weifanggongshangzhiye'
newList = []
url = "http://cwkj.wfgsxy.com/"
succ = 'xwzx.htm'

for i in range(1):
    url = url + succ
    page = urllib.request.urlopen(url)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)

    succ = doc.find("a.Next").attr('href')
    if succ.find("xwzx") == -1:
        url = "http://cwkj.wfgsxy.com/xwzx/"
    else:
        url = "http://cwkj.wfgsxy.com/"
    actItems = doc.find("div.main_conRCb>ul>li")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.find("span").text()
        title = o.find("em").text()
        detailUrl = "http://cwkj.wfgsxy.com/" + o.find("a").attr('href').replace('../', '')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 潍坊工商职业 end

# 山东理工职业 begin
collectionName = 'sdligongzhiye'
newList = []
url = "http://jrkj.sdpu.edu.cn/list2.jsp?a6t=17&a6p="
succ = "&a6c=15&urltype=tree.TreeTempUrl&wbtreeid=1012"
for i in range(3):
    url1 = url + str(i + 1) + succ
    page = urllib.request.urlopen(url1)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)
    actItems = doc.find("td.timestyle16375")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.text().replace("/", "-")
        title = o.prev().find("a").text()
        detailUrl = "http://jrkj.sdpu.edu.cn" + o.prev().find("a").attr('href')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 山东理工职业 end.

# 德州职业技术学院 begin
collectionName = 'dezhouzhiye'
newList = []
url = "http://jjglx.dzvc.edu.cn/xwdt/"
succ = "index.html"
for i in range(2):
    if i > 0:
        succ = str(i + 1) + ".html"
    url1 = url + succ
    page = urllib.request.urlopen(url1)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('GBK')
    doc = pq(htmlcode)
    actItems = doc.find("div.list_xb>ul>li")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.find("span.rt").text().split(" ")[0]
        title = o.find("a").text()
        detailUrl = o.find("a").attr('href')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 德州职业技术学院 end.

# 东营科技职业 begin
collectionName = 'dongyingkejizhiye'
newList = []
url = "http://www.dycollege.net/dept/newsList?type=news&did=7&p="
for i in range(3):
    succ = str(i + 1)
    url1 = url + succ
    page = urllib.request.urlopen(url1)
    htmlcode = page.read()
    htmlcode = htmlcode.decode("utf-8")
    doc = pq(htmlcode)
    actItems = doc.find("table.columnStyle")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.find("td[align='right']").text().split(" ")[0]
        title = o.find("a").text()
        detailUrl = "http://www.dycollege.net" + o.find("a").attr('href')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 东营科技职业 end.

# 烟台工程职业技术学院 begin
collectionName = 'yantaigongchengzhiye'
newList = []
url = "https://jg.ytetc.edu.cn/news/221/"
for i in range(1):
    succ = str(i + 1) + ".html"
    url1 = url + succ
    page = urllib.request.urlopen(url1)
    htmlcode = page.read()
    htmlcode = htmlcode.decode("utf-8")
    doc = pq(htmlcode)
    actItems = doc.find("span.column-news-title")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.find("span.fr").text()
        title = o.find("span.title").text()
        detailUrl = "https://jg.ytetc.edu.cn" + o.find("a").attr('href')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 烟台工程职业技术学院 end.

# 枣庄职业 begin
collectionName = 'zaozhuangzhiye'
newList = []
url = "http://jxx.sdzzvc.edu.cn/news_list.asp?pageno=1&pagesize=15&edition=&id=&Ntype=%BD%CC%D3%FD%BD%CC%D1%A7&lanmuName=%BD%CC%D1%A7%B6%AF%CC%AC"
page = urllib.request.urlopen(url)
htmlcode = page.read()
htmlcode = htmlcode.decode("GBK")
doc = pq(htmlcode)
actItems = doc.find("div.mid11").find("div.midn-n11 a[title]")
# 每页n条
for ii in range(actItems.length):
    o = pq(actItems[ii])
    date = o.parent().next().text().replace("年", "-").replace("月", "-").replace("日", "-")
    title = o.text()
    detailUrl = "http://jxx.sdzzvc.edu.cn" + o.attr('href')
    # 关键字筛选
    keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 枣庄职业 end.

# 临沂职业学院 begin
collectionName = 'linyizhiye'
newList = []
url = "http://kjjrxy.lyvc.edu.cn/xydt/"
succ = 'xydt.htm'

for i in range(7):
    url = url + succ
    page = urllib.request.urlopen(url)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)

    succ = doc.find("span.p_next a").attr('href')
    if succ.find("xydt") == -1:
        url = "http://kjjrxy.lyvc.edu.cn/xydt/xydt/"
    else:
        url = "http://kjjrxy.lyvc.edu.cn/xydt/"
    actItems = doc.find("ul.ss>li")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.find("span").text()
        title = o.find("a").text()
        detailUrl = "http://kjjrxy.lyvc.edu.cn/" + o.find("a").attr('href').replace('../', '')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 临沂职业学院 end

# 威海海洋职业学院 begin
collectionName = 'weihaihaiyangzhiye'
newList = []
url = "http://www.whovc.cn/OceanJingJi/JingJi_ListNews.aspx?CId1=0&CId2=114&page="

for i in range(2):
    url = url + str(i+1)
    page = urllib.request.urlopen(url)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)
    actItems = doc.find("div.xinwen_content>ul>li")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.find("span.time").text()
        title = o.find("a").text()
        detailUrl = "http://www.whovc.cn/" + o.find("a").attr('href')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 威海海洋职业学院 end

# 烟台黄金 begin
collectionName = 'yantaihuangjin'
newList = []
url = "http://xb.ytgc.edu.cn/ythjjjx/list.jsp?a6t=2&a6p="
succ = "&a6c=10&urltype=tree.TreeTempUrl&wbtreeid=1005"
for i in range(2):
    url1 = url + str(i + 1) + succ
    page = urllib.request.urlopen(url1)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)
    actItems = doc.find("td.timestyle186083")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.text()
        title = o.prev().find("a").text()
        detailUrl = "http://xb.ytgc.edu.cn/" + o.prev().find("a").attr('href')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)
# 烟台黄金 end.

# 资讯整理 Begin

# 烟台黄金 begin
collectionName = 'yantaihuangjin'
newList = []
url = "http://xb.ytgc.edu.cn/ythjjjx/list.jsp?a6t=2&a6p="
succ = "&a6c=10&urltype=tree.TreeTempUrl&wbtreeid=1005"
for i in range(2):
    url1 = url + str(i + 1) + succ
    page = urllib.request.urlopen(url1)
    htmlcode = page.read()
    htmlcode = htmlcode.decode('utf-8')
    doc = pq(htmlcode)
    actItems = doc.find("td.timestyle186083")
    # 每页n条
    for ii in range(actItems.length):
        o = pq(actItems[ii])
        date = o.text()
        title = o.prev().find("a").text()
        detailUrl = "http://xb.ytgc.edu.cn/" + o.prev().find("a").attr('href')
        # 关键字筛选
        keywordhandle(keyWords, title, date, detailUrl, newList)
insandexclude(collectionName, newList)

# 资讯整理 End.

# # 创建表格
excelTabel= xlwt.Workbook()
sheet1 = excelTabel.add_sheet('竞品', cell_overwrite_ok=True)
sheet1.write(0, 2, '日期')
sheet1.write(0, 1, '标题')
sheet1.write(0, 0, '所属机构')
sheet1.write(0, 3, '浏览地址')

sheet2 = excelTabel.add_sheet('企业', cell_overwrite_ok=True)
sheet2.write(0, 2, '日期')
sheet2.write(0, 1, '标题')
sheet2.write(0, 0, '所属机构')
sheet2.write(0, 3, '浏览地址')

sheet3 = excelTabel.add_sheet('院校', cell_overwrite_ok=True)
sheet3.write(0, 2, '日期')
sheet3.write(0, 1, '标题')
sheet3.write(0, 0, '所属机构')
sheet3.write(0, 3, '浏览地址')


collist = mydb. list_collection_names()

# 导出时的查询条件
cond = {'isSend': 0}

index = 1
# 生成资讯文件
for collectionKey in jingpinMap1:
    if collectionKey in collist:
        for obj in mydb[collectionKey].find(cond):
            sheet1.write(index, 2, obj['date'])
            sheet1.write(index, 1, obj['title'])
            sheet1.write(index, 0, jingpinMap1[collectionKey])
            sheet1.write(index, 3, obj['url'])
            index += 1
            mydb[collectionKey].update_one({'title': obj['title']}, {"$set": {"isSend": "1"}})
    if index == 1:
        sheet1.write(index, 0, "暂无新资讯")

index = 1
# 生成资讯文件
for collectionKey in jingpinMap2:
    if collectionKey in collist:
        for obj in mydb[collectionKey].find(cond):
            sheet2.write(index, 2, obj['date'])
            sheet2.write(index, 1, obj['title'])
            sheet2.write(index, 0, jingpinMap2[collectionKey])
            sheet2.write(index, 3, obj['url'])
            index += 1
            mydb[collectionKey].update_one({'title': obj['title']}, {"$set": {"isSend": "1"}})
    if index == 1:
        sheet2.write(index, 0, "暂无新资讯")

index = 1
# 生成资讯文件
for collectionKey in jingpinMap3:
    if collectionKey in collist:
        for obj in mydb[collectionKey].find(cond):
            sheet3.write(index, 2, obj['date'])
            sheet3.write(index, 1, obj['title'])
            sheet3.write(index, 0, jingpinMap3[collectionKey])
            sheet3.write(index, 3, obj['url'])
            index += 1
            mydb[collectionKey].update_one({'title': obj['title']}, {"$set": {"isSend": "1"}})
    if index == 1:
        sheet3.write(index, 0, "暂无新资讯")

fileName = time.strftime("%Y-%m-%d", time.localtime()) + 'Messages.xls'
excelTabel.save(filePath + fileName)

#
# 自动发送邮件
sender = 'onlineeval_inspur@163.com'
to_reciver = ['2681666570@qq.com']
cc_reciver = ['onlineeval_inspur@163.com']
reciver = to_reciver + cc_reciver
#receiver = '2681666570@qq.com'  # 接收邮箱
# 创建一个带附件的实例
message = MIMEMultipart()
message['From'] = "Harland<onlineeval_inspur@163.com>"
message['To'] = ";".join(to_reciver)
message['Cc'] = ";".join(cc_reciver)
message['Subject'] = Header(time.strftime("%Y-%m-%d", time.localtime()) + '竞品及院校资讯', 'utf-8')

# 邮件正文内容
message.attach(MIMEText(time.strftime("%Y-%m-%d", time.localtime()) + '竞品及院校资讯', 'plain', 'utf-8'))

# 构造附件1，传送当前目录下的 test.txt 文件
att1 = MIMEApplication(open(filePath + fileName, 'rb').read())
att1["Content-Type"] = 'application/octet-stream'
# 这里的filename可以任意写，写什么名字，邮件中显示什么名字
att1["Content-Disposition"] = 'attachment; filename="' + fileName + '"'
message.attach(att1)

smtpObj = smtplib.SMTP()
smtpObj.connect("smtp.163.com", 25)
smtpObj.login(sender, "masterkey1107")
smtpObj.sendmail(sender, reciver, message.as_string())
print("邮件发送成功")
smtpObj.quit()

