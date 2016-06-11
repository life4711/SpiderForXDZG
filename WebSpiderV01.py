#-*- coding: UTF-8 -*-
import os
from xlutils.copy import copy
import xlrd as ExcelRead
import urllib
import urllib2
import requests
import re
import codecs
import time
import xlwt

#作用就是将html文件里的一些标签去掉，只保留文字部分
class HTML_Tool:
    BgnCharToNoneRex = re.compile("(\t|\n| |<a.*?>|<img.*?>)")
    EndCharToNoneRex = re.compile("<.*?>")
    BgnPartRex = re.compile("<p.*?>")
    CharToNewLineRex = re.compile("<br/>|</p>|<tr>|<div>|</div>")
    CharToNextTabRex = re.compile("<td>")

    def Replace_Char(self,x):
        x=self.BgnCharToNoneRex.sub("",x)
        x=self.BgnPartRex.sub("\n   ",x)
        x=self.CharToNewLineRex.sub("\n",x)
        x=self.CharToNextTabRex.sub("\t",x)
        x=self.EndCharToNoneRex.sub("",x)
        return x

#爬虫类
class Spider:
    def __init__(self,headers,num_start,num_end,num1,num2):
        self.num_start = num_start
        self.num_end = num_end
        self.num1=num1
        self.num2=num2+1
        self.headers = headers
        self.s = requests.session()
        self.login = ""
        self.datas = []
        self.myTool = HTML_Tool()
        #s = time.strftime("%Y-%m-%d(%H_%M_%S)")
        self.file_name = 'XDZG.xls'
        print u'会刊网爬虫程序已启动，正在开始加载...'

    def Solve(self):
        self.get_data()
        print u'爬虫报告：会刊目录已保存至当前文件夹下"XDZG.xls"文件'
        print u'加载完成，按Enter键退出爬虫程序'
        raw_input()

    def wExcel(self):
        #print self.datas
        r_xls = ExcelRead.open_workbook(self.file_name)
        r_sheet = r_xls.sheet_by_index(0)
        rows = r_sheet.nrows
        w_xls = copy(r_xls)
        sheet_write = w_xls.get_sheet(0)
        for i in range(len(self.datas)):
            for j in range(len(self.datas[i])):
                sheet_write.write(rows + i, j, self.datas[i][j].decode('utf8'))
        w_xls.save(self.file_name)

    def get_data(self):
        for i in range(self.num1,self.num2):
            print u'爬虫报告，正在加载第%d页...' %i
            afterURL1 = "http://wmcrm.baidu.com/crm?qt=orderlist&order_status=0&start_time="
            afterURL2 = "&end_time="
            afterURL3 = "&pay_type=2&is_asap=0&page="
            response = self.s.get(afterURL1+self.num_start+afterURL2+self.num_end+afterURL3+str(i),headers = headers)
            mypage = response.content
            self.deal_data(mypage.decode('utf-8'))
            #self.wExcel()
            #print mypage
            #raw_input()


    def deal_data(self,mypage):
        myItems = re.findall('<div class="list-item".*?#.*?</span>(.*?)</div>.*?<div>(.*?)</div>.*?<div>(.*?)</div>.*?<div>(.*?)<span.*?class="info-div-div">(.*?)<span class="orderlist".*?class="info-div-div">(.*?)</div>.*?<div  class="info-div-div">(.*?)</div>.*?<div class="userAddr info-div-div common-list-item">(.*?)<span class="mapToken.*?class="info-div-margin">(.*?)</div>.*?<span class="menutotal">(.*?)</span>',mypage,re.S)
        #print myItems
        for item in myItems:
            items = []
            #print item[0]+' '+item[1]+' '+item[2]+' '+item[3]+' '+item[4]+' '+item[5]+' '+item[6]+' '+item[7]
            items.append(item[0].replace("\n",""))
            items.append(item[1].replace("\n",""))
            items.append(item[2].replace("\n", ""))
            items.append(item[3].replace("\n", ""))
            items.append(item[4].replace("\n", ""))
            items.append(item[5].replace("\n", ""))
            items.append(item[6].replace("\n", ""))
            items.append(item[7].replace("\n", ""))
            items.append(item[8].replace("\n", ""))
            items.append(item[9].replace("\n", ""))
            #print items
            mid = []
            for i in items:
                mid.append(self.myTool.Replace_Char(i.replace("\n","").encode('utf-8')))
            #print mid
            self.datas = []
            self.datas.append(mid)
            self.wExcel()
            #print self.datas


# def get_data():
#     s = requests.session()
#     afterURL = "http://wmcrm.baidu.com/crm?qt=neworderlist"
#     response = s.get(afterURL)
#     mypage = response.content
#     # self.deal_data(mypage.decode('utf-8'))
#     # self.wExcel()
#     print mypage

if __name__ == "__main__":
    # get_data()
    #----------- 程序的入口处 -----------
    print u"""
    *-----------------------------------------------------------------------------------------------
    * 程序：网络爬虫
    * 版本：V01
    * 作者：lvshubao
    * 日期：2016-06-03
    * 语言：Python 2.7
    * 功能：将小度掌柜的订单信息以文件追加写入的方式保存到当前目录下的"xdzg.xls"文件
    * 操作：按Enter键开始执行程序，加载时间可能会很长，请耐心等待
    * 提示：1、程序运行之初请保证当前目录下存在"xdzg.xls"文件，并保证没有其他程序正在使用该文件
    *       2、程序运行过程中需要连接互联网，期间允许中断，建议选取合适大小的时间段进行爬取
    *
    *-----------------------------------------------------------------------------------------------
    """
    raw_input()
    headers = {'Connection': 'keep-alive','Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8','Upgrade-Insecure-Requests': '1','User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36','Accept-Encoding': 'gzip, deflate, sdch', 'Accept-Language': 'zh-CN,zh;q=0.8','Cookie': 'BAIDUID=BA4B5E2D8A87B2E6C2FEAE0A2BA2A4AB:FG=1; BIDUPSID=BA4B5E2D8A87B2E6C2FEAE0A2BA2A4AB; PSTM=1462755380; BDUSS=xQd1hhN0l6cFJzZWJMWTVLeGtVUW52SnZ6RzE4eGtFMzRISGtWOUFueE5zVmhYQVFBQUFBJCQAAAAAAAAAAAEAAAAH2IY9tOfNwc3Q1vEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAE0kMVdNJDFXTm; plus_cv=0::m:1-nav:646c305e-hotword:10105b16; ULOG_UID=ioyyjftc901; WMUSS=BAAAtDmcqD3o7JT8uM1BNHmYRHxNtKExBKl47VlMhbQUdGxIFRCF-fkU-LXcPJBYXegAAeRZ1CzQIeFtTOQYxdFwFFkoKMmQbNyg7Nid%7Ec1N3A00Nd1hBf1AdUwoiBX1-ChNIE0UtLXtUaDF7cgbhAPoXjg8oSEEM1TX%7EDMCIUBGl5gEOIgEAANEaCCYAAME; new_order_time=1464913256; new_remind_time=1464913227; BDRCVFR[feWj1Vr5u3D]=I67x6TjHwwYf0; H_PS_PSSID=20144_1435_18240_17001_14971_12347_19372; _wmcrmpush=39165BD3781D43E8B116B79391DAF02E_1464918299_10_1616631776_a181045182820; newyear=open'}
    print u'请输入开始时间，格式：2016-05-03：'
    num_start = raw_input()
    print u'请输入结束时间，格式：2016-05-03：'
    num_end = raw_input()
    print u'请输入起始页数：'
    num1 = raw_input()
    print u'请输入结束页数：'
    num2 = raw_input()
    MySpider = Spider(headers,num_start,num_end,int(num1),int(num2))
    MySpider.Solve()