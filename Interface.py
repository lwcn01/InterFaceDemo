#!/usr/bin/env python
# -*- coding: utf-8 -*-
__author__ = 'new'

import os,sys
import re
import xdrlib
import xlrd
import json
import requests
import smtplib
import configparser
from email import encoders
from email.header import Header
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr

def tc_path(path):
    if os.path.exists(path.strip()):
        if os.path.isfile(path.strip()):
            return path.strip()
    else:
        return False

def logger(path):
    import logging
    mylog = logging.getLogger('')
    log_file = os.path.join(os.path.dirname(path),'test.log')
    console = logging.StreamHandler()
    console.setLevel(logging.DEBUG)
    log_format = '[%(asctime)s] [%(levelname)s] %(message)s'     #配置log格式
    logging.basicConfig(format=log_format, filename=log_file, filemode='w', level=logging.DEBUG)
    formatter = logging.Formatter(log_format)
    console.setFormatter(formatter)
    mylog.addHandler(console)

    return mylog

# 通过index读取excel中数据
def read_excel(path):
    global mylog
    try:
        file_path = tc_path(path)
        mylog = logger(path)
    except Exception as e:
        mylog.error(str(e))
        sys.exit()
    else:
        tc = xlrd.open_workbook(file_path)
        table = tc.sheet_by_index(0)   #获取excell文件中第一张表格sheet1
        #table = data.sheets()[by_index] #by_name
        rows = table.nrows  #返回excel行数，整型
        cols = table.ncols  #返回excel列数，整型
        colnames = table.row_values(0) #以列表形式返回第一行的所有value值

        #将每一行数据以字典返回，将每一行字段组成的字典添加生成一个list
        xls = []
        for r in range(1,rows):
            xls_row = {}
            encryption = table.cell(r,8).value             #加密字段          
            for c in range(cols):
                if encryption == 'MD5':           #如果数据采用md5加密，便先将数据加密
                    if c == 7:
                        request_data = table.cell(r,7).value           #请求参数
                        request_data = json.loads(request_data)
                        request_data['pwd'] = md5Encode(request_data['pwd'])
                        xls_row[colnames[7]] = request_data
                    else:
                        table_value = table.cell(r,c).value  
                        xls_row.setdefault(colnames[c],table_value)
                else: 
                    xls_row.setdefault(colnames[c],table.cell(r,c).value)
            xls.append(xls_row)
           
        return xls            
      
def apicall(method,url,params):
    s = requests.Session()
    headers = { 'Content-Type' : 'text/html; charset=UTF-8',
                'X-Requested-With' : 'XMLHttpRequest',
                'Connection' : 'keep-alive',
                'Referer' : 'http://192.168.1.3:8080',
                'User-Agent' : 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML,\
                 like Gecko) Chrome/47.0.2526.111 Safari/537.36'
                }
    if method == 'GET':
        if params != '':
            page = s.get(url,params,timeout = 0.01)
        else:
            page = s.get(url,timeout = 0.01)
    if method == 'POST':
        if params != '':
            page = s.post(url,params,headers,timeout = 0.01)
  
    #page_info = json.dumps(page.text)  #把一个Python对象编码转换成Json字符串 

    return page

def get_page():
    xls = read_excel(path)
    
    for row_xls in xls:
        if row_xls in xls:
            url = "http://" + row_xls["API Host"] + ":" + str(int(row_xls["Port"]))
            method = row_xls["Request Method"]
            params = row_xls["Request Data"]
            try:
                result = apicall(method,url,params)
            except Exception as e:
                mylog.error(row_xls["NO."] + ' ' + row_xls["API Purpose"] + " : " + str(e))
            else:
                status = result.status_code
                errorCase = []                #用于保存接口返回的内容和HTTP状态码
                if status == 200 :
                    if re.search(row_xls["Check Point"], str(result.text)):
                        mylog.info(row_xls["NO."] + ' ' + row_xls["API Purpose"] + ' Suceuss，' + str(status) + ', ' + str(result.text))
                        
                    else:
                        mylog.error(row_xls["NO."] + ' ' + row_xls["API Purpose"] + ' Failed，' + str(status) + ', ' + str(result.text))
                        errorCase.append((row_xls["NO."] + ' ' + row_xls["API Purpose"],  str(status), 'http://'+ row_xls["API Host"] + row_xls["Request Address"], result.text))
                        return errorCase
                else:
                    mylog.error(row_xls["NO."] + ' ' + row_xls["API Purpose"] + ' Failed，' + str(status) + ', ' + str(result.text))
                    errorCase.append((row_xls["NO."] + ' ' + row_xls["API Purpose"], str(status), 'http://'+ row_xls["API Host"] + row_xls["Request Address"], result.text))
                    return errorCase
                    
def mail_conf():
    #通过自带的configparser模块，读取发送邮件的配置文件，作为字典返回
    conf_file = configparser.RawConfigParser()
    try:
      conf_file.read(os.path.join(os.path.dirname(path),'Config.conf'))
    except:
      mylog.error("conf.ini not found")
    else:
      conf_info  = []
      for i in conf_file.sections():   #返回所有的section
          #conf_file.options(i)        #以列表返回一个section中的所有key
          mail = {}
          for j in conf_file.items(i):
            mail[j[0]] = j[1]
          conf_info.append(mail)
      
      if len(conf_file.sections()) == 1:
          return conf_info[0]
      else:
          return conf_info             #以列表返回各section信息
            
    #conf['sender'] = conf_file.get("email","sender")  #获取指定section的option 信息
def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(), addr))

def send_mail(text):         
    #SMTP发送邮件
    conf_info =  mail_conf()
    msg = MIMEText(text,'html','utf-8')
    msg['From'] = _format_addr('new <%s>' % conf_info['sender'])
    msg['To'] = _format_addr('<%s>' % conf_info['receiver'])
    msg['Subject'] = Header('接口自动化测试报告','utf-8').encode()
    server = smtplib.SMTP(conf_info['smtpserver'],25)
    server.set_debuglevel(1)
    server.login(conf_info['uername'], conf_info['password'])
    server.sendmail(conf_info['sender'], [conf_info['receiver']], msg.as_string())
    server.quit()

def main():
    global path
    path = "G:\\job\\work\\jk.xlsx"
    xls = read_excel(path)
    test_result = get_page()
    if len(test_result) > 0:
        html = '<html><body>接口自动化扫描，共 ' + str(len(xls)) + ' 条接口测试用例，其中 ' + str(len(test_result)) + ' 个接口异常，列表如下：' + '</p><table><tr><th style="width:100px;text-align:left">接口</th><th style="width:50px;text-align:left">状态</th><th style="width:200px;text-align:left">接口地址</th><th   style="text-align:left">接口返回值</th></tr>'
        for test in test_result:
            print(test)
            html = html + '<tr><td style="text-align:left">' + test[0] + '</td><td style="text-align:left">' + test[1] + '</td><td style="text-align:left">' + test[2] + '</td><td style="text-align:left">' + test[3] + '</td></tr>'
        send_mail(html)

if __name__ == '__main__':
    main()