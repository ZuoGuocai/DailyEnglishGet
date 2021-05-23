######################################################################################
#
# Author: zuoguocai@126.com
#
# Function: Get daily english TO DOCX file Version 2.0
#
# Modified Time:  2021年4月11日
# 
# Help:  need install  requests,lxml and python-docx
#        pip3 install requests lxml python-docx -i https://pypi.douban.com/simple
#
######################################################################################
#!/usr/bin/env python



import requests

# 忽略请求https，客户端没有证书警告信息
requests.packages.urllib3.disable_warnings()


from lxml import etree


from docx import Document
from docx.shared import Inches

import time
import random

import concurrent.futures

url = "https://bj.wendu.com/zixun/yingyu/6697.html"

headers = {
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.114 Safari/537.36'
}



# 获取表格中 所有a标签链接

r = requests.get(url,headers=headers,verify=False,timeout=120)
html = etree.HTML(r.text)


html_urls = html.xpath("//tr//a/@href")

num_count = len(html_urls)

print("总共发现:   {}句".format(num_count))



# 获取链接下内容, 去除广告内容

def craw(url):

  r = requests.get(url, headers=headers, verify=False,timeout=120)
  
  result_html = etree.HTML(r.content, parser=etree.HTMLParser(encoding='utf8'))

  html_data = result_html.xpath('//div[@class="article-body"]/p//text()')
  
  # 获取标题
  head = html_data[1]
  
  # 句子和问题
  juzi = '\n'.join(html_data[2:4])
  # 选项
  xuanxiang = '\n'.join(html_data[4:10])
  # 分析
  fengxi = '\n'.join(html_data[10:-4])
  
  # 合并为一篇内容，中间以换行符分隔
  content = "\n\n\n".join((head,juzi,xuanxiang,fengxi))
  
  file_name = 'C:\\Users\\zuoguocai\\Desktop\\pachong\\docs\\' + head + '.docx'

  # 写入word文档 
  document = Document()
  paragraph = document.add_paragraph(content)
  document.save(file_name)
  
  
  Message = "正在处理===>" + head  + "  "+ url + "  处理完成..."
  return Message
  

  
  
# 使用线程池加速IO操作, 缺点:可能因为网络问题或者网站限制，导致出现空文件
with concurrent.futures.ThreadPoolExecutor() as pool:
    futures =  [ pool.submit(craw,url) for url in html_urls ]
    for future in futures:
        print(future.result())
  
  
  
  
 



