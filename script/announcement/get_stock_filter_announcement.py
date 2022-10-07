'''
Author: chanyeow && chanyeow1995@gmail.com
Date: 2022-10-07 15:57:55
LastEditors: chanyeow 767138730@qq.com
LastEditTime: 2022-10-07 21:25:16
Description: 根据特殊字眼，筛选公告选股
Copyright (c) 2022 by chanyeow 767138730@qq.com, All Rights Reserved. 
'''

import openpyxl
import os
import time
import requests
import os.path
import json
import sys
sys.path.append(os.getcwd() + "/script/util")
import util

from pathlib import Path
from util import get_all_stock_code
from util import get_marketcode
from util import get_market_id

#const
FILE_NAME_PDFURL = '/financeData/Announcement_PdfUrl.xlsx'
MAX_PAGE_COUNT = 1000

TITLE_LIST = ['short_name', 'stock_code', 'art_code', 'notice_date']

KEY_WORDS = [
    "重组",
    "发行股份",
    "支付现金",
    "购买资产",
]

URL= 'https://np-anotice-stock.eastmoney.com/api/security/ann'
HEADER = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36',
    # 'X-Requested-With': 'XMLHttpRequest'
}

DATA = {
    # 'cb' : 'jQuery11230933578843087203_1665129761622',
    'sr' : -1,
    'page_size' : 50,
    'page_index' : 1,
    'ann_type' : 'SHA,CYB,SZA,BJA',
    'client_source' : 'web',
    'f_node' : 0,
    's_node' : 0,
}

# https://np-anotice-stock.eastmoney.com/api/security/ann?
# cb=jQuery11230933578843087203_1665129761622&sr=-1&page_size=50&page_index=2&ann_type=SHA,CYB,SZA,BJA&client_source=web&f_node=0&s_node=0

def spideringPdfUrl(pageIndex):
    DATA['page_index'] = pageIndex
    ret = False
    try:
        r = requests.get(URL, headers=HEADER, params=DATA)
    except requests.exceptions.RequestException as e:
        print(e)
    r.encoding = 'UTF-8'
    if r.status_code == requests.codes.ok and r.text != '':
        data = r.json()
        ret = True
    else:
        data = {}
        ret = False
    try:
        r.close()
    except Exception as e:
        print(e)
    return data["data"]["list"], ret

def getPdfUrl():
    sFileName = os.getcwd() + FILE_NAME_PDFURL
    file = Path(sFileName)
    if not file.is_file():
        wb = openpyxl.Workbook()
        wb.save(sFileName)
    wb = openpyxl.load_workbook(sFileName)
    allSheest = wb.get_sheet_names()
    ws = wb.get_sheet_by_name(allSheest[0])
    ws.append(TITLE_LIST)

    print('开始爬取公告链接...')
    for pageIndex in range(1, 2):
        print('尝试爬取第 %d/%d 页...'%(pageIndex, MAX_PAGE_COUNT))
        mData, bRet = spideringPdfUrl(pageIndex)
        if bRet:
            for cell in mData:
                stockData = []
                stockData.append(cell["codes"][0]["short_name"])
                stockData.append(cell["codes"][0]["stock_code"])
                stockData.append(cell["art_code"])
                stockData.append(cell["notice_date"])
                ws.append(stockData)

        time.sleep(0.1)

    wb.save(sFileName)
    wb.close
    print('爬取公告链接完毕...')

def main():
    getPdfUrl()

if __name__ == '__main__':
    main()



