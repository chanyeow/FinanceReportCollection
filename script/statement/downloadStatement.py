'''
Author: chanyeow && chanyeow1995@gmail.com
Date: 2022-01-02 16:14:35
LastEditors: chanyeow 767138730@qq.com
LastEditTime: 2022-10-06 14:20:19
Description: 下载全部股票的基础信息，保存在./financeData下
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

RESPONSE_TIMEOUT = 10

URL= 'http://emweb.securities.eastmoney.com/PC_HSF10/NewFinanceAnalysis/ZYZBAjaxNew'
HEADER = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.25 Safari/537.36 Core/1.70.3878.400 QQBrowser/10.8.4518.400',
    'X-Requested-With': 'XMLHttpRequest'
}

DATA = {
    'type' : '0',
    'code' : 'SH600780',
}

def spidering(marketId, code):
    DATA['type'] = marketId
    DATA["code"] = code
    ret = False
    try:
        # r = requests.post(URL, DATA, HEADER, timeout=RESPONSE_TIMEOUT)
        r = requests.get(URL,headers=HEADER,params=DATA)
    except requests.exceptions.RequestException as e:
        print(e)
    r.encoding = 'UTF-8'
    if r.status_code == requests.codes.ok and r.text != '':
        data = r.text
        ret = True
    else:
        data = {}
        ret = False
    try:
        r.close()
    except Exception as e:
        print(e)
    return ret, data
    
def main(filePath):
    mData = get_all_stock_code()

    iTotalCount = len(mData)
    if iTotalCount == 0:
        print("stock list is empty")
        return

    iIndex = 0
    vErrorCode = []
    for code, name in mData.items():
        iIndex += 1

        fileName = filePath + str(code)
        file = Path(fileName)
        if file.is_file():
            print("正在爬取 %s - %s 的数据(%d/%d)"%(code, name, iIndex, iTotalCount))
            continue

        marketId = get_market_id(code)
        marketCode = get_marketcode(code)
        ret, allData = spidering(marketId, marketCode)
        if ret:
            with open(fileName, 'w') as f:
                f.write(allData)
                f.close()   
        else:
            vErrorCode.append(code)    
        print("正在爬取 %s - %s 的数据(%d/%d)"%(code, name, iIndex, iTotalCount))
        time.sleep(0.01)
        
    print("全部数据爬取完毕, 一共 %d 个(%s)"%(iTotalCount, filePath))
    sErr = ""
    for i in vErrorCode:
        sErr = "%s , %s"%(sErr, vErrorCode[i])
    print(sErr)

def loadDataFromFileTest(code, path):
    fileName = path + str(code)
    file = open(fileName, 'r')
    data = file.read()
    dict = json.loads(data).get("data")
    print(dict[0]['SECUCODE'])
    file.close()

if __name__ == '__main__':
    filePath = os.getcwd() + "/financeData/allStockStatement" + '/'
    myPath = Path(filePath)
    if not myPath.exists():
        os.makedirs(myPath)

    main(filePath)
    loadDataFromFileTest('000001', filePath)

