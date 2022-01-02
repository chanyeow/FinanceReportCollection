import openpyxl
import os
import time
import requests
import os.path
import json

from pathlib import Path
from util.util import standardize_dir
from util.util import get_all_stock_code

RESPONSE_TIMEOUT = 10

URL= 'http://90.push2his.eastmoney.com/api/qt/stock/kline/get'
HEADER = {
    #'User-Agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36',
    #'X-Requested-With': 'XMLHttpRequest'
    'User-Agent: Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.25 Safari/537.36 Core/1.70.3878.400 QQBrowser/10.8.4518.400',
    'X-Requested-With: XMLHttpRequest'
}

QUERY = {
    #'cb' : 'jQuery112402816109881259974_1638893884237',
    'secid' : '0.000001',
    'ut' : 'fa5fd1943c7b386f172d6893dbfba10b',
    'fields1' : 'f1,f2,f3,f4,f5,f6',
    'fields2' : 'f51,f52,f53,f54,f55,f56,f57,f58,f59,f60,f61',
    'klt' : 103,
    'fqt' : 0,
    'beg' : '20170101',
    'end' : '20500101',
    'smplmt' : 460,
    'lmt' : 1000000,
    '_' : '1638893884280',
}

def get_code_secid(code):
    if code != '':
        if code[0] == '6':
            marketid = 1
        else:
            marketid = 0
        return '%d.%s'%(marketid, code)
    else:
        return ''

PROXIES = {
    # "http": "http://101.132.189.87:9090", 
    "http": "http://101.133.239.96:8080", 
}

def spidering(secid):
    QUERY['secid'] = secid
    try:
        # r = requests.post(URL, QUERY, HEADER, proxies=PROXIES, timeout=RESPONSE_TIMEOUT)
        r = requests.post(URL, QUERY, HEADER, timeout=RESPONSE_TIMEOUT)
    except requests.exceptions.RequestException as e:
        print(e)

    r.encoding = 'utf-8'
    if r.status_code == requests.codes.ok and r.text != '':
        # data = r.json()
        # data =json.loads(r.text).get('data')
        data = r.text
    else:
        data = {}
    try:
        r.close()
    except Exception as e:
        print(e)
    return data

def main(filePath):
    mData = get_all_stock_code()

    iTotalCount = len(mData)
    iIndex = 0
    for code, name in mData.items():
        iIndex += 1

        fileName = filePath + str(code)
        file = Path(fileName)
        if file.is_file():
            print("正在爬取 %s - %s 的数据(%d/%d)"%(code, name, iIndex, iTotalCount))
            continue

        secid = get_code_secid(code)
        allData = spidering(secid)

        with open(fileName, 'w') as f:
            f.write(allData)
            f.close()

        print("正在爬取 %s - %s 的数据(%d/%d)"%(code, name, iIndex, iTotalCount))
        time.sleep(0.01)

    print("全部数据爬取完毕, 一共 %d 个(%s)"%(iTotalCount, filePath))

def loadDataFromFileTest(code, path):
    fileName = path + str(code)
    file = open(fileName, 'r')
    data = file.read()
    dict = json.loads(data).get("data")
    print(dict)
    file.close()

if __name__ == '__main__':
    filePath = os.getcwd() + '/' + "allStockKline" + '/'
    if not os.path.exists(filePath):
        os.makedirs(filePath)

    filePath = standardize_dir(filePath)
    main(filePath)
    loadDataFromFileTest('000001', filePath)
