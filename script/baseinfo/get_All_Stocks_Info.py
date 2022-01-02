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
from util import standardize_dir
from util import get_all_stock_code

RESPONSE_TIMEOUT = 10

URL= 'http://push2.eastmoney.com/api/qt/stock/get'
HEADER = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.25 Safari/537.36 Core/1.70.3878.400 QQBrowser/10.8.4518.400',
    'X-Requested-With': 'XMLHttpRequest'
}

QUERY = {
    #'cb' : 'jQuery112401545144085503931_1639893352029',
    'secid' : '0.000001',
    'ut' : 'fa5fd1943c7b386f172d6893dbfba10b',
    'invt' : 2,
    'fltt' : 2,
    'fields' : 'f43,f57,f58,f169,f170,f46,f44,f51,f168,f47,f164,f163,f116,f60,f45,f52,f50,f48,f167,f117,f71,f161,f49,f530,f135,f136,f137,f138,f139,f141,f142,f144,f145,f147,f148,f140,f143,f146,f149,f55,f62,f162,f92,f173,f104,f105,f84,f85,f183,f184,f185,f186,f187,f188,f189,f190,f191,f192,f107,f111,f86,f177,f78,f110,f260,f261,f262,f263,f264,f267,f268,f250,f251,f252,f253,f254,f255,f256,f257,f258,f266,f269,f270,f271,f273,f274,f275,f127,f199,f128,f193,f196,f194,f195,f197,f80,f280,f281,f282,f284,f285,f286,f287,f292,f293,f181,f294,f295,f279,f288',
    '_' : '1639893352185',
}

class Stock:
    def __init__(self, code, name, jsonData):
        self.code = code
        self.name = name
        self.jsonData = jsonData

    def get_All_info(self):
        return self.code, self.name, self.jsonData

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
    # except Exception as e:
    except requests.exceptions.RequestException as e:
        print(e)
    r.encoding = 'UTF-8'
    if r.status_code == requests.codes.ok and r.text != '':
        # data = r.json()
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
        # allData = filter_illegal_string(allData)

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
    filePath = os.getcwd() + '/' + "allStockInfo" + '/'
    filePath = standardize_dir(filePath)
    main(filePath)
    loadDataFromFileTest('000001', filePath)

