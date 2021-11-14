import requests
import random
import time
import pandas as pd

download_path = 'http://static.cninfo.com.cn/'
saving_path = 'D:/2020年报'

User_Agent = [
    ###这里自建一个User_Agent列表
]  # User_Agent的集合

headers = {'Accept': 'application/json, text/javascript, */*; q=0.01',
           "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
           "Accept-Encoding": "gzip, deflate",
           "Accept-Language": "zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7,zh-HK;q=0.6,zh-TW;q=0.5",
           'Host': 'www.cninfo.com.cn',
           'Origin': 'http://www.cninfo.com.cn',
           'Referer': 'http://www.cninfo.com.cn/new/commonUrl?url=disclosure/list/notice',
           'X-Requested-With': 'XMLHttpRequest'
           }

###巨潮要获取数据，需要ordid字段，具体post的形式是'stock':'证券代码,ordid;'
def get_orgid(Namelist):
    orglist = []
    url = 'http://www.cninfo.com.cn/new/information/topSearch/detailOfQuery'
    hd = {
        'Host': 'www.cninfo.com.cn',
        'Origin': 'http://www.cninfo.com.cn',
        'Pragma': 'no-cache',
        'Accept-Encoding': 'gzip,deflate',
        'Connection': 'keep-alive',
        'Content-Length': '70',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Accept': 'application/json,text/plain,*/*',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8'}
    for name in Namelist:
        data = {'keyWord': name,
                'maxSecNum': 10,
                'maxListNum': 5,
				}
        r = requests.post(url, headers=hd, data=data)
        org_id = r.json()['keyBoardList'][0]['orgId']
        #print(org_id+'****'+name)
        orglist.append(org_id)
    ##对列表去重
    formatlist = list(set(orglist))
    formatlist.sort(key=orglist.index)
    return formatlist


def single_page(stock):
    query_path = 'http://www.cninfo.com.cn/new/hisAnnouncement/query'
    headers['User-Agent'] = random.choice(User_Agent)  # 定义User_Agent
    print(stock)
    
    query = {'pageNum': 1,  # 页码
             'pageSize': 30,
             'tabName': 'fulltext',
             'column': 'szse',  
             'stock': stock,
             'searchkey': '',
             'secid': '',
             'plate': '',   
             'category': 'category_ndbg_szsh;',  # 年度报告
             'trade': '',   #行业
             'seDate': '2020-11-27~2021-05-28'  # 时间区间
             }
    namelist = requests.post(query_path, headers=headers, data=query)
    single_page = namelist.json()['announcements']
    print(len(single_page))
    return single_page  # json中的年度报告信息


def saving(single_page):  # 下载年报
    headers = {'Host': 'static.cninfo.com.cn',
               'Connection': 'keep-alive',
               'Upgrade-Insecure-Requests': '1',
               'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36 Edg/90.0.818.66',
               'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
               'Accept-Encoding': 'gzip, deflate',
               'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
               'Cookie': 'routeId=.uc1'
               }
    for i in single_page:
        if ('2020年年度报告（更新后）' in i['announcementTitle'])  or ('2020年年度报告' in i['announcementTitle']) or ('2020年年度报告（修订版）' in i['announcementTitle']) :
            download = download_path + i["adjunctUrl"]
            name = i["secCode"] + '_' + i['secName'] + '_' + i['announcementTitle'] + '.pdf'
            file_path = saving_path + '/' + name
            print(file_path)
            time.sleep(random.random() * 2)
            headers['User-Agent'] = random.choice(User_Agent)
            r = requests.get(download, headers=headers)
            time.sleep(10)
            print(r.status_code)
            f = open(file_path, "wb")
            f.write(r.content)
            f.close()
        else:
            continue


if __name__ == '__main__':
    org_list = get_orgid(['日海智能'])

    # Sec = pd.read_excel('D:/dict.xlsx',dtype = {'code':'object'})  #读取excel,证券代码+证券简称
    # Seclist = list(Sec['code'])  #证券代码转换成list
    # Namelist = list(Sec['name'])
    # org_list = get_orgid(Namelist)
    # Sec['orgid'] = org_list
    # Sec.to_excel('D:/dict.xlsx',sheet_name='sheet-2',index=False)
    # stock = ''
    # count = 0
    # ##按行遍历
    # for rows in Sec.iterrows():
    #     stock = str(rows[1]['code'])+','+str(rows[1]['orgid'])+';'
    #     try:
    #         page_data = single_page(stock)
    #     except :
    #         print('page error, retrying')
    #         try:
    #             page_data = single_page(stock)
    #         except:
    #             print('page error!') 
    #     saving(page_data)
    #     count = count + 1
    # print('共有',count,'家券商')

