import os
import urllib.request as urllib
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import os

def make_format(url, fileNM):
    keylist = ['갱신요구권사용','건축년도','계약구분','계약기간','년', '법정동','보증액','아파트','월','일','월세금액','전용면적','종전계약보증금','종전계약월세','지번','지역코드','층']
    res = urllib.urlopen(url)
    result = res.read()
    soup = BeautifulSoup(result, 'lxml-xml')
    if True:
        items = soup.findAll('item')
        itemList = []
        for v in items:
            item = {}
            for key in keylist:
                try:
                    item[key] = v.find(key).text
                except:
                    continue
            itemList.append(item)
        writer = pd.ExcelWriter('crawl_data_2/'+fileNM, engine='xlsxwriter')
        df = pd.DataFrame(columns=keylist)
        for idx, row in enumerate(itemList):
            df.loc[idx] = row
        df.to_excel(writer)
        writer.save()
        print(f'MAKE({len(itemList)}) || {fileNM}')


def get_data_all(serviceKey, start_yymm, region_cd_dict, numrow):
    for i in range(1000):
        yymm = (datetime.strptime(start_yymm, '%Y%m') - relativedelta(months=i)).strftime('%Y%m')
        for region in region_cd_dict.keys():
            url = f'http://openapi.molit.go.kr:8081/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcAptRent?LAWD_CD={region_cd_dict[region]}&DEAL_YMD={yymm}&serviceKey={serviceKey}&numOfRows={numrow}'
            fileNM = f'{yymm}_{region}({region_cd_dict[region]}).xlsx'
            if fileNM not in os.listdir('crawl_data_2/'):
                make_format(url, fileNM)
            else:
                if i!=0 and len(pd.read_excel(f'crawl_data_2/{fileNM}'))<1:
                    make_format(url, fileNM)
                else:
                    print(f'EXIST || {fileNM}')


serviceKey = 'ojb2HfXlX3ufbp9fOwLMCXOCYljZo9GydOnoG6cGJD3A7caoxtSatynePR5JvP11Gg%2FUQVSwq0LR1M%2BEviWgGQ%3D%3D'
region_cd = pd.read_excel('region_code5.xlsx').drop_duplicates(subset=['code5'])
region_cd_dict = region_cd.set_index('region')['code5'].to_dict()
numrow = 1000
start_yymm = '201403'
get_data_all(serviceKey, start_yymm, region_cd_dict, numrow)

print(1)