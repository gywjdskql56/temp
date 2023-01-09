import os
import urllib
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import os

def make_format(url, fileNM):
    keylist = ['거래금액','거래유형','건축년도','년','도로명','도로명건물본번호코드','도로명건물부번호코드','도로명시군구코드','도로명일련번호코드','도로명코드',
        '법정동','법정동본번코드','법정동부번코드','법정동시군구코드','법정동읍면동코드','법정동지번코드','아파트',
        '월','일','일련번호','전용면적','중개사소재지','지번','지역코드','층','해제사유발생일','해제여부']
    res = urllib.request.urlopen(url)
    result = res.read()
    soup = BeautifulSoup(result, 'lxml-xml')
    if soup.find('resultMsg').text == 'NORMAL SERVICE.':
        print('EXCEED')
    else:
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
        writer = pd.ExcelWriter('crawl_data/'+fileNM, engine='xlsxwriter')
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
            url = f'http://openapi.molit.go.kr/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcAptTradeDev?LAWD_CD={region_cd_dict[region]}&DEAL_YMD={yymm}&serviceKey={serviceKey}&numOfRows={numrow}'
            fileNM = f'{yymm}_{region}({region_cd_dict[region]}).xlsx'
            if fileNM not in os.listdir('crawl_data/'):
                make_format(url, fileNM)
            else:
                if i!=0 and len(pd.read_excel(f'crawl_data/{fileNM}'))<1:
                    make_format(url, fileNM)
                else:
                    print(f'EXIST || {fileNM}')




serviceKey = 'ojb2HfXlX3ufbp9fOwLMCXOCYljZo9GydOnoG6cGJD3A7caoxtSatynePR5JvP11Gg%2FUQVSwq0LR1M%2BEviWgGQ%3D%3D'
region_cd = pd.read_excel('region_code5.xlsx').drop_duplicates(subset=['code5'])
region_cd_dict = region_cd.set_index('region')['code5'].to_dict()
numrow = 10000
start_yymm = '201001'
get_data_all(serviceKey, start_yymm, region_cd_dict, numrow)

# # start_yymm = '202301'
# # landcode = '11440'
# #
# # url = f'http://openapi.molit.go.kr/OpenAPI_ToolInstallPackage/service/rest/RTMSOBJSvc/getRTMSDataSvcAptTradeDev?LAWD_CD={landcode}&DEAL_YMD={yymm}&serviceKey={serviceKey}&numOfRows={numrow}'
#
#
# yymm = (datetime.strptime(start_yymm,'%Y%m') - relativedelta(months=1)).strftime('%Y%m')
# res = urllib.request.urlopen(url)
# result = res.read()
# soup = BeautifulSoup(result, 'lxml-xml')
# items = soup.findAll('item')
#
#
# keylist = [
#     '거래금액',
#     '거래유형',
#     '건축년도',
#     '년',
#     '도로명',
#     '도로명건물본번호코드',
#     '도로명건물부번호코드',
#     '도로명시군구코드',
#     '도로명일련번호코드',
#     '도로명코드',
#     '법정동',
#     '법정동본번코드',
#     '법정동부번코드',
#     '법정동시군구코드',
#     '법정동읍면동코드',
#     '법정동지번코드',
#     '아파트',
#     '월',
#     '일',
#     '일련번호',
#     '전용면적',
#     '중개사소재지',
#     '지번',
#     '지역코드',
#     '층',
#     '해제사유발생일',
#     '해제여부'
# ]
#
# itemList = []
# for v in items:
#     item = {}
#     for key in keylist:
#         item[key] = v.find(key).text
#     itemList.append(item)
#
# writer = pd.ExcelWriter(f'data_{yymm}_{landcode}_{numrow}.xlsx', engine='xlsxwriter')
# df = pd.DataFrame(columns=keylist)
# for idx,row in enumerate(itemList):
#     df.loc[idx] = row
# df.to_excel(writer)
# writer.save()
# print(f'data_{yymm}_{landcode}_{numrow}.xlsx')
# print(len(itemList))
print(1)