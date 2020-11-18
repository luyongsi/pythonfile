import sys
import json
import os
import urllib.request
import urllib.parse
import hashlib
import time
import xlwt
import xlrd 
from time import strftime, localtime

def getlocation(name,city):
    result={}
    
    try:
        query_uri='http://api.map.baidu.com'
        # 以get请求为例http://api.map.baidu.com/geocoder/v2/?address=百度大厦&output=json&ak=yourak
        query_str = '/place/v2/search?query='+name+'&tag=教育培训&region='+city+'&output=json&ak=AxRi1wkuhtAKLABnMGRk5XUNTzOyWGsx'

        # 对queryStr进行转码，safe内的保留字符不转换
        encoded_str = urllib.parse.quote(query_str, safe="/:=&?#+!$,;'@()*[]")

        # 在最后直接追加上yoursk
        raw_str = encoded_str + 'hHjd9tWMkgyB03UMmN23UzZ2jVoEVsaB'
        # md5计算出的sn值7de5a22212ffaa9e326444c75a58f9a0
        # 最终合法请求url是http://api.map.baidu.com/geocoder/v2/?address=百度大厦&output=json&ak=yourak&sn=7de5a22212ffaa9e326444c75a58f9a0

        query_sn = hashlib.md5(urllib.parse.quote_plus(raw_str).encode("utf8")).hexdigest()

        finalurl=query_uri+encoded_str+'&sn='+ query_sn

        res = urllib.request.urlopen(finalurl)
        result_string = res.read().decode("utf-8")
        result_json = json.loads(result_string)
        
        if result_json["status"] == 0:
            result = result_json["results"][0]
    except Exception as e:
        print (e)
    
    return result

def prase_excel():

    txt_file=open("file/test.txt","w")
    filename='file/highschool.xlsx'
    data = xlrd.open_workbook(filename)
    sheet1  = data.sheets()[0]
    for index in range(4,sheet1.nrows):
        row = sheet1.row_values(index)
        localtion_detail = getlocation(row[1],row[4])
        province,city,area,address,telephone,location=('---','---','---','---','---','---')
        if 'province' in localtion_detail:
            province = localtion_detail["province"]
        if 'city' in localtion_detail:
            city = localtion_detail["city"]
        if 'area' in localtion_detail:
            area = localtion_detail["area"]
        if 'address' in localtion_detail:
            address = localtion_detail["address"]
        if 'telephone' in localtion_detail:
            telephone = localtion_detail["telephone"]
        if 'location' in localtion_detail:
            location = json.dumps(localtion_detail["location"])

        message_list=[row[1],row[5],province,city,area,address,telephone,location,'\n']
        #print(' '.join(message_list))
        txt_file.write(' '.join(message_list))
        if index % 50 == 0:
            print (index,strftime("%Y-%m-%d %H:%M:%S", localtime()))
    txt_file.close()

    
   


    
    

if __name__ == "__main__":
    #getlocation('','')
    prase_excel()