# -*- coding: utf-8 -*-
import json,datetime
import urllib,requests
import calendar

def nowtime():
    nowTime = datetime.datetime.now().strftime('%Y%m%d')
    return nowTime

def what_day():#判断节假日

    try:
        date = nowtime()
        # date = '20180909'

        server_url='http://api.goseek.cn/Tools/holiday?date='

        vop_url_request = urllib.request.Request(server_url + date)
        vop_response = urllib.request.urlopen(vop_url_request)
        vop_data = json.loads(vop_response.read())
        # print(vop_data)

        #0：weekday、1：weekend、2：holiday
        if vop_data['data'] ==0:
            today ="weekday"
        elif vop_data['data'] == 1 or vop_data['data'] ==2 :
            today ='holiday'
        else:
            today=None

    except Exception as e:
        today ='Error'
    # print(today)
    return today


#当月最后一天
def last_day():
    year = int(datetime.datetime.now().strftime('%Y'))
    month = int(datetime.datetime.now().strftime('%m'))
    day = calendar.monthrange(year,month)[1]
    f_day=str(year) + "-" + str(month).rjust(2, '0') + "-" + str(day)
    # print(f_day)
    nowTime = str(datetime.datetime.now().strftime('%Y-%m-%d'))
    # nowTime='2018-12-31'
    if nowTime==f_day:
        f_day='do'
    # print(f_day)
    return f_day

#当月第一天
def first_day():
     l_day= str(datetime.datetime.now().strftime('%Y-%m-'))+'01'
     nowTime = str(datetime.datetime.now().strftime('%Y-%m-%d'))
     # nowTime='2018-12-01'
     if nowTime == l_day:
         l_day = 'do'
     # print(l_day)
     return l_day


# last_day()
# first_day()



# print(what_day())
