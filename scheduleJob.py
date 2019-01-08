# -*- coding: utf-8 -*-
import datetime,time,itchat,xlrd,imp,re
import schedule
import CRR.crrInfo as ci
import pyConfig.xnJobReport as xnJR
import pyConfig.timeConfig as tc

def nowTime():
    nowTime = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    return nowTime

#发送消息配置
def sendChat(content=None,toRoom=None,file=None):
    room=itchat.search_chatrooms(name=toRoom)[0]['UserName']
    try:
        if content:
            itchat.send(content, toUserName=room)
        if file:
            itchat.send_image(file, toUserName=room)
    except Exception as e:
        print(nowTime() +' '+toRoom+' sendChat error：' + str(e))

#读取schedule指定发送群
def chatRoom():
    fn=xlrd.open_workbook(r'.\fileConfig\scheduleRoom.xlsx')
    tb=fn.sheet_by_index(0)
    row=tb.nrows
    app={}
    for i in range(1,row):
        sp=tb.row_values(i)[0] #项目名
        spn=tb.row_values(i)[1].split('/') #指定群名
        app[sp]=spn
    return app


#读取schedule指定时间执行PY脚本
def fileClean():
    imp.load_source('fileClean','fileClean.py')
    room = chatRoom()['fileClean']
    for i in room:
        sendChat(content=nowTime() + ' fileClean done! ',toRoom=i)
# fileClean()
def DSP():
    imp.load_source('DSP','DailySalesPerformance_Affinity.py')
    room = chatRoom()['DSP']
    for i in room:
        sendChat(content=nowTime() + " DSP done!",toRoom=i)

def monitor():
    room=chatRoom()['monitor']
    for i in room:
        sendChat(content=nowTime() + " 运行正常报告.", toRoom=i)

#CRR推送任务
def crrInfo(nt):
    day = datetime.datetime.now().weekday()
    if tc.what_day() == 'weekday' and day in [0, 2, 4]:
        content=ci.runInfo(excel=r'.\CRR\crrFile\information.xls',nt=nt)
        if content :
            room = chatRoom()['CRR收集群'][0]
            sendChat(content=content, toRoom=room)
#报表推送任务
def sendReport():
    try:
        rep=xnJR.jobTran()
        if rep:
            fn=xlrd.open_workbook(r'.\fileConfig\xnReportConfig.xlsx')
            row=fn.sheet_by_index(0).nrows
            for i  in range(1,row):
                job=fn.sheet_by_index(0).row_values(i)[0]
                group=str(fn.sheet_by_index(0).row_values(i)[1]).split('/')
                for k in rep:
                    if k==job:
                        for n in group:
                            room =chatRoom()[n]
                            for j in room:
                                sendChat(file=rep[k], toRoom=j)
    except Exception as e:
        print(nowTime() + ' sendReport error：' + str(e))

def run():
    time.sleep(5)
    print(nowTime() + " 定时任务启动")
    # schedule.every(5).seconds.do(DSP)  # test用
    # schedule.every(59).seconds.do(crrInfo)  #CCR接收通知

    schedule.every(30).minutes.do(monitor)  # 微信机器人运行报告
    schedule.every(10).minutes.do(sendReport)  #报表推送任务
    schedule.every().day.at("01:00").do(fileClean)  # 定时清理数据
    schedule.every().day.at("23:59").do(DSP)  # DSP

    #CCR通知
    schedule.every().day.at('10:30').do(crrInfo, '10:30')
    schedule.every().day.at('17:00').do(crrInfo, '17:00')
    schedule.every().day.at('17:30').do(crrInfo, '17:30')

    while True:
        hms= datetime.datetime.now().strftime('%H:%M:%S')
        try:
            schedule.run_pending()  # 定时任务执行

        except Exception as e:
            if re.search("Bad CRC-32 for file",str(e)):
                pass
            else:
                print(nowTime() + ' schedule wrong:' + str(e))
