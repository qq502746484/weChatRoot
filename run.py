# -*- coding: utf-8 -*-

import itchat
import chatRoom,scheduleJob
import threading,datetime

def nowTime():
    nowTime = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    return nowTime

def wechatRun():
    chatRoom.run()

def schedule():
    scheduleJob.run()
    # pass

if __name__ == '__main__':

    # itchat.send('ceshi', group_id('小梁测试组'))
    while True:
        try:
            t1 = threading.Thread(target=schedule, args=())  # 定时任务
            t2 = threading.Thread(target=wechatRun, args=())  # 微信机器人
            # 守护进程，防止未执行
            t1.setDaemon(True)
            t2.setDaemon(True)
            # 进程运行
            t1.start()
            t2.start()
            t1.join()
            t2.join()
        except Exception as e:
            print( nowTime()+' threading run error：：' + str(e))
