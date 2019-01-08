# -*- coding: utf-8 -*-

import itchat,datetime,xlrd,re
from itchat.content import *
import nationGrade.room as ngRoom
import cmbwh.room as cmbwhRoom
import CRR.room as crrRoom
import  dtcsz.room as dtcszRoom
import dtcsh.room as dtcshRoom

def nowTime():
    nowTime = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    return nowTime

def textBackup(txt,text): #文本备份
    f = open(txt, "a+", encoding='gb18030')
    try:
        f.writelines(str(text) + '\n')
    except:
        f.writelines('\n')
    f.close()

#定义群名称方法
def group_id(room):
    try:
        df = itchat.search_chatrooms(name=room)[0]['UserName']
    except Exception as e :
        df = '异常'
    return df

# 登入登出提示配置
def lc():
    print(nowTime()+' itchat finish login')
def ec():
    print(nowTime()+' itchat exit')
    run()

def chatRoom(excel):
    fn=xlrd.open_workbook(excel)
    tb=fn.sheet_by_index(0)
    row=tb.nrows
    app={}
    for i in range(1,row):
        sp=tb.row_values(i)[0] #项目名
        spn=tb.row_values(i)[1] #指定群名
        app[sp]=spn
    return app

# room=chatRoom(excel=r'.\fileConfig\chatRoom.xlsx')


#接收微信并处理微信
@itchat.msg_register([TEXT, MAP, CARD, NOTE, SHARING], isGroupChat=True)
def tuling_reply(msg):
    try:
        wkhti = r'.\fileConfig\wkhtmltopdf\bin\wkhtmltoimage.exe'
        text=msg.text.replace('@小月儿', '').strip()   # 拿掉@对象
        name=msg.actualNickName
        room=chatRoom(excel=r'.\fileConfig\chatRoom.xlsx') #访问群配置文件
        # print(text)
        '''群配置'''
        if msg.isAt:
            item = group_id(room['全国业绩群'])
            ngRoom.config(msg=msg,name=name,item=item,
                          wkhti=wkhti,
                          excel=r'.\nationGrade\gradeFile\grade.xls',
                          txt=r'.\nationGrade\gradeFile\grade.txt',
                          pathBy=r'.\nationGrade\gradeFile',
                          reExcel=r'.\nationGrade\config\reply.xlsx',
                          image=r'.\nationGrade\gradeFile\grade.jpg')


            # item = group_id(room['cmbwh业绩群'])
            # cmbwhRoom.config(name=name,item=item,noAtMsg=text,msg=msg,
            #                  file=r'.\cmbwh\gradeFile\performance.xls',
            #                  fileBy=r'.\cmbwh\gradeFile',
            #                  file_save=r'.\cmbwh\gradeFile\performance_save.xls',
            #                  txt=r'.\cmbwh\gradeFile\performance_save.txt',
            #                  wkhti=wkhti,
            #                  reExcel=r'.\cmbwh\config\reply.xlsx',
            #                  image=r'.\cmbwh\gradeFile\performance_save.jpg')
            #
            #
            # item = group_id(room['CRR收集群'])
            # crrRoom.config(excel=r'.\CRR\crrFile\information.xls',
            #                 fileBy=r'.\CRR\crrFile',
            #                txt=r'.\CRR\crrFile\crr_statu.txt',
            #                reExcel=r'.\CRR\config\reply.xlsx',
            #                image=r'.\CRR\crrFile\crr_statu.jpg',
            #                wkhti=wkhti,item=item,name=name,msg=msg)
            #
            # item = group_id(room['测试群'])
            # ngRoom.config(msg=msg, name=name, item=item,
            #               wkhti=wkhti,
            #               excel=r'.\nationGrade\gradeFile\grade.xls',
            #               txt=r'.\nationGrade\gradeFile\grade.txt',
            #               pathBy=r'.\nationGrade\gradeFile',
            #               reExcel=r'.\nationGrade\config\reply.xlsx',
            #               image=r'.\nationGrade\gradeFile\grade.jpg')
            # item = group_id(room['DTCSZ业绩群'])
            # dtcszRoom.config(name=name,item=item,noAtMsg=text,msg=msg,
            #                  file=r'.\dtcsz\gradeFile\performance.xls',
            #                  fileBy=r'.\dtcsz\gradeFile',
            #                  file_save=r'.\dtcsz\gradeFile\performance_save.xls',
            #                  txt=r'.\dtcsz\gradeFile\performance_save.txt',
            #                  wkhti=wkhti,
            #                  reExcel=r'.\dtcsz\config\reply.xlsx',
            #                  image=r'.\dtcsz\gradeFile\performance_save.jpg')
            # item = group_id(room['DTCSH业绩群'])
            # dtcshRoom.config(name=name, item=item, noAtMsg=text, msg=msg,
            #                  file=r'.\dtcsh\gradeFile\performance.xls',
            #                  fileBy=r'.\dtcsh\gradeFile',
            #                  file_save=r'.\dtcsh\gradeFile\performance_save.xls',
            #                  txt=r'.\dtcsh\gradeFile\performance_save.txt',
            #                  wkhti=wkhti,
            #                  reExcel=r'.\dtcsh\config\reply.xlsx',
            #                  image=r'.\dtcsh\gradeFile\performance_save.jpg')




    except Exception as e:
        if re.search('list index out of range',str(e)):
            pass
        else:
            print(nowTime() + ' itchat run error：' + str(e))


def run():
    itchat.auto_login(hotReload=True, loginCallback=lc, exitCallback=ec)
    itchat.run()