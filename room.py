# -*- coding: utf-8 -*-
import os,xlrd,itchat,re
import nationGrade.gradeDeal as gd
import nationGrade.gradeUpdate as gu



def replyFile(reExcel):
    fn = xlrd.open_workbook(reExcel)
    tb = fn.sheet_by_index(0)
    row = tb.nrows
    app = {}
    for i in range(1, row):
        sp = tb.row_values(i)[0]  # 项目名
        content = tb.row_values(i)[1]  # 内容
        app[sp] = content
    return app
# print(replyFile(excel=r'.\reply.xlsx'))



def config(msg,name,item,wkhti,image,excel,txt,pathBy,reExcel):
    if msg['FromUserName'] == item:
        #业绩处理
        text=msg.text

        # 读取回复配置文件
        ref=replyFile(reExcel)

        if re.search('AFUP',text):
            gu.nationGradeUpdate(order='AFUP')
            itchat.send(u'@%s : %s' % (name, 'AF Update Done！'), item)
        elif re.search('DTCUP', text):
            gu.nationGradeUpdate(order='DTCUP')
            itchat.send(u'@%s : %s' % (name, 'DTC Update Done！'), item)
        else:
            reGarde = gd.run(excel, text, txt, pathBy, wkhti)
            for key in ref:
                if key == reGarde:
                    if ref[key] == '推图':
                        itchat.send_image(image, item)
                    elif key == '输入有误':
                        itchat.send(u'@%s : %s' % (name, ref[key]), item)
                    elif key == '自动回复':
                        itchat.send(u'@%s : %s' % (name, ref[key]), item)
                    else:
                        itchat.send_image(image, item)
                        itchat.send(u'@%s : %s' % (name, ref[key]), item)
                    break

    else:
        pass
