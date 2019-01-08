# -*- coding: utf-8 -*-

import xlrd,xlutils.copy,datetime



def jobTran(file=r'D:\scheduleJob\fileConfig\jobRecord.xls'):
    nowTime = datetime.datetime.now().strftime('%Y-%m-%d  %H:%M:%S')
    try:
        app={}
        fn = xlrd.open_workbook(file)
        table=fn.sheet_by_index(0)
        nrow= table.nrows
        wb=xlutils.copy.copy(fn)
        ws=wb.get_sheet(0)
        for i in range(1,nrow):
            sp = table.row_values(i)[0] #项目
            status = table.row_values(i)[2] #状态
            path=table.row_values(i)[3] #jpg路径
            if status=='transformed':
                ws.write(i, 2, 'sent')  # status

                app[sp]=path
                wb.save(file)
                print(nowTime+' '+sp+' ready to send.')
    except Exception as e:
        print(nowTime+' xnJobTran running error:' + str(e))
        app=None
    # print(app)
    return app


# jobTran(file=r'D:\scheduleJob\fileConfig\jobRecord.xls')
