# -*- coding: utf-8 -*-
import  xlrd,xlutils.copy,datetime,shutil,re
import pyConfig.timeConfig as tc

def nt():
    nowTime = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    return nowTime

def nationGrade(excel):
    fh = xlrd.open_workbook(excel)
    nrows = fh.sheets()[0].nrows
    wb = xlutils.copy.copy(fh)  # 写入不覆盖
    ws = wb.get_sheet(0)
    for i in range(nrows):
        ws.write(i, 2, " ")
    wb.save(excel)
    print(nt() + ' nationGrade clean done!')

def cmbwhGrade(excel):
    fn = xlrd.open_workbook(excel)
    nrow = fn.sheet_by_index(0).nrows
    ncol = fn.sheet_by_index(0).ncols
    wb = xlutils.copy.copy(fn)
    ws = wb.get_sheet(0)
    for i in range(1, nrow):
        for j in range(2, ncol):
            ws.write(i, j, '')
    wb.save(excel)
    print(nt() + ' cmbwhGrade clean done!')

def crrData(excel):    # crr_file配置
    fh = xlrd.open_workbook(excel)
    table = fh.sheet_by_name('反馈状态')
    nrows = table.nrows  # 行数
    wb = xlutils.copy.copy(fh)  # 写入不覆盖
    ws = wb.get_sheet(1)  # 获取sheet对象

    for i in range(1,nrows):
        ws.write(i, 2, " ")
    wb.save(excel)
    try:
        # shutil.copyfile(excel, r'\\tsclient\D\scheduleJob\fileConfig\CRRFile\information.xls')
        shutil.copyfile(excel, r'D:\scheduleJob\fileConfig\CRRFile\information.xls')
    except Exception as e:
        print(nt() + ' crrData error ：'+str(e))
    print(nt() + ' crrData clean done!')

def dtcszGrade(excel):
    fh = xlrd.open_workbook(excel)
    nrow= fh.sheets()[0].nrows
    ncol   = fh.sheets()[0].ncols
    wb = xlutils.copy.copy(fh)  # 写入不覆盖
    ws = wb.get_sheet(0)
    for i in range(1,nrow):
        for j in range(3,ncol):
            ws.write(i, j, "")
    wb.save(excel)
    print(nt() + ' dtcSZGrade clean done!')

def dtcshGrade(excel):
    fh = xlrd.open_workbook(excel)
    nrow= fh.sheets()[0].nrows
    ncol   = fh.sheets()[0].ncols
    wb = xlutils.copy.copy(fh)  # 写入不覆盖
    ws = wb.get_sheet(0)
    for i in range(1,nrow):
        for j in range(3,ncol):
            ws.write(i, j, "")
    wb.save(excel)
    print(nt() + ' dtcSHGrade clean done!')


def run():
    nationGrade(excel=r'.\nationGrade\gradeFile\grade.xls')
    cmbwhGrade(excel=r'.\cmbwh\gradeFile\performance.xls')
    crrData(excel=r'.\CRR\crrFile\information.xls')
    dtcszGrade(excel=r'.\dtcsz\gradeFile\performance.xls')
    dtcshGrade(excel=r'.\dtcsh\gradeFile\performance.xls')

run()






























