# -*- coding: utf-8 -*-
import  datetime,shutil,re

def nt():
    nowTime = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    return nowTime

def nationGradeUpdate(AF=r'.\nationGrade\gradeFile\gradeAF.xls',
                      DTC=r'.\nationGrade\gradeFile\gradeDTC.xls',
                      order=None):
    try:
        if order=='AFUP':
            shutil.copyfile(AF, r'.\nationGrade\gradeFile\grade.xls')
            print(nt() + ' AF nation grade update done!')
        elif order=='DTCUP':
            shutil.copyfile(DTC, r'.\nationGrade\gradeFile\grade.xls')
            print(nt() + ' DTC nation grade update done!')
    except Exception as e:
        print(nt() + ' nationGradeUpdate error ：' + str(e))