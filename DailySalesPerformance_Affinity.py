# -*- coding: utf-8 -*-


import xlsxwriter
import pandas as pd
import datetime,time
import numpy as np
import os


def mkd(path):
    folder = os.path.exists(path)

    if not folder:         #判断是否存在文件夹如果不存在则创建为文件夹
        os.makedirs(path)
    else:
        pass
    return path
def DailySalesPerformanceAffinity(grade,dsp,dsp_by):


    df1 = pd.read_excel(grade)   #读取grade表
    df2=pd.read_excel(dsp,header =None)   #读取dsp表，不把第一行作列属性


    #对grade表的项目格式转为dsp表可接收的格式
    df1["项目"] = np.where(df1["项目"] == "CMBSH", "CMB CC(SH)",
                        (np.where(df1["项目"] == "CGBSH", "CGB(SH)",
                        (np.where(df1["项目"] == "CIB", "CIB",
                        (np.where(df1["项目"] == "CMBWH", "CMB CC(WH)",
                        (np.where(df1["项目"] == "BOCOMWH", "BoCom WH",
                        (np.where(df1["项目"] == "BOCOMHQ", "BoCom CC(JS)",
                        (np.where(df1["项目"] == "CITIC", "CITIC",
                        (np.where(df1["项目"] == "CGBGZ", "CGB(GZ)",
                        (np.where(df1["项目"] == "GZB", "GZBCC",
                        (np.where(df1["项目"] == "MSCC", "MS CC",
                        (np.where(df1["项目"] == "CMBCD", "CMB CC(CD)",
                        (np.where(df1["项目"] == "CEB", "CEB",
                        (np.where(df1["项目"] == "HXB", "HXB(BJ)",
                        (np.where(df1["项目"] == "CGBSY", "CGB(SY)",
                        (np.where(df1["项目"] == "CITIC211", "CITIC-211",
                        (np.where(df1["项目"] == "GZB211", "GZB-211",
                        (np.where(df1["项目"] == "CMBWH211", "CMBWH-211","null"
                                        )))))))))))))))))))))))))))))))))

    # 转list
    list_df1=np.array(df1)
    list_df2=np.array(df2 )


    nowtime=datetime.datetime.now().strftime('%Y/%m/%d')  #用于写入excel待日期格式1
    date_time = datetime.datetime.strptime(nowtime, '%Y/%m/%d')  #用于写入excel待日期格式2
    mor_time=int(datetime.datetime.now().strftime('%y%m')) #用于定义mor
    nowtime2=datetime.datetime.now().strftime('%Y%m%d')   #用于定义文件名
    nowtime3 = datetime.datetime.now().strftime('%Y-%m-%d') #用于定义文件夹名



    # 项目对比，业绩格式转换写入
    for i in range(1,len(list_df2)):
        list_df2[i][1] = nowtime   #获取系统时间
        for j in range(len(list_df1)):
            if list_df1[j][0]==list_df2[i][3]:
                try:
                    list_df2[i][7]=float(list_df1[j][2])*10000   #业绩还原实际数额
                except:
                    pass



    # dfgout=pd.DataFrame(list_df2)   #转为DataFrame格式
    #
    # #写入excel，取消第一行和第一列标识
    # dfgout.to_excel\
    #     (r"D:\itchat\file\DSP_AF\DailySalesPerformance_Affinity_"+nowtime2+'.xlsx',header=0,index=False)
    try:
        folder=mkd(path=dsp_by+nowtime3)
    except Exception as e:
        folder = mkd(path=r'.\fileConfig\DSP_AF\\' + nowtime3)
        print(nt() + ' DailySalesPerformanceAffinity error：' + str(e))
    workbook = xlsxwriter.Workbook(folder+'/DailySalesPerformance_Affinity_'+str(nowtime2)+'.xlsx')
    worksheet = workbook.add_worksheet("Daily Sales PerFormance")
    date_format = workbook.add_format({'num_format': 'YYYY/m/d'})

    #重新写入excel，使用xlsxwriter循环写入
    for i in range(len(list_df2)):
        for j in range(len(list_df2[i])):
            try:
                if i!=0 and j==1:
                    worksheet.write_datetime(i, j, date_time, date_format)
                elif i!=0 and j==2:
                    worksheet.write(i, j, mor_time)
                else:
                    worksheet.write(i, j, list_df2[i][j])


            except Exception as e:
                pass
    workbook.close()
    print(nt()+' DailySalesPerformanceAffinity done!')

def nt():
    nowTime = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    return nowTime

#dsp_by=r"\\tsclient\D\scheduleJob\fileConfig\DSP"
try:
    DailySalesPerformanceAffinity(grade=r'.\nationGrade\gradeFile\grade.xls',
                             dsp=r'.\fileConfig\DSP_AF\DailySalesPerformance_Affinity.xlsx',
        dsp_by=r"\\SZVMDBSQETL01\TM ShareFolder\tm_mis_day2\07_Pre Prodution\Milestone1\\")
except Exception as e:
    print( nt()+' DailySalesPerformanceAffinity error：'+str(e))
