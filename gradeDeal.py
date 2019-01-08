# -*- coding: utf-8 -*-

import datetime,re,xlrd,os,imgkit
import pandas as pd
import xlutils.copy
import pyConfig.timeConfig as tc


# 业绩数据处理
def write_grade(deta_num, app_num, ws, j):
    if re.search("\D", deta_num) and re.search("w|W|万", deta_num):  # 判断包含非数字带单位
        for k in deta_num:
            if re.search("\d|\.", k):  # 判断包含数字和英文句号
                app_num.append(k)
            else:
                break
        num = float(" ".join(app_num).replace(" ", ""))  # 含非数字业绩整合
        num = round(num, 1)  # 小数保留位
        ws.write(j, 2, num)  # 当日业绩写入excel

    elif re.search("\D", deta_num) and re.search("[^w|W|万]", deta_num):  # 判断包含非数字不带单位
        for k in deta_num:
            if re.search("\d|\.", k):  # 判断包含数字和英文句号
                app_num.append(k)
            else:
                break
        num = float(" ".join(app_num).replace(" ", ""))  # 含非数字业绩整合

        num = float(num) / 10000
        num = round(num, 1)  # 小数保留位
        ws.write(j, 2, num)

    else:  # 若为纯数字
        try:
            num = float(deta_num) / 10000
            num = round(num, 1)  # 小数保留位
            ws.write(j, 2, num)
        except Exception as e:
            print(deta_num, "数据异常", e)

    return num


def rw_xls(file, gradeMsg):  # re1_detail为本函数返回值，用以控制微信端回复模式

    try:
        nowTime = datetime.datetime.now().strftime('%m/%d')
        fh = xlrd.open_workbook(file)
        table = fh.sheets()[0]
        nrows = table.nrows  # 行数

        # 写入不覆盖
        wb = xlutils.copy.copy(fh)
        ws = wb.get_sheet(0)  # 获取sheet对象
        ws.write(0, 2, nowTime + "业绩")  # 写入系统日期

        #  拆分接收数据，已“：”隔开，分中文英文格式判断
        if re.search(":", gradeMsg):
            detail = gradeMsg.replace(" ", "").split(":")
        else:
            detail = gradeMsg.replace(" ", "").split("：")

        # print(gradeMsg)
        # print(detail[0])

        # 遍历每一行
        for j in range(nrows):

            value = table.row_values(j)  # 读取表格每一行

            # 直接@情况下，回复图片
            if detail[0] == '@小月儿 ' or detail[0] == '@小月儿':
                re1_detail = "pass"
                # print(re1_detail)
                break


            # 判断项目名是否一致 ， re.IGNORECASE忽略字母大小写
            elif re.search(value[0], detail[0], re.IGNORECASE):

                # print(value[0], detail[0])

                app_num = []  # 非数字业绩中挑取数字存放
                strr = " "  # 用于连接app_num的业绩
                deta_num = detail[1]  # 业绩数据

                # 判断项目是否包含211，识别业绩写入
                if re.search('211', detail[0]) and re.search('211', value[0]):
                    # print('211',value[0], detail[0])


                    num = write_grade(deta_num, app_num, ws, j)


                # 211匹配错误（接收数据有211，对应表格为非211），跳过本次循环，进入下一个循环
                elif re.search('211', detail[0]) and re.search('[^211]', value[0]):
                    continue

                else:  # 非211项目业绩识别写入，通上述211处理相同
                    # print('非211',value[0], detail[0])

                    num = write_grade(deta_num, app_num, ws, j)

                value_target = table.row_values(j)[1]  # 目标业绩
                try:
                    #jie
                    if tc.what_day() == 'holiday':
                        num = num / 0.7
                    else:
                        pass

                    # 业绩超额识别
                    day = datetime.datetime.now().weekday()
                    if day in [5,6]:  #周末取消鼓励答复
                        re1_detail='pass'
                    else:
                        if num >= value_target:
                            re1_detail = "达标"
                        else:
                            if num / value_target <= 0.5:
                                re1_detail = "低于0.5"
                            elif 0.5 < num / value_target < 0.8:
                                re1_detail = "低于0.8"
                            elif num / value_target >= 0.8:
                                re1_detail = "勉强合格"



                except:  # 处理没有目标的数据
                    re1_detail = "esale"
                break  # 有匹配则跳出循环
            # 其余情况默认自动回复
            else:
                re1_detail = "自动回复"

        wb.save(file)


    except Exception as e:
        re1_detail = "输入有误"

    return re1_detail


# 业绩AF，DTC,网电汇总，并写入excel
def summary(file):
    fh1 = xlrd.open_workbook(file)
    table1 = fh1.sheets()[0]
    nrows1 = table1.nrows
    wb1 = xlutils.copy.copy(fh1)
    ws1 = wb1.get_sheet(0)  # 获取sheet对象

    # 自助寻找汇总行
    for j in range(1, nrows1):
        grade_sum = table1.row_values(j)[0]
        if re.search('AF总计', grade_sum):
            sum_af = j
        elif re.search('DTC总计', grade_sum):
            sum_dtc = j
        elif re.search('网电事业部总计', grade_sum):
            sum_tm = j

    n = 0  # AF汇总
    m = 0  # dtc汇总
    # 汇总计算及写入
    for k in range(1, nrows1):
        num = table1.row_values(k)[2]
        # print(num)

        try:

            if k < sum_af:

                n = n + num
                # print(n)

            elif sum_af < k < sum_dtc:
                m = m + num
                # print(m)
        except Exception as e:

            pass

    ws1.write(sum_af, 2, n)  # AF汇总
    ws1.write(sum_dtc, 2, m)  # dtc汇总
    ws1.write(sum_tm, 2, m + n)  # 电销汇总
    wb1.save(file)

def convertToHtml(file_path_excel,file_path_txt):
    #将excel转为html再保存为txt格式

    f = pd.read_excel(file_path_excel)
    df = pd.DataFrame(f)
    h = df.to_html(file_path_txt,index=False)
    return h

def open_txt(file):#将pandas转换成的业绩txt（html）写入list

    app_value=[]
    f = open(file, "r+")
    fr = f.readlines()
    linenum = len(fr)

    for i in range(linenum):
        app_value.append(fr[i])
    return app_value

def open_excel(file):
    app_value = []  #存放excel所有业绩
    app_value_1 = [] #超标业绩list
    fh = xlrd.open_workbook(file)
    table = fh.sheets()[0]
    nrows = table.nrows  # 行数

    #遍历excel文件
    for i in range(nrows):

        app_value.append(table.row_values(i))

    #遍历excel存放的app list，筛选达标或超标的业绩的项目名
    for j in range(1,len(app_value)):
        try:
            if float(app_value[j][2])>=float(app_value[j][1]):
                app_value_1.append(app_value[j][0])

        except:
            pass

    return app_value_1 #返回超标业绩list


def change_txt(app_excel,app_txt):#超标业绩list和txt内容


    if len(app_excel)==0:#对超额业绩list判断,避免无超额业绩，导致html修改不进行
        app_excel.append("限制器")   #使下方至少循环一次
    else:
        pass
    # print(app_excel)
    #遍历excel存放的超额业绩list
    for i in range(len(app_excel)):#至少循环一次
        for j in range(len(app_txt)):#遍历txt转成的list
            app_excel_i = app_excel[i] #超标业绩的项目list
            if re.search(str(app_excel_i), str(app_txt[j])):#识别超标业绩

                # print(app_excel_i,app_txt[j])
                app_txt_sp = str(app_txt[j]).split("<td>")[1].split("</td>")[0]

                if app_excel_i ==app_txt_sp:  #判断excel和txt的项目是否统一，避免citic对citic-211等
                    app_txt_num = str(app_txt[j + 2]).split("<td>")[1].split("</td>")[0]# 摘取app TXT超标业绩数值
                    # html编写
                    app_txt[j + 2] = ' ' * 6 + '<td font color="red" bgcolor="red">' + app_txt_num + '</td>' + '\n'


            # 修改html显示内容


            elif j==2 :#table——head修正
                app_txt[j] = ' '*4+'<tr  align="center" bgcolor="DarkOrange">'
                app_txt[j+1] =' ' * 6 + '<th width=160>项目</th>'
                app_txt[j+2] =' ' * 6 + '<th width=160>日平台目标</th>'
                at=app_txt[j+3].split(">")[1].split('<')[0]  #日期截取
                app_txt[j+3]=' ' * 6 + '<th width=160>'+at+'</th>'
            elif j==0: #修正整个表格式
                app_txt[j] ='<table border="1" cellspacing="0" cellpadding="0" class="dataframe">'

            elif re.search('<tbody>',app_txt[j]):#table——body修正
                app_txt[j] =" "*2+'<tbody align="center" bgcolor="Bisque">'

            elif re.search('<td>AF总计</td>|<td>DTC总计</td>|<td>网电事业部总计</td>',app_txt[j]):  #AF/DTC/网电汇总一行底色
                app_txt[j-1] = " " * 4 +'<tr bgcolor="DarkOrange" >'

            elif re.search('NaN',app_txt[j]) :
                app_txt[j] =' ' * 6 + '<td > </td>' + '\n'

    #对整个list做插入，修改整个html布局
    app_txt.insert(0,"<style>")
    app_txt.insert(1, 'table{line-height:28px;font-size:20;font-weight:bold;font-family:Arial, "宋体";border-collapse:collapse;border-spacing:0;border-left:1px solid #888;border-top:1px solid #888;background:#efefef;}')
    app_txt.insert(2, "th,td{border-right:1px solid #888;border-bottom:1px solid #888;}")
    app_txt.insert(3, "</style>")


#将修改后的txtapp（html）写入txt
def appToTxt(app_txt,file):
    f = open(file, "w")
    for j in range(len(app_txt)):
        f.writelines(app_txt[j])

#html转图片工具
def htmlToImage(file_path_by,wkhti):
    config = imgkit.config(wkhtmltoimage=wkhti)
    '''直接截图处理模式，图像不是很清晰'''
    options={ 'format': 'jpg',
    'crop-h': '730',
    'crop-w': '489',
    'crop-x': '6',
    'crop-y': '6',
    'encoding': "gb18030",
}
    imgkit.from_file(file_path_by + r'\grade.html', file_path_by + r"\grade.jpg",options,
                     config=config)

def run(excel,text,txt,pathBy,wkhti=None):
    reply=rw_xls(file=excel, gradeMsg=text)# 接收msg，识别项目业绩,并获取回复
    summary(file=excel)  # 业绩AF，DTC,网电汇总
    convertToHtml(file_path_excel=excel, file_path_txt=txt)#将excel转为html再保存为txt格式
    appTxt = open_txt(file=txt) # 读取excel转成txt的文件
    appExcel =open_excel(file=excel)  # 读取excel的文件，返回超标业绩
    change_txt(app_excel=appExcel, app_txt=appTxt)  # 对txt内的html文件批量编辑
    appToTxt(app_txt=appTxt, file=txt)
    try:
        os.remove(pathBy + '/grade.html')  # 移除html文件
    except:
        pass
    os.rename(txt, pathBy + '/grade.html')  # txt转html
    htmlToImage(file_path_by=pathBy,wkhti=wkhti)  # html转图片

    return reply
# reGarde=run(excel=r'.\gradeFile\grade.xls',
#            text=r'@小月儿 cmbsh：80W',
#            txt=r'.\gradeFile\grade.txt',
#            pathBy=r'.\gradeFile',
#             wkhti=r'D:\weChat\fileConfig\wkhtmltopdf\bin\wkhtmltoimage.exe'
#            )
# print(reGarde)
