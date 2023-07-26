import requests
import base64
import os
import xlwt
import shutil
import smtplib
import time
import zipfile
import cv2
from datetime import datetime
from py2neo import *
from pymongo import MongoClient
from collections import Counter
from pymongo import MongoClient
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication 

# globle
Traders = []
num_all = 0
# top-K
k=10


# neo4j的部分
graph = Graph('http://127.0.0.1:7474', auth=("neo4j", "root"))
# 清空neo4j
graph.delete_all()
# label1 = "复核人"
label2 = "购买方"
# label3 = "收款人"
# label4 = "开票人"
# label5 = "发票编号"
label6 = "销售方"
def createNode(a, b, c, d, e, f,propertiesDate, propertiesCost):
    # node_1 = Node(label1, name=a)
    node_2 = Node(label2, name=b)
    # node_3 = Node(label3, name=c)
    # node_4 = Node(label4, name=d)
    # node_5 = Node(label5, name=e)
    node_6 = Node(label6, name=f)
    # graph.merge(node_1, label1, "name")
    graph.merge(node_2, label2, "name")
    # graph.merge(node_3, label3, "name")
    # graph.merge(node_4, label4, "name")
    # graph.merge(node_5, label5, "name")
    graph.merge(node_6, label6, "name")
    propertiesDate = propertiesDate
    propertiesCost = propertiesCost
    # ship1 = Relationship(node_5, '发票买方', node_2, **properties)
    # graph.create(ship1)
    # ship2 = Relationship(node_5, '发票卖方', node_6)
    # graph.create(ship2)
    # ship5 = Relationship(node_5, '发票复核人', node_1)
    # graph.create(ship5)
    # ship6 = Relationship(node_5, '发票开票人', node_4)
    # graph.create(ship6)
    ship3 = Relationship(node_2, '买卖双方', node_6,**propertiesDate)
    graph.create(ship3)
    ship3 = Relationship(node_2, '买卖双方', node_6,**propertiesCost)
    graph.create(ship3)
    # ship4 = Relationship(node_6, '销售方收款人', node_3)
    # graph.create(ship4)

# mongodb的部分
host = '127.0.0.1'   # 你的ip地址
client = MongoClient(host, 27017)  # 建立客户端对象
db = client['sehw_02']  # 连接mydb数据库，没有则自动创建
myset = db['fapiao_a']   # 使用fapiao集合，没有则自动创建  

'''
增值税发票识别
'''

# 看数据是否符合标准
def isOk(type,a):
    if a == '':
        return False;
    if type == 2:
        try:
            datetime.strptime(a, "%Y年%m月%d日")
        except ValueError:
            return False;
    elif type == 3:
        try:
            float(a)
        except ValueError:
            return False;
    return True;

# 看是否符合通过的标准
def isPass(a, b, c):
    # b的标准
    # # 先判断是否需要转人工
    # print("a is:",a, isOk(1,a))
    # print("b is:",b, isOk(2,b))
    # print("c is:",c, isOk(3,c))
    # if isOk(1, a)==False or isOk(2, b)==False or isOk(3, c)==False:
    #     return "转人工";
    # # 分a,b
    # # 付款方
    # payer = "深圳市购机汇网络有限公司"
    # # 时间范围
    # need_year = 2016
    # need_month = 6
    # need_day = 12
    # # 审批金额
    # cost = 2700
    
    # # b是年月日
    # date_str = b
    # date = datetime.strptime(date_str, "%Y年%m月%d日")
    # year = date.year
    # month = date.month
    # day = date.day
    
    # print("年月日is:",year, month,day)
    # if (a == payer) and (float(c) <= cost) and (year == need_year) and (month == need_month) and (day == need_day):
    #     return "通过"
    # else:
    #     return "不通过"
    
    
    # a的标准
    # 先判断是否需要转人工
    print("a is:",a, isOk(1,a))
    print("b is:",b, isOk(2,b))
    print("c is:",c, isOk(3,c))
    if isOk(1, a)==False or isOk(3, c)==False:
        return "转人工";
    # 分a,b
    # 付款方
    payer = "浙江大学"
    # 时间范围
    need_year = 2015
    # 审批金额
    cost = 1600
    
    # b是年月日或年
    date_str = b
    isItOk = False;
    try:
        date = datetime.strptime(date_str, "%Y年%m月%d日")
        isItOk = True
    except ValueError:
        isItOk =False
    if(isItOk == False):
        try:
            date = datetime.strptime(date_str, "%Y年")
            isItOk = True
        except ValueError:
            isItOk =False
    if(isItOk == False):
        try:
            date = datetime.strptime(date_str, "%Y-%M-%D")
            isItOk = True
        except ValueError:
            isItOk =False
    if(isItOk == False):
        return "转人工"
    year = date.year
    print("年is:",year)
    if (a == payer) and (float(c) <= cost) and (year <= need_year):
        return "通过"
    else:
        return "不通过"

# 获取发票正文内容
def get_context(pic):
    global Traders
    # print('正在获取图片正文内容！')
    data = {}
    try:
        url = "baiduOCR url"
        payload = ""
        headers = {
            'Content-Type': 'application/json',
            'Accept': 'application/json'
        }
        response = requests.request("POST", url, headers=headers, data=payload)
        # 获取token
        access_token=response.json().get("access_token")
        print("access toke is:",access_token)
        request_url = "baiduOCR invoice url"
        # 二进制方式打开图片文件
        f = open(pic, 'rb')
        img = base64.b64encode(f.read())
        params = {"image":img}
        request_url = request_url + "?access_token=" + access_token
        headers = {'content-type': 'application/x-www-form-urlencoded'}
        response = requests.post(request_url, data=params, headers=headers)
        if response:
            if response.json()['words_result'] == {}:
                print("wrong data");
                return None;
            print("now is in the ocr")
            print ("json is:",response.json())
            json1 = response.json()
            print("now is normal part:")
            # PurchaserAddress 购买方地址电话
            data['PurchaserAddress'] = json1['words_result']['PurchaserAddress']
            # TotalAmount 总金额
            data['TotalAmount'] = json1['words_result']['TotalAmount']
            # Checker 复核人
            data['Checker'] = json1['words_result']['Checker']
            # PurchaserBank 购买方开户人及帐号
            data['PurchaserBank'] = json1['words_result']['PurchaserBank']
            # InvoiceTypeOrg 发票名称(头顶)
            data['InvoiceTypeOrg'] = json1['words_result']['InvoiceTypeOrg']
            # InvoiceNumConfirm 发票编号
            data['InvoiceNumConfirm'] = json1['words_result']['InvoiceNumConfirm']
            # TotalTax 总税额
            data['TotalTax'] = json1['words_result']['TotalTax']
            # SellerBank 销售方开户行及电话
            data['SellerBank'] = json1['words_result']['SellerBank']
            # SellerAddress 销售方地址电话
            data['SellerAddress'] = json1['words_result']['SellerAddress']
            # NoteDrawer 开票人
            data['NoteDrawer'] = json1['words_result']['NoteDrawer']
            # Payee 收款人
            data['Payee'] = json1['words_result']['Payee']
            # AmountInWords 价税合计大写
            data['AmountInWords'] = json1['words_result']['AmountInWords']
            # AmountInFiguers 价税合计小写
            data['AmountInFiguers'] = json1['words_result']['AmountInFiguers']
            # InvoiceType 发票类型
            data['InvoiceType'] = json1['words_result']['InvoiceType']
            # PurchaserName 购买方名称
            data['PurchaserName'] = json1['words_result']['PurchaserName']
            # InvoiceDate 开票日期
            data['InvoiceDate'] = json1['words_result']['InvoiceDate']
            # 卖家姓名 销售方名称
            data['SellerName'] = json1['words_result']['SellerName']
            # Province 省份
            data['Province'] = json1['words_result']['Province']
            # SellerRegisterNum 销售方纳税人识别号
            data['SellerRegisterNum'] = json1['words_result']['SellerRegisterNum']
            # 是否通过还是转人工
            data["isPass"] =isPass(json1['words_result']['PurchaserName'], json1['words_result']['InvoiceDate'], json1['words_result']['AmountInFiguers'])
            print("通过情况：", data["isPass"])
            Traders.append(data['PurchaserName'])
            Traders.append(data['SellerName'])
            myset.insert_one({"image":img, 
                              "PurchaserAddress":data["PurchaserAddress"],
                              "TotalAmount":data["TotalAmount"],
                              "Checker":data["Checker"],
                              "PurchaserBank":data["PurchaserBank"],
                              "InvoiceTypeOrg":data["InvoiceTypeOrg"],
                              "InvoiceNumConfirm":data["InvoiceNumConfirm"],
                              "TotalTax":data["TotalTax"],
                              "SellerBank":data["SellerBank"],
                              "SellerAddress":data["SellerAddress"],
                              "NoteDrawer":data["NoteDrawer"],
                              "Payee":data["Payee"],
                              "AmountInWords":data["AmountInWords"],
                              "AmountInFiguers":data["AmountInFiguers"],
                              "InvoiceType":data["InvoiceType"],
                              "PurchaserName":data["PurchaserName"],
                              "InvoiceDate":data["InvoiceDate"],
                              "SellerName":data["SellerName"],
                              "Province":data["Province"],
                              "SellerRegisterNum":data["SellerRegisterNum"],
                              "isPass":data["isPass"],
                                  })
            print("写入mongodb成功")
        return data
    except Exception as e:
        print(e)
    return None

# 定义生成图片路径的函数
def pics(path):
    global num_all
    print('正在生成图片路径')
    #生成一个空列表用于存放图片路径
    pics = []
    # 遍历文件夹，找到后缀为jpg和png的文件，整理之后加入列表
    for filename in os.listdir(path):
        if filename.endswith('jpg') or filename.endswith('png'):
            pic = path + '/' + filename
            pics.append(pic)
    print('图片路径生成成功！')
    num_all = len(pics)
    return pics

# 定义一个获取文件夹内所有文件正文内容的函数，每次返回一个字典，把返回的所有字典存放在一个列表里
def datas(pics):
    datas = []
    # 转人工的发票地址
    datas_trans = []
    Data_all = []
    for p in pics:
        data = get_context(p)
        if(data != None):
            # data就是OCR识别出来一个pic的内容
            # data = get_context(p)
            print(data);
            datas.append(data)
        else:
            # datas_trans就是OCR识别不出来需要转人工的内容
            print(p)
            datas_trans.append(p)
            print("转人工")
    Data_all.append(datas)
    Data_all.append(datas_trans)
    return Data_all

# 定义一个写入将数据excel表格的函数
def save(datas):
    print('正在写入数据！')
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet('增值税发票内容登记', cell_overwrite_ok=True)
    # 设置表头
    title = ['购买方地址电话', '总金额', '复核人', '购买方开户人及帐号', '发票名称','发票编号','总税额','销售方开户行及电话','销售方地址电话','开票人','收款人','价税合计大写','价税合计小写','发票类型','购买方名称','开票日期','销售方名称','省份','销售方纳税人识别号','通过与否']
    for i in range(len(title)):
        sheet.write(0, i, title[i])
    for d in range(len(datas)):
        print(d, datas[d])
        sheet.write(d + 1, 0, datas[d]['PurchaserAddress'])
        sheet.write(d + 1, 1, datas[d]['TotalAmount'])
        sheet.write(d + 1, 2, datas[d]['Checker'])
        sheet.write(d + 1, 3, datas[d]['PurchaserBank'])
        sheet.write(d + 1, 4, datas[d]['InvoiceTypeOrg'])
        sheet.write(d + 1, 5, datas[d]['InvoiceNumConfirm'])
        sheet.write(d + 1, 6, datas[d]['TotalTax'])
        sheet.write(d + 1, 7, datas[d]['SellerBank'])
        sheet.write(d + 1, 8, datas[d]['SellerAddress'])
        sheet.write(d + 1, 9, datas[d]['NoteDrawer'])
        sheet.write(d + 1, 10, datas[d]['Payee'])
        sheet.write(d + 1, 11, datas[d]['AmountInWords'])
        sheet.write(d + 1, 12, datas[d]['AmountInFiguers'])
        sheet.write(d + 1, 13, datas[d]['InvoiceType'])
        sheet.write(d + 1, 14, datas[d]['PurchaserName'])
        sheet.write(d + 1, 15, datas[d]['InvoiceDate'])
        sheet.write(d + 1, 16, datas[d]['SellerName'])
        sheet.write(d + 1, 17, datas[d]['Province'])
        sheet.write(d + 1, 18, datas[d]['SellerRegisterNum'])
        sheet.write(d + 1, 19, datas[d]["isPass"])
        name1 = datas[d]['Checker'];
        name2 = datas[d]['PurchaserName'],
        name3 = datas[d]['Payee']
        name4 = datas[d]['NoteDrawer']
        name5 = datas[d]['InvoiceNumConfirm']
        name6 = datas[d]['SellerName']
        str = datas[d]['InvoiceNumConfirm'] + "号订单Date"
        str2 = datas[d]['InvoiceNumConfirm'] + "号订单Cost"
        propertyDate = {str : datas[d]['InvoiceDate']}
        propertyCost = {str2 : datas[d]['TotalAmount']}
        createNode(name1, name2, name3, name4, name5, name6, propertyDate, propertyCost)
        print(name5, "号订单写入neo4j")     
    print('数据写入成功！')
    book.save('OCR可识别发票.xls')

# 获取top-k个交易主体
def getTopK(Tranders):
    counts = Counter(Tranders)
    unique_str = set(Tranders)
    print(counts)
    print(unique_str)
    trades = []
    num_trade = 0;
    for s, count in sorted(counts.items(), key=lambda x: x[1], reverse=True):
        # print(f"{s}: {count}")
        print(s," ",count)
        trade = []
        trade.append(s);
        trade.append(count);
        trades.append(trade)
        num_trade = num_trade + 1;
        if(num_trade == k):
            break;
    return trades;

# save——需要转人工的发票
def save_unpass(datas):
    print('正在写入未通过数据！')
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet('未通过登记', cell_overwrite_ok=True)
    # 设置表头
    title = '发票图片名称'
    sheet.write(0, 0, title)
    for i in range(len(datas)):
        sheet.write(i + 1, 0, datas[i])
    print('数据写入成功！')
    book.save('转人工发票.xls')

# 此轮处理的统计信息
def getAllInfo(trades,num_all, num_pass, num_not_pass, num_trans):
    print('正在写入统计信息！')
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet('此轮处理统计信息', cell_overwrite_ok=True)
    # 设置表头
    title = ['发票数量','审批通过数量','审批不通过数量','转人工数量','审批通过比例','审批不通过比例','转人工比例']
    for i in range(len(title)):
        sheet.write(0, i, title[i])
    print("now is some statics number:",num_all,num_trans,num_not_pass,num_pass)
    sheet.write(1, 0, num_all)
    sheet.write(1, 1, num_pass)
    sheet.write(1, 2, num_not_pass)
    sheet.write(1, 3, num_trans)
    sheet.write(1, 4, str(num_pass / num_all * 100) + "%")
    sheet.write(1, 5, str(num_not_pass/num_all * 100) + "%")
    sheet.write(1, 6, str(num_trans/num_all * 100) + "%")
    sheet.write(2, 0, "Top-K交易主体")
    for i in range(len(trades)):
        sheet.write(3 + i, 0, trades[i][0])
        sheet.write(3 + i, 1, trades[i][1])
    print('统计数据写入成功！')
    book.save('统计数据.xls')
    
# 获取通过、不通过的数量
def getNumData(Datas):
    num_pass = 0;
    num_not_pass = 0;
    num_trans = 0;
    ret =[]
    for i in range(len(Datas)):
        if(Datas[i]["isPass"] == '通过'):
            num_pass = num_pass + 1
        elif (Datas[i]["isPass"] == '不通过'):
            num_not_pass = num_not_pass + 1
        else:
            num_trans = num_trans + 1
    ret.append(num_pass)
    ret.append(num_not_pass)
    ret.append(num_trans)    
    return ret;


# 移动目标文件夹的根目录
movabs_path = "人工处理发票" 
# 将需要转人工的文件copy到当前目录的人工文件夹下面
def movePicTrans(datas):
    for i in range(len(datas)):
        # 移动操作
        fileName = os.path.basename(datas[i])
        shutil.copy(datas[i],movabs_path+"/"+ fileName)
        
def zipDir(dirpath, outFullName):
    """
    压缩指定文件夹
    :param dirpath: 目标文件夹路径
    :param outFullName: 压缩文件保存路径+xxxx.zip
    :return: 无
    """
    zip = zipfile.ZipFile(outFullName, "w", zipfile.ZIP_DEFLATED)
    for path, dirnames, filenames in os.walk(dirpath):
        # 去掉目标跟路径，只对目标文件夹下边的文件及文件夹进行压缩
        fpath = path.replace(dirpath, '')
        for filename in filenames:
            zip.write(os.path.join(path, filename), os.path.join(fpath, filename))
    zip.close()
       
#   发邮件
def sentMail(num_all, num_pass, num_not_pass, num_trans):
    fromaddr = 'fromaddr'
    password = 'password'
        
    toaddrs = ['mail1','mail2']
        
    t = time.localtime()
    print(t)
    content = str(t.tm_year)+'年'+str(t.tm_mon)+'月'+str(t.tm_mday)+'日'+str(t.tm_hour)+'时'+str(t.tm_min)+'分'+str(t.tm_sec)+'秒'+'完成的该批次发票处理结果见附件。此此处理中，共处理：'+ str(num_all)+"张发票，其中，通过审批的有：" +str(num_pass)+"张发票，未通过审批的有：" +str(num_not_pass)+"张发票，需要转人工的有：" +str(num_trans)+"张发票。其中，需要转人工的发票已经在附件压缩包中"
    textApart = MIMEText(content)
 
    excel1File = 'OCR可识别发票.xls'
    excel1Apart = MIMEApplication(open(excel1File, 'rb').read())
    excel1Apart.add_header('Content-Disposition', 'attachment', filename=excel1File)
    
    excel2File = '转人工发票.xls'
    excel2Apart = MIMEApplication(open(excel2File, 'rb').read())
    excel2Apart.add_header('Content-Disposition', 'attachment', filename=excel2File)
        
    excel3File = '统计数据.xls'
    excel3Apart = MIMEApplication(open(excel3File, 'rb').read())
    excel3Apart.add_header('Content-Disposition', 'attachment', filename=excel3File)
    
    zipFile = '人工处理发票.zip'
    zipApart = MIMEApplication(open(zipFile, 'rb').read())
    zipApart.add_header('Content-Disposition', 'attachment', filename=zipFile)
 
    m = MIMEMultipart()
    m.attach(textApart)
    m.attach(excel1Apart)
    m.attach(excel2Apart)
    m.attach(excel3Apart)
    m.attach(zipApart)
    m['Subject'] = str(t.tm_year)+'年'+str(t.tm_mon)+'月'+str(t.tm_mday)+'日'+str(t.tm_hour)+'时'+str(t.tm_min)+'分'+str(t.tm_sec)+'秒' + '处理的批次发票'
    try:
        server = smtplib.SMTP('smtp.163.com')
        server.login(fromaddr,password)
        server.sendmail(fromaddr, toaddrs, m.as_string())
        print('success')
        server.quit()
    except smtplib.SMTPException as e:
        print('error:',e) #打印错误

# 获取需要转人工的发票zip
def getZipTrans(datas):
    movePicTrans(datas)
    input_path = "人工处理发票"
    output_path = "人工处理发票.zip"
    zipDir(input_path, output_path)
    
# 定义生成图片路径的函数
def getNormalPics(path):
    print('正在生成图片路径')
    #生成一个空列表用于存放图片路径
    pics = []
    # 遍历文件夹，找到后缀为jpg和png的文件，整理之后加入列表
    for filename in os.listdir(path):
        if filename.endswith('jpg') or filename.endswith('png'):
            pic = path + '/' + filename
            pics.append(pic)
    print('图片路径生成成功！')
    print(pics)
    return pics

def getNewPic(path, path_after):
    Pics = getNormalPics(path)
    for i in range (len(Pics)):
        img = cv2.imread(Pics[i]) #读图
        print(img.shape)
        height,width = img.shape[:2]  #获取原图像的水平方向尺寸和垂直方向尺寸。
        height = int(height / 1.25)
        width = int(width / 1.25)
        print(height,width)
        res = cv2.resize(img, (width, height) ,interpolation=cv2.INTER_CUBIC) 
        cv2.imshow('res',res)
        fileName = os.path.basename(Pics[i])
        output_path = path_after + "/" + fileName
        print(output_path)
        cv2.imwrite(output_path, res)

def main():
    global Traders;
    print('开始执行！！！')
    # 发票的存放地址
    path_before = 'fapiao'
    # 发票预处理完成后的存放地址
    path = 'fapiao1'
    # 图片预处理
    getNewPic(path_before, path)
    # pics图片路径
    Pics1 = pics(path)
    
    # 一批次处理的发票数量
    batch_size = 50
    # 开始分批次处理发票
    subarrays = [Pics1[i:i+batch_size] for i in range(0, len(Pics1), batch_size)]
    if len(subarrays[-1]) < batch_size:
        last_subarray = subarrays.pop()
        subarrays[-1].extend(last_subarray)
        
    print(len(subarrays))
    print(subarrays)
    # subarray就是分批次的数组
    for numarr in range(len(subarrays)):
        Traders = [];
        Pics = subarrays[numarr]
        print(Pics)
        transpath='人工处理发票'
        folder = os.path.exists(transpath)
        if not folder:                   #判断是否存在文件夹如果不存在则创建为文件夹
            os.makedirs(transpath)            #makedirs 创建文件时如果路径不存在会创建这个路径
            print("---  new folder...  ---")
            print("---  OK  ---")
        else:
            print("---  There is this folder!  ---")
        
        # 接着就是正常的处理
        # 得到Data_all
        Data_all = datas(Pics)
        Datas = Data_all[0]
        # 需要转人工的Data
        Datas_trans = Data_all[1]
        print("能够识别的是:",Datas)
        print("不能识别的是:",Datas_trans)
        
        save(Datas)
        print("能识别的输出完成")
        # 
        save_unpass(Datas_trans)
        print("不能识别的输出完成")
        # 得到一些统计的数据
        trades = getTopK(Traders);
        num_trans = len(Datas_trans)
        numData = getNumData(Datas)
        num_pass = numData[0]
        num_not_pass = numData[1]
        num_trans = num_trans + numData[2]
        num_all = len(Pics)
        getAllInfo(trades,num_all, num_pass, num_not_pass, num_trans)
        print("统计信息输出完成")
        # 得到需要转人工的发票集合zip
        getZipTrans(Datas_trans)
        # 发邮件
        sentMail(num_all, num_pass, num_not_pass, num_trans)
        print('执行结束！')
        
        print("清空人工处理发票以便进行下一个批次的发送任务")
        shutil.rmtree('人工处理发票') # 能删除该文件夹和文件夹下所有文件

if __name__ == '__main__':
    main()
