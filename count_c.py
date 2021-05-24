#-*- codeing = utf-8 -*- 
#@Time: 2021/5/6 13:59
#@Author : dapao
#@File : count_c.py
#@Software: PyCharm

import xlwt
import cv2
import os
import numpy as np

def count(path):
    '''
    计算周长和面积
    :param path: 
    :return: 周长，面积
    '''
    image_name = []#存放文件名
    perimeter=[]#存放面积的列表 list
    area = []#存放面积的列表 list
    for id in os.listdir(path):
        print(id)
        image_name.append(id)
        image_path = os.path.join(path,id)

        image = cv2.imread(image_path)#读入图片
        #cv2.imshow("image",image)
        img_gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)#转化成灰度图像
        ret, img_bin = cv2.threshold(img_gray, 127, 255, cv2.THRESH_BINARY)
        # cv2.imshow("images",img_bin)

        kernel = cv2.getStructuringElement(
            cv2.MORPH_RECT, (15, 15))  # 构造一个用于形态学使用的核函数，
        binary = cv2.morphologyEx(img_bin, cv2.MORPH_OPEN, kernel)
        #binary = cv2.morphologyEx(img_bin, cv2.MORPH_CLOSE, kernel, anchor=(2, 0), iterations=5)
        # cv2.imshow("imagesss", binary)
        contours, hierarchy = cv2.findContours(binary,cv2.RETR_TREE,cv2.CHAIN_APPROX_SIMPLE)#边缘
        #print(len(contours))
        # cv2.imshow("gray",img_gray)
        contours= sorted(contours, key=cv2.contourArea, reverse=True)  # 找到面积最大的
        c = cv2.arcLength(contours[1],True)#计算周长，
        print('周长：',c)
        s = cv2.contourArea(contours[1])#计算面积
        print('面积：',s)
        print('-'*30)
        imgnew = cv2.drawContours(image, contours, -1, (0, 255, 255), 3)  # 把所有轮廓画出来
        perimeter.append(c)
        area.append(s)
    return image_name,perimeter,area
        # cv2.imshow("canny",imgnew)
        # # cv2.waitKey(0)
        # cv2.destroyAllWindows()


def write_excel(save_excel_path,image_name,perimeter,area):
    '''
    输入图片的名称；周长；面积。他们均以list的形式保存，并将他们写入到excel中
    '''
    #首先判断我们保存路径是否已经存在这个表格，如果存在这里的做法是删除，如果你想保存多个表格，可以修改保存路径下的文件
    if os.path.isfile(save_excel_path):
        os.remove(save_excel_path)
        print("已经删除表格{}".format(save_excel_path))
    index = len(image_name)  # 获取需要写入数据的行数
    workbook = xlwt.Workbook()  # 新建一个工作簿
    sheet_name = 'count'
    sheet = workbook.add_sheet(sheet_name,cell_overwrite_ok=True)  # 在工作簿中新建一个表格
    sheet.write(0,0,"图片名")  # 像表格中写入数据（对应的行和列）
    sheet.write(0,1,"周长")
    sheet.write(0,2,"面积")
    for i in range(0, index):#行       
        sheet.write(i+1, 0, image_name[i])  # 像表格中写入数据（对应的行和列）
        sheet.write(i+1, 1, perimeter[i])  # 像表格中写入数据（对应的行和列）
        sheet.write(i+1, 2, area[i])  # 像表格中写入数据（对应的行和列）
    workbook.save(save_excel_path)  # 保存工作簿
    print("xls格式表格写入数据成功！")



path = './images'#图片存放的文件路径
save_excel_path = './计算表.xls'
image_name,perimeter,area = count(path)
write_excel(save_excel_path,image_name,perimeter,area)