# coding=gbk

import os
import xlwt
import string

#改变工作目标到gucy文件下
work_dir = "D:\gucy"
os.chdir(work_dir)
gucy_dir_list = os.listdir(work_dir)
for i in gucy_dir_list:
    #获得gucy文件夹下的Resultname目录文件路径
    if(os.path.isdir(i)):
        res_dir_abspath = os.getcwd()+"\\"+i
        #列出Resultname目录下文件名
        Resultname_list = os.listdir(res_dir_abspath)
        name=0;
        txt=0;
        ssim=0
        filename = xlwt.Workbook ()
        sheet = filename.add_sheet(i)
        for j in Resultname_list:
            #txt文件的绝对路径
            file_path = res_dir_abspath +"\\"+j
            #print file_path
            dic={}
            for line in open(file_path):
                #line的形式：'57698.txt\t0.10871917032516225\n'
                tmp = line.split('\n')
                str_txt_value= tmp[0].split('\t')
                dic[str_txt_value[0]]=string.atof(str_txt_value[1])
            print j
            #print dic
            sort_dict = sorted(dic.iteritems(), key=lambda d:d[1], reverse = True)
            print sort_dict
            #排序后格式如下：[('57698.txt', 0.10871917032516225), ('83600.txt', 0.07924852550386113), ('9835.txt', 0.06039286652151702)]
            for k in sort_dict:
                #写入文件名
                sheet.write(name,0,j)
                name +=1
                #写入文本名
                sheet.write(txt,1,k[0])
                txt +=1
                #写入相似度
                sheet.write(ssim,2,k[1])
                ssim +=1
        filename.save(i+'.xls')
