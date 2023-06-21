

import win32com.client as win32
import os
def transform(parent_path,out_path):
    fileList = os.listdir(parent_path)  #文件夹下面所有的文件
    num = len(fileList)
    for i in range(num):
        file_Name = os.path.splitext(fileList[i])   #文件和格式分开
        if file_Name[1] == '.xls':
            transfile1 = parent_path+'\\'+fileList[i]  #要转换的excel
            transfile2 = out_path+'\\'+file_Name[0]    #转换出来excel
            excel=win32.gencache.EnsureDispatch('excel.application')
            pro=excel.Workbooks.Open(transfile1)   #打开要转换的excel
            pro.SaveAs(transfile2 + ".xlsx", FileFormat=51)  # 另存为xlsx格式
            pro.Close()
            excel.Application.Quit()
def

if __name__=='__main__':
    path1=r"E:\untitled1\test_report_demo"  #待转换文件所在目录
    path2=r"E:\untitled1\test_data"  #转换文件存放目录
    transform(path1, path2)