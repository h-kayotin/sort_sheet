"""
copy_sheet_quick - 用xlrd复制带格式的sheet

Author: hanayo
Date： 2023/6/25
"""

import xlrd
import xlwt
from xlutils.copy import copy


# def read_format_info(file):
#     # 打开要读取的 Excel 文件
#     fileAll = xlrd.open_workbook(file, formatting_info=True)
#     # 读取数据和样式，保存在data中
#     data = copy(fileAll)
#     # 获取sheet1表格
#     table = data.get_sheet(0)
#     # 把默认输出的页眉页脚删除
#     table.header_str = b''
#     table.footer_str = b''
#     # 新建一个样式，因为此库虽然保留整体格式，但是在输出时不保留原字体格式，需要手动设置,在此处设置一个宋体，14号字体
#     song14 = xlwt.XFStyle()
#     song14.font.name = u'宋体'
#     song14.font.height = 280  # 字号*20
#     # 参数说明:x,y：索引列，行 value：此处写入的值 song14：样式
#     table.write(x, y, value, song14)
#     # 保存 path:保存路径
#     data.save("复制数据.xls")
#
#
#
# if __name__ == '__main__':
#     file_path = r"C:\Users\JiangHai江海\Desktop\test111\test1.xls"
#     read_format_info(file_path)
