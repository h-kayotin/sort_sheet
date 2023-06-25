"""
trans_xlsx - 把xls转成xlsx

Author: hanayo
Date： 2023/6/25
"""

import pathlib
import win32com.client as win32


def trans_xlsx(file):
    """
    转换xls到xlsx
    :param file:文件的完整路径
    :return:返回成功或失败的原因，以及布尔值
    """
    file_name = pathlib.Path(file).name
    save_path = pathlib.Path(file).parent
    # 创建 Excel Application 对象
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    try:
        wb = excel.Workbooks.Open(file)
        new_name = file_name + "x"
        wb.SaveAs(f"{save_path}/{new_name}", FileFormat=51)
        wb.Close(SaveChanges=1)
        txt_log = "转换成功"
        return txt_log, True
    except Exception as e:
        txt_log = f'转换失败！原因：{e}'
        return txt_log, False
    finally:
        # 退出 Excel
        excel.Application.Quit()
