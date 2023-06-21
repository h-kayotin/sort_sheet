import pathlib
import win32com.client as win32


def trans_xlsx():
    """将Excel转换成xlsx"""
    file_name = input("输入文件路径：")

    # 创建 Excel Application 对象
    excel = win32.gencache.EnsureDispatch('Excel.Application')

    try:
        # 使用 with 语法打开 .xls 文件
        with excel.Workbooks.Open("res/exp.xls") as wb:
            # 将 .xls 文件另存为 .xlsx 文件
            wb.SaveAs(f'transed.xlsx', FileFormat=51)

        print('转换成功！')

    except Exception as e:
        print(f'转换失败！原因：{e}')

    finally:
        # 退出 Excel
        excel.Application.Quit()



trans_xlsx()
