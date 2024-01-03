# sort_sheet
给Excel里的sheet排序

根据xls文件中指定单元格的值对所有sheet进行排序，然后保存到新的文件

### 背景介绍

财务导出的某个xls文件包含多个相同格式的sheet，需要根据指定单元格的值进行排序，然后保存为新的工作簿。

本来以为很简单的一个需求，但其实难点在于原sheet的格式不能改变。

不管是openpyxl，还是xlrd，xlwt，都没有带格式的复制的方法。

本工具用xlrd读取sheet名称和单元格的值，进行排序。

然后利用openpyxl读取值和格式，写入到新的工作簿

界面化用ttk库实现

### 如何使用 how to use

直接运行sort_main.py 即可，所需库见requirements.txt

### 适用场景

单个xls文件中有多个sheet，每个sheet的大体格式是一样的，根据某个指定单元格作为依据对所有sheet进行排序。

注意：目前的排序依据是Q6，请修改如下代码来指定单元格

```python
sheet_obj["sheet_index"] = int(sheet.cell(5, col_index("Q") - 1).value)
```
