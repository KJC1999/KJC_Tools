import xlrd
import shutil

# 打开Excel文件
workbook = xlrd.open_workbook('C:/Users/Administrator/Desktop/111/对话框修改相关.xls')

# 遍历每个sheet
for sheet_name in workbook.sheet_names():
    # 获取当前sheet
    sheet = workbook.sheet_by_name(sheet_name)
    # 获取编号所在列的索引，假设是第一列
    index = 0
    # 遍历每一行
    for row_num in range(sheet.nrows):
        # 获取当前行的编号
        num = sheet.cell_value(row_num, index)
        pName = num.split('-')[1]
        # 如果当前行没有编号，则跳过
        if not num:
            continue
        # 拼接远程路径
        remote_dir = u'F:/arsenal/通过库/ET{}{}'.format(pName, num)
        # 拼接本地路径
        local_dir = u'F:/111/{}/{}'.format(sheet_name, num)
        # 创建本地目录
        # os.makedirs(local_dir, exist_ok=True)
        # 复制文件
        shutil.copytree(remote_dir, local_dir)
