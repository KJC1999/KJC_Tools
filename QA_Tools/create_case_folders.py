# -- coding: utf-8 --
import os
import xlrd
import zipfile
from datetime import datetime


def zip_folder(folder_path):
    """
    压缩文件夹到指定路径
    :param folder_path: 要压缩的文件夹路径
    """
    output_path = os.path.join(os.path.dirname(folder_path), os.path.basename(folder_path) + '.zip')
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                zipf.write(file_path, os.path.relpath(file_path, folder_path))


def get_result(xlsPath, condition=1):
    workbook = xlrd.open_workbook(xlsPath)
    sheet = workbook.sheet_by_index(0)
    nrows = sheet.nrows
    result = []
    for i in range(1, nrows):
        # 获取测试用例编号
        case_id = sheet.cell_value(i, 0)
        rule = sheet.cell_value(i, 9)
        if not case_id:
            continue
        if condition == 0 and rule:
            continue
        if condition == 1 and not rule:
            continue
        # 获取测试步骤，若存在多个步骤，先进行切割保存
        steps = sheet.cell_value(i, 6).replace('"', '”').split('\n')
        # 获取文字描述的预期结果&去除文字描述的预期结果中的换行符
        expect = sheet.cell_value(i, 7).replace('"', '”').replace('\n', '')
        # 将每行数据存入result中
        result.append([case_id, steps, expect])
    return result


def create_case_folders(folderPath, xlsPath, condition):
    # 遍历result二维数组
    for row in get_result(xlsPath, condition):
        case_id = row[0]
        steps = row[1]
        desc = row[2]
        # 获取当前用例前缀（WPS/ET/WPP）并存放大小写
        cPrefix = case_id.split('-')[0]
        lPrefix = case_id.split('-')[0][0] + case_id.split('-')[0][1:].lower()
        # 创建测试用例目录
        case_dir = os.path.join(folderPath, case_id)
        if not os.path.exists(case_dir):
            os.makedirs(case_dir)

        # 创建case.py文件，并写入数据
        case_path = os.path.join(case_dir, 'case.py')
        with open(case_path, 'w', encoding='utf-8') as f:
            f.write('# -- coding: utf-8 --\n')
            f.write('import sys\n')
            f.write('import KatyushaDriver\n\n')
            f.write('args = sys.argv\n')
            f.write('driver = KatyushaDriver.Katyusha(args)\n')
            f.write('# 测试点或预期\n')
            f.write(f"driver.stage('{desc}', '{desc}')\n")
            f.write("driver.execute_action(action_name='LaunchWps', desc='启动WPS进程（路径启动）', app_type='EXTEND', is_must_execute=False, is_must_pass=False, args={")
            f.write(f"'ExePath': '%{lPrefix}Path%'")
            f.write("})\n")
            f.write(f"driver.execute_action(action_name='NewDocument', desc='新建空白文档', app_type='{cPrefix}', is_must_execute=False, is_must_pass=False, ")
            f.write("args={})\n")
            f.write('# 测试步骤\n')
            for step in steps:
                f.write(f"driver.stage('{step}', '{step}')\n")
            f.write(
                f"driver.execute_checkpoint(checkpoint_name='ExecuteJsFile', desc='运行 JS File文件中的Function检查点', app_type='{cPrefix}', ")
            f.write(
                "args={'jsFile': '%CaseDir%/test.js', 'runFunctionName': 'Macro1', 'Expected': '', 'arg1': '', 'arg2': '', 'arg3': '', 'arg4': '', 'arg5': '', 'arg6': '', 'arg7': '', 'arg8': '', 'arg9': '', 'arg10': ''})\n")
            f.write(f"driver.execute_action(action_name='QuitWps', desc='退出WPS进程', app_type='{cPrefix}', ")
            f.write("is_must_execute=True, is_must_pass=True, args={})\n")
            f.write('driver.quit()\n')
        # 创建test.js文件
        js_path = os.path.join(case_dir, 'test.js')
        with open(js_path, 'w', encoding='utf-8') as f:
            pass
    # 压缩文件夹
    zip_folder(folderPath)



if __name__ == '__main__':
    print('请注意！！！\n'
          '在输入前请确保Cases目录已上传用例表格文件\n'
          '默认输出仅含内核（即第三个输入可直接回车跳过）')
    xlsName = input('请输入用例xls名称（包括后缀）：')
    xlsPath = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Cases', xlsName)
    folderName = input('请输入生成后的文件夹名称（例：WPP主题）：')
    while len(folderName) == 0:
        folderName = input('文件夹名不得为空！！！！请输入生成后的文件夹名称：')
    formatted_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    folderPath = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Case_Folders', folderName+formatted_time)
    condition = int(input('是否含内核逻辑（0为非内核，1为仅内核，2为全部）：'))
    create_case_folders(folderPath, xlsPath, condition)
    print('Finish!!!')

