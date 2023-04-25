import os


def create_folders(start_num, num_folders, directory):
    current_num = int(start_num.split("-")[-1])
    prefix = "-".join(start_num.split("-")[:-1])
    for i in range(num_folders):
        folder_name = f"{prefix}-{str(current_num).zfill(8)}"
        folder_path = os.path.join(directory, folder_name)
        try:
            os.makedirs(folder_path)
        except FileExistsError:
            print('当文件已存在时，无法创建该文件。')
        with open(os.path.join(folder_path, "test.js"), "w") as f:
            f.write("")
        with open(os.path.join(folder_path, "case.py"), "w") as f:
            f.write("")
        current_num += 1


start_num = "WPP-BJ-00000253"
num_folders = 2
directory = "E:\KJC的案例\WPP背景"

create_folders(start_num, num_folders, directory)
