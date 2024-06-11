import os
import shutil

# 定义A文件夹和B文件夹的路径
source_folder = 'D:/PycharmProjects/ChartQA-main/ChartQA-main/ChartQA Dataset/test/tables'
destination_folder = 'D:/PycharmProjects/ChartQA-main/ChartQA-main/ChartQA Dataset/test/test_2_col/tables'

# 确保B文件夹存在，不存在则创建
if not os.path.exists(destination_folder):
    os.makedirs(destination_folder)

# 遍历A文件夹中的所有文件
for filename in os.listdir(source_folder):
    # 检查文件是否以"2_col"开头
    if filename.startswith("two_col"):
        source_file = os.path.join(source_folder, filename)
        destination_file = os.path.join(destination_folder, filename)
        # 复制文件到B文件夹
        shutil.copy(source_file, destination_file)
        print(f"已复制: {filename}")

print("所有文件复制完成。")
