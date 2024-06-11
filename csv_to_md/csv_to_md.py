import os
import pandas as pd

# 定义A文件夹和B文件夹的路径
source_folder = 'D:/PycharmProjects/ChartQA-main/ChartQA-main/ChartQA Dataset/2_col/test/tables'
destination_folder = 'D:/PycharmProjects/ChartQA-main/ChartQA-main/ChartQA Dataset/2_col/test/markdown'

# 确保B文件夹存在，不存在则创建
if not os.path.exists(destination_folder):
    os.makedirs(destination_folder)

# 遍历A文件夹中的所有文件
for filename in os.listdir(source_folder):
    if filename.endswith('.csv'):
        source_file = os.path.join(source_folder, filename)

        # 读取CSV文件
        df = pd.read_csv(source_file)

        # 获取标题和标签
        title = df.columns[1]
        label1 = df.columns[0]
        label2 = "Value"

        # 创建Markdown表格
        markdown_table = f"# {title}\n\n"
        markdown_table += f"| {label1} | {label2} |\n"
        markdown_table += "|----------------|-------|\n"

        # 填充Markdown表格
        for index, row in df.iterrows():
            characteristic = row[label1]
            value = row[title]
            markdown_table += f"| {characteristic} | {value} |\n"

        # 输出Markdown文件到B文件夹
        output_file = os.path.join(destination_folder, filename.replace('.csv', '.md'))
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(markdown_table)

print("所有文件处理完成。")
