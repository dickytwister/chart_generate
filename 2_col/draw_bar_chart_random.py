import pandas as pd
from pyecharts.charts import Bar
from pyecharts import options as opts
from pyecharts.render import make_snapshot
from snapshot_selenium import snapshot
from PIL import Image
import os
import random

def get_bar_chart(data):
    # 假设Excel文件有两列，第一列为类别，第二列为数值
    x_label = data.iloc[0, 0]  # 第一列的标题
    y_label = data.iloc[0, 1]  # # 第二列的标题
    categories = data.iloc[1:, 0].tolist()  # 第一列（去掉标题行）
    values = data.iloc[1:, 1].tolist()  # 第二列（去掉标题行）

    # 随机参数
    font_size = random.randint(20, 30)
    item_width = random.randint(20, 80)
    item_height = random.randint(10, 40)
    pos_left = f"{random.randint(10, 50)}%"
    pos_bottom = f"{random.randint(1, 10)}%"
    rotate_angle = random.choice([0, 45, 90])
    bar_width = random.randint(10, 90)  # 将百分比设为10到90之间的随机值
    x_axis_font_size = random.randint(15, 25)
    y_axis_font_size = random.randint(15, 25)
    label_font_size = random.randint(10, 20)
    x_label_font_size = random.randint(15, 25)  # x 轴标签字体大小随机化
    y_label_font_size = random.randint(15, 25)  # y 轴标签字体大小随机化
    legend_orient = random.choice(["horizontal", "vertical"])

    # 随机颜色
    bar_colors = [f"#{random.randint(0, 0xFFFFFF):06x}" for _ in range(len(categories))]

    # 创建柱状图
    bar = Bar()
    bar.add_xaxis(categories)
    bar.add_yaxis("", values, itemstyle_opts=opts.ItemStyleOpts(color=random.choice(bar_colors)),
                  category_gap=f'{bar_width}%')

    # 设置全局选项，包括图例位置
    bar.set_global_opts(
        legend_opts=opts.LegendOpts(orient=legend_orient, pos_bottom=pos_bottom, item_width=item_width, item_height=item_height),
        xaxis_opts=opts.AxisOpts(name=x_label, axislabel_opts=opts.LabelOpts(rotate=rotate_angle, font_size=x_axis_font_size),
                                 name_textstyle_opts=opts.TextStyleOpts(font_size=x_label_font_size)),
        yaxis_opts=opts.AxisOpts(name=y_label, axislabel_opts=opts.LabelOpts(font_size=y_axis_font_size),
                                 name_textstyle_opts=opts.TextStyleOpts(font_size=y_label_font_size)),
        title_opts=opts.TitleOpts(title=''.join(data.columns[0]), is_show=True, text_align='center',
                                  pos_left=pos_left, title_textstyle_opts=opts.TextStyleOpts(font_size=font_size))
    )

    # 设置柱体内标签的字体大小
    bar.set_series_opts(label_opts=opts.LabelOpts(font_size=label_font_size))

    # 渲染成HTML文件
    make_snapshot(snapshot, bar.render(), "bar_chart.png")

    # 打开图片
    img = Image.open("bar_chart.png")

    # 创建白色背景
    white_bg = Image.new("RGB", img.size, "white")

    # 将原图粘贴到白色背景上
    white_bg.paste(img, (0, 0), img)

    # 返回白色背景图片
    return white_bg


def process_files(excel_path, output_path):
    # 遍历excel_path下的所有文件
    for filename in os.listdir(excel_path):
        if filename.endswith(".xlsx"):
            filepath = os.path.join(excel_path, filename)
            # 读取Excel文件
            df = pd.read_excel(filepath, sheet_name="Sheet1")
            # 保存为JSON文件
            json_path = os.path.join(output_path, filename.replace(".xlsx", ".json"))
            df.to_json(json_path, orient='records', force_ascii=False)
            # 获取柱状图
            img = get_bar_chart(df)
            # 保存图片
            img.save(os.path.join(output_path, filename.replace(".xlsx", ".png")))


if __name__ == "__main__":
    excel_path = "./"
    output_path = "./output"
    if not os.path.exists(output_path):
        os.makedirs(output_path)
    process_files(excel_path, output_path)
