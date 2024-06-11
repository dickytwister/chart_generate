import pandas as pd
from pyecharts.charts import Line
from pyecharts import options as opts
from pyecharts.render import make_snapshot
from snapshot_selenium import snapshot
from PIL import Image
import os
import random

def get_line_chart(data):
    # 假设Excel文件有两列，第一列为类别，第二列为数值
    # print(data.iloc[0, 0])
    # print(data.iloc[0, 1])
    x_label = data.iloc[0, 0]  # 第一列的标题
    y_label = data.iloc[0, 1]  #  # 第二列的标题
    categories = data.iloc[1:, 0].tolist()  # 第一列（去掉标题行）
    values = data.iloc[1:, 1].tolist()  # 第二列（去掉标题行）

    # 随机参数
    line_colors = ["#c23531", "#2f4554", "#61a0a8", "#d48265", "#91c7ae", "#749f83", "#ca8622", "#bda29a", "#6e7074",
                   "#546570"]
    legend_orients = ["horizontal", "vertical"]

    line_color = random.choice(line_colors)
    legend_orient = random.choice(legend_orients)
    pos_left = f"{random.randint(20, 80)}%"
    pos_bottom = f"{random.randint(1, 10)}%"
    font_size = random.randint(20, 30)
    line_width = random.randint(1, 5)
    x_label_font_size = random.randint(15, 25)  # x 轴标签字体大小随机化
    y_label_font_size = random.randint(15, 25)  # y 轴标签字体大小随机化
    label_font_size = random.randint(15, 20)  # 折线上方数字字体大小随机化

    # 创建折线图
    line = Line()
    line.add_xaxis(categories)
    line.add_yaxis("", values, linestyle_opts=opts.LineStyleOpts(color=line_color, width=line_width),
                   label_opts=opts.LabelOpts(font_size=label_font_size, position="top"))

    # 设置全局选项，包括图例位置
    line.set_global_opts(
        legend_opts=opts.LegendOpts(orient=legend_orient, pos_bottom=pos_bottom, item_width=56, item_height=20),
        xaxis_opts=opts.AxisOpts(name=x_label, axislabel_opts=opts.LabelOpts(rotate=0, font_size=font_size),
                                 name_textstyle_opts=opts.TextStyleOpts(font_size=x_label_font_size)),
        yaxis_opts=opts.AxisOpts(name=y_label, axislabel_opts=opts.LabelOpts(font_size=font_size),
                                 name_textstyle_opts=opts.TextStyleOpts(font_size=y_label_font_size)),
        title_opts=opts.TitleOpts(title=''.join(data.columns[0]), is_show=True, text_align='center',
                                  pos_left=pos_left, title_textstyle_opts=opts.TextStyleOpts(font_size=font_size))
    )

    # 将折线图渲染成图片
    make_snapshot(snapshot, line.render(), "line_chart.png")

    # 打开图片
    img = Image.open("line_chart.png")

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
            # 获取折线图
            img = get_line_chart(df)
            # 保存图片
            img.save(os.path.join(output_path, filename.replace(".xlsx", ".png")))


if __name__ == "__main__":
    excel_path = "./"
    output_path = "./output"
    if not os.path.exists(output_path):
        os.makedirs(output_path)
    process_files(excel_path, output_path)
