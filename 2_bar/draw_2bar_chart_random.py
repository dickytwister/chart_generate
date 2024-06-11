import pandas as pd
from pyecharts.charts import Bar, Grid
from pyecharts import options as opts
from pyecharts.render import make_snapshot
from snapshot_selenium import snapshot
from PIL import Image
import os
import random

def get_bar_chart(data):
    # 提取数据
    title = data.columns[0]
    x_axis_name = data.iloc[0, 0]
    bar1_y_axis_name = data.iloc[0, 1]
    bar2_y_axis_name = data.iloc[0, 2]

    categories = data.iloc[1:, 0].tolist()
    bar1_data = data.iloc[1:, 1].tolist()
    bar2_data = data.iloc[1:, 2].tolist()

    # 生成随机样式
    colors = ["#"+''.join([random.choice('0123456789ABCDEF') for j in range(6)]) for i in range(2)]
    bar1_color = colors[0]
    bar2_color = colors[1]
    bar_width = random.uniform(0.2, 0.8)
    label_font_size = random.randint(10, 20)
    title_font_size = random.randint(20, 30)
    legend_font_size = random.randint(10, 20)
    x_label_rotation = random.randint(0, 90)
    bar_label_font_size = random.randint(10, 20)  # 数字字体大小

    # 创建柱状图
    bar = (
        Bar()
        .add_xaxis(categories)
        .add_yaxis(bar1_y_axis_name, bar1_data, category_gap=f"{bar_width * 100}%", itemstyle_opts=opts.ItemStyleOpts(color=bar1_color))
        .add_yaxis(bar2_y_axis_name, bar2_data, category_gap=f"{bar_width * 100}%", itemstyle_opts=opts.ItemStyleOpts(color=bar2_color))
        .set_global_opts(
            title_opts=opts.TitleOpts(title=title, title_textstyle_opts=opts.TextStyleOpts(font_size=title_font_size)),
            legend_opts=opts.LegendOpts(pos_left='45%', textstyle_opts=opts.TextStyleOpts(font_size=legend_font_size)),
            tooltip_opts=opts.TooltipOpts(trigger='axis', axis_pointer_type='shadow'),
            toolbox_opts=opts.ToolboxOpts(is_show=False),
            xaxis_opts=opts.AxisOpts(name=x_axis_name, axislabel_opts=opts.LabelOpts(rotate=x_label_rotation, font_size=label_font_size)),
            yaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(font_size=label_font_size))
        )
        .set_series_opts(
            label_opts=opts.LabelOpts(position='inside', font_size=bar_label_font_size)
        )
    )

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
