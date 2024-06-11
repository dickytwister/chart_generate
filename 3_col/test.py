import pandas as pd
from pyecharts.charts import Bar, Line, Grid
from pyecharts import options as opts
from pyecharts.render import make_snapshot
from snapshot_selenium import snapshot
from PIL import Image
import os


def get_bar_line_chart(data):
    # 提取数据
    title = data.columns[0]
    x_axis_name = data.iloc[0, 0]
    bar_y_axis_name = data.iloc[0, 1]
    line_y_axis_name = data.iloc[0, 2]

    categories = data.iloc[1:, 0].tolist()
    bar_data = data.iloc[1:, 1].tolist()
    line_data = data.iloc[1:, 2].tolist()

    # 创建柱状图
    bar = (
        Bar()
        .add_xaxis(categories)
        .add_yaxis(bar_y_axis_name, bar_data, z=0, label_opts=opts.LabelOpts(is_show=False))
        .extend_axis(
            yaxis=opts.AxisOpts(
                name=bar_y_axis_name,
                type_="value",
                position="left",
            )
        )
        .extend_axis(
            yaxis=opts.AxisOpts(
                name=line_y_axis_name,
                type_="value",
                position="right",
            )
        )
        .set_global_opts(
            title_opts=opts.TitleOpts(title=title),
            legend_opts=opts.LegendOpts(pos_left='45%'),
            tooltip_opts=opts.TooltipOpts(trigger='axis', axis_pointer_type='cross'),
            toolbox_opts=opts.ToolboxOpts(is_show=False),
        )
    )

    # 创建折线图
    line = (
        Line()
        .add_xaxis(categories)
        .add_yaxis(
            series_name=line_y_axis_name,
            y_axis=line_data,
            yaxis_index=2,
            z=2
        )
    )

    bar.overlap(line)

    # 渲染成HTML文件
    make_snapshot(snapshot, bar.render(), "bar_line_chart.png")

    # 打开图片
    img = Image.open("bar_line_chart.png")

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
            img = get_bar_line_chart(df)
            # 保存图片
            img.save(os.path.join(output_path, filename.replace(".xlsx", ".png")))


if __name__ == "__main__":
    excel_path = "../3_col"
    output_path = "../output"
    if not os.path.exists(output_path):
        os.makedirs(output_path)
    process_files(excel_path, output_path)