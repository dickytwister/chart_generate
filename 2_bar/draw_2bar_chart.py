import pandas as pd
from pyecharts.charts import Bar, Grid
from pyecharts import options as opts
from pyecharts.render import make_snapshot
from snapshot_selenium import snapshot
from PIL import Image
import os

def get_bar_chart(data):
    # 提取数据
    title = data.columns[0]
    x_axis_name = data.iloc[0, 0]
    bar1_y_axis_name = data.iloc[0, 1]
    bar2_y_axis_name = data.iloc[0, 2]

    categories = data.iloc[1:, 0].tolist()
    bar1_data = data.iloc[1:, 1].tolist()
    bar2_data = data.iloc[1:, 2].tolist()

    # 创建柱状图
    bar = (
        Bar()
        .add_xaxis(categories)
        .add_yaxis(bar1_y_axis_name, bar1_data, category_gap="40%")
        .add_yaxis(bar2_y_axis_name, bar2_data, category_gap="40%")
        .set_global_opts(
            title_opts=opts.TitleOpts(title=title),
            legend_opts=opts.LegendOpts(pos_left='45%'),
            tooltip_opts=opts.TooltipOpts(trigger='axis', axis_pointer_type='shadow'),
            toolbox_opts=opts.ToolboxOpts(is_show=False),
            xaxis_opts=opts.AxisOpts(name=x_axis_name),  # 添加x轴名称
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
