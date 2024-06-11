import pandas as pd
from pyecharts.charts import Pie
from pyecharts import options as opts
from pyecharts.render import make_snapshot
from snapshot_selenium import snapshot
from PIL import Image
import os
import random

def get_pie_chart(data):
    # 假设Excel文件有两列，第一列为类别，第二列为数值
    categories = data.iloc[:, 0].tolist()
    values = data.iloc[:, 1].tolist()

    # 随机参数
    font_size = random.randint(15, 25)
    item_width = random.randint(20, 60)
    item_height = random.randint(10, 40)
    pos_left = f"{random.randint(20, 80)}%"
    pos_bottom = f"{random.randint(1, 10)}%"
    radius_inner = f"{random.randint(20, 50)}%"
    radius_outer = f"{random.randint(60, 70)}%"
    legend_font_size = random.randint(15, 25)  # 图例字体大小随机化

    # 创建饼图
    pie = Pie()
    pie.add(
        "",
        [list(z) for z in zip(categories, values)],
        radius=[radius_inner, radius_outer],
        label_opts=opts.LabelOpts(is_show=True, formatter="{d}%", font_size=font_size),  # {b}: {c} ({d}%)
    )

    # 设置全局选项，包括图例位置
    pie.set_global_opts(
        legend_opts=opts.LegendOpts(orient="horizontal", pos_bottom=pos_bottom, item_width=item_width, item_height=item_height, textstyle_opts=opts.TextStyleOpts(font_size=legend_font_size)),
        # 注意：如果希望图例在底部，应使用pos_bottom而非pos_top
    )
    # print(data.columns[0:2])
    # 设置全局选项
    pie.set_global_opts(title_opts=opts.TitleOpts(title=''.join(data.columns[0]),
                                                  is_show=True,
                                                  text_align='center',
                                                  pos_left=pos_left,
                                                  title_textstyle_opts=opts.TextStyleOpts(font_size=30)
                                                  )
                        )

    # 将饼图渲染成图片
    make_snapshot(snapshot, pie.render(), "pie_chart.png")

    # 打开图片
    img = Image.open("pie_chart.png")

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
            # 获取饼图
            img = get_pie_chart(df)
            # 保存图片
            img.save(os.path.join(output_path, filename.replace(".xlsx", ".png")))

if __name__ == "__main__":
    excel_path = "./"
    output_path = "./output"
    process_files(excel_path, output_path)
