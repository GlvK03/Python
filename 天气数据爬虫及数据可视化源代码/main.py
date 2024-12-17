import os
import xlrd
import xlwt
from lxml import etree
from xlutils.copy import copy
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time


class TianQi:
    def __init__(self):
        self.city_dict = {
            "青岛": "qingdao"
        }
        self.headers = {
            'authority': 'lishi.tianqi.com',
            'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
            'accept-language': 'zh-CN,zh;q=0.9',
            'cache-control': 'no-cache',
            'cookie': 'Hm_lvt_7c50c7060f1f743bccf8c150a646e90a=1701184759; Hm_lvt_30606b57e40fddacb2c26d2b789efbcb=1701184793; Hm_lpvt_30606b57e40fddacb2c26d2b789efbcb=1701184932; Hm_lpvt_7c50c7060f1f743bccf8c150a646e90a=1701185017',
            'pragma': 'no-cache',
            'referer': 'https://lishi.tianqi.com/ankang/202409.html',
            'sec-ch-ua': '"Google Chrome";v="119", "Chromium";v="119", "Not?A_Brand";v="24"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'document',
            'sec-fetch-mode': 'navigate',
            'sec-fetch-site': 'same-origin',
            'sec-fetch-user': '?1',
            'upgrade-insecure-requests': '1',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
        }

    def clickjs(self, url):
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--disable-gpu')
        browser = webdriver.Chrome(options=chrome_options)
        browser.implicitly_wait(4)
        browser.get(url)
        time.sleep(1)

        # 检查是否存在“查询更多”的按钮并点击
        try:
            more_button = browser.find_element(By.XPATH, '//div[@class="lishidesc2"]')
            if more_button:
                more_button.click()
                time.sleep(2)  # 等待页面加载
        except Exception as e:
            print("没有找到'查询更多'按钮或无法点击:", e)

        source = browser.page_source
        browser.quit()
        return source

    def spider(self, url=None):
        if url is None:
            city = '青岛'
            city_code = self.city_dict[city]
            year = '2024'
            month = '11'
            url = f'https://lishi.tianqi.com/{city_code}/{year}{month}.html'

        # 使用 Selenium 获取页面内容
        source = self.clickjs(url)
        tree = etree.HTML(source)
        datas = tree.xpath(
            "/html/body/div[@class='main clearfix']/div[@class='main_left inleft']/div[@class='tian_three']/ul[@class='thrui']/li")
        weizhi = tree.xpath(
            "/html/body/div[@class='main clearfix']/div[@class='main_left inleft']/div[@class='inleft_tian']/div[@class='tian_one']/div[@class='flex'][1]/h3/text()")[
            0]

        self.parase(datas, weizhi, year, month)

    def parase(self, datas, weizhi, year, month):
        for data in datas:
            datetime = data.xpath("./div[@class='th200']/text()")[0]
            max_qiwen = data.xpath("./div[@class='th140'][1]/text()")[0]
            min_qiwen = data.xpath("./div[@class='th140'][2]/text()")[0]
            tianqi = data.xpath("./div[@class='th140'][3]/text()")[0]
            fengxiang = data.xpath("./div[@class='th140'][4]/text()")[0]
            dict_tianqi = {
                '日期': datetime,
                '最高气温': max_qiwen,
                '最低气温': min_qiwen,
                '天气': tianqi,
                '风向': fengxiang
            }
            data_excel = {
                f'{weizhi}【{year}年{month}月】': [datetime, max_qiwen, min_qiwen, tianqi, fengxiang]
            }
            self.chucun_excel(data_excel, weizhi, year, month)
            print(dict_tianqi)

        print(f"Data length: {len(datas)}")

    def chucun_excel(self, data, weizhi, year, month):
        file_name = f'{weizhi}【{year}年{month}月】.xls'
        if not os.path.exists(file_name):
            wb = xlwt.Workbook(encoding='utf-8')
            sheet = wb.add_sheet(f'{weizhi}【{year}年{month}月】', cell_overwrite_ok=True)
            borders = xlwt.Borders()
            borders.left = xlwt.Borders.THIN
            borders.right = xlwt.Borders.THIN
            borders.top = xlwt.Borders.THIN
            borders.bottom = xlwt.Borders.THIN
            borders.left_colour = 0x40
            borders.right_colour = 0x40
            borders.top_colour = 0x40
            borders.bottom_colour = 0x40
            style = xlwt.XFStyle()
            style.borders = borders
            align = xlwt.Alignment()
            align.horz = 0x02
            align.vert = 0x01
            style.alignment = align
            header = ('日期', '最高气温', '最低气温', '天气', '风向')
            for i in range(len(header)):
                sheet.col(i).width = 2560 * 3
                sheet.write(0, i, header[i], style)
            wb.save(file_name)

        wb = xlrd.open_workbook(file_name)
        worksheet = wb.sheet_by_name(f'{weizhi}【{year}年{month}月】')
        rows_old = worksheet.nrows
        new_workbook = copy(wb)
        new_worksheet = new_workbook.get_sheet(f'{weizhi}【{year}年{month}月】')

        # 检查日期是否已存在，如果存在则覆盖数据
        date_exists = False
        for row in range(1, rows_old):
            if worksheet.cell_value(row, 0) == data[f'{weizhi}【{year}年{month}月】'][0]:
                date_exists = True
                for num in range(len(data[f'{weizhi}【{year}年{month}月】'])):
                    new_worksheet.write(row, num, data[f'{weizhi}【{year}年{month}月】'][num])
                break

        if not date_exists:
            for num in range(len(data[f'{weizhi}【{year}年{month}月】'])):
                new_worksheet.write(rows_old, num, data[f'{weizhi}【{year}年{month}月】'][num])

        new_workbook.save(file_name)


if __name__ == '__main__':
    t = TianQi()
    t.spider()

import pandas as pd
from pyecharts.charts import Scatter, Pie
from pyecharts import options as opts
from scipy import stats
import calendar

# 读取数据
df = pd.read_excel('青岛历史天气【2024年11月】.xls')
print(df.columns)  # 打印列名以检查

# 获取指定月份的天数
year = 2024
month = 11
_, num_days = calendar.monthrange(year, month)

if len(df) != num_days:
    print(f"警告：数据不完整，只有 {len(df)} 天的数据。")
else:
    print(f"数据完整，共有 {num_days} 天的数据。")

# 使用 jieba 处理数据，去除 "C"
if '最高气温' in df.columns and '最低气温' in df.columns:
    df['最高气温'] = df['最高气温'].str.replace('℃', '').astype(float)
    df['最低气温'] = df['最低气温'].str.replace('℃', '').astype(float)

    # 创建散点图
    scatter = Scatter()
    scatter.add_xaxis(df['最低气温'].tolist())
    scatter.add_yaxis("最高气温", df['最高气温'].tolist())
    scatter.set_global_opts(title_opts=opts.TitleOpts(title="最低气温与最高气温的散点图"))
    scatter_html = scatter.render_embed()

    slope, intercept, r_value, p_value, std_err = stats.linregress(df['最低气温'], df['最高气温'])
    analysis_text_scatter = f"回归方程为：y = {slope:.2f}x + {intercept:.2f}"

    # 生成散点图HTML文件
    scatter_html_content = f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>天气数据分析 - 散点图</title>
    <style>
     body {{
            background-color: #f0f0f0;
            font-family: 'Arial', sans-serif;
        }}
        .navbar {{
            background-color: #4CAF50;
            overflow: hidden;
            padding: 10px;
            margin-bottom: 20px;
            display: flex;
            justify-content: space-around;
        }}
        .navbar a {{
            float: left;
            display: block;
            color: white;
            text-align: center;
            padding: 14px 16px;
            text-decoration: none;
            font-size: 18px;
            transition: background-color 0.3s;
        }}
        .navbar a:hover {{
            background-color: #45a049;
        }}
        .chart-container {{
            background-color: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            width: 80%;
            max-width: 1200px;
        }}
    </style>
</head>
<body>
    <div class="navbar">
        <a href="青岛历史天气【{year}年{month}月】温度散点图.html">散点图分析</a>
        <a href="青岛历史天气【{year}年{month}月】天气分布饼图.html">天气分布饼图</a>
        <a href="青岛历史天气【{year}年{month}月】词云图.html">词云图分析</a>
        <a href="青岛历史天气【{year}年{month}月】温度趋势图.html">温度趋势图</a>
    </div>
    <div class="chart-container">
        <div>{scatter_html}</div>
         <h2>最低气温与最高气温的散点图</h2>
        <p>{analysis_text_scatter}</p>
    </div>
</body>
</html>
"""
    with open("青岛历史天气【2024年11月】温度散点图.html", "w", encoding="utf-8") as file:
        file.write(scatter_html_content)

    # 生成饼图
    gender_counts = df['天气'].value_counts()
    total = gender_counts.sum()
    percentages = {gender: count / total * 100 for gender, count in gender_counts.items()}
    analysis_parts = []
    for gender, percentage in percentages.items():
        analysis_parts.append(f"{gender}天气占比为{percentage:.2f}%，")
    analysis_text_pie = "天气比例饼状图显示，" + ''.join(analysis_parts)

    pie = Pie(init_opts=opts.InitOpts(bg_color='#e4cf8e'))  # 移除了 ThemeType
    pie.add(
        series_name="青岛市天气分布",
        data_pair=[list(z) for z in zip(gender_counts.index.tolist(), gender_counts.values.tolist())],
        rosetype="radius",
        radius=["40%", "70%"],
        label_opts=opts.LabelOpts(is_show=True, position="outside", font_size=14,
                                  formatter="{a}<br/>{b}: {c} ({d}%)")
    )
    pie.set_global_opts(
        title_opts=opts.TitleOpts(title="青岛市11月份天气分布", pos_right="50%"),
        legend_opts=opts.LegendOpts(orient="vertical", pos_top="15%", pos_left="2%"),
        toolbox_opts=opts.ToolboxOpts(is_show=True)
    )
    pie.set_series_opts(label_opts=opts.LabelOpts(formatter="{b}: {c} ({d}%)"))
    pie_html = pie.render_embed()

    # 生成饼图HTML文件
    pie_html_content = f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>天气数据分析 - 饼图</title>
    <style>
        body {{
            background-color: #f0f0f0;
            font-family: 'Arial', sans-serif;
        }}
        .navbar {{
            background-color: #4CAF50;
            overflow: hidden;
            padding: 10px;
            margin-bottom: 20px;
            display: flex;
            justify-content: space-around;
        }}
        .navbar a {{
            float: left;
            display: block;
            color: white;
            text-align: center;
            padding: 14px 16px;
            text-decoration: none;
            font-size: 18px;
            transition: background-color 0.3s;
        }}
        .navbar a:hover {{
            background-color: #45a049;
        }}
        .chart-container {{
            background-color: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            width: 80%;
            max-width: 1200px;
        }}
    
    </style>
</head>
<body>
    <div class="navbar">
        <a href="青岛历史天气【{year}年{month}月】温度散点图.html">散点图分析</a>
        <a href="青岛历史天气【{year}年{month}月】天气分布饼图.html">天气分布饼图</a>
        <a href="青岛历史天气【{year}年{month}月】词云图.html">词云图分析</a>
        <a href="青岛历史天气【{year}年{month}月】温度趋势图.html">温度趋势图</a>
    </div>
    <div class="chart-container">
        <div>{pie_html}</div>
        <h3>分析报告：</h3>
        <h2>青岛市11月份天气分布</h2>
        <p>{analysis_text_pie}</p>
    </div>
</body>
</html>
"""
    with open("青岛历史天气【2024年11月】天气分布饼图.html", "w", encoding="utf-8") as file:
        file.write(pie_html_content)

else:
    print("列名不匹配，请检查Excel文件的列名。")

from pyecharts.charts import WordCloud
from pyecharts import options as opts
from pyecharts.globals import SymbolType
import pandas as pd
from collections import Counter
import jieba

# 读取Excel文件
df = pd.read_excel('青岛历史天气【2024年11月】.xls')
# 提取商品名
word_names = df["风向"].tolist() + df["天气"].tolist()
# 提取关键字
seg_list = [jieba.lcut(text) for text in word_names]
words = [word for seg in seg_list for word in seg if len(word) > 1]
word_counts = Counter(words)
word_cloud_data = [(word, count) for word, count in word_counts.items()]

# 创建词云图
wordcloud = (
    WordCloud(init_opts=opts.InitOpts(bg_color='#00FFFF'))
    .add("", word_cloud_data, word_size_range=[20, 100], shape=SymbolType.DIAMOND,
         word_gap=5, rotate_step=45,
         textstyle_opts=opts.TextStyleOpts(font_family='cursive', font_size=15))
    .set_global_opts(title_opts=opts.TitleOpts(title="青岛历史天气【2024年11月】词云图", pos_top="5%", pos_left="center"),
                     toolbox_opts=opts.ToolboxOpts(
                         is_show=True,
                         feature={
                             "saveAsImage": {},
                             "dataView": {},
                             "restore": {},
                             "refresh": {}
                         }
                     )

                     )
)
# 渲染词图到HTML文件
wordcloud.render("青岛历史天气【2024年11月】词云图.html")

import pandas as pd
from pyecharts.charts import Line
from pyecharts import options as opts

# 读取数据
df = pd.read_excel('青岛历史天气【2024年11月】.xls')

# 使用 jieba 处理数据，去除 "C"
df['最高气温'] = df['最高气温'].str.replace('℃', '').astype(float)
df['最低气温'] = df['最低气温'].str.replace('℃', '').astype(float)

# 创建折线图
line = Line()
line.add_xaxis(df['日期'].tolist())
line.add_yaxis("最高气温", df['最高气温'].tolist(), color="red")
line.add_yaxis("最低气温", df['最低气温'].tolist(), color="skyblue")
line.set_global_opts(
    title_opts=opts.TitleOpts(title="2024年11月最高气温与最低气温趋势"),
    xaxis_opts=opts.AxisOpts(name="日期"),
    yaxis_opts=opts.AxisOpts(name="温度（单位°C）"),
    legend_opts=opts.LegendOpts(pos_top="10%"),
)

# 渲染图表到 HTML 文件
line_html = line.render_embed()

# 生成 HTML 文件内容
html_content = f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>天气数据分析 - 温度趋势图</title>
    <style>
    body {{
            background-color: #f0f0f0;
            font-family: 'Arial', sans-serif;
        }}
        .navbar {{
            background-color: #4CAF50;
            overflow: hidden;
            padding: 10px;
            margin-bottom: 20px;
            display: flex;
            justify-content: space-around;
        }}
        .navbar a {{
            float: left;
            display: block;
            color: white;
            text-align: center;
            padding: 14px 16px;
            text-decoration: none;
            font-size: 18px;
            transition: background-color 0.3s;
        }}
        .navbar a:hover {{
            background-color: #45a049;
        }}
        .chart-container {{
            background-color: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            width: 80%;
            max-width: 1200px;
        }}
    </style>

</head>
<body>
    <div class="navbar">
        <a href="青岛历史天气【{year}年{month}月】温度散点图.html">散点图分析</a>
        <a href="青岛历史天气【{year}年{month}月】天气分布饼图.html">天气分布饼图</a>
        <a href="青岛历史天气【{year}年{month}月】词云图.html">词云图分析</a>
        <a href="青岛历史天气【{year}年{month}月】温度趋势图.html">温度趋势图</a>
    </div>
    <div class="chart-container">
        <div>{line_html}</div>
    </div>
</body>
</html>
"""

# 保存 HTML 文件
with open("青岛历史天气【2024年11月】温度趋势图.html", "w", encoding="utf-8") as file:
    file.write(html_content)

import pandas as pd
from pyecharts.charts import WordCloud
from pyecharts import options as opts
from pyecharts.globals import SymbolType
from collections import Counter
import jieba
import re

# 读取Excel文件
df = pd.read_excel("青岛历史天气【2024年11月】.xls")

# 提取天气状况、风速和温度作为词云的数据源
text_data = df['天气'].astype(str).tolist() + df['风向'].astype(str).tolist()

# 清洗数据，移除特殊字符
cleaned_data = [re.sub(r'[^\w\s]', '', line) for line in text_data]

# 使用jieba进行中文分词
words = []
for line in cleaned_data:
    words.extend(jieba.lcut(line))

# 过滤掉包含“级”的词条
words = [word for word in words if '级' not in word]

# 统计词频
word_counts = Counter(words)

# 准备词云图的数据
word_cloud_data = [(word, count) for word, count in word_counts.items()]

# 创建词云图
wordcloud = (
    WordCloud(init_opts=opts.InitOpts(bg_color='#e4cf8e'))
    .add("", word_cloud_data, word_size_range=[20, 100], shape=SymbolType.DIAMOND,
         word_gap=5, rotate_step=45,
         textstyle_opts=opts.TextStyleOpts(font_family='cursive', font_size=15))
    .set_global_opts(title_opts=opts.TitleOpts(title="青岛历史天气【2024年11月】词云图", pos_top="5%", pos_left="center"),
                     toolbox_opts=opts.ToolboxOpts(
                         is_show=True,
                         feature=opts.ToolBoxFeatureOpts(
                             save_as_image=opts.ToolBoxFeatureSaveAsImageOpts(),
                             data_view=opts.ToolBoxFeatureDataViewOpts(),
                             restore=opts.ToolBoxFeatureRestoreOpts(),
                         )
                     )
    )
)

# 渲染词云图为HTML内容
html_content = wordcloud.render_embed()

# 生成完整的HTML文件内容
complete_html = f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <title>天气数据分析 - 词云图</title>
    <style>
         body {{
            background-color: #f0f0f0;
            font-family: 'Arial', sans-serif;
        }}
        .navbar {{
            background-color: #4CAF50;
            overflow: hidden;
            padding: 10px;
            margin-bottom: 20px;
            display: flex;
            justify-content: space-around;
        }}
        .navbar a {{
            float: left;
            display: block;
            color: white;
            text-align: center;
            padding: 14px 16px;
            text-decoration: none;
            font-size: 18px;
            transition: background-color 0.3s;
        }}
        .navbar a:hover {{
            background-color: #45a049;
        }}
        .chart-container {{
            background-color: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            width: 80%;
            max-width: 1200px;
        }}
    </style>
</head>
<body>
    <div class="navbar">
        <a href="青岛历史天气【{year}年{month}月】温度散点图.html">散点图分析</a>
        <a href="青岛历史天气【{year}年{month}月】天气分布饼图.html">天气分布饼图</a>
        <a href="青岛历史天气【{year}年{month}月】词云图.html">词云图分析</a>
        <a href="青岛历史天气【{year}年{month}月】温度趋势图.html">温度趋势图</a>
    </div>
    <!-- 包含图表 -->
    
        <div class="chart-container">{html_content}</div>
    
</body>
</html>
"""

# 保存HTML文件
with open("青岛历史天气【2024年11月】词云图.html", "w", encoding="utf-8") as file:
    file.write(complete_html)