import os
import webbrowser
import time
import datetime
import openpyxl
from pyecharts import options as opts
from pyecharts.charts import Map
lastModifyDate = time.strftime("%Y-%m-%d", time.localtime(os.path.getmtime('疫情通报.xlsx')))
lastValidDate = ((datetime.datetime.strptime(lastModifyDate, "%Y-%m-%d")) + datetime.timedelta(days = -1)).strftime('%Y-%m-%d')
firstDate = time.strptime(lastValidDate, "%Y-%m-%d")
finalDate = time.strptime('2020-05-16', "%Y-%m-%d")
date1 = datetime.datetime.strptime(lastValidDate, "%Y-%m-%d").date()

province = ['日期', '本土', '福建', '北京', '河北', '山西', '辽宁', '吉林', '黑龙江', '江苏', '浙江', '安徽', '江西', '山东', '河南', '湖北', '湖南', '广东', '海南', '四川',
            '贵州', '云南', '陕西', '甘肃', '内蒙古', '青海', '广西', '西藏', '宁夏', '新疆', '天津', '上海', '重庆', '兵团']
gat = ['日期', '香港', '澳门', '台湾']
wb = openpyxl.load_workbook('疫情通报.xlsx')
def MAP(date):
    date2 = datetime.datetime.strptime(date, "%Y-%m-%d").date()

    dateIndex = int((date1 - date2).days) + 2
    sheet = wb['本土每日新增确诊']
    newDiagnosed1 = [(province[x], sheet[dateIndex][x].value) for x in range(len(province))]
    sheet = wb['港澳台每日新增确诊']
    newDiagnosed2 = [(gat[x], sheet[dateIndex][x].value) for x in [1,2,3]]
    print('新增确诊', end='')
    print(newDiagnosed1[1:-1] + newDiagnosed2)
    sheet = wb['本土每日新增无症状']
    newAsymptomatic = [(province[x], sheet[dateIndex][x].value) for x in range(len(province))]
    print("新增无症状", end='')
    print(newAsymptomatic[1:-1])
    html = '{}疫情通报.html'.format(date)
    if os.path.exists(html) == False:
        map = (
            Map()
            .add("新增确诊", newDiagnosed1[2:-1] + newDiagnosed2, "china")
            .add("新增无症状", newAsymptomatic[2:-1], "china")
            .set_global_opts(title_opts=opts.TitleOpts(title=('{}疫情通报'.format(newDiagnosed1[0][1])),
                                                       subtitle="兵团确诊{}例， 无症状{}例\n无症状无统计港澳台".format(newDiagnosed1[-1][1], newAsymptomatic[-1][1])),
                             visualmap_opts=opts.VisualMapOpts(
                                 is_piecewise=True,
                                 pieces=[
                                     {"max": 10, "min": 1, "label": "1-10人"},
                                     {"max": 50, "min": 11, "label": "11-50人"},
                                     {"max": 100, "min": 51, "label": "51-100人"},
                                     {"max": 500, "min": 101, "label": "101-500人"},
                                     {"max": 1000, "min": 501, "label": "501-1000人"},
                                     {"max": 5000, "min": 1001, "label": "1001-5000人"},
                                     {"max": 9999999999999, "min": 5001, "label": "5000人以上"},
                                 ])
                             )
            .render(html)
        )
    webbrowser.open_new(html)

def visualize():
    while True:
        date = input('请输入日期(2020-05-16至{})，格式为yyyy-mm-dd:'.format(lastValidDate))
        try:
            t = time.strptime(date, "%Y-%m-%d")
        except:
            print('无效输入，请重试')
            continue
        if t > firstDate or t < finalDate:
            print('无效输入，请重试')
            continue
        print(date)
        MAP(date)
