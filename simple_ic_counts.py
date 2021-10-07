import pandas as pd
import time
from numpy import nan as NA


# 定义将数据转化为时间戳的函数
def timetrans(s):
    return (s.apply(lambda x: str(x)).apply(lambda x: time.strptime(
        x, "%Y-%m-%d %H:%M:%S")).apply(lambda x: time.mktime(x)))


# 选择ic卡数据、到离站时间数据、站点数据、车牌号、写入的文件
path_ic = '96/96_ic.xlsx'
path_dl = '96/0412_dl_96.xlsx'
path_station = '96/0412_station_96.xlsx'
writer = pd.ExcelWriter('test_data.xlsx')
# 选择日期
d = 12

# 输入车牌号及车辆对应序号
car = pd.DataFrame({
    '12': ['00C270', '00D015', '00D050', '00D601', NA, NA, NA, NA, NA],
    '13': ['00C270', '00D015', '00D321', '00D570', '00D601', NA, NA, NA, NA],
    '14': [
        '00C270', '00D015', '00D321', '00D570', '00E560', '00E590', '00E780',
        '00E783', '00E878'
    ],
    '15': [
        '00C270', '00D015', '00D050', '00D570', '00E560', '00E590', '00E780',
        '00E783', '00E878'
    ]
})
num = pd.DataFrame({
    '12': [0, 1, 2, 5, NA, NA, NA, NA, NA],
    '13': [0, 1, 3, 4, 5, NA, NA, NA, NA],
    '14': [0, 1, 3, 4, 6, 7, 8, 9, 10],
    '15': [0, 1, 2, 4, 5, 6, 7, 9, 10]
})
# 选择对应日期的车辆
car_number = [x for x in car['{}'.format(d)].dropna()]
number = [int(x) for x in num['{}'.format(d)].dropna()]

# 导入ic卡数据
ic_data = pd.read_excel(path_ic, 'Sheet1', usecols=number)
ic_data.columns = list(range(len(car_number)))
# 导入到离站时间数据
dl = pd.read_excel(path_dl, 'Sheet1')
# 导入站点标签数据
label1 = pd.read_excel(path_station, 'Sheet1')

# 循环计数并写入文件
for n in range(len(car_number)):

    # 转化日期格式
    timetrans(ic_data[n].dropna())
    timetrans(dl[n].dropna())
    # 以站点数据作为分割
    lb1 = list(label1[n].dropna())
    data1 = pd.cut(ic_data[n], dl[n], labels=lb1, ordered=False)
    # 计数
    datas = data1.value_counts(sort=False)
    # 显示结果
    pd.set_option('display.max_rows', None)
    print(datas)
    # 将数据写入Excel
    datas.to_excel(writer, sheet_name=car_number[n])
    # 设置索引列的列宽
    writer.sheets[car_number[n]].column_dimensions['A'].width = 30

writer.save()
