import pandas as pd
import time


# 定义将数据转化为时间戳的函数
def timetrans(s):
    return (s.apply(lambda x: str(x))
            .apply(lambda x: time.strptime(x, "%Y-%m-%d %H:%M:%S"))
            .apply(lambda x: time.mktime(x)))


# 选择ic卡数据、到离站时间数据、站点数据、车牌号、写入的文件
path_ic = '96/0412_ic_96.xlsx'
path_dl = '96/0412_dl_96.xlsx'
path_station = '96/0412_station_96.xlsx'
car_numbers = ['00C270', '00D015', '00D050', '00D601']
writer = pd.ExcelWriter('test_data.xlsx')

# 导入ic卡数据
ic_data = pd.read_excel(path_ic, 'Sheet1').dropna()
# 导入到离站时间数据
dl = pd.read_excel(path_dl, 'Sheet1').dropna()
# 导入站点标签数据
label1 = pd.read_excel(path_station, 'Sheet1').dropna()

for n in range(len(car_numbers)):

    # 转化日期格式
    timetrans(ic_data[n])
    timetrans(dl[n])
    # 以站点数据作为分割
    lb1 = list(label1[n])
    data1 = pd.cut(ic_data[n], dl[n], labels=lb1, ordered=False)
    # 计数
    datas = pd.value_counts(data1, sort=False)
    # 显示结果
    pd.set_option('display.max_rows', None)
    print(datas)
    # 将数据写入Excel
    datas.to_excel(writer, sheet_name=car_numbers[n])
    # 设置索引列的列宽
    writer.sheets[car_numbers[n]].column_dimensions['A'].width = 30

writer.save()
writer.close()
