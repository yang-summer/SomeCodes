import pandas as pd
import time

# 选择ic卡数据、到离站时间数据、站点数据
path_ic = '96/all_ic.xlsx'
path_dl = '96/18_dl.xlsx'
path_station = '96/18_station.xlsx'

# 导入ic卡数据
ic_data = pd.read_excel(path_ic, 'Sheet1')
# 转换时间数据为时间戳
ic_data[0] = ic_data[0].apply(lambda x: str(x))
ic_data[0] = ic_data[0].apply(lambda x: time.strptime(x, "%Y-%m-%d %H:%M:%S"))
ic_data[0] = ic_data[0].apply(lambda x: time.mktime(x))

# 导入到离站时间数据
dl = pd.read_excel(path_dl, 'Sheet1')
# 转换时间数据为时间戳
dl[0] = dl[0].apply(lambda x: str(x))
dl[0] = dl[0].apply(lambda x: time.strptime(x, "%Y-%m-%d %H:%M:%S"))
dl[0] = dl[0].apply(lambda x: time.mktime(x))

# 以站点数据作为分割
label1 = pd.read_excel(path_station, 'Sheet1')
lb1 = list(label1[0])
data1 = pd.cut(ic_data[0], dl[0], labels=lb1, ordered=False)
# 计数
datas = pd.value_counts(data1, sort=False)

# 显示结果
pd.set_option('display.max_rows', None)
print(datas)
