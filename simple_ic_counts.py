import pandas as pd

import time

path_ic = '96/all_ic.xlsx'
path_dl = '96/18_dl.xlsx'
path_station = '96/18_station.xlsx'

ic_data = pd.read_excel(path_ic, 'Sheet1')

ic_data[0] = ic_data[0].apply(lambda x: str(x))

ic_data[0] = ic_data[0].apply(lambda x: time.strptime(x, "%Y-%m-%d %H:%M:%S"))

ic_data[0] = ic_data[0].apply(lambda x: time.mktime(x))

dl = pd.read_excel(path_dl, 'Sheet1')

dl[0] = dl[0].apply(lambda x: str(x))

dl[0] = dl[0].apply(lambda x: time.strptime(x, "%Y-%m-%d %H:%M:%S"))

dl[0] = dl[0].apply(lambda x: time.mktime(x))

label1 = pd.read_excel(path_station, 'Sheet1')

lb1 = list(label1[0])

data1 = pd.cut(ic_data[0], dl[0], labels=lb1, ordered=False)

datas = pd.value_counts(data1, sort=False)

pd.set_option('display.max_rows', None)

print(datas)
