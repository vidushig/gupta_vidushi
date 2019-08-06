import psutil
import pandas
from pandas import *

addrs = psutil.net_if_addrs()

port_data = []
port_list = []

for key, value in addrs.items():
    for i in range(len(value)):
        data = key, list(value[i])
        port_data.append(data[0])
        for val in data[1]:
            port_data.append(val)
        port_list.append(port_data)
        port_data = []

df = pandas.DataFrame(port_list)
df.columns = ["Port Name","Family","IP Address","Broadcast","Netmask","PTP"]
df_final = df.fillna("None")
df_final.to_csv("Psutil.csv", index=None)
