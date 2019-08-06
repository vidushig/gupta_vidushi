"""Extract ip address from the system info.
Modify the data from dictionary to a list format"""

import psutil

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

for item in port_list:
    print(item)

