import pandas
from pandas import *

##Read data from the excel sheet
filename = pandas.ExcelFile(r'V:\\vi_moorea\\users\\c_vidush\\hana_configmon\\convex\\drivers\\vi_ddr\\platform\\vi_hana_v2\\src\\DDRSS_SNS_V3P5\\tools\\target\\SDM855\\internal\\SDM855_V2_ASIC.xlsm')

swc_sheet=pandas.read_excel(filename, 'SWCs', skiprows=9)
swc_data = pandas.DataFrame(swc_sheet, columns = ['Block','Address','Register','Recommended','PoR Source'])
swc_data_list = swc_data.values

freq_sheet=pandas.read_excel(filename, 'STRUCTs')
freq_data = pandas.DataFrame(freq_sheet, columns = ['Frequency','Block','Register','Recommended','PoR Source'])
freq_data_list = freq_data.values

alc_sheet=pandas.read_excel(filename, 'ALCs')
alc_data = pandas.DataFrame(alc_sheet, columns = ['Block','Register','Recommended','PoR Source'])
alc_data_list = alc_data.values

##Sorting the values from swc sheet and the alc sheet
for alc_data in alc_data_list:
    for swc_data in swc_data_list:
        if alc_data[1] == swc_data[2]:
            swc_data[3] = hex((int(swc_data[3],16) & int(swc_data[4],16)) | (int(alc_data[2],16) & int(alc_data[3],16)))
            swc_data[4] = hex(int(swc_data[4],16) | int(alc_data[3],16))

##Sorting the values from swc sheet and the structs sheet
for freq_data in freq_data_list:
    if freq_data[0] == 200.0:
        for swc_data in swc_data_list:
            if freq_data[1] == swc_data[2]:
                if freq_data[4] == 'FFFFFFFF':
                    swc_data[3]=freq_data[3]
                    swc_data[4]=freq_data[4]
                else:
                    swc_data[3] = hex((int(swc_data[3],16) & int(swc_data[4],16)) | (int(freq_data[3],16) & int(freq_data[4],16)))
                    swc_data[4] = hex(int(swc_data[4],16) | int(freq_data[4],16))
                    
swc_list_updated = pandas.DataFrame(swc_data_list, columns = ['Block','Address','Register','Recommended','PoR Source'])

##Create scorecard data from SWC and Frequency sheet
scorecard_sheet=pandas.read_excel(filename, 'SCORECARD', skiprows=33)
scorecard_data = pandas.DataFrame(scorecard_sheet, columns = ['Frequency','Block','Address','Register','Recommended','PoR Source', 'Read', 'Delta'])
scorecard_data_list = scorecard_data.values

scorecard_list = []
##Remove DDR_CLK_PERIOD entry from the list if frequency is equal to the Silicon value
for line in scorecard_data_list:
    if 'DDR_CLK_PERIOD' in line[3]:
        if line[0] == ((1/int(line[6],16))*1000000):
            continue
    elif 'SEC_ADDR' in line[3]:
        continue
    elif 'APM_OPT_CFG' in line[3]:
        continue
    elif 'SHKE_CMD_SET_CFG' in line[3]:
        continue
    else:
        scorecard_list.append(line)

writer = ExcelWriter('V:\\vi_moorea\\users\\c_vidush\\hana_configmon\\convex\\drivers\\vi_ddr\\platform\\vi_hana_v2\\src\\DDRSS_SNS_V3P5\\tools\\target\\SDM855\\internal\\SWC_Sheet_updates.xlsx')
swc_list_updated.to_excel(writer, 'SWCs', index=False)
scorecard_df = pandas.DataFrame(scorecard_list, columns = ['Frequency','Block','Address','Register','Recommended','PoR Source', 'Read', 'Delta'])
scorecard_data.to_excel(writer, 'Scorecard_Default', index=False)
scorecard_df.to_excel(writer, 'Scorecard', index=False)
writer.save()
