import pandas, csv, re, os
from pandas import *
from csv import *

cwd = os.getcwd()

##Extracting SWC settings from the excel sheet
filename = "Qranium_V1_RUMI.xlsm"
SWC_Sheet = pandas.read_excel(os.path.join(cwd,filename), sheet_name='SWCs', skiprows=9)
SWC_Initial_Data = pandas.DataFrame(SWC_Sheet, columns= ['Register','Recommended'])
SWC_Data_List = SWC_Initial_Data.values.tolist()
        
##Extracting the data from the register dump files for 8 clusters
##Cluster0 Data
filename = "hwiosave_ddrss0.txt"
cluster0_Name = []
cluster0_Value = []
cluster0_file = open(os.path.join(cwd,filename),"r")
for line in cluster0_file:
    line = [line.strip('\n') for line in cluster0_file if line.startswith('DDR')]
    for item in line:
        if ('SHRM_MEM_SHRM' not in item.split(' ')[0]) and ('MEMNOC_HM_MEM_NOC' not in item.split(' ')[0]) and ('AHB2PHY_BROADCAST_ADDRESS_SPACE' not in item.split(' ')[0]):
            item = re.sub(r" (.+) "," ",item)
            cluster0_Name.append(item.split(' ')[0])
            cluster0_Value.append(item.split(' ')[-1])

##Cluster1 Data
filename = "hwiosave_ddrss1.txt"
cluster1_Name = []
cluster1_Value = []                          
cluster1_file = open(os.path.join(cwd,filename),"r")
for line in cluster1_file:
    line = [line.strip('\n') for line in cluster1_file if line.startswith('DDR')]
    for item in line:
        if ('SHRM_MEM_SHRM' not in item.split(' ')[0]) and ('MEMNOC_HM_MEM_NOC' not in item.split(' ')[0]) and ('AHB2PHY_BROADCAST_ADDRESS_SPACE' not in item.split(' ')[0]):
            item = re.sub(r" (.+) "," ",item)
            cluster1_Name.append(item.split(' ')[0])
            cluster1_Value.append(item.split(' ')[-1])

##Cluster2 Data
filename = "hwiosave_ddrss2.txt"
cluster2_Name = []
cluster2_Value = []                          
cluster2_file = open(os.path.join(cwd,filename),"r")
for line in cluster2_file:
    line = [line.strip('\n') for line in cluster2_file if line.startswith('DDR')]
    for item in line:
        if ('SHRM_MEM_SHRM' not in item.split(' ')[0]) and ('MEMNOC_HM_MEM_NOC' not in item.split(' ')[0]) and ('AHB2PHY_BROADCAST_ADDRESS_SPACE' not in item.split(' ')[0]):
            item = re.sub(r" (.+) "," ",item)
            cluster2_Name.append(item.split(' ')[0])
            cluster2_Value.append(item.split(' ')[-1])

##Cluster3 Data
filename = "hwiosave_ddrss3.txt"
cluster3_Name = []
cluster3_Value = []                    
cluster3_file = open(os.path.join(cwd,filename),"r")
for line in cluster3_file:
    line = [line.strip('\n') for line in cluster3_file if line.startswith('DDR')]
    for item in line:
        if ('SHRM_MEM_SHRM' not in item.split(' ')[0]) and ('MEMNOC_HM_MEM_NOC' not in item.split(' ')[0]) and ('AHB2PHY_BROADCAST_ADDRESS_SPACE' not in item.split(' ')[0]):
            item = re.sub(r" (.+) "," ",item)
            cluster3_Name.append(item.split(' ')[0])
            cluster3_Value.append(item.split(' ')[-1])

##Cluster4 Data
filename = "hwiosave_ddrss4.txt"
cluster4_Name = []
cluster4_Value = []                     
cluster4_file = open(os.path.join(cwd,filename),"r")
for line in cluster4_file:
    line = [line.strip('\n') for line in cluster4_file if line.startswith('DDR')]
    for item in line:
        if ('SHRM_MEM_SHRM' not in item.split(' ')[0]) and ('MEMNOC_HM_MEM_NOC' not in item.split(' ')[0]) and ('AHB2PHY_BROADCAST_ADDRESS_SPACE' not in item.split(' ')[0]):
            item = re.sub(r" (.+) "," ",item)
            cluster4_Name.append(item.split(' ')[0])
            cluster4_Value.append(item.split(' ')[-1])

##Cluster5 Data
filename = "hwiosave_ddrss5.txt"
cluster5_Name = []
cluster5_Value = []                         
cluster5_file = open(os.path.join(cwd,filename),"r")
for line in cluster5_file:
    line = [line.strip('\n') for line in cluster5_file if line.startswith('DDR')]
    for item in line:
        if ('SHRM_MEM_SHRM' not in item.split(' ')[0]) and ('MEMNOC_HM_MEM_NOC' not in item.split(' ')[0]) and ('AHB2PHY_BROADCAST_ADDRESS_SPACE' not in item.split(' ')[0]):
            item = re.sub(r" (.+) "," ",item)
            cluster5_Name.append(item.split(' ')[0])
            cluster5_Value.append(item.split(' ')[-1])

##Cluster6 Data
filename = "hwiosave_ddrss6.txt"
cluster6_Name = []
cluster6_Value = []                  
cluster6_file = open(os.path.join(cwd,filename),"r")
for line in cluster6_file:
    line = [line.strip('\n') for line in cluster6_file if line.startswith('DDR')]
    for item in line:
        if ('SHRM_MEM_SHRM' not in item.split(' ')[0]) and ('MEMNOC_HM_MEM_NOC' not in item.split(' ')[0]) and ('AHB2PHY_BROADCAST_ADDRESS_SPACE' not in item.split(' ')[0]):
            item = re.sub(r" (.+) "," ",item)
            cluster6_Name.append(item.split(' ')[0])
            cluster6_Value.append(item.split(' ')[-1])

##Cluster7 Data
filename = "hwiosave_ddrss7.txt"
cluster7_Name = []
cluster7_Value = []                   
cluster7_file = open(os.path.join(cwd,filename),"r")
for line in cluster7_file:
    line = [line.strip('\n') for line in cluster7_file if line.startswith('DDR')]
    for item in line:
        if ('SHRM_MEM_SHRM' not in item.split(' ')[0]) and ('MEMNOC_HM_MEM_NOC' not in item.split(' ')[0]) and ('AHB2PHY_BROADCAST_ADDRESS_SPACE' not in item.split(' ')[0]):
            item = re.sub(r" (.+) "," ",item)
            cluster7_Name.append(item.split(' ')[0])
            cluster7_Value.append(item.split(' ')[-1])

cluster_list = list(zip(cluster0_Name,cluster0_Value,cluster1_Value,cluster2_Value,cluster3_Value,cluster4_Value,cluster5_Value,cluster6_Value,cluster7_Value))
comparison_list = []
for cluster_reg_name in cluster_list:
    for swc_reg_name in SWC_Data_List:
        if swc_reg_name[0] in cluster_reg_name[0]:
            swc_value = '0x' + swc_reg_name[-1]
            result = list(cluster_reg_name)
            result.append(swc_value)
            comparison_list.append(result)
     
comparison_data = pandas.DataFrame(comparison_list, columns=['Register Name','Cluster0','Cluster1','Cluster2','Cluster3','Cluster4','Cluster5','Cluster6','Cluster7','Recommended Setting'])

filename = "RegDump_Comparison.xlsx"
writer = ExcelWriter(os.path.join(cwd,filename))
comparison_data.to_excel(writer, 'Data', index=None)

workbook = writer.book
worksheet = writer.sheets['Data']

color_format = workbook.add_format({"bg_color": "#FF0000"})

for index, line in comparison_data.iterrows():
    if bin(int(line['Cluster0'],16)) != bin(int(line['Recommended Setting'],16)):
        worksheet.write(index+1,1,line['Cluster0'],color_format)
    if bin(int(line['Cluster1'],16)) != bin(int(line['Recommended Setting'],16)):
        worksheet.write(index+1,2,line['Cluster1'],color_format)
    if bin(int(line['Cluster2'],16)) != bin(int(line['Recommended Setting'],16)):
        worksheet.write(index+1,3,line['Cluster2'],color_format)
    if bin(int(line['Cluster3'],16)) != bin(int(line['Recommended Setting'],16)):
        worksheet.write(index+1,4,line['Cluster3'],color_format)
    if bin(int(line['Cluster4'],16)) != bin(int(line['Recommended Setting'],16)):
        worksheet.write(index+1,5,line['Cluster4'],color_format)
    if bin(int(line['Cluster5'],16)) != bin(int(line['Recommended Setting'],16)):
        worksheet.write(index+1,6,line['Cluster5'],color_format)
    if bin(int(line['Cluster6'],16)) != bin(int(line['Recommended Setting'],16)):
        worksheet.write(index+1,7,line['Cluster6'],color_format)
    if bin(int(line['Cluster7'],16)) != bin(int(line['Recommended Setting'],16)):
        worksheet.write(index+1,8,line['Cluster7'],color_format)

writer.save()
