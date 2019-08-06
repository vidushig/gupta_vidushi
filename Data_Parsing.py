import pandas, os, glob
from pandas import *

cwd = os.getcwd()

for excel_filename in glob.glob("*.xlsm"):
    excel_data = pandas.read_excel(excel_filename,skiprows=3)
    excel_data_col_filter = excel_data.drop(columns=['Channel','Rank','Byte','Vref'])
    excel_data_row_filter = excel_data_col_filter.dropna()
    excel_data_format = excel_data_row_filter.astype(int)
    excel_data_format.to_csv(r'C:\\Users\\c_vidush\\Documents\\Qranium\\data.txt',header=None,index=None,sep=' ')


