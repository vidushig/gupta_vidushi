"""Get a file as input using TK.
Copy the data from xlsm file as a dataframe and format the data to a text file"""

import tkinter
from tkinter import filedialog
import pandas, os, glob, numpy
from pandas import *
from os import *


root = tkinter.Tk()
root.withdraw()

excel_filename = filedialog.askopenfilename()

excel_data = pandas.read_excel(excel_filename,skiprows=3)
excel_data_col_filter = excel_data.drop(excel_data.columns[[0,1,2,3]], axis=1, inplace=True)
excel_data_col_filter = excel_data
excel_data_row_filter = excel_data_col_filter.dropna()
excel_data_format = excel_data_row_filter.astype(int)
excel_matrix = excel_data_format.values
search_matrix = np.zeros((3,3))
search_matrix[1,1] = 3
sm_rows = len(search_matrix)
sm_cols = len(search_matrix[0])
dm_rows = len(excel_matrix)
dm_cols = len(excel_matrix[0])
   
for i in range(dm_rows - sm_rows):
    for j in range(dm_cols - sm_cols):
        if (np.equal(search_matrix,excel_matrix[i:i+sm_rows,j:j+sm_cols]).all()):
            excel_matrix[i:i+sm_rows,j:j+sm_cols] = 0

filename, ext = os.path.splitext(excel_filename)
text_filename = str(filename) + ".txt"
numpy.savetxt(text_filename,excel_matrix,fmt="%d")
